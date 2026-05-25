"""API REST para el Asistente Excel + bot de Telegram integrado.

Modos de ejecución (se detectan automáticamente por las variables de entorno):

  Modo WEBHOOK  (Render / cloud)
    → WEBHOOK_URL definida en .env
    → Telegram envía los mensajes al endpoint /telegram/webhook
    → Ideal para despliegues en la nube

  Modo POLLING  (servidor empresa / pruebas locales con API)
    → WEBHOOK_URL vacía o ausente, TELEGRAM_TOKEN presente
    → El bot hace polling a Telegram cada pocos segundos
    → Sin puertos entrantes, ideal para servidores internos

  Para uso personal sin API (solo bot, sin Add-in):
    → Usa bot.py directamente en lugar de api.py
"""
import asyncio
import logging
import os
import re
from contextlib import asynccontextmanager

import pandas as pd
import uvicorn
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, Header, HTTPException, Query, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from telegram import Update

from excel.analyzer import resumir
from excel.editor import EditorError, aplicar_edicion
from excel.query_engine import QueryError, ejecutar_query
from logging_config import configurar_logging
from services.llm import (extraer_operacion_edicion, extraer_query_dsl,
                          extraer_estructura_excel, obtener_respuesta)
from config import SYSTEM_PROMPT_ADDIN

load_dotenv()
configurar_logging()

logger = logging.getLogger(__name__)

_API_KEY     = os.getenv("API_KEY", "")
_ADMIN_KEY   = os.getenv("ADMIN_KEY", "") or _API_KEY   # fallback a API_KEY si no hay ADMIN_KEY
_WEBHOOK_URL = os.getenv("WEBHOOK_URL", "").rstrip("/")
_BOT_TOKEN   = os.getenv("TELEGRAM_TOKEN", "")

# Módulos activables (por defecto ambos activos)
_ENABLE_TELEGRAM = os.getenv("ENABLE_TELEGRAM", "true").lower() == "true"
_ENABLE_ADDIN    = os.getenv("ENABLE_ADDIN",    "true").lower() == "true"

# ID de Telegram al que enviar alertas del sistema (por defecto el primer AUTHORIZED_USER)
_ids_autorizados  = [u.strip() for u in os.getenv("AUTHORIZED_USERS", "").split(",") if u.strip()]
_ALERT_TELEGRAM_ID = int(os.getenv("ALERT_TELEGRAM_ID", _ids_autorizados[0] if _ids_autorizados else "0") or "0")

# Umbrales de alerta (% sobre el límite de Render free)
_ALERTA_PCT       = 80.0    # rojo a partir del 80 %
_ALERTA_COOLDOWN  = 3600    # segundos entre avisos del mismo tipo (evita spam)
_MONITOR_INTERVALO = 600    # comprobar cada 10 minutos

# ---------------------------------------------------------------------------
# Bot de Telegram — modo webhook o polling según configuración
# ---------------------------------------------------------------------------

_ptb_app  = None
_bot_mode = "none"

if _ENABLE_TELEGRAM and _BOT_TOKEN:
    from telegram_app import crear_aplicacion
    _ptb_app = crear_aplicacion()
    _bot_mode = "webhook" if _WEBHOOK_URL else "polling"
elif not _ENABLE_TELEGRAM:
    logger.info("Módulo Telegram desactivado (ENABLE_TELEGRAM=false)")


async def _notificar_telegram(texto: str) -> None:
    """Envía un mensaje de alerta a todos los suscriptores activos.

    Si la tabla alert_subs está vacía usa ALERT_TELEGRAM_ID del .env como fallback
    para no romper instalaciones existentes.
    """
    if not _ptb_app:
        return
    from utils.alert_subs import ids_activos
    destinos = ids_activos()
    if not destinos and _ALERT_TELEGRAM_ID:
        destinos = [_ALERT_TELEGRAM_ID]
    for tid in destinos:
        try:
            await _ptb_app.bot.send_message(chat_id=tid, text=texto, parse_mode="Markdown")
        except Exception as exc:
            logger.warning("No se pudo enviar alerta a %s: %s", tid, exc)


async def _monitor_alertas_sistema() -> None:
    """Comprueba RAM y disco cada 10 min y avisa si superan el umbral rojo."""
    import time
    _ultima: dict[str, float] = {}

    await asyncio.sleep(60)   # espera inicial para dejar que el servidor arranque
    while True:
        try:
            ahora   = time.time()
            sistema = _obtener_info_sistema()
            alertas = []

            ram_pct = sistema.get("ram_pct_render")
            if ram_pct and ram_pct >= _ALERTA_PCT:
                if ahora - _ultima.get("ram", 0) > _ALERTA_COOLDOWN:
                    alertas.append(
                        f"🔴 *RAM al {ram_pct:.0f}%*\n"
                        f"   {sistema['ram_usado_mb']} MB / {_RENDER_RAM_MB} MB (límite Render)"
                    )
                    _ultima["ram"] = ahora

            # Alerta si data/ supera 400 MB (runtime creciente; código ocupa ~300 MB fijos)
            data_mb = sistema["data_mb"]
            if data_mb >= 400:
                if ahora - _ultima.get("disco", 0) > _ALERTA_COOLDOWN:
                    alertas.append(
                        f"🔴 *Carpeta data/ en {data_mb:.0f} MB*\n"
                        f"   La base de datos o los logs están creciendo demasiado."
                    )
                    _ultima["disco"] = ahora

            if alertas:
                await _notificar_telegram(
                    "⚠️ *Alerta del servidor — Asistente Excel*\n\n" + "\n\n".join(alertas)
                )
                logger.warning("Alerta sistema: %s", alertas)

        except asyncio.CancelledError:
            raise
        except Exception as exc:
            logger.warning("Monitor alertas sistema: %s", exc)

        await asyncio.sleep(_MONITOR_INTERVALO)


_errores_bot_vistos: dict[str, float] = {}
_COOLDOWN_BOT_ERR = 300   # segundos entre alertas del mismo tipo de error

async def _manejador_error_bot(update: object, context: object) -> None:
    """Handler de errores de python-telegram-bot — notifica al administrador."""
    import time
    error = context.error  # type: ignore[attr-defined]
    logger.error("Error PTB: %s", error, exc_info=error)
    nombre = type(error).__name__
    ahora = time.time()
    if ahora - _errores_bot_vistos.get(nombre, 0) > _COOLDOWN_BOT_ERR:
        _errores_bot_vistos[nombre] = ahora
        asyncio.create_task(_notificar_telegram(
            f"🤖 *Error en el bot de Telegram*\n`{nombre}: {str(error)[:350]}`"
        ))


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Arranca y detiene el bot PTB junto con FastAPI."""
    if _ptb_app and _bot_mode == "webhook":
        await _ptb_app.initialize()
        await _ptb_app.start()
        # set_webhook va DESPUÉS de start() para que los handlers estén activos
        # cuando Telegram empiece a entregar updates.
        # drop_pending_updates=True descarta mensajes encolados durante el cold start.
        await _ptb_app.bot.set_webhook(
            url=f"{_WEBHOOK_URL}/telegram/webhook",
            allowed_updates=Update.ALL_TYPES,
            drop_pending_updates=True,
        )
        logger.info("Bot Telegram arrancado en modo WEBHOOK → %s/telegram/webhook", _WEBHOOK_URL)

    elif _ptb_app and _bot_mode == "polling":
        await _ptb_app.initialize()
        await _ptb_app.start()
        await _ptb_app.updater.start_polling(drop_pending_updates=True)
        logger.info("Bot Telegram arrancado en modo POLLING (sin puertos entrantes)")

    else:
        logger.info("TELEGRAM_TOKEN no definido — bot desactivado en esta instancia")

    if _ptb_app:
        _ptb_app.add_error_handler(_manejador_error_bot)

    # Monitor de alertas del sistema (RAM / disco)
    _task_monitor = asyncio.create_task(_monitor_alertas_sistema())

    # Notificación de arranque — también avisa de reinicios/caídas previas
    modo_str = _bot_mode.upper() if _ptb_app else "SIN BOT"
    ia_str   = os.getenv("LLM_PROVIDER", "?").upper()
    asyncio.create_task(_notificar_telegram(
        f"🟢 *Asistente Excel — servidor online*\n"
        f"Modo: `{modo_str}` · IA: `{ia_str}`"
    ))

    yield   # ← la app está corriendo

    _task_monitor.cancel()
    try:
        await _task_monitor
    except asyncio.CancelledError:
        pass

    if _ptb_app and _bot_mode == "polling":
        await _ptb_app.updater.stop()
        await _ptb_app.stop()
        await _ptb_app.shutdown()
    elif _ptb_app and _bot_mode == "webhook":
        await _ptb_app.stop()
        await _ptb_app.shutdown()


# ---------------------------------------------------------------------------
# FastAPI
# ---------------------------------------------------------------------------

app = FastAPI(title="Asistente Excel API", version="1.0", lifespan=lifespan)


# Middleware: detecta errores HTTP 500 y avisa por Telegram (cooldown 5 min por ruta)
_errores_500_vistos: dict[str, float] = {}
_COOLDOWN_500 = 300  # segundos

@app.middleware("http")
async def _middleware_errores_500(request: Request, call_next):
    import time
    response = await call_next(request)
    if response.status_code == 500:
        clave = f"{request.method} {request.url.path}"
        ahora = time.time()
        if ahora - _errores_500_vistos.get(clave, 0) > _COOLDOWN_500:
            _errores_500_vistos[clave] = ahora
            asyncio.create_task(_notificar_telegram(
                f"💥 *Error 500 en el servidor*\n`{clave}`"
            ))
    return response


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def _verificar_clave(x_api_key: str = Header(...)) -> None:
    if not _API_KEY:
        raise HTTPException(status_code=500, detail="API_KEY no configurada en el servidor")
    if x_api_key != _API_KEY:
        raise HTTPException(status_code=403, detail="Clave de API inválida")


# ---------------------------------------------------------------------------
# Modelos de petición
# ---------------------------------------------------------------------------

class PeticionPregunta(BaseModel):
    datos: list[list] | None = None
    pregunta: str
    historial: list[dict] = []
    device_id: str | None = None


class PeticionAnalisis(BaseModel):
    datos: list[list]


class PeticionEdicion(BaseModel):
    datos: list[list]
    instruccion: str
    historial: list[dict] = []
    device_id: str | None = None


class PeticionEnviarAlBot(BaseModel):
    datos: list[list]
    nombre_archivo: str = "datos.xlsx"
    email: str | None = None        # flujo A: SSO / Azure AD
    device_id: str | None = None    # flujo B: emparejamiento por código


class PeticionVerificarCodigo(BaseModel):
    device_id: str
    codigo: str


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _a_dataframe(datos: list[list]) -> pd.DataFrame:
    if len(datos) < 2:
        raise HTTPException(
            status_code=400,
            detail="Se necesitan al menos una fila de cabeceras y una de datos",
        )
    cabeceras = [str(c) for c in datos[0]]
    return pd.DataFrame(datos[1:], columns=cabeceras)


def _df_a_matriz(df: pd.DataFrame) -> list[list]:
    import json
    cabeceras = [str(c) for c in df.columns]
    filas_json = json.loads(df.to_json(orient="values", default_handler=str))
    filas = [["" if v is None else v for v in fila] for fila in filas_json]
    return [cabeceras] + filas


def _resultado_a_texto(resultado, descripcion: str) -> str:
    if isinstance(resultado, pd.DataFrame):
        return f"{descripcion}:\n{resultado.to_string(index=False)}"
    return f"{descripcion}: {resultado}"


# ---------------------------------------------------------------------------
# Webhook de Telegram
# ---------------------------------------------------------------------------

@app.post("/telegram/webhook")
async def telegram_webhook(request: Request):
    if not _ptb_app:
        raise HTTPException(status_code=503, detail="Bot no activo en este modo")
    data = await request.json()
    update = Update.de_json(data, _ptb_app.bot)
    await _ptb_app.process_update(update)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Endpoints del Add-in
# ---------------------------------------------------------------------------

@app.get("/health")
def health():
    return {
        "status": "ok",
        "modo_bot":  _bot_mode,
        "telegram":  _ENABLE_TELEGRAM,
        "addin":     _ENABLE_ADDIN,
        "llm":       os.getenv("LLM_PROVIDER", "groq"),
    }


# user_id sintético para peticiones del Add-in sin vincular a Telegram
_ADDIN_ANON_ID = 0


def _usuario_addin(device_id: str | None) -> int:
    """Devuelve el telegram_id vinculado al device_id, o _ADDIN_ANON_ID si no hay vínculo."""
    if device_id:
        from utils.user_links import obtener_device_link
        link = obtener_device_link(device_id)
        if link:
            return link["telegram_id"]
    return _ADDIN_ANON_ID


def _registrar_addin(device_id: str | None, texto: str) -> None:
    """Escribe la pregunta del Add-in en historial para que el usuario aparezca en el panel."""
    try:
        from utils.history import agregar_mensaje
        agregar_mensaje(_usuario_addin(device_id), "user", texto[:2000])
    except Exception as exc:
        logger.warning("No se pudo registrar uso del Add-in: %s", exc)


def _verificar_addin_activo():
    if not _ENABLE_ADDIN:
        raise HTTPException(status_code=503, detail="Módulo Add-in desactivado (ENABLE_ADDIN=false)")


@app.post("/ask")
def ask(peticion: PeticionPregunta, _: None = Depends(_verificar_clave),
        __: None = Depends(_verificar_addin_activo)) -> dict:
    _registrar_addin(peticion.device_id, peticion.pregunta)
    # Sin datos: pregunta general o creación desde cero
    if not peticion.datos or len(peticion.datos) < 2:
        if _RE_CREAR_TABLA_ADDIN.search(peticion.pregunta):
            estructura = extraer_estructura_excel(peticion.pregunta)
            if estructura:
                columnas = estructura.get("columnas", [])
                datos_filas = estructura.get("datos", [])
                matriz = [columnas] + [
                    [("" if v is None else v) for v in fila]
                    for fila in datos_filas
                ]
                titulo = estructura.get("titulo", "Nueva tabla")
                return {
                    "tipo": "datos",
                    "datos_modificados": matriz,
                    "descripcion": f"Tabla '{titulo}' creada ({len(datos_filas)} filas)",
                }
        return {"respuesta": obtener_respuesta(
            peticion.historial, peticion.pregunta,
            system_override=SYSTEM_PROMPT_ADDIN,
        )}

    df = _a_dataframe(peticion.datos)

    query = extraer_query_dsl(df, peticion.pregunta)
    if query:
        try:
            resultado, descripcion = ejecutar_query(df, query)
            return {"respuesta": _resultado_a_texto(resultado, descripcion)}
        except QueryError as error:
            logger.warning("DSL falló, usando LLM libre: %s", error)

    columnas = ", ".join(str(c) for c in df.columns)
    muestra  = df.head(5).to_string(index=False)
    contexto = (
        f"El usuario tiene una tabla con columnas: {columnas}.\n"
        f"Primeras filas:\n{muestra}\n\n"
        f"Pregunta: {peticion.pregunta}"
    )
    return {"respuesta": obtener_respuesta(
        peticion.historial, contexto,
        system_override=SYSTEM_PROMPT_ADDIN,
    )}


@app.post("/edit")
def edit(peticion: PeticionEdicion, _: None = Depends(_verificar_clave),
         __: None = Depends(_verificar_addin_activo)) -> dict:
    _registrar_addin(peticion.device_id, peticion.instruccion)
    df = _a_dataframe(peticion.datos)

    op = extraer_operacion_edicion(df, peticion.instruccion)
    if op:
        if "op" not in op and "operacion" in op:
            op["op"] = op.pop("operacion")
        try:
            df_mod, descripcion, _extras = aplicar_edicion(df, op)
            return {
                "tipo": "edicion",
                "datos_modificados": _df_a_matriz(df_mod),
                "descripcion": descripcion,
            }
        except EditorError as error:
            logger.warning("Editor falló, intentando query DSL: %s", error)

    # La petición no es una edición → intentar como consulta de datos
    query = extraer_query_dsl(df, peticion.instruccion)
    if query and not query.get("aclaracion_necesaria"):
        try:
            resultado, descripcion = ejecutar_query(df, query)
            if isinstance(resultado, pd.DataFrame) and not resultado.empty:
                # Resultado tabular → se puede escribir en celdas
                return {
                    "tipo": "edicion",
                    "datos_modificados": _df_a_matriz(resultado),
                    "descripcion": descripcion,
                }
            # Resultado escalar → mostrar como texto
            return {"tipo": "texto", "respuesta": _resultado_a_texto(resultado, descripcion)}
        except QueryError as error:
            logger.warning("Query DSL falló, usando LLM libre: %s", error)

    columnas = ", ".join(str(c) for c in df.columns)
    muestra  = df.head(5).to_string(index=False)
    contexto = (
        f"El usuario tiene una tabla con columnas: {columnas}.\n"
        f"Primeras filas:\n{muestra}\n\n"
        f"Pregunta: {peticion.instruccion}"
    )
    return {"tipo": "texto", "respuesta": obtener_respuesta(
        peticion.historial, contexto,
        system_override=SYSTEM_PROMPT_ADDIN,
    )}


@app.post("/analizar")
def analizar(peticion: PeticionAnalisis, _: None = Depends(_verificar_clave),
             __: None = Depends(_verificar_addin_activo)) -> dict:
    df = _a_dataframe(peticion.datos)
    return {"resumen": resumir(df, "Datos de Excel")}


@app.get("/addin-config")
def addin_config(_: None = Depends(_verificar_clave),
                 __: None = Depends(_verificar_addin_activo)) -> dict:
    """Configuración dinámica del Add-in.

    El Add-in lo consulta al arrancar para:
    - Saber si el módulo Telegram está activo (oculta el botón "Enviar al bot" si no)
    - Obtener el nombre de empresa configurado en el servidor (sin necesidad de recompilar)
    """
    return {
        "telegram_habilitado": _ENABLE_TELEGRAM,
        "nombre_empresa":      os.getenv("COMPANY_NAME", ""),
    }


@app.get("/tiene-vinculo")
def tiene_vinculo(
    email: str | None = Query(None),
    device_id: str | None = Query(None),
    _: None = Depends(_verificar_clave),
    __: None = Depends(_verificar_addin_activo),
) -> dict:
    """Comprueba si hay vínculo activo.

    Acepta dos modos:
    - device_id: flujo B (emparejamiento por código efímero)
    - email: flujo A (SSO / Azure AD, actualmente sin uso en desktop)
    """
    from utils.user_links import obtener_telegram_id, obtener_device_link

    if device_id:
        link = obtener_device_link(device_id)
        if not link:
            return {"vinculado": False}
        # Verificar que el email sigue activo en user_links
        return {"vinculado": obtener_telegram_id(link["email"]) is not None}

    if email:
        return {"vinculado": obtener_telegram_id(email) is not None}

    return {"vinculado": False}


@app.post("/verificar-codigo")
def verificar_codigo(
    peticion: PeticionVerificarCodigo,
    _: None = Depends(_verificar_clave),
    __: None = Depends(_verificar_addin_activo),
) -> dict:
    """Valida un código efímero generado por /codigo en Telegram.

    Si es válido, crea el vínculo device_id ↔ telegram_id en device_links.
    El código queda marcado como usado para que no pueda reutilizarse.
    """
    from utils.user_links import (
        obtener_codigo_dispositivo, marcar_codigo_usado, guardar_device_link,
    )
    from datetime import datetime

    if not peticion.codigo.isdigit() or len(peticion.codigo) != 6:
        raise HTTPException(status_code=400, detail="Formato de código inválido")

    registro = obtener_codigo_dispositivo(peticion.codigo)
    if not registro:
        raise HTTPException(status_code=404, detail="Código no encontrado")
    if registro["usado"]:
        raise HTTPException(status_code=410, detail="Código ya utilizado")
    if datetime.fromisoformat(registro["expiry"]) < datetime.now():
        raise HTTPException(status_code=410, detail="Código expirado")

    marcar_codigo_usado(peticion.codigo)
    guardar_device_link(peticion.device_id, registro["telegram_id"], registro["email"])

    logger.info(
        "Device '%s' vinculado a email '%s' (telegram_id %s)",
        peticion.device_id, registro["email"], registro["telegram_id"],
    )
    return {"ok": True}


@app.post("/enviar-al-bot")
async def enviar_al_bot(peticion: PeticionEnviarAlBot,
                        _: None = Depends(_verificar_clave),
                        __: None = Depends(_verificar_addin_activo)) -> dict:
    """Recibe datos del Add-in y los envía como archivo .xlsx al chat de Telegram.

    Acepta device_id (flujo B) o email (flujo A/SSO).
    """
    from utils.user_links import obtener_telegram_id, obtener_device_link
    from excel.editor import exportar_xlsx
    import io

    telegram_id = None

    if peticion.device_id:
        link = obtener_device_link(peticion.device_id)
        if link and obtener_telegram_id(link["email"]):
            telegram_id = link["telegram_id"]
    elif peticion.email:
        telegram_id = obtener_telegram_id(peticion.email)

    if not telegram_id:
        raise HTTPException(
            status_code=404,
            detail=(
                "No hay vínculo activo con Telegram. "
                "Introduce el código del bot en el Add-in para vincular este dispositivo."
            ),
        )

    if not _ptb_app:
        raise HTTPException(status_code=503, detail="Bot de Telegram no activo en este servidor")

    df = _a_dataframe(peticion.datos)

    def _crear_xlsx() -> bytes:
        buf = exportar_xlsx(df, peticion.nombre_archivo.rsplit(".", 1)[0])
        return buf.read()

    try:
        xlsx_bytes = await asyncio.to_thread(_crear_xlsx)
        buf = io.BytesIO(xlsx_bytes)
        nombre = peticion.nombre_archivo if peticion.nombre_archivo.endswith(".xlsx") else peticion.nombre_archivo + ".xlsx"
        buf.name = nombre

        await _ptb_app.bot.send_document(
            chat_id=telegram_id,
            document=buf,
            filename=nombre,
            caption=(
                f"📎 Archivo enviado desde el Add-in de Excel\n"
                f"📄 {nombre} · {len(df):,} filas × {len(df.columns)} columnas"
            ),
        )
        logger.info("Archivo '%s' enviado a telegram_id %s desde email %s",
                    nombre, telegram_id, peticion.email)
        return {"ok": True, "mensaje": f"Archivo enviado a tu chat de Telegram ({nombre})"}

    except Exception as error:
        logger.error("Error enviando archivo al bot: %s", error)
        raise HTTPException(status_code=500, detail=f"Error al enviar el archivo: {error}")


# ---------------------------------------------------------------------------
# Panel de administración
# ---------------------------------------------------------------------------

def _verificar_admin(key: str = Query(..., alias="key")) -> None:
    if not _ADMIN_KEY:
        raise HTTPException(status_code=500, detail="ADMIN_KEY no configurada")
    if key != _ADMIN_KEY:
        raise HTTPException(status_code=403, detail="Clave de administrador inválida")


@app.get("/admin/stats")
def admin_stats(_: None = Depends(_verificar_admin)) -> dict:
    """Estadísticas de uso en formato JSON."""
    from utils.stats import obtener_estadisticas
    return obtener_estadisticas()



@app.delete("/admin/vinculos")
def admin_eliminar_vinculo(telegram_id: int = Query(...),
                           email: str = Query(...),
                           _: None = Depends(_verificar_admin)) -> dict:
    """Elimina un vínculo Telegram ↔ email desde el panel de administración."""
    from utils.user_links import desvincular
    eliminados = desvincular(telegram_id, email)
    if not eliminados:
        raise HTTPException(status_code=404, detail="Vínculo no encontrado")
    return {"ok": True}


class PeticionVinculo(BaseModel):
    telegram_id: int
    email: str


@app.post("/admin/vinculos")
def admin_crear_vinculo(peticion: PeticionVinculo,
                        _: None = Depends(_verificar_admin)) -> dict:
    """Crea un vínculo Telegram ↔ email desde el panel de administración."""
    from utils.user_links import vincular
    vincular(peticion.telegram_id, peticion.email)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Suscriptores de notificaciones
# ---------------------------------------------------------------------------

class PeticionSub(BaseModel):
    telegram_id: int
    etiqueta: str = ""


@app.get("/admin/notificaciones")
def admin_notificaciones_list(_: None = Depends(_verificar_admin)) -> dict:
    from utils.alert_subs import obtener_subs
    return {"subs": obtener_subs()}


@app.post("/admin/notificaciones")
def admin_notificaciones_add(peticion: PeticionSub,
                             _: None = Depends(_verificar_admin)) -> dict:
    from utils.alert_subs import agregar_sub
    agregar_sub(peticion.telegram_id, peticion.etiqueta)
    return {"ok": True}


@app.delete("/admin/notificaciones/{telegram_id}")
def admin_notificaciones_del(telegram_id: int,
                             _: None = Depends(_verificar_admin)) -> dict:
    from utils.alert_subs import eliminar_sub
    if not eliminar_sub(telegram_id):
        raise HTTPException(status_code=404, detail="Suscriptor no encontrado")
    return {"ok": True}


@app.patch("/admin/notificaciones/{telegram_id}/toggle")
def admin_notificaciones_toggle(telegram_id: int,
                                _: None = Depends(_verificar_admin)) -> dict:
    from utils.alert_subs import toggle_sub
    nuevo = toggle_sub(telegram_id)
    if nuevo is None:
        raise HTTPException(status_code=404, detail="Suscriptor no encontrado")
    return {"ok": True, "activo": nuevo}


@app.get("/admin", response_class=HTMLResponse)
def admin_panel(_: None = Depends(_verificar_admin)):
    """Panel de administración con estadísticas de uso."""
    from utils.stats import obtener_estadisticas, obtener_stats_usuarios_avanzadas
    from utils.llm_stats import obtener_stats_ia
    stats    = obtener_estadisticas()
    stats_ia = obtener_stats_ia()
    stats_ses = obtener_stats_usuarios_avanzadas()
    sistema  = _obtener_info_sistema()
    logs     = _leer_logs_recientes(150)
    return _renderizar_admin_html(stats, stats_ia, sistema, logs, stats_ses)


# ── Helpers del panel ────────────────────────────────────────────────────────

_RENDER_RAM_MB  = 512

# Detección de intención de creación de tabla desde el Add-in (hoja en blanco)
_RE_CREAR_TABLA_ADDIN = re.compile(
    r"\b("
    r"crea[rm]?\s+(?:un[ao]?\s+)?(?:nuevo\s+)?(?:tabla|hoja|excel|archivo|libro)|"
    r"hazme\s+(?:un[ao]?\s+)?(?:tabla|hoja|excel|archivo|plantilla)|"
    r"haz\s+(?:un[ao]?\s+)?(?:tabla|hoja|excel|archivo)|"
    r"genera[rm]?\s+(?:un[ao]?\s+)?(?:tabla|hoja|excel|archivo)|"
    r"pon[er]?\s+(?:una?\s+)?(?:tabla|datos|ejemplo|muestra)|"
    r"inserta[r]?\s+(?:una?\s+)?(?:tabla|datos)|"
    r"añade\s+(?:datos|filas?|registros?|ejemplos?|muestras?)|"
    r"añadir\s+(?:datos|filas?|registros?|ejemplos?)|"
    r"dame\s+(?:un[ao]?\s+)?(?:tabla|lista|ejemplo|muestra)\s+(?:con|de|para)|"
    r"necesito\s+(?:un[ao]?\s+)?(?:tabla|hoja|lista)\s+(?:con|para|de)|"
    r"quiero\s+(?:un[ao]?\s+)?(?:tabla|hoja|lista)\s+(?:con|para|de)|"
    r"escribe\s+(?:los?\s+)?(?:datos|valores?|resultados?)|"
    r"introduce\s+(?:los?\s+)?(?:datos|valores?)|"
    r"rellena\s+(?:las?\s+)?celdas?|"
    r"ponme\s+(?:un[ao]?\s+)?(?:tabla|ejemplo|muestra|lista)"
    r")\b",
    re.IGNORECASE,
)
_RENDER_DISK_MB = 1024


def _obtener_info_sistema() -> dict:
    import sys
    from datetime import datetime

    def _mb_dir(ruta: str) -> float:
        total = 0
        try:
            for r, _, files in os.walk(ruta):
                for f in files:
                    try:
                        total += os.path.getsize(os.path.join(r, f))
                    except OSError:
                        pass
        except OSError:
            pass
        return round(total / (1024 * 1024), 1)

    # Solo medimos las carpetas que la app controla — el disco raíz en Render
    # corresponde al host físico compartido (cientos de GB) y no es útil.
    data_mb = _mb_dir("data")
    resultado: dict = {
        "data_mb":        data_mb,
        "logs_mb":        _mb_dir("data/logs"),
        "temp_mb":        _mb_dir("data/temp"),
        "python_version": sys.version.split()[0],
        "ram_usado_mb":   None,
        "ram_pct_render": None,
        "cpu_pct":        None,
        "uptime_seg":     None,
    }
    try:
        import psutil
        # RSS del proceso actual — virtual_memory() lee la RAM del host físico
        # de Render (varios GB), no del contenedor de 512 MB.
        proc   = psutil.Process()
        rss_mb = proc.memory_info().rss // (1024 * 1024)
        resultado["ram_usado_mb"]   = rss_mb
        resultado["ram_pct_render"] = min(round(rss_mb / _RENDER_RAM_MB * 100, 1), 100)
        resultado["cpu_pct"]        = psutil.cpu_percent(interval=0.2)
        resultado["uptime_seg"]     = int(datetime.now().timestamp() - proc.create_time())
    except Exception:
        pass
    return resultado


def _formato_uptime(seg: int | None) -> str:
    if seg is None:
        return "—"
    d, rem  = divmod(seg, 86400)
    h, rem  = divmod(rem, 3600)
    m       = rem // 60
    if d:
        return f"{d}d {h}h {m}m"
    if h:
        return f"{h}h {m}m"
    return f"{m}m"


def _color_pct(pct: float | None) -> str:
    if pct is None:
        return "#aaa"
    if pct < 60:
        return "#27ae60"
    if pct < 80:
        return "#e67e22"
    return "#e74c3c"


def _leer_logs_recientes(n: int = 150) -> list[str]:
    ruta = os.path.join("data", "logs", "bot.log")
    if not os.path.exists(ruta):
        return []
    try:
        with open(ruta, encoding="utf-8", errors="replace") as f:
            lineas = f.readlines()
        return [l.rstrip() for l in lineas[-n:]]
    except Exception:
        return []


def _html_linea_log(linea: str) -> str:
    import html as _html
    if "ERROR" in linea:
        cls, nivel = "log-error", "error"
    elif "WARNING" in linea or "WARN" in linea:
        cls, nivel = "log-warn", "warn"
    elif "DEBUG" in linea:
        cls, nivel = "log-debug", "other"
    else:
        cls, nivel = "log-info", "other"
    return f'<div class="{cls}" data-lvl="{nivel}">{_html.escape(linea)}</div>'


# ── Renderizado HTML ─────────────────────────────────────────────────────────

def _fmt_min(m: float) -> str:
    """Convierte minutos (float) a cadena legible."""
    if m < 1:
        return "< 1 min"
    if m < 60:
        return f"{int(m)} min"
    h, rest = divmod(int(m), 60)
    return f"{h}h {rest}m" if rest else f"{h}h"


def _renderizar_admin_html(
    stats: dict,
    stats_ia: dict,
    sistema: dict,
    logs: list[str],
    stats_ses: dict | None = None,
) -> str:
    from datetime import datetime, timezone
    from utils.user_links import obtener_todos_los_vinculos
    from utils.db import estado as _estado_db
    from utils.alert_subs import obtener_subs

    # ── Gráfico de actividad ────────────────────────────────────────────────
    max_n = max((d["n"] for d in stats["mensajes_por_dia"]), default=1)
    barras_html = ""
    for d in stats["mensajes_por_dia"]:
        pct = int(d["n"] / max_n * 100)
        dia = d["dia"][5:]
        barras_html += (
            f'<div class="bar-item">'
            f'  <div class="bar-fill" style="height:{pct}%" title="{d["n"]} mensajes"></div>'
            f'  <div class="bar-label">{dia}</div>'
            f'  <div class="bar-val">{d["n"]}</div>'
            f'</div>'
        )

    # ── Tabla de usuarios ───────────────────────────────────────────────────
    hoy_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    usuarios_hoy = sum(
        1 for u in stats["usuarios"]
        if u["ultima_actividad"] != "—" and u["ultima_actividad"][:10] == hoy_str
    )
    filas_usuarios = ""
    for u in stats["usuarios"]:
        uid  = u["user_id"]
        uid_label = "<em style='color:#888'>Add-in</em>" if uid == _ADDIN_ANON_ID else f"<code>{uid}</code>"
        email = u["email"] or "—"
        ver  = u["version_excel"] or "—"
        modo = "🔊" if u["modo_respuesta"] == "voz" else "💬"
        priv = "🔒" if u["modo_privado"] else ""
        ts   = u["ultima_actividad"][:16].replace("T", " ") if u["ultima_actividad"] != "—" else "—"
        filas_usuarios += (
            f"<tr>"
            f"  <td>{uid_label}</td>"
            f"  <td>{email}</td>"
            f"  <td class='num'>{u['mensajes_enviados']}</td>"
            f"  <td class='num'>{u['total_mensajes']}</td>"
            f"  <td>{ts}</td>"
            f"  <td>{ver}</td>"
            f"  <td class='centro'>{modo} {priv}</td>"
            f"</tr>"
        )

    # ── Suscriptores de notificaciones ──────────────────────────────────────
    subs = obtener_subs()
    filas_subs = ""
    for s in subs:
        tid   = s["telegram_id"]
        etiq  = s["etiqueta"] or "—"
        activo = bool(s["activo"])
        badge  = '<span class="badge-ok">✅ Activo</span>' if activo else '<span class="badge-warn">⏸ Pausado</span>'
        clase_btn  = "btn-sm active" if activo else "btn-sm"
        texto_btn  = "Pausar" if activo else "Activar"
        filas_subs += (
            f"<tr>"
            f"  <td><code>{tid}</code></td>"
            f"  <td>{etiq}</td>"
            f"  <td class='centro'>{badge}</td>"
            f"  <td class='centro'>"
            f"    <button class='{clase_btn}' onclick='toggleSub({tid})'>{texto_btn}</button> "
            f"    <button class='btn-del' onclick='eliminarSub({tid})'>✕</button>"
            f"  </td>"
            f"</tr>"
        )
    # Nota si la tabla está vacía y se usa el fallback del .env
    nota_fallback = ""
    if not subs and _ALERT_TELEGRAM_ID:
        nota_fallback = (
            f"<div style='padding:10px 20px;font-size:.8rem;color:#888'>"
            f"Sin suscriptores configurados — usando <code>ALERT_TELEGRAM_ID={_ALERT_TELEGRAM_ID}</code> del .env como fallback."
            f"</div>"
        )

    # ── Vínculos ────────────────────────────────────────────────────────────
    vinculos = obtener_todos_los_vinculos()
    filas_vinculos = ""
    for v in vinculos:
        tid  = v["telegram_id"]
        mail = v["email"]
        fecha = v["creado_en"][:10] if v["creado_en"] else "—"
        filas_vinculos += (
            f"<tr>"
            f"  <td><code>{tid}</code></td>"
            f"  <td>{mail}</td>"
            f"  <td>{fecha}</td>"
            f"  <td class='centro'>"
            f"    <button class='btn-del' onclick=\"eliminarVinculo({tid}, '{mail}')\">✕</button>"
            f"  </td>"
            f"</tr>"
        )

    # ── Base de datos ───────────────────────────────────────────────────────
    db_inf = _estado_db()
    db_modo_label  = "☁️ Turso (cloud)" if db_inf["modo"] == "turso" else "💾 SQLite local"
    db_estado_badge = (
        '<span class="badge-ok">✅ Conectada</span>'
        if db_inf["ok"] else
        f'<span class="badge-warn">❌ {db_inf["error"]}</span>'
    )

    # ── Sistema — barras de progreso ────────────────────────────────────────
    ram_mb   = sistema["ram_usado_mb"]
    ram_pct  = sistema["ram_pct_render"]
    ram_col  = _color_pct(ram_pct)
    ram_lbl  = f"{ram_mb} MB / {_RENDER_RAM_MB} MB" if ram_mb is not None else "No disponible"
    ram_bar  = f'<div class="prog-fill" style="width:{ram_pct or 0}%;background:{ram_col}"></div>'

    cpu_str    = f'{sistema["cpu_pct"]}%' if sistema["cpu_pct"] is not None else "—"
    uptime_str = _formato_uptime(sistema["uptime_seg"])

    # ── Bot / IA ────────────────────────────────────────────────────────────
    proveedor = os.getenv("LLM_PROVIDER", "groq").upper()
    modelo    = os.getenv("LLM_MODEL", "—")
    tg_status = "✅ Activo" if _ENABLE_TELEGRAM else "⚪ Desactivado"
    webhook   = "🌐 Webhook" if _WEBHOOK_URL else "📡 Polling"

    # ── Estadísticas IA ─────────────────────────────────────────────────────
    def _badge_pct(pct):
        col = "#27ae60" if pct >= 95 else "#e67e22" if pct >= 80 else "#e74c3c"
        return f'<span style="color:{col};font-weight:700">{pct}%</span>'

    filas_proveedores = ""
    for p in stats_ia["por_proveedor"]:
        es_fallback = any(r["n"] > 0 for r in [{"n": p["fallbacks"]}])
        tag_fallback = ' <small style="color:#e67e22">(secundario)</small>' if p["fallbacks"] > 0 else ""
        filas_proveedores += (
            f"<tr>"
            f"<td><strong>{p['proveedor']}</strong>{tag_fallback}</td>"
            f"<td class='num'>{p['total']}</td>"
            f"<td class='num' style='color:#27ae60'>{p['exitosas']}</td>"
            f"<td class='num' style='color:#e74c3c'>{p['errores']}</td>"
            f"<td class='num' style='color:#e67e22'>{p['fallbacks']}</td>"
            f"<td class='num'>{_badge_pct(p['pct_ok'])}</td>"
            f"</tr>"
        )

    filas_errores_ia = ""
    _etiquetas_error = {
        "rate_limit": "⏳ Límite de tasa",
        "timeout":    "⏰ Timeout",
        "conexion":   "🌐 Conexión",
        "auth":       "🔑 Autenticación",
        "limite":     "📏 Tokens excedidos",
        "generico":   "⚠️ Genérico",
    }
    for e in stats_ia["errores_por_tipo"]:
        label = _etiquetas_error.get(e["tipo"], e["tipo"])
        filas_errores_ia += (
            f"<tr><td>{label}</td><td class='num'>{e['n']}</td></tr>"
        )
    if not filas_errores_ia:
        filas_errores_ia = "<tr><td colspan='2' style='color:#27ae60;text-align:center'>✅ Sin errores registrados</td></tr>"

    ia_sin_datos = stats_ia["total"] == 0
    ia_nota = "<div style='padding:16px;color:#999;text-align:center'>Sin llamadas registradas aún — los datos aparecerán tras las primeras peticiones.</div>" if ia_sin_datos else ""

    # ── Análisis de sesiones ─────────────────────────────────────────────────
    if stats_ses and stats_ses["por_usuario"]:
        res_ses = stats_ses["resumen"]
        filas_ses = ""
        for u in stats_ses["por_usuario"]:
            uid = u["user_id"]
            uid_label = "<em style='color:#888'>Add-in</em>" if uid == _ADDIN_ANON_ID else f"<code>{uid}</code>"
            dias = u["dias_inactivo"]
            if dias <= 7:
                estado = '<span style="color:#27ae60;font-weight:600">● Activo</span>'
            elif dias <= 30:
                estado = '<span style="color:#e67e22;font-weight:600">● Reciente</span>'
            else:
                estado = '<span style="color:#e74c3c;font-weight:600">● Inactivo</span>'
            if u["es_ocasional"]:
                estado += ' <small style="color:#aaa">ocasional</small>'
            ult = u["ultima_sesion"][:16].replace("T", " ") if u["ultima_sesion"] != "—" else "—"
            dias_str = "hoy" if dias < 1 else f"{int(dias)}d"
            filas_ses += (
                f"<tr>"
                f"<td>{uid_label}</td>"
                f"<td class='num'>{u['total_sesiones']}</td>"
                f"<td class='num'>{_fmt_min(u['duracion_media_min'])}</td>"
                f"<td class='num'>{_fmt_min(u['sesion_mas_larga_min'])}</td>"
                f"<td>{ult}</td>"
                f"<td class='num'>{dias_str}</td>"
                f"<td class='centro'>{estado}</td>"
                f"</tr>"
            )
        html_sesiones = f"""
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;padding:16px 20px 8px">
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.4rem;color:#27ae60">{res_ses['activos_7d']}</div>
        <div class="lbl">Activos 7 días</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.4rem">{res_ses['activos_30d']}</div>
        <div class="lbl">Activos 30 días</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.4rem">{res_ses['total_sesiones']}</div>
        <div class="lbl">Sesiones totales</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.4rem">{_fmt_min(res_ses['duracion_media_global_min'])}</div>
        <div class="lbl">Dur. media sesión</div>
      </div>
    </div>
    <table>
      <thead>
        <tr>
          <th>ID</th><th>Sesiones</th><th>Dur. media</th>
          <th>Más larga</th><th>Última sesión</th><th>Sin acceso</th><th>Estado</th>
        </tr>
      </thead>
      <tbody>{filas_ses}</tbody>
    </table>
    <div style="padding:6px 20px 12px;font-size:.72rem;color:#aaa">
      Sesión = bloque de mensajes con ≤ 30 min de silencio entre ellos. Ordenados por último acceso.
    </div>"""
    else:
        html_sesiones = '<div style="padding:20px;color:#999;text-align:center">Sin datos de sesiones todavía.</div>'

    # ── Logs ────────────────────────────────────────────────────────────────
    n_errores  = sum(1 for l in logs if "ERROR" in l)
    lineas_log = "".join(_html_linea_log(l) for l in reversed(logs)) if logs else \
                 '<div style="color:#999;padding:12px">Sin entradas de log</div>'

    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Asistente Excel — Administración</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
       background:#f0f2f5;color:#1a1a2e;min-height:100vh}}
  .topbar{{background:#2c3e7a;color:#fff;padding:14px 24px;
           display:flex;align-items:center;justify-content:space-between}}
  .topbar h1{{font-size:1.1rem;font-weight:600}}
  .topbar span{{font-size:.8rem;opacity:.7}}
  .main{{padding:24px;max-width:1200px;margin:0 auto}}

  /* Tarjetas */
  .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
          gap:14px;margin-bottom:24px}}
  .card{{background:#fff;border-radius:10px;padding:18px 14px;
         box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}}
  .card .val{{font-size:1.9rem;font-weight:700;color:#2c3e7a}}
  .card .lbl{{font-size:.75rem;color:#666;margin-top:4px;text-transform:uppercase;
              letter-spacing:.03em}}

  /* Secciones */
  .section{{background:#fff;border-radius:10px;
            box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:20px;overflow:hidden}}
  .section-head{{padding:13px 20px;background:#f8f9fb;
                 border-bottom:1px solid #eee;font-weight:600;font-size:.9rem;
                 display:flex;align-items:center;justify-content:space-between}}
  .section-head .head-actions{{display:flex;gap:8px}}

  /* Grid de 2 columnas */
  .two-col{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}}
  @media(max-width:700px){{.two-col{{grid-template-columns:1fr}}}}

  /* Métricas simples */
  .metric-list{{padding:16px 20px;display:flex;flex-direction:column;gap:12px}}
  .metric-row{{display:flex;justify-content:space-between;align-items:center;
               font-size:.85rem}}
  .metric-row .mk{{color:#666}}
  .metric-row .mv{{font-weight:600;color:#1a1a2e}}

  /* Barra de progreso */
  .prog-wrap{{margin-top:4px;background:#eee;border-radius:4px;height:10px;overflow:hidden}}
  .prog-fill{{height:100%;border-radius:4px;transition:width .4s}}
  .prog-label{{font-size:.75rem;color:#666;margin-top:3px;display:flex;
               justify-content:space-between}}

  /* Gráfico */
  .chart{{display:flex;align-items:flex-end;gap:8px;
          padding:20px 20px 10px;height:140px}}
  .bar-item{{display:flex;flex-direction:column;align-items:center;flex:1;height:100%}}
  .bar-fill{{background:#2c3e7a;border-radius:3px 3px 0 0;width:100%;
             min-height:4px;transition:height .3s}}
  .bar-label,.bar-val{{font-size:.68rem;color:#666;margin-top:3px}}
  .bar-val{{color:#2c3e7a;font-weight:600}}

  /* Tablas */
  table{{width:100%;border-collapse:collapse;font-size:.85rem}}
  th{{padding:10px 14px;text-align:left;font-weight:600;
      background:#f8f9fb;border-bottom:2px solid #eee;color:#444}}
  td{{padding:9px 14px;border-bottom:1px solid #f0f0f0;vertical-align:middle}}
  tr:last-child td{{border-bottom:none}}
  tr:hover td{{background:#fafbff}}
  .num{{text-align:right;font-variant-numeric:tabular-nums}}
  .centro{{text-align:center}}
  code{{background:#f0f0f0;padding:2px 5px;border-radius:4px;font-size:.8rem}}
  .badge-ok{{color:#27ae60;font-weight:600}}
  .badge-warn{{color:#e67e22}}

  /* Botones */
  .btn-del{{background:#e74c3c;color:#fff;border:none;border-radius:4px;
            padding:3px 9px;cursor:pointer;font-size:.8rem}}
  .btn-del:hover{{background:#c0392b}}
  .btn-sm{{padding:4px 12px;border:1px solid #ddd;border-radius:6px;
           background:#fff;cursor:pointer;font-size:.8rem;color:#444}}
  .btn-sm.active{{background:#2c3e7a;color:#fff;border-color:#2c3e7a}}
  .btn-sm:hover:not(.active){{background:#f0f2f5}}

  /* Formulario vínculos */
  .form-vincular{{display:flex;gap:8px;padding:14px 20px;
                  border-top:1px solid #eee;flex-wrap:wrap}}
  .form-vincular input{{flex:1;min-width:140px;padding:6px 10px;
                        border:1px solid #ddd;border-radius:6px;font-size:.85rem}}
  .form-vincular button{{padding:6px 16px;background:#2c3e7a;color:#fff;
                         border:none;border-radius:6px;cursor:pointer;font-size:.85rem}}
  .form-vincular button:hover{{background:#1a2a5e}}

  /* Visor de logs */
  .log-box{{height:320px;overflow-y:auto;font-family:'Courier New',monospace;
            font-size:.75rem;line-height:1.5;padding:12px 16px;background:#fafafa}}
  .log-error{{color:#c0392b;background:#fdecea;padding:1px 4px;border-radius:2px;margin:1px 0}}
  .log-warn{{color:#8b6914;background:#fef9e7;padding:1px 4px;border-radius:2px;margin:1px 0}}
  .log-info{{color:#2c3e7a}}
  .log-debug{{color:#aaa}}
  .log-box.only-errors .log-info,
  .log-box.only-errors .log-debug{{display:none}}
  .log-count-err{{background:#e74c3c;color:#fff;border-radius:10px;
                  padding:1px 7px;font-size:.72rem;font-weight:700;margin-left:6px}}
</style>
</head>
<body>
<div class="topbar">
  <h1>🤖 Asistente Excel — Panel de administración</h1>
  <span>Actualizado: {ahora}</span>
</div>
<div class="main">

  <!-- ① Tarjetas KPI -->
  <div class="cards">
    <div class="card">
      <div class="val">{stats['total_usuarios']}</div>
      <div class="lbl">Usuarios totales</div>
    </div>
    <div class="card">
      <div class="val">{stats['total_mensajes']:,}</div>
      <div class="lbl">Mensajes totales</div>
    </div>
    <div class="card">
      <div class="val {'badge-ok' if stats['mensajes_hoy'] > 0 else ''}">{stats['mensajes_hoy']}</div>
      <div class="lbl">Mensajes hoy</div>
    </div>
    <div class="card">
      <div class="val">{usuarios_hoy}</div>
      <div class="lbl">Usuarios activos hoy</div>
    </div>
    <div class="card">
      <div class="val">{len(vinculos)}</div>
      <div class="lbl">Vínculos Telegram</div>
    </div>
    <div class="card">
      <div class="val">{uptime_str}</div>
      <div class="lbl">Uptime servidor</div>
    </div>
  </div>

  <!-- ② Sistema + Bot/IA en dos columnas -->
  <div class="two-col">

    <!-- Sistema -->
    <div class="section">
      <div class="section-head">💻 Sistema <small style="font-weight:400;color:#888">(límites Render free)</small></div>
      <div class="metric-list">

        <div>
          <div class="metric-row">
            <span class="mk">RAM usada</span>
            <span class="mv">{ram_lbl}</span>
          </div>
          <div class="prog-wrap">{ram_bar}</div>
          <div class="prog-label"><span>0</span><span>límite {_RENDER_RAM_MB} MB</span></div>
        </div>

        <div class="metric-row"><span class="mk">CPU</span><span class="mv">{cpu_str}</span></div>
        <div class="metric-row"><span class="mk">Python</span><span class="mv">{sistema['python_version']}</span></div>
        <div class="metric-row" style="margin-top:4px">
          <span class="mk">data/ (BD + logs)</span><span class="mv">{sistema['data_mb']} MB</span>
        </div>
        <div class="metric-row"><span class="mk">└ logs/</span><span class="mv">{sistema['logs_mb']} MB</span></div>
        <div class="metric-row"><span class="mk">└ temp/</span><span class="mv">{sistema['temp_mb']} MB</span></div>
        <div class="metric-row" style="margin-top:4px;font-size:.72rem;color:#aaa">
          <span>Disco raíz no medible (Render comparte host físico)</span>
        </div>

      </div>
    </div>

    <!-- Bot / IA -->
    <div class="section">
      <div class="section-head">🤖 Bot e IA</div>
      <div class="metric-list">
        <div class="metric-row"><span class="mk">Telegram</span><span class="mv">{tg_status}</span></div>
        <div class="metric-row"><span class="mk">Modo conexión</span><span class="mv">{webhook}</span></div>
        <div class="metric-row"><span class="mk">Proveedor IA</span><span class="mv">{proveedor}</span></div>
        <div class="metric-row"><span class="mk">Modelo</span><span class="mv" style="font-size:.8rem">{modelo}</span></div>
        <div class="metric-row"><span class="mk">Base de datos</span><span class="mv">{db_modo_label}</span></div>
        <div class="metric-row"><span class="mk">Estado BD</span><span class="mv">{db_estado_badge}</span></div>
        <div class="metric-row"><span class="mk">URL BD</span>
          <span class="mv"><code style="font-size:.72rem;word-break:break-all">{db_inf['url']}</code></span>
        </div>
      </div>
    </div>

  </div>

  <!-- ③ Actividad 7 días -->
  <div class="section">
    <div class="section-head">📊 Actividad — últimos 7 días</div>
    <div class="chart">
      {barras_html if barras_html else '<p style="padding:20px;color:#999">Sin datos</p>'}
    </div>
  </div>

  <!-- ④ Estadísticas IA -->
  <div class="section">
    <div class="section-head">🧠 Inteligencia Artificial — estadísticas</div>
    {ia_nota}
    {f'''
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;padding:16px 20px 8px">
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.5rem">{stats_ia["total"]:,}</div>
        <div class="lbl">Llamadas totales</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.5rem;color:{"#27ae60" if stats_ia["tasa_exito"]>=95 else "#e67e22" if stats_ia["tasa_exito"]>=80 else "#e74c3c"}">{stats_ia["tasa_exito"]}%</div>
        <div class="lbl">Tasa de éxito</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.5rem;color:#e67e22">{stats_ia["fallbacks"]}</div>
        <div class="lbl">Saltos a secundaria</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.5rem;color:#e74c3c">{stats_ia["errores"]}</div>
        <div class="lbl">Errores totales</div>
      </div>
      <div class="card" style="padding:12px 10px">
        <div class="val" style="font-size:1.5rem">{stats_ia["total_hoy"]}</div>
        <div class="lbl">Llamadas hoy</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:0;border-top:1px solid #eee">
      <div style="padding:14px 20px;border-right:1px solid #eee">
        <div style="font-weight:600;font-size:.85rem;margin-bottom:8px">Por proveedor</div>
        <table>
          <thead>
            <tr>
              <th>Proveedor</th><th>Total</th><th>OK</th><th>Err</th><th>Fallb.</th><th>%OK</th>
            </tr>
          </thead>
          <tbody>{filas_proveedores or "<tr><td colspan='6' style='color:#999;text-align:center'>—</td></tr>"}</tbody>
        </table>
      </div>
      <div style="padding:14px 20px">
        <div style="font-weight:600;font-size:.85rem;margin-bottom:8px">Tipos de error</div>
        <table>
          <thead><tr><th>Tipo</th><th>Veces</th></tr></thead>
          <tbody>{filas_errores_ia}</tbody>
        </table>
      </div>
    </div>
    ''' if not ia_sin_datos else ""}
  </div>

  <!-- ⑤ Usuarios -->
  <div class="section">
    <div class="section-head">👤 Usuarios ({stats['total_usuarios']})</div>
    <table>
      <thead>
        <tr>
          <th>Telegram ID</th><th>Email (Add-in)</th>
          <th>Msgs enviados</th><th>Total msgs</th>
          <th>Última actividad</th><th>Versión Excel</th><th>Modo</th>
        </tr>
      </thead>
      <tbody>
        {filas_usuarios or '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px">Sin usuarios</td></tr>'}
      </tbody>
    </table>
  </div>

  <!-- ⑤b Análisis de sesiones -->
  <div class="section">
    <div class="section-head">📈 Análisis de sesiones</div>
    {html_sesiones}
  </div>

  <!-- ⑥ Vínculos -->
  <div class="section">
    <div class="section-head">🔗 Vínculos Telegram ↔ Add-in ({len(vinculos)})</div>
    <table>
      <thead>
        <tr><th>Telegram ID</th><th>Email</th><th>Vinculado el</th><th></th></tr>
      </thead>
      <tbody>
        {filas_vinculos or '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px">Sin vínculos</td></tr>'}
      </tbody>
    </table>
    <form class="form-vincular" onsubmit="agregarVinculo(event)">
      <input type="number" id="inp-tid"   placeholder="Telegram ID" required>
      <input type="email"  id="inp-email" placeholder="email@empresa.com" required>
      <button type="submit">+ Añadir vínculo</button>
    </form>
  </div>

  <!-- ⑦ Notificaciones -->
  <div class="section">
    <div class="section-head">🔔 Notificaciones del sistema ({len(subs)})</div>
    <table>
      <thead>
        <tr><th>Telegram ID</th><th>Etiqueta</th><th>Estado</th><th></th></tr>
      </thead>
      <tbody>
        {filas_subs or '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px">Sin suscriptores</td></tr>'}
      </tbody>
    </table>
    {nota_fallback}
    <form class="form-vincular" onsubmit="agregarSub(event)">
      <input type="number" id="inp-sub-tid"   placeholder="Telegram ID" required>
      <input type="text"   id="inp-sub-label" placeholder="Etiqueta (ej: Roberto)">
      <button type="submit">+ Añadir</button>
    </form>
  </div>

  <!-- ⑧ Logs -->
  <div class="section">
    <div class="section-head">
      <span>📋 Logs recientes
        {'<span class="log-count-err">' + str(n_errores) + ' errores</span>' if n_errores else ''}
      </span>
      <div class="head-actions">
        <button class="btn-sm active" id="btn-todos"   onclick="filtrarLogs('todos')">Todos</button>
        <button class="btn-sm"        id="btn-errores" onclick="filtrarLogs('errores')">Solo errores</button>
        <button class="btn-sm" onclick="location.reload()">↻ Refrescar</button>
      </div>
    </div>
    <div class="log-box" id="log-box">
      {lineas_log}
    </div>
  </div>

</div>
<script>
  const _key = new URLSearchParams(window.location.search).get("key") || "";

  // Scroll al final de los logs al cargar
  const lb = document.getElementById("log-box");
  if (lb) lb.scrollTop = lb.scrollHeight;

  function filtrarLogs(modo) {{
    const box = document.getElementById("log-box");
    document.getElementById("btn-todos").classList.toggle("active", modo === "todos");
    document.getElementById("btn-errores").classList.toggle("active", modo === "errores");
    box.classList.toggle("only-errors", modo === "errores");
  }}

  async function eliminarVinculo(tid, email) {{
    if (!confirm("¿Eliminar el vínculo de " + email + "?")) return;
    const resp = await fetch(
      "/admin/vinculos?key=" + encodeURIComponent(_key) +
      "&telegram_id=" + tid + "&email=" + encodeURIComponent(email),
      {{ method: "DELETE" }}
    );
    if (resp.ok) location.reload();
    else alert("Error al eliminar el vínculo");
  }}

  async function agregarVinculo(e) {{
    e.preventDefault();
    const tid   = document.getElementById("inp-tid").value.trim();
    const email = document.getElementById("inp-email").value.trim();
    const resp  = await fetch("/admin/vinculos?key=" + encodeURIComponent(_key), {{
      method: "POST",
      headers: {{ "Content-Type": "application/json" }},
      body: JSON.stringify({{ telegram_id: parseInt(tid), email }})
    }});
    if (resp.ok) location.reload();
    else alert("Error al añadir el vínculo");
  }}

  async function agregarSub(e) {{
    e.preventDefault();
    const tid      = document.getElementById("inp-sub-tid").value.trim();
    const etiqueta = document.getElementById("inp-sub-label").value.trim();
    const resp = await fetch("/admin/notificaciones?key=" + encodeURIComponent(_key), {{
      method: "POST",
      headers: {{ "Content-Type": "application/json" }},
      body: JSON.stringify({{ telegram_id: parseInt(tid), etiqueta }})
    }});
    if (resp.ok) location.reload();
    else alert("Error al añadir suscriptor");
  }}

  async function toggleSub(tid) {{
    const resp = await fetch(
      "/admin/notificaciones/" + tid + "/toggle?key=" + encodeURIComponent(_key),
      {{ method: "PATCH" }}
    );
    if (resp.ok) location.reload();
    else alert("Error al cambiar estado");
  }}

  async function eliminarSub(tid) {{
    if (!confirm("¿Eliminar el suscriptor " + tid + " de las notificaciones?")) return;
    const resp = await fetch(
      "/admin/notificaciones/" + tid + "?key=" + encodeURIComponent(_key),
      {{ method: "DELETE" }}
    );
    if (resp.ok) location.reload();
    else alert("Error al eliminar suscriptor");
  }}
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Ficheros estáticos del Add-in (solo en cloud — dist/ generado por webpack)
# ---------------------------------------------------------------------------

_DIST = os.path.join(os.path.dirname(__file__), "excel-addin", "dist")
if os.path.isdir(_DIST):
    app.mount("/", StaticFiles(directory=_DIST, html=True), name="addin")
    logger.info("Add-in estático servido desde %s", _DIST)


# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
