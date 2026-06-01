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
import html as _html_mod
import logging
import os
import re
import time
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
                          extraer_estructura_excel, extraer_regla_formato,
                          extraer_peticion_grafico, extraer_params_pivote,
                          extraer_formula, _col_letra,
                          obtener_respuesta, LLMError)
from config import SYSTEM_PROMPT_ADDIN, ENABLE_TELEGRAM as _ENABLE_TELEGRAM, ENABLE_ADDIN as _ENABLE_ADDIN
from utils.macros import listar_macros as _listar_macros_db, obtener_macro as _obtener_macro_db
from utils.feature_config import (obtener_config as _obtener_feature_config,
                                   esta_activo as _feature_activa,
                                   toggle_feature as _toggle_feature)

load_dotenv()
configurar_logging()

logger = logging.getLogger(__name__)

_API_KEY     = os.getenv("API_KEY", "")
_ADMIN_KEY   = os.getenv("ADMIN_KEY", "")
_WEBHOOK_URL = os.getenv("WEBHOOK_URL", "").rstrip("/")
_BOT_TOKEN   = os.getenv("TELEGRAM_TOKEN", "")

# Módulos activables — leídos de config.py (que aplica los defaults según APP_MODE)
# Los tests pueden sobreescribir api_mod._ENABLE_TELEGRAM directamente si lo necesitan

# ID de Telegram al que enviar alertas del sistema (por defecto el primer AUTHORIZED_USER)
_ids_autorizados  = [u.strip() for u in os.getenv("AUTHORIZED_USERS", "").split(",") if u.strip()]
# Solo se envían alertas si ALERT_TELEGRAM_ID está explícitamente en .env.
# El fallback automático a AUTHORIZED_USERS[0] se eliminó — causaba spam no deseado.
_ALERT_TELEGRAM_ID = int(os.getenv("ALERT_TELEGRAM_ID", "0") or "0")

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
    from utils.alert_config import esta_activo as _alerta_activa
    _ultima: dict[str, float] = {}

    await asyncio.sleep(60)   # espera inicial para dejar que el servidor arranque
    while True:
        try:
            ahora   = time.time()
            sistema = _obtener_info_sistema()
            alertas = []

            ram_pct = sistema.get("ram_pct_render")
            if ram_pct and ram_pct >= _ALERTA_PCT and _alerta_activa("ram"):
                if ahora - _ultima.get("ram", 0) > _ALERTA_COOLDOWN:
                    alertas.append(
                        f"🔴 *RAM al {ram_pct:.0f}%*\n"
                        f"   {sistema['ram_usado_mb']} MB / {_RENDER_RAM_MB} MB (límite Render)"
                    )
                    _ultima["ram"] = ahora

            # Alerta si data/ supera 400 MB (runtime creciente; código ocupa ~300 MB fijos)
            data_mb = sistema["data_mb"]
            if data_mb >= 400 and _alerta_activa("disco"):
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

# Detección de petición de análisis estadístico en /edit (fallback cuando el LLM devuelve RESPUESTA_LIBRE)
_RE_ANALISIS = re.compile(
    r"\b(analiza[r]?|análisis|calidad\s+de\s+datos|estadísticas?|estadisticas?|"
    r"correlaciones?|resumen\s+estadístico|resumen\s+estadistico|describe\s+(?:el|los|la|las)\s+datos?)\b",
    re.IGNORECASE,
)
_RE_GRAFICO_PEDIDO = re.compile(r"\bgráfico|grafico|chart|gr[aá]fica\b", re.IGNORECASE)

async def _manejador_error_bot(update: object, context: object) -> None:
    """Handler de errores de python-telegram-bot — notifica al administrador."""
    from utils.alert_config import esta_activo as _alerta_activa
    error = context.error  # type: ignore[attr-defined]
    logger.error("Error PTB: %s", error, exc_info=error)
    nombre = type(error).__name__
    ahora = time.time()
    if ahora - _errores_bot_vistos.get(nombre, 0) > _COOLDOWN_BOT_ERR and _alerta_activa("bot_error"):
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
    from utils.alert_config import esta_activo as _alerta_activa
    if _alerta_activa("arranque"):
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
async def _middleware_http(request: Request, call_next):
    from utils.alert_config import esta_activo as _alerta_activa
    t0 = time.perf_counter()
    response = await call_next(request)
    ms = (time.perf_counter() - t0) * 1000
    ruta = f"{request.method} {request.url.path}"
    logger.info("HTTP %-30s → %d  (%.0f ms)", ruta, response.status_code, ms)
    if response.status_code == 500 and _alerta_activa("error_500"):
        ahora = time.time()
        if ahora - _errores_500_vistos.get(ruta, 0) > _COOLDOWN_500:
            _errores_500_vistos[ruta] = ahora
            asyncio.create_task(_notificar_telegram(
                f"💥 *Error 500 en el servidor*\n`{ruta}`"
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
    user_email: str | None = None
    display_name: str | None = None
    excel_version: str | None = None


class PeticionAnalisis(BaseModel):
    datos: list[list]


class PeticionEdicion(BaseModel):
    datos: list[list]
    instruccion: str
    historial: list[dict] = []
    device_id: str | None = None
    user_email: str | None = None
    display_name: str | None = None
    excel_version: str | None = None


class PeticionFormato(BaseModel):
    datos: list[list]
    instruccion: str
    device_id: str | None = None
    user_email: str | None = None
    display_name: str | None = None
    excel_version: str | None = None


class PeticionFeedback(BaseModel):
    device_id: str | None = None
    pregunta: str
    respuesta: str
    tipo: str = "positivo"


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


def _usuario_addin(device_id: str | None, display_name: str | None = None) -> int:
    """Devuelve el user_id para este device_id.

    Prioridad:
    1. telegram_id si el dispositivo está vinculado (flujo B / por nombre / por email)
    2. user_id negativo estable asignado por device_emails (usuario Add-in propio)
    3. _ADDIN_ANON_ID (0) si no hay device_id en absoluto
    """
    if not device_id:
        return _ADDIN_ANON_ID
    from utils.user_links import obtener_device_link
    link = obtener_device_link(device_id)
    if link:
        return link["telegram_id"]
    from utils.device_emails import obtener_o_crear_usuario
    return obtener_o_crear_usuario(device_id, display_name=display_name)


def _registrar_addin(
    device_id: str | None,
    texto: str,
    user_email: str | None = None,
    display_name: str | None = None,
    excel_version: str | None = None,
) -> None:
    """Escribe la pregunta del Add-in en historial para que el usuario aparezca en el panel."""
    try:
        from utils.history import agregar_mensaje
        agregar_mensaje(_usuario_addin(device_id, display_name), "user", texto[:2000])
        if device_id:
            from utils.device_emails import guardar_email
            guardar_email(device_id, user_email, display_name, excel_version)
    except Exception as exc:
        logger.warning("No se pudo registrar uso del Add-in: %s", exc)


def _verificar_addin_activo():
    if not _ENABLE_ADDIN:
        raise HTTPException(status_code=503, detail="Módulo Add-in desactivado (ENABLE_ADDIN=false)")


@app.post("/ask")
def ask(peticion: PeticionPregunta, _: None = Depends(_verificar_clave),
        __: None = Depends(_verificar_addin_activo)) -> dict:
    _registrar_addin(peticion.device_id, peticion.pregunta, peticion.user_email, peticion.display_name, peticion.excel_version)
    # Sin datos: el LLM decide si crear tabla o responder como chat
    _uid_ask = _usuario_addin(peticion.device_id, peticion.display_name)
    if not peticion.datos or len(peticion.datos) < 2:
        estructura = extraer_estructura_excel(peticion.pregunta)
        if estructura:
            columnas = estructura.get("columnas", [])
            datos_filas = estructura.get("datos", [])
            matriz = [columnas] + [
                [("" if v is None else v) for v in fila]
                for fila in datos_filas
            ]
            titulo = estructura.get("titulo", "Nueva tabla")
            nueva_hoja = estructura.get("nueva_hoja", False)
            return {
                "tipo": "datos",
                "datos_modificados": matriz,
                "nombre_hoja": titulo if nueva_hoja else None,
                "nueva_hoja": nueva_hoja,
                "descripcion": f"Tabla '{titulo}' creada ({len(datos_filas)} filas)",
            }
        return {"respuesta": obtener_respuesta(
            peticion.historial, peticion.pregunta,
            system_override=SYSTEM_PROMPT_ADDIN,
            user_id=_uid_ask,
        )}

    df = _a_dataframe(peticion.datos)

    try:
        query = extraer_query_dsl(df, peticion.pregunta)
        if query:
            try:
                resultado, descripcion = ejecutar_query(df, query)
                return {"respuesta": _resultado_a_texto(resultado, descripcion)}
            except QueryError as error:
                logger.warning("DSL falló, usando LLM libre: %s", error)

        # Detectar intención de edición — redirigir internamente a la lógica de /edit
        ops = extraer_operacion_edicion(df, peticion.pregunta)
        if isinstance(ops, list) and ops:
            pasos = _ejecutar_pipeline(df, ops, peticion.pregunta)
            if len(pasos) == 1:
                return pasos[0]
            return {
                "tipo": "pipeline",
                "pasos": pasos,
                "descripcion": "; ".join(p.get("descripcion", "") for p in pasos if p.get("descripcion")),
            }

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
            user_id=_uid_ask,
        )}
    except LLMError as exc:
        logger.warning("LLMError en /ask: %s", exc)
        return {"respuesta": str(exc)}


def _ejecutar_pipeline(df: "pd.DataFrame", ops: list[dict], instruccion: str) -> list[dict]:
    """Ejecuta una lista de operaciones DSL secuencialmente.

    Las operaciones de datos (ordenar, filtrar, etc.) modifican df_actual en orden.
    Las operaciones visuales (formato_condicional, grafico, tabla_dinamica) se ejecutan
    con el df_actual resultante de todos los pasos de datos anteriores.

    Devuelve la lista de pasos con sus resultados, lista para devolver al cliente.
    """
    pasos: list[dict] = []
    df_actual = df.copy()  # una sola copia defensiva para todo el pipeline

    for op in ops:
        # Normalizar nombre del campo op
        if "op" not in op and "operacion" in op:
            op = dict(op)
            op["op"] = op.pop("operacion")

        nombre_op = op.get("op", "")

        # ── Ops visuales: no modifican df, usan df_actual hasta este punto ──────
        if nombre_op == "formato_condicional":
            _re_decimales = re.compile(r'\b(decimal|decimales|\d+\s*decimal)\b', re.IGNORECASE)
            if _re_decimales.search(instruccion):
                # El LLM usó formato_condicional para pedir formato numérico — corregir
                m_decs = re.search(r'(\d+)\s*decimal', instruccion, re.IGNORECASE)
                decimales = int(m_decs.group(1)) if m_decs else 2
                col_fmt = op.get("col") or op.get("columna")
                if not col_fmt:
                    # Inferir del último añadir_columna del pipeline
                    for prev_op in ops:
                        if prev_op.get("op") == "añadir_columna":
                            col_fmt = prev_op.get("nombre")
                if col_fmt and col_fmt in df_actual.columns:
                    col_idx = list(df_actual.columns).index(col_fmt)
                    pasos.append({
                        "tipo": "formato_numero",
                        "col": col_fmt,
                        "col_idx": col_idx,
                        "decimales": decimales,
                        "descripcion": f"Columna '{col_fmt}' formateada con {decimales} decimal{'es' if decimales != 1 else ''}",
                    })
            else:
                reglas = extraer_regla_formato(df_actual, instruccion)
                if reglas:
                    pasos.append({
                        "tipo": "formato",
                        "reglas": reglas,
                        "descripcion": _describir_reglas_formato(reglas),
                    })

        elif nombre_op == "analisis":
            from excel.analyzer import analisis_estadistico_completo, analisis_correlaciones
            texto = analisis_estadistico_completo(df_actual)
            texto_corr, _ = analisis_correlaciones(df_actual)
            if texto_corr:
                texto += "\n\n" + texto_corr

            # Construir matriz estructurada para la hoja Excel
            cols_num = [c for c in df_actual.columns if pd.api.types.is_numeric_dtype(df_actual[c])]
            hoja_datos: list[list] = []

            hoja_datos.append(["CALIDAD DE DATOS", ""])
            hoja_datos.append(["Total filas", len(df_actual)])
            hoja_datos.append(["Total columnas", len(df_actual.columns)])
            hoja_datos.append(["Filas duplicadas", int(df_actual.duplicated().sum())])
            hoja_datos.append(["", ""])
            hoja_datos.append(["Valores nulos por columna", ""])
            for col in df_actual.columns:
                hoja_datos.append([f"  {col}", int(df_actual[col].isnull().sum())])
            hoja_datos.append(["", ""])

            if cols_num:
                hoja_datos.append(["ESTADÍSTICAS", "Media", "Mediana", "Mín", "Máx", "Desv. std"])
                for col in cols_num:
                    serie = df_actual[col].dropna()
                    if not serie.empty:
                        hoja_datos.append([
                            col,
                            round(float(serie.mean()), 2),
                            round(float(serie.median()), 2),
                            round(float(serie.min()), 2),
                            round(float(serie.max()), 2),
                            round(float(serie.std()), 2),
                        ])
                hoja_datos.append(["", "", "", "", "", ""])

            if len(cols_num) >= 2:
                corr = df_actual[cols_num].corr()
                hoja_datos.append(["CORRELACIONES"] + cols_num)
                for col in cols_num:
                    hoja_datos.append(
                        [col] + [round(float(corr.loc[col, c2]), 2) for c2 in cols_num]
                    )

            pasos.append({
                "tipo": "analisis_hoja",
                "texto": texto,
                "hoja_datos": hoja_datos,
                "descripcion": "Análisis estadístico completo",
            })

        elif nombre_op == "grafico":
            params = extraer_peticion_grafico(df_actual, instruccion)
            if params:
                pasos.append(_preparar_respuesta_grafico(df_actual, params))

        elif nombre_op == "tabla_dinamica":
            params = extraer_params_pivote(df_actual, instruccion)
            if params:
                pasos.append({
                    "tipo": "tabla_dinamica",
                    "params": params,
                    "descripcion": (
                        f"Tabla dinámica: {', '.join(params.get('filas', []))} "
                        f"→ {params.get('valores')} ({params.get('funcion', 'suma')})"
                    ),
                })

        elif nombre_op == "formula":
            params = extraer_formula(df_actual, instruccion)
            if params:
                formula_tmpl = params.get("formula", "")
                col_nueva    = params.get("col_nueva", "Fórmula")
                formulas = [
                    [formula_tmpl.replace("{row}", str(row))]
                    for row in range(2, len(df_actual) + 2)
                ]
                pasos.append({
                    "tipo": "formula",
                    "col_nueva": col_nueva,
                    "formulas": formulas,
                    "descripcion": (
                        f"Columna '{col_nueva}' con fórmula "
                        f"{formula_tmpl.replace('{row}', 'N')}"
                    ),
                })

        elif nombre_op == "query":
            pregunta_q = op.get("pregunta", "")
            if pregunta_q:
                try:
                    dsl_q = extraer_query_dsl(df_actual, pregunta_q)
                    es_aclaracion = isinstance(dsl_q, dict) and dsl_q.get("aclaracion_necesaria")
                    if dsl_q and not es_aclaracion:
                        ops_q = dsl_q if isinstance(dsl_q, list) else [dsl_q]
                        partes_q: list[str] = []
                        partes_escalares: list[str] = []   # solo resultados no-tabla
                        ultimo_df_q: pd.DataFrame | None = None
                        descripcion_q = ""
                        for q in ops_q:
                            resultado_q, descripcion_q = ejecutar_query(df_actual, q)
                            texto_q = _resultado_a_texto(resultado_q, descripcion_q)
                            partes_q.append(texto_q)
                            if isinstance(resultado_q, pd.DataFrame) and not resultado_q.empty:
                                ultimo_df_q = resultado_q
                            else:
                                partes_escalares.append(texto_q)
                        paso_q: dict = {
                            "tipo": "query_resultado",
                            "pregunta": pregunta_q,
                            "resultado": "\n\n".join(partes_q),
                            "texto_resumen": "\n\n".join(partes_escalares),
                            "descripcion": descripcion_q,
                        }
                        if ultimo_df_q is not None:
                            paso_q["datos_tabla"] = _df_a_matriz(ultimo_df_q)
                        pasos.append(paso_q)
                except (QueryError, Exception) as e_q:
                    logger.warning("Query inline falló en pipeline: %s", e_q)

        # ── Ops de datos: modifican df_actual ────────────────────────────────────
        else:
            try:
                cols_antes = list(df_actual.columns)
                df_actual, descripcion, _extras = aplicar_edicion(df_actual, op, copy_df=False)
                pasos.append({
                    "tipo": "edicion",
                    "operacion": nombre_op,
                    "destino": op.get("destino") or None,
                    "datos_modificados": _df_a_matriz(df_actual),
                    "descripcion": descripcion,
                })
                # Para añadir_columna aritmética: excluir la columna de datos_modificados
                # y emitir un paso formula_columna para que el frontend la escriba como
                # fórmula Excel viva en lugar de valor calculado.
                if nombre_op == "añadir_columna":
                    col1       = op.get("col1")
                    col2       = op.get("col2")
                    operador   = op.get("operador")
                    valor_fijo = op.get("valor_fijo")
                    cols_nuevas = [c for c in df_actual.columns if c not in cols_antes]
                    _ops_aritmeticas = {"+", "-", "*", "/"}
                    if cols_nuevas and col1 and operador and operador in _ops_aritmeticas:
                        col_nueva = cols_nuevas[0]
                        cols_list = list(df_actual.columns)
                        if col2 and col1 in cols_list and col2 in cols_list:
                            # col op col → fórmula con dos referencias (=E2*F2)
                            pasos[-1]["datos_modificados"] = _df_a_matriz(
                                df_actual.drop(columns=[col_nueva])
                            )
                            pasos.append({
                                "tipo": "formula_columna",
                                "col": col_nueva,
                                "col1_idx": cols_list.index(col1),
                                "col2_idx": cols_list.index(col2),
                                "operador": operador,
                                "descripcion": f"{col_nueva} = {col1} {operador} {col2}",
                            })
                        elif valor_fijo is not None and col1 in cols_list:
                            # col op constante → fórmula con referencia + literal (=E2*1.21)
                            pasos[-1]["datos_modificados"] = _df_a_matriz(
                                df_actual.drop(columns=[col_nueva])
                            )
                            pasos.append({
                                "tipo": "formula_columna",
                                "col": col_nueva,
                                "col1_idx": cols_list.index(col1),
                                "valor_fijo": valor_fijo,
                                "operador": operador,
                                "descripcion": f"{col_nueva} = {col1} {operador} {valor_fijo}",
                            })
                    # Funciones (redondear, abs, raiz…): los valores ya están en
                    # datos_modificados como estáticos — no se emite formula_columna.
            except (EditorError, Exception) as error:
                logger.warning("Op '%s' falló en pipeline: %s", nombre_op, error)

    return pasos


@app.post("/edit")
def edit(peticion: PeticionEdicion, _: None = Depends(_verificar_clave),
         __: None = Depends(_verificar_addin_activo)) -> dict:
    user_id = _usuario_addin(peticion.device_id, peticion.display_name)
    _registrar_addin(peticion.device_id, peticion.instruccion, peticion.user_email, peticion.display_name, peticion.excel_version)
    df = _a_dataframe(peticion.datos)

    # Inyectar macros disponibles solo si el módulo está activo
    _macros_on = _feature_activa("macros")
    nombres_macros = [m["nombre"] for m in _listar_macros_db(user_id)] if _macros_on else []
    resultado = extraer_operacion_edicion(df, peticion.instruccion,
                                          macros_disponibles=nombres_macros or None)
    logger.info("DSL resultado: %s", resultado)

    # Aclaración → devolver al frontend antes de ejecutar nada
    if isinstance(resultado, dict) and resultado.get("aclaracion_necesaria"):
        return {"tipo": "aclaracion", **resultado}

    if isinstance(resultado, list) and resultado:
        # Verificar si algún op de la lista necesita aclaración
        for op in resultado:
            if isinstance(op, dict) and op.get("aclaracion_necesaria"):
                return {"tipo": "aclaracion", **op}

        # Expandir ops de macro → sustituir por sus operaciones almacenadas
        ops_expandidas: list[dict] = []
        for op in resultado:
            if isinstance(op, dict) and op.get("op") == "macro":
                nombre_m = op.get("nombre", "").lower()
                macro_ops = _obtener_macro_db(user_id, nombre_m)
                if macro_ops:
                    logger.info("Macro '%s' expandida: %d pasos", nombre_m, len(macro_ops))
                    ops_expandidas.extend(macro_ops)
                else:
                    logger.warning("Macro '%s' no encontrada para user_id=%s — op ignorada",
                                   nombre_m, user_id)
            else:
                ops_expandidas.append(op)

        pasos = _ejecutar_pipeline(df, ops_expandidas, peticion.instruccion)
        logger.info("Pipeline pasos tipos: %s", [p.get("tipo") for p in pasos])

        if pasos:
            # Un solo paso → devolver directamente (retrocompatibilidad con frontend)
            if len(pasos) == 1:
                return pasos[0]
            return {
                "tipo": "pipeline",
                "pasos": pasos,
                "descripcion": "; ".join(p.get("descripcion", "") for p in pasos),
            }

    # Detección de análisis estadístico — tiene prioridad sobre query_dsl
    if _RE_ANALISIS.search(peticion.instruccion):
        ops_an: list[dict] = [{"op": "analisis"}]
        if _RE_GRAFICO_PEDIDO.search(peticion.instruccion):
            ops_an.append({"op": "grafico"})
        try:
            pasos_an = _ejecutar_pipeline(df, ops_an, peticion.instruccion)
        except Exception as e_an:
            logger.error("Error en pipeline analisis: %s", e_an, exc_info=True)
            pasos_an = []
        logger.info("Análisis por patrón — %d paso(s)", len(pasos_an))
        if pasos_an:
            if len(pasos_an) == 1:
                return pasos_an[0]
            return {
                "tipo": "pipeline",
                "pasos": pasos_an,
                "descripcion": "; ".join(p.get("descripcion", "") for p in pasos_an),
            }

    # La petición no es análisis ni edición → intentar como consulta de datos
    query = extraer_query_dsl(df, peticion.instruccion)
    if query and not query.get("aclaracion_necesaria"):
        try:
            resultado_q, descripcion = ejecutar_query(df, query)
            if isinstance(resultado_q, pd.DataFrame) and not resultado_q.empty:
                return {
                    "tipo": "edicion",
                    "datos_modificados": _df_a_matriz(resultado_q),
                    "descripcion": descripcion,
                }
            return {"tipo": "texto", "respuesta": _resultado_a_texto(resultado_q, descripcion)}
        except QueryError as error:
            logger.warning("Query DSL falló, usando LLM libre: %s", error)

    # Último recurso: detectar si el usuario quiere crear una tabla nueva (aunque haya datos seleccionados)
    estructura = extraer_estructura_excel(peticion.instruccion)
    if estructura:
        columnas_e = estructura.get("columnas", [])
        datos_filas = estructura.get("datos", [])
        matriz = [columnas_e] + [
            [("" if v is None else v) for v in fila]
            for fila in datos_filas
        ]
        titulo = estructura.get("titulo", "Nueva tabla")
        nueva_hoja = estructura.get("nueva_hoja", False)
        return {
            "tipo": "datos",
            "datos_modificados": matriz,
            "nombre_hoja": titulo if nueva_hoja else None,
            "nueva_hoja": nueva_hoja,
            "descripcion": f"Tabla '{titulo}' creada ({len(datos_filas)} filas)",
        }

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


def _preparar_respuesta_grafico(df: pd.DataFrame, params: dict) -> dict:
    """Prepara los datos para que el Add-in cree un gráfico nativo en Office.js."""
    col_y   = params.get("col_y")
    col_x   = params.get("col_x")
    tipo    = params.get("tipo", "barras")
    agregar = params.get("agregar")

    _AGG = {"suma": "sum", "promedio": "mean", "contar": "count", "max": "max", "min": "min"}

    try:
        if agregar and col_x and col_x in df.columns and col_y in df.columns:
            df_chart = (
                df.groupby(col_x)[col_y]
                .agg(_AGG.get(agregar, "sum"))
                .reset_index()
            )
        elif col_x and col_x in df.columns and col_y in df.columns:
            df_chart = df[[col_x, col_y]].copy()
        elif col_y in df.columns:
            df_chart = df[[col_y]].copy()
        else:
            df_chart = df.copy()
    except Exception:
        df_chart = df.copy()

    titulo = f"{col_y} por {col_x}" if col_x else str(col_y)

    return {
        "tipo": "grafico",
        "tipo_grafico": tipo,
        "datos_chart": _df_a_matriz(df_chart),
        "titulo": titulo,
        "descripcion": f"Gráfico de {tipo}: {titulo}",
    }


def _describir_una_regla(regla: dict) -> str:
    tipo = regla.get("tipo", "")
    col  = regla.get("col") or ""
    if tipo == "valor":
        return f"'{col}' {regla.get('op')} {regla.get('valor')} → {regla.get('color', '')}"
    if tipo == "top_bottom":
        dir_ = "superiores" if regla.get("direccion") == "top" else "inferiores"
        n    = regla.get("n", 10)
        pct  = "%" if regla.get("porcentaje") else ""
        return f"'{col}': {n}{pct} valores {dir_} → {regla.get('color', '')}"
    if tipo == "escala":
        return f"Escala de color en '{col}'"
    if tipo == "barra":
        return f"Barra de datos en '{col}'"
    if tipo == "icono":
        return f"Iconos ({regla.get('estilo', '')}) en '{col}'"
    if tipo == "texto":
        return f"'{col}' {regla.get('op')} '{regla.get('valor')}' → {regla.get('color', '')}"
    if tipo == "formula":
        return f"Fórmula en '{col or 'rango'}' → {regla.get('color', '')}"
    return "Regla de formato"


def _describir_reglas_formato(reglas: list[dict]) -> str:
    if len(reglas) == 1:
        return "Formato condicional: " + _describir_una_regla(reglas[0])
    partes = "; ".join(_describir_una_regla(r) for r in reglas)
    return f"Formato condicional ({len(reglas)} reglas): {partes}"


@app.post("/format")
def format_condicional(peticion: PeticionFormato, _: None = Depends(_verificar_clave),
                       __: None = Depends(_verificar_addin_activo)) -> dict:
    _registrar_addin(peticion.device_id, peticion.instruccion, peticion.user_email, peticion.display_name, peticion.excel_version)
    df = _a_dataframe(peticion.datos)
    reglas = extraer_regla_formato(df, peticion.instruccion)
    if not reglas:
        return {
            "tipo": "texto",
            "respuesta": (
                "No pude interpretar la regla de formato. "
                "Describe qué columna colorear, con qué condición y con qué color."
            ),
        }
    return {
        "tipo": "formato",
        "reglas": reglas,
        "descripcion": _describir_reglas_formato(reglas),
    }


@app.post("/feedback")
def feedback(peticion: PeticionFeedback, _: None = Depends(_verificar_clave),
             __: None = Depends(_verificar_addin_activo)) -> dict:
    """Guarda un ejemplo de respuesta valorada positivamente por el usuario (RAG)."""
    user_id = _usuario_addin(peticion.device_id)
    if not peticion.pregunta.strip() or not peticion.respuesta.strip():
        raise HTTPException(status_code=400, detail="pregunta y respuesta no pueden estar vacías")
    tipo = peticion.tipo if peticion.tipo in ("positivo", "negativo") else "positivo"
    try:
        from utils.rag import guardar_ejemplo
        guardar_ejemplo(user_id, peticion.pregunta, peticion.respuesta, tipo=tipo)
    except Exception as exc:
        logger.warning("No se pudo guardar ejemplo RAG: %s", exc)
    return {"ok": True}


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
        raise HTTPException(status_code=503, detail="Panel admin no disponible: ADMIN_KEY no configurada")
    if key != _ADMIN_KEY:
        raise HTTPException(status_code=403, detail="Clave de administrador inválida")


@app.get("/admin/stats")
def admin_stats(_: None = Depends(_verificar_admin)) -> dict:
    """Estadísticas de uso en formato JSON."""
    from utils.stats import obtener_estadisticas
    return obtener_estadisticas()


@app.get("/admin/stats.csv")
def admin_stats_csv(_: None = Depends(_verificar_admin)):
    """Exporta la tabla de usuarios en formato CSV."""
    import csv, io
    from utils.stats import obtener_estadisticas
    from fastapi.responses import StreamingResponse

    stats = obtener_estadisticas()
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["user_id", "email", "msgs_enviados", "total_msgs",
                     "ultima_actividad", "version_excel", "modo_respuesta"])
    for u in stats["usuarios"]:
        writer.writerow([
            u["user_id"],
            u.get("email") or "",
            u.get("msgs_enviados", 0),
            u.get("total_msgs", 0),
            u.get("ultima_actividad", ""),
            u.get("version_excel") or "",
            u.get("modo_respuesta") or "",
        ])
    buf.seek(0)
    return StreamingResponse(
        iter([buf.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=usuarios.csv"},
    )



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
# Configuración de alertas por tipo
# ---------------------------------------------------------------------------

@app.get("/admin/alert-config")
def admin_alert_config_list(_: None = Depends(_verificar_admin)) -> dict:
    from utils.alert_config import obtener_config
    return {"tipos": obtener_config()}


@app.patch("/admin/alert-config/{tipo}/toggle")
def admin_alert_config_toggle(tipo: str,
                               _: None = Depends(_verificar_admin)) -> dict:
    from utils.alert_config import toggle_tipo
    nuevo = toggle_tipo(tipo)
    if nuevo is None:
        raise HTTPException(status_code=404, detail="Tipo de alerta no encontrado")
    return {"ok": True, "activo": nuevo}


@app.patch("/admin/feature-config/{feature}/toggle")
def admin_feature_config_toggle(feature: str,
                                 _: None = Depends(_verificar_admin)) -> dict:
    nuevo = _toggle_feature(feature)
    if nuevo is None:
        raise HTTPException(status_code=404, detail="Módulo no encontrado")
    logger.info("Módulo '%s' → %s", feature, "activo" if nuevo else "pausado")
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
    # Enriquecer usuarios Add-in (user_id negativo) con email, display_name y excel_version
    from utils.device_emails import obtener_info_devices, obtener_info_por_user_id
    devices_addin = {d["user_id"]: d for d in obtener_info_devices() if d.get("user_id")}
    for u in stats["usuarios"]:
        uid = u["user_id"]
        if uid < 0 and uid in devices_addin:
            d = devices_addin[uid]
            if not u["email"]:
                u["email"] = d["email"]
            if not u.get("display_name"):
                u["display_name"] = d.get("display_name")
            if not u["version_excel"] and d.get("excel_version"):
                u["version_excel"] = d["excel_version"]
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
    from utils.alert_config import obtener_config as _obtener_alert_config

    # ── Gráfico de actividad (px para evitar barras uniformes con % en flex) ─
    _CHART_MAX_PX = 110
    max_n = max((d["n"] for d in stats["mensajes_por_dia"]), default=1)
    barras_html = ""
    labels_html = ""
    for d in stats["mensajes_por_dia"]:
        bar_px = max(4, int(d["n"] / max_n * _CHART_MAX_PX))
        dia    = d["dia"][5:]
        barras_html += (
            f'<div class="bar-col">'
            f'<div class="bar-num">{d["n"]}</div>'
            f'<div class="bar-fill" style="height:{bar_px}px" title="{d["n"]} msgs el {d["dia"]}"></div>'
            f'</div>'
        )
        labels_html += f'<div class="bar-lbl">{dia}</div>'

    # ── Helper de usuario: identidad + tipo ─────────────────────────────────
    # Definido aquí para que esté disponible en filas_usuarios Y en los tops
    email_por_uid = {u["user_id"]: u.get("email") or "" for u in stats["usuarios"]}

    def _usuario_celda(uid):
        """Devuelve (identidad_html, tipo_badge) para tablas de usuarios."""
        u_data = next((u for u in stats["usuarios"] if u["user_id"] == uid), {})
        email  = _html_mod.escape(u_data.get("email") or email_por_uid.get(uid, ""))
        nombre = _html_mod.escape(u_data.get("display_name") or "")
        if uid <= 0:  # Add-in: uid negativo (propio) o 0 (legacy anónimo)
            if nombre and email and not email.startswith("("):
                identidad = (
                    f"<span style='font-size:.85rem'>{nombre}</span>"
                    f"<br><small style='color:#aaa'>{email}</small>"
                )
            elif nombre:
                identidad = f"<span style='font-size:.85rem'>{nombre}</span>"
            elif email:
                identidad = f"<span style='font-size:.85rem'>{email}</span>"
            else:
                identidad = "<em style='color:#999;font-size:.85rem'>Add-in anónimo</em>"
            tipo = "<span class='badge-addin'>Add-in</span>"
        else:  # Telegram
            if email:
                identidad = (
                    f"<span style='font-size:.85rem'>{email}</span>"
                    f"<br><small style='color:#aaa'>{uid}</small>"
                )
            else:
                identidad = f"<code>{uid}</code>"
            tipo = "<span class='badge-tg'>📱 TG</span>"
        return identidad, tipo

    # ── Tabla de usuarios ───────────────────────────────────────────────────
    hoy_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    usuarios_hoy = sum(
        1 for u in stats["usuarios"]
        if u["ultima_actividad"] != "—" and u["ultima_actividad"][:10] == hoy_str
    )
    filas_usuarios = ""
    for u in stats["usuarios"]:
        uid  = u["user_id"]
        ident_u, tipo_u = _usuario_celda(uid)
        email = u["email"] or "—"
        ver  = u["version_excel"] or "—"
        modo_label = "🔊 Voz" if u["modo_respuesta"] == "voz" else "💬 Texto"
        priv_badge = (' <span style="color:#e67e22;font-weight:600;font-size:.75rem">🔒 Priv</span>'
                      if u["modo_privado"] else "")
        ts   = u["ultima_actividad"][:16].replace("T", " ") if u["ultima_actividad"] != "—" else "—"
        filas_usuarios += (
            f"<tr>"
            f"  <td>{ident_u}</td>"
            f"  <td class='centro'>{tipo_u}</td>"
            f"  <td class='num'>{u['mensajes_enviados']}</td>"
            f"  <td class='num'>{u['total_mensajes']}</td>"
            f"  <td>{ts}</td>"
            f"  <td>{ver}</td>"
            f"  <td class='centro'><span style='font-size:.8rem'>{modo_label}</span>{priv_badge}</td>"
            f"</tr>"
        )

    # ── Tops de usuarios ────────────────────────────────────────────────────
    top_por_msgs = sorted(stats["usuarios"], key=lambda u: u["mensajes_enviados"], reverse=True)[:5]
    por_usuario_ses = (stats_ses or {}).get("por_usuario", [])
    top_por_ses  = sorted(por_usuario_ses, key=lambda u: u["total_sesiones"], reverse=True)[:5]
    inactivos_mas = sorted(
        [u for u in por_usuario_ses if u["dias_inactivo"] > 30],
        key=lambda u: u["dias_inactivo"], reverse=True,
    )[:5]

    filas_top_msgs = ""
    for i, u in enumerate(top_por_msgs, 1):
        ident, tipo = _usuario_celda(u["user_id"])
        filas_top_msgs += (
            f"<tr><td style='color:#aaa;font-size:.78rem'>{i}</td>"
            f"<td>{ident}</td>"
            f"<td class='centro'>{tipo}</td>"
            f"<td class='num'><strong>{u['mensajes_enviados']}</strong></td></tr>"
        )

    filas_top_ses = ""
    for i, u in enumerate(top_por_ses, 1):
        ident, tipo = _usuario_celda(u["user_id"])
        filas_top_ses += (
            f"<tr><td style='color:#aaa;font-size:.78rem'>{i}</td>"
            f"<td>{ident}</td>"
            f"<td class='centro'>{tipo}</td>"
            f"<td class='num'><strong>{u['total_sesiones']}</strong></td></tr>"
        )

    filas_inactivos = ""
    for u in inactivos_mas:
        dias = int(u["dias_inactivo"])
        ident, tipo = _usuario_celda(u["user_id"])
        filas_inactivos += (
            f"<tr><td>{ident}</td>"
            f"<td class='centro'>{tipo}</td>"
            f"<td class='num' style='color:#e74c3c'>{dias}d</td></tr>"
        )
    if not filas_inactivos:
        filas_inactivos = "<tr><td colspan='3' style='color:#27ae60;text-align:center;padding:12px'>✅ Ninguno</td></tr>"

    # ── Configuración de alertas ─────────────────────────────────────────────
    alert_tipos = _obtener_alert_config()
    filas_alertas = ""
    for a in alert_tipos:
        tipo   = a["tipo"]
        label  = a["label"]
        activo = bool(a["activo"])
        badge  = '<span class="badge-ok">✅ Activo</span>' if activo else '<span class="badge-warn">⏸ Pausado</span>'
        clase_btn = "btn-sm active" if activo else "btn-sm"
        texto_btn = "Pausar" if activo else "Activar"
        filas_alertas += (
            f"<tr>"
            f"  <td><code>{tipo}</code></td>"
            f"  <td style='font-size:.82rem'>{label}</td>"
            f"  <td class='centro'>{badge}</td>"
            f"  <td class='centro'>"
            f"    <button class='{clase_btn}' onclick='toggleAlerta(\"{tipo}\")'>{texto_btn}</button>"
            f"  </td>"
            f"</tr>"
        )
    nota_alerta_id = ""
    if _ALERT_TELEGRAM_ID:
        nota_alerta_id = (
            f"<div class='nota-info'>"
            f"Destinatario: <code>ALERT_TELEGRAM_ID={_ALERT_TELEGRAM_ID}</code> "
            f"(configurado en .env)"
            f"</div>"
        )
    else:
        nota_alerta_id = (
            "<div class='nota-info' style='color:#e74c3c'>"
            "⚠️ <code>ALERT_TELEGRAM_ID</code> no configurado — las alertas no se enviarán."
            "</div>"
        )

    # ── Módulos funcionales ──────────────────────────────────────────────────
    feature_tipos = _obtener_feature_config()
    filas_modulos = ""
    for f in feature_tipos:
        feat   = f["feature"]
        label  = f["label"]
        activo = bool(f["activo"])
        badge     = '<span class="badge-ok">✅ Activo</span>' if activo else '<span class="badge-warn">⏸ Pausado</span>'
        clase_btn = "btn-sm active" if activo else "btn-sm"
        texto_btn = "Pausar" if activo else "Activar"
        filas_modulos += (
            f"<tr>"
            f"  <td><code>{feat}</code></td>"
            f"  <td style='font-size:.82rem'>{label}</td>"
            f"  <td class='centro'>{badge}</td>"
            f"  <td class='centro'>"
            f"    <button class='{clase_btn}' onclick='toggleModulo(\"{feat}\")'>{texto_btn}</button>"
            f"  </td>"
            f"</tr>"
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
            ident_s, tipo_s = _usuario_celda(uid)
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
                f"<td>{ident_s}</td>"
                f"<td class='centro'>{tipo_s}</td>"
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
          <th>Identificador</th><th>Tipo</th><th>Sesiones</th><th>Dur. media</th>
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

    # ── Colores IA para f-strings anidadas ──────────────────────────────────
    _col_ia = "#27ae60" if stats_ia["tasa_exito"] >= 95 else "#e67e22" if stats_ia["tasa_exito"] >= 80 else "#e74c3c"
    _err_badge = f'<span class="log-count-err">{n_errores} errores</span>' if n_errores else ""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Asistente Excel — Admin</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
       background:#f0f2f5;color:#1a1a2e;min-height:100vh}}
  .topbar{{background:#2c3e7a;color:#fff;padding:13px 22px;
           display:flex;align-items:center;justify-content:space-between}}
  .topbar h1{{font-size:1.05rem;font-weight:600}}
  .topbar span{{font-size:.78rem;opacity:.7}}
  .main{{padding:18px;max-width:1240px;margin:0 auto}}

  /* KPI cards */
  .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));
          gap:11px;margin-bottom:16px}}
  .card{{background:#fff;border-radius:10px;padding:15px 10px;
         box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}}
  .card .val{{font-size:1.75rem;font-weight:700;color:#2c3e7a}}
  .card .lbl{{font-size:.7rem;color:#666;margin-top:3px;text-transform:uppercase;
              letter-spacing:.04em}}

  /* Tabs */
  .tab-nav{{display:flex;gap:0;background:#fff;border-radius:10px;
            padding:5px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:16px}}
  .tab-btn{{flex:1;padding:8px 10px;border:none;background:transparent;
            cursor:pointer;border-radius:7px;font-size:.83rem;font-weight:500;
            color:#555;transition:all .15s;white-space:nowrap}}
  .tab-btn:hover{{background:#f0f2f5}}
  .tab-btn.active{{background:#2c3e7a;color:#fff}}
  .tab-panel{{display:none}}
  .tab-panel.active{{display:block}}

  /* Sections */
  .section{{background:#fff;border-radius:10px;
            box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:14px;overflow:hidden}}
  .section-head{{padding:11px 16px;background:#f8f9fb;border-bottom:1px solid #eee;
                 font-weight:600;font-size:.87rem;
                 display:flex;align-items:center;justify-content:space-between}}
  .head-actions{{display:flex;gap:6px}}

  /* Grids */
  .two-col{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
  .three-col{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:14px}}
  @media(max-width:750px){{.two-col,.three-col{{grid-template-columns:1fr}}}}

  /* Metric list */
  .metric-list{{padding:12px 16px;display:flex;flex-direction:column;gap:9px}}
  .metric-row{{display:flex;justify-content:space-between;align-items:center;font-size:.83rem}}
  .metric-row .mk{{color:#666}}
  .metric-row .mv{{font-weight:600}}

  /* Progress */
  .prog-wrap{{margin-top:3px;background:#eee;border-radius:4px;height:7px;overflow:hidden}}
  .prog-fill{{height:100%;border-radius:4px}}
  .prog-label{{font-size:.7rem;color:#666;margin-top:2px;display:flex;justify-content:space-between}}

  /* Chart — pixel heights, bottom-aligned */
  .chart-wrap{{padding:12px 16px 0}}
  .chart-area{{display:flex;align-items:flex-end;height:120px;gap:4px}}
  .bar-col{{flex:1;display:flex;flex-direction:column;align-items:center}}
  .bar-num{{font-size:.6rem;color:#555;margin-bottom:2px;min-height:12px;display:flex;align-items:flex-end}}
  .bar-fill{{background:#4472C4;border-radius:3px 3px 0 0;width:82%;min-height:4px}}
  .chart-labels{{display:flex;gap:4px;padding:3px 16px 12px;border-bottom:1px solid #f0f0f0}}
  .bar-lbl{{flex:1;text-align:center;font-size:.6rem;color:#888}}

  /* Tables */
  table{{width:100%;border-collapse:collapse;font-size:.83rem}}
  th{{padding:8px 13px;text-align:left;font-weight:600;
      background:#f8f9fb;border-bottom:2px solid #eee;color:#444;white-space:nowrap}}
  td{{padding:7px 13px;border-bottom:1px solid #f0f0f0;vertical-align:middle}}
  tr:last-child td{{border-bottom:none}}
  tr:hover td{{background:#fafbff}}
  .num{{text-align:right;font-variant-numeric:tabular-nums}}
  .centro{{text-align:center}}
  code{{background:#f0f0f0;padding:2px 5px;border-radius:4px;font-size:.77rem}}
  .badge-ok{{color:#27ae60;font-weight:600}}
  .badge-warn{{color:#e67e22}}
  .badge-tg{{background:#2196f3;color:#fff;border-radius:4px;padding:1px 6px;font-size:.72rem;white-space:nowrap}}
  .badge-addin{{background:#7b1fa2;color:#fff;border-radius:4px;padding:1px 6px;font-size:.72rem;white-space:nowrap}}

  /* Buttons */
  .btn-del{{background:#e74c3c;color:#fff;border:none;border-radius:4px;
            padding:2px 7px;cursor:pointer;font-size:.77rem}}
  .btn-del:hover{{background:#c0392b}}
  .btn-sm{{padding:4px 10px;border:1px solid #ddd;border-radius:6px;
           background:#fff;cursor:pointer;font-size:.77rem;color:#444}}
  .btn-sm.active{{background:#2c3e7a;color:#fff;border-color:#2c3e7a}}
  .btn-sm:hover:not(.active){{background:#f0f2f5}}

  /* Forms */
  .form-row{{display:flex;gap:7px;padding:11px 16px;border-top:1px solid #eee;flex-wrap:wrap}}
  .form-row input{{flex:1;min-width:120px;padding:5px 9px;
                   border:1px solid #ddd;border-radius:6px;font-size:.83rem}}
  .form-row button{{padding:5px 13px;background:#2c3e7a;color:#fff;
                    border:none;border-radius:6px;cursor:pointer;font-size:.83rem}}
  .form-row button:hover{{background:#1a2a5e}}

  /* Logs */
  .log-box{{height:280px;overflow-y:auto;font-family:'Courier New',monospace;
            font-size:.72rem;line-height:1.5;padding:10px 13px;background:#fafafa}}
  .log-error{{color:#c0392b;background:#fdecea;padding:1px 3px;border-radius:2px;margin:1px 0}}
  .log-warn{{color:#8b6914;background:#fef9e7;padding:1px 3px;border-radius:2px;margin:1px 0}}
  .log-info{{color:#2c3e7a}}
  .log-debug{{color:#bbb}}
  .log-box.only-errors .log-info,
  .log-box.only-errors .log-debug{{display:none}}
  .log-count-err{{background:#e74c3c;color:#fff;border-radius:10px;
                  padding:1px 6px;font-size:.7rem;font-weight:700;margin-left:5px}}
  .nota-info{{padding:10px 16px;font-size:.78rem;color:#888;background:#fffbf0;
              border-top:1px solid #f0e6c8}}
</style>
</head>
<body>
<div class="topbar">
  <h1>🤖 Asistente Excel — Panel de administración</h1>
  <span>Actualizado: {ahora}</span>
</div>
<div class="main">

  <!-- KPI cards (siempre visibles) -->
  <div class="cards">
    <div class="card"><div class="val">{stats['total_usuarios']}</div><div class="lbl">Usuarios</div></div>
    <div class="card"><div class="val">{stats['total_mensajes']:,}</div><div class="lbl">Msgs totales</div></div>
    <div class="card"><div class="val">{stats['mensajes_hoy']}</div><div class="lbl">Msgs hoy</div></div>
    <div class="card"><div class="val">{usuarios_hoy}</div><div class="lbl">Activos hoy</div></div>
    <div class="card"><div class="val">{len(vinculos)}</div><div class="lbl">Vínculos</div></div>
    <div class="card"><div class="val">{uptime_str}</div><div class="lbl">Uptime</div></div>
  </div>

  <!-- Navegación por pestañas -->
  <div class="tab-nav">
    <button class="tab-btn active" onclick="mostrarTab('resumen',this)">📊 Resumen</button>
    <button class="tab-btn" onclick="mostrarTab('usuarios',this)">👤 Usuarios</button>
    <button class="tab-btn" onclick="mostrarTab('ia',this)">🧠 IA</button>
    <button class="tab-btn" onclick="mostrarTab('sistema',this)">⚙️ Sistema</button>
  </div>

  <!-- ═══ Tab: RESUMEN ════════════════════════════════════════════════════════ -->
  <div id="tab-resumen" class="tab-panel active">

    <div class="section">
      <div class="section-head">📊 Actividad — últimos 7 días</div>
      {'<div class="chart-wrap"><div class="chart-area">' + barras_html + '</div><div class="chart-labels">' + labels_html + '</div></div>'
        if barras_html else
       '<div style="padding:20px;color:#999;text-align:center">Sin datos de actividad aún</div>'}
    </div>

    <div class="two-col">
      <div class="section">
        <div class="section-head">💻 Sistema <small style="font-weight:400;color:#888">(Render free)</small></div>
        <div class="metric-list">
          <div>
            <div class="metric-row"><span class="mk">RAM usada</span><span class="mv">{ram_lbl}</span></div>
            <div class="prog-wrap">{ram_bar}</div>
            <div class="prog-label"><span>0</span><span>límite {_RENDER_RAM_MB} MB</span></div>
          </div>
          <div class="metric-row"><span class="mk">CPU</span><span class="mv">{cpu_str}</span></div>
          <div class="metric-row"><span class="mk">Python</span><span class="mv">{sistema['python_version']}</span></div>
          <div class="metric-row"><span class="mk">data/ (BD+logs)</span><span class="mv">{sistema['data_mb']} MB</span></div>
          <div class="metric-row"><span class="mk">└ logs/</span><span class="mv">{sistema['logs_mb']} MB</span></div>
          <div class="metric-row"><span class="mk">└ temp/</span><span class="mv">{sistema['temp_mb']} MB</span></div>
        </div>
      </div>
      <div class="section">
        <div class="section-head">🤖 Bot e IA</div>
        <div class="metric-list">
          <div class="metric-row"><span class="mk">Telegram</span><span class="mv">{tg_status}</span></div>
          <div class="metric-row"><span class="mk">Modo</span><span class="mv">{webhook}</span></div>
          <div class="metric-row"><span class="mk">Proveedor IA</span><span class="mv">{proveedor}</span></div>
          <div class="metric-row"><span class="mk">Modelo</span><span class="mv" style="font-size:.78rem">{modelo}</span></div>
          <div class="metric-row"><span class="mk">Base de datos</span><span class="mv">{db_modo_label}</span></div>
          <div class="metric-row"><span class="mk">Estado BD</span><span class="mv">{db_estado_badge}</span></div>
          <div class="metric-row"><span class="mk">URL BD</span>
            <span class="mv"><code style="font-size:.68rem;word-break:break-all">{db_inf['url'].split('?')[0] + ('?authToken=***' if '?' in db_inf['url'] else '')}</code></span>
          </div>
        </div>
      </div>
    </div>

  </div>

  <!-- ═══ Tab: USUARIOS ══════════════════════════════════════════════════════ -->
  <div id="tab-usuarios" class="tab-panel">

    <div class="three-col">
      <div class="section">
        <div class="section-head">🏆 Más activos (msgs)</div>
        <table>
          <thead><tr><th>#</th><th>Identificador</th><th>Tipo</th><th>Msgs</th></tr></thead>
          <tbody>{filas_top_msgs or "<tr><td colspan='4' style='color:#999;text-align:center;padding:12px'>—</td></tr>"}</tbody>
        </table>
      </div>
      <div class="section">
        <div class="section-head">🔁 Más sesiones</div>
        <table>
          <thead><tr><th>#</th><th>Identificador</th><th>Tipo</th><th>Ses.</th></tr></thead>
          <tbody>{filas_top_ses or "<tr><td colspan='4' style='color:#999;text-align:center;padding:12px'>—</td></tr>"}</tbody>
        </table>
      </div>
      <div class="section">
        <div class="section-head">😴 Inactivos &gt;30 días</div>
        <table>
          <thead><tr><th>Identificador</th><th>Tipo</th><th>Días</th></tr></thead>
          <tbody>{filas_inactivos}</tbody>
        </table>
      </div>
    </div>

    <div class="section">
      <div class="section-head" style="display:flex;align-items:center;justify-content:space-between;">
        <span>👤 Todos los usuarios ({stats['total_usuarios']})</span>
        <a onclick="const k=new URLSearchParams(location.search).get('key');location.href='/admin/stats.csv?key='+encodeURIComponent(k)"
           style="font-size:.75rem;padding:4px 10px;background:#217346;color:#fff;border-radius:3px;text-decoration:none;cursor:pointer;">
          ⬇ CSV
        </a>
      </div>
      <table>
        <thead>
          <tr>
            <th>Identificador</th><th>Tipo</th>
            <th>Msgs env.</th><th>Total</th>
            <th>Última actividad</th><th>Ver. Excel</th><th>Respuesta</th>
          </tr>
        </thead>
        <tbody>
          {filas_usuarios or '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px">Sin usuarios</td></tr>'}
        </tbody>
      </table>
    </div>

    <div class="section">
      <div class="section-head">📈 Análisis de sesiones</div>
      {html_sesiones}
    </div>

  </div>

  <!-- ═══ Tab: IA ════════════════════════════════════════════════════════════ -->
  <div id="tab-ia" class="tab-panel">
    <div class="section">
      <div class="section-head">🧠 Inteligencia Artificial</div>
      {ia_nota}
      {'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;padding:14px 16px 8px">'
        f'<div class="card" style="padding:11px 8px"><div class="val" style="font-size:1.4rem">{stats_ia["total"]:,}</div><div class="lbl">Llamadas totales</div></div>'
        f'<div class="card" style="padding:11px 8px"><div class="val" style="font-size:1.4rem;color:{_col_ia}">{stats_ia["tasa_exito"]}%</div><div class="lbl">Tasa de éxito</div></div>'
        f'<div class="card" style="padding:11px 8px"><div class="val" style="font-size:1.4rem;color:#e67e22">{stats_ia["fallbacks"]}</div><div class="lbl">Fallbacks</div></div>'
        f'<div class="card" style="padding:11px 8px"><div class="val" style="font-size:1.4rem;color:#e74c3c">{stats_ia["errores"]}</div><div class="lbl">Errores</div></div>'
        f'<div class="card" style="padding:11px 8px"><div class="val" style="font-size:1.4rem">{stats_ia["total_hoy"]}</div><div class="lbl">Hoy</div></div>'
        '</div>'
        '<div style="display:grid;grid-template-columns:1fr 1fr;border-top:1px solid #eee">'
        '<div style="padding:12px 16px;border-right:1px solid #eee">'
        '<div style="font-weight:600;font-size:.83rem;margin-bottom:7px">Por proveedor</div>'
        '<table><thead><tr><th>Proveedor</th><th>Total</th><th>OK</th><th>Err</th><th>Fallb.</th><th>%OK</th></tr></thead>'
        f'<tbody>{filas_proveedores or "<tr><td colspan=5 style=color:#999;text-align:center>—</td></tr>"}</tbody></table>'
        '</div><div style="padding:12px 16px">'
        '<div style="font-weight:600;font-size:.83rem;margin-bottom:7px">Tipos de error</div>'
        f'<table><thead><tr><th>Tipo</th><th>Veces</th></tr></thead><tbody>{filas_errores_ia}</tbody></table>'
        '</div></div>'
        if not ia_sin_datos else ""}
    </div>
  </div>

  <!-- ═══ Tab: SISTEMA ═══════════════════════════════════════════════════════ -->
  <div id="tab-sistema" class="tab-panel">

    <div class="section">
      <div class="section-head">🧩 Módulos funcionales</div>
      <table>
        <thead><tr><th>Módulo</th><th>Descripción</th><th>Estado</th><th></th></tr></thead>
        <tbody>{filas_modulos}</tbody>
      </table>
    </div>

    <div class="section">
      <div class="section-head">🔔 Notificaciones del sistema</div>
      <table>
        <thead><tr><th>Tipo</th><th>Descripción</th><th>Estado</th><th></th></tr></thead>
        <tbody>{filas_alertas}</tbody>
      </table>
      {nota_alerta_id}
    </div>

    <div class="section">
      <div class="section-head">🔗 Vínculos Telegram ↔ Add-in ({len(vinculos)})</div>
      <table>
        <thead><tr><th>Telegram ID</th><th>Email</th><th>Vinculado el</th><th></th></tr></thead>
        <tbody>
          {filas_vinculos or '<tr><td colspan="4" style="text-align:center;color:#999;padding:16px">Sin vínculos</td></tr>'}
        </tbody>
      </table>
      <form class="form-row" onsubmit="agregarVinculo(event)">
        <input type="number" id="inp-tid"   placeholder="Telegram ID" required>
        <input type="email"  id="inp-email" placeholder="email@empresa.com" required>
        <button type="submit">+ Añadir vínculo</button>
      </form>
    </div>

    <div class="section">
      <div class="section-head">📋 Logs recientes {_err_badge}</div>
      <div style="display:flex;gap:6px;padding:9px 13px;border-bottom:1px solid #eee;align-items:center">
        <button class="btn-sm active" id="btn-todos"   onclick="filtrarLogs('todos')">Todos</button>
        <button class="btn-sm"        id="btn-errores" onclick="filtrarLogs('errores')">Solo errores</button>
        <button class="btn-sm" onclick="location.reload()" style="margin-left:auto">↻ Refrescar</button>
      </div>
      <div class="log-box" id="log-box">{lineas_log}</div>
    </div>

  </div>

</div>
<script>
  const _key = new URLSearchParams(window.location.search).get("key") || "";

  document.getElementById("log-box").scrollTop = 9999;

  function mostrarTab(id, btn) {{
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('tab-' + id).classList.add('active');
    btn.classList.add('active');
  }}

  function filtrarLogs(modo) {{
    const box = document.getElementById("log-box");
    document.getElementById("btn-todos").classList.toggle("active", modo === "todos");
    document.getElementById("btn-errores").classList.toggle("active", modo === "errores");
    box.classList.toggle("only-errors", modo === "errores");
  }}

  async function toggleAlerta(tipo) {{
    const resp = await fetch(
      "/admin/alert-config/" + encodeURIComponent(tipo) + "/toggle?key=" + encodeURIComponent(_key),
      {{ method: "PATCH" }}
    );
    if (resp.ok) location.reload();
    else alert("Error al cambiar estado de la alerta");
  }}

  async function toggleModulo(feature) {{
    const resp = await fetch(
      "/admin/feature-config/" + encodeURIComponent(feature) + "/toggle?key=" + encodeURIComponent(_key),
      {{ method: "PATCH" }}
    );
    if (resp.ok) location.reload();
    else alert("Error al cambiar estado del módulo");
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
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Ficheros estáticos del Add-in (solo en cloud — dist/ generado por webpack)
# ---------------------------------------------------------------------------

_DIST = os.path.join(os.path.dirname(__file__), "excel-addin", "dist")
if os.path.isdir(_DIST):
    # no-store en local para que WebView2 no cachee JS/CSS entre recargas
    if not os.getenv("WEBHOOK_URL"):
        @app.middleware("http")
        async def _no_cache_static(request: Request, call_next):
            response = await call_next(request)
            if request.url.path.endswith((".js", ".css", ".html")):
                response.headers["Cache-Control"] = "no-store"
            return response
    app.mount("/", StaticFiles(directory=_DIST, html=True), name="addin")
    logger.info("Add-in estático servido desde %s", _DIST)


# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
