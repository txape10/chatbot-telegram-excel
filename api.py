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
from services.llm import extraer_operacion_edicion, extraer_query_dsl, obtener_respuesta

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

    yield   # ← la app está corriendo

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


class PeticionAnalisis(BaseModel):
    datos: list[list]


class PeticionEdicion(BaseModel):
    datos: list[list]
    instruccion: str


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


def _verificar_addin_activo():
    if not _ENABLE_ADDIN:
        raise HTTPException(status_code=503, detail="Módulo Add-in desactivado (ENABLE_ADDIN=false)")


@app.post("/ask")
def ask(peticion: PeticionPregunta, _: None = Depends(_verificar_clave),
        __: None = Depends(_verificar_addin_activo)) -> dict:
    # Sin datos: pregunta general o creación desde cero
    if not peticion.datos or len(peticion.datos) < 2:
        return {"respuesta": obtener_respuesta([], peticion.pregunta)}

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
    return {"respuesta": obtener_respuesta([], contexto)}


@app.post("/edit")
def edit(peticion: PeticionEdicion, _: None = Depends(_verificar_clave),
         __: None = Depends(_verificar_addin_activo)) -> dict:
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
            logger.warning("Editor falló, usando LLM libre: %s", error)

    columnas = ", ".join(str(c) for c in df.columns)
    muestra  = df.head(5).to_string(index=False)
    contexto = (
        f"El usuario tiene una tabla con columnas: {columnas}.\n"
        f"Primeras filas:\n{muestra}\n\n"
        f"Pregunta: {peticion.instruccion}"
    )
    return {"tipo": "texto", "respuesta": obtener_respuesta([], contexto)}


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


@app.get("/admin/test-db")
def test_db(_: None = Depends(_verificar_admin)) -> dict:
    """Diagnóstico: escribe una fila de prueba y la lee de vuelta."""
    import traceback
    from utils.db import conectar
    resultado: dict = {}

    # 1. Escribir
    try:
        with conectar() as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS _test_write (ts TEXT, val TEXT)"
            )
            conn.execute(
                "INSERT INTO _test_write (ts, val) VALUES (datetime('now'), 'ping')"
            )
        resultado["escritura"] = "ok"
    except Exception:
        resultado["escritura"] = traceback.format_exc()

    # 2. Leer (conexión nueva)
    try:
        with conectar() as conn:
            filas = conn.execute(
                "SELECT ts, val FROM _test_write ORDER BY rowid DESC LIMIT 3"
            ).fetchall()
        resultado["lectura"] = [{"ts": f[0], "val": f[1]} for f in filas]
    except Exception:
        resultado["lectura"] = traceback.format_exc()

    # 3. Contar filas en historial
    try:
        with conectar() as conn:
            n = conn.execute("SELECT COUNT(*) FROM historial").fetchone()
            resultado["historial_count"] = n[0] if n else "tabla no existe"
    except Exception:
        resultado["historial_count"] = traceback.format_exc()

    return resultado


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


@app.get("/admin", response_class=HTMLResponse)
def admin_panel(_: None = Depends(_verificar_admin)):
    """Panel de administración con estadísticas de uso."""
    from utils.stats import obtener_estadisticas
    stats = obtener_estadisticas()
    return _renderizar_admin_html(stats)


def _renderizar_admin_html(stats: dict) -> str:
    from datetime import datetime
    from utils.user_links import obtener_todos_los_vinculos
    from utils.db import estado as _estado_db

    # Barra de mensajes por día (últimos 7 días)
    max_n = max((d["n"] for d in stats["mensajes_por_dia"]), default=1)

    barras_html = ""
    for d in stats["mensajes_por_dia"]:
        pct  = int(d["n"] / max_n * 100)
        dia  = d["dia"][5:]   # MM-DD
        barras_html += (
            f'<div class="bar-item">'
            f'  <div class="bar-fill" style="height:{pct}%" title="{d["n"]} mensajes"></div>'
            f'  <div class="bar-label">{dia}</div>'
            f'  <div class="bar-val">{d["n"]}</div>'
            f'</div>'
        )

    # Filas de usuarios
    filas_usuarios = ""
    for u in stats["usuarios"]:
        uid   = u["user_id"]
        email = u["email"] or "—"
        ver   = u["version_excel"] or "—"
        modo  = "🔊" if u["modo_respuesta"] == "voz" else "💬"
        priv  = "🔒" if u["modo_privado"] else ""
        ts    = u["ultima_actividad"][:16].replace("T", " ") if u["ultima_actividad"] != "—" else "—"
        filas_usuarios += (
            f"<tr>"
            f"  <td><code>{uid}</code></td>"
            f"  <td>{email}</td>"
            f"  <td class='num'>{u['mensajes_enviados']}</td>"
            f"  <td class='num'>{u['total_mensajes']}</td>"
            f"  <td>{ts}</td>"
            f"  <td>{ver}</td>"
            f"  <td class='centro'>{modo} {priv}</td>"
            f"</tr>"
        )

    # Filas de vínculos Telegram ↔ email
    vinculos = obtener_todos_los_vinculos()
    filas_vinculos = ""
    for v in vinculos:
        tid   = v["telegram_id"]
        mail  = v["email"]
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

    ahora  = datetime.now().strftime("%d/%m/%Y %H:%M")
    db_inf = _estado_db()
    db_modo_label = "☁️ Turso (cloud)" if db_inf["modo"] == "turso" else "💾 SQLite local"
    db_estado_badge = (
        '<span class="badge-ok">✅ Conectada</span>'
        if db_inf["ok"] else
        f'<span class="badge-warn">❌ Error: {db_inf["error"]}</span>'
    )

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
  .main{{padding:24px;max-width:1100px;margin:0 auto}}
  .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
          gap:16px;margin-bottom:28px}}
  .card{{background:#fff;border-radius:10px;padding:20px;
         box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}}
  .card .val{{font-size:2rem;font-weight:700;color:#2c3e7a}}
  .card .lbl{{font-size:.78rem;color:#666;margin-top:4px}}
  .section{{background:#fff;border-radius:10px;
            box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:24px;overflow:hidden}}
  .section-head{{padding:14px 20px;background:#f8f9fb;
                 border-bottom:1px solid #eee;font-weight:600;font-size:.9rem}}
  .chart{{display:flex;align-items:flex-end;gap:8px;
          padding:20px 20px 10px;height:140px}}
  .bar-item{{display:flex;flex-direction:column;align-items:center;
             flex:1;height:100%}}
  .bar-fill{{background:#2c3e7a;border-radius:3px 3px 0 0;width:100%;
             min-height:4px;transition:height .3s}}
  .bar-label,.bar-val{{font-size:.68rem;color:#666;margin-top:3px}}
  .bar-val{{color:#2c3e7a;font-weight:600}}
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
  .btn-del{{background:#e74c3c;color:#fff;border:none;border-radius:4px;
            padding:3px 9px;cursor:pointer;font-size:.8rem}}
  .btn-del:hover{{background:#c0392b}}
  .form-vincular{{display:flex;gap:8px;padding:14px 20px;
                  border-top:1px solid #eee;flex-wrap:wrap}}
  .form-vincular input{{flex:1;min-width:140px;padding:6px 10px;
                        border:1px solid #ddd;border-radius:6px;font-size:.85rem}}
  .form-vincular button{{padding:6px 16px;background:#2c3e7a;color:#fff;
                         border:none;border-radius:6px;cursor:pointer;font-size:.85rem}}
  .form-vincular button:hover{{background:#1a2a5e}}
</style>
</head>
<body>
<div class="topbar">
  <h1>🤖 Asistente Excel — Panel de administración</h1>
  <span>Actualizado: {ahora}</span>
</div>
<div class="main">

  <!-- Tarjetas resumen -->
  <div class="cards">
    <div class="card">
      <div class="val">{stats['total_usuarios']}</div>
      <div class="lbl">Usuarios activos</div>
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
      <div class="val">{sum(1 for u in stats['usuarios'] if u['email'])}</div>
      <div class="lbl">Add-in vinculados</div>
    </div>
  </div>

  <!-- Gráfico de actividad -->
  <div class="section">
    <div class="section-head">📊 Actividad — últimos 7 días</div>
    <div class="chart">
      {barras_html if barras_html else '<p style="padding:20px;color:#999">Sin datos</p>'}
    </div>
  </div>

  <!-- Tabla de usuarios -->
  <div class="section">
    <div class="section-head">👤 Usuarios ({stats['total_usuarios']})</div>
    <table>
      <thead>
        <tr>
          <th>Telegram ID</th>
          <th>Email (Add-in)</th>
          <th>Msgs enviados</th>
          <th>Total msgs</th>
          <th>Última actividad</th>
          <th>Versión Excel</th>
          <th>Modo</th>
        </tr>
      </thead>
      <tbody>
        {filas_usuarios if filas_usuarios else
         '<tr><td colspan="7" style="text-align:center;color:#999;padding:20px">Sin usuarios</td></tr>'}
      </tbody>
    </table>
  </div>

  <!-- Vinculaciones Telegram ↔ Add-in -->
  <div class="section">
    <div class="section-head">🔗 Vinculaciones Telegram ↔ Add-in ({len(vinculos)})</div>
    <table>
      <thead>
        <tr>
          <th>Telegram ID</th>
          <th>Email</th>
          <th>Vinculado el</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        {filas_vinculos if filas_vinculos else
         '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px">Sin vínculos</td></tr>'}
      </tbody>
    </table>
    <form class="form-vincular" onsubmit="agregarVinculo(event)">
      <input type="number" id="inp-tid"   placeholder="Telegram ID" required>
      <input type="email"  id="inp-email" placeholder="email@empresa.com" required>
      <button type="submit">+ Añadir vínculo</button>
    </form>
  </div>

  <!-- Estado de la base de datos -->
  <div class="section">
    <div class="section-head">🗄 Base de datos</div>
    <table>
      <thead>
        <tr><th>Modo</th><th>URL / Ruta</th><th>Estado</th></tr>
      </thead>
      <tbody>
        <tr>
          <td>{db_modo_label}</td>
          <td><code style="font-size:.75rem;word-break:break-all">{db_inf["url"]}</code></td>
          <td>{db_estado_badge}</td>
        </tr>
      </tbody>
    </table>
  </div>

</div>
<script>
  const _key = new URLSearchParams(window.location.search).get("key") || "";

  async function eliminarVinculo(tid, email) {{
    if (!confirm("¿Eliminar el vínculo de " + email + "?")) return;
    const resp = await fetch(
      "/admin/vinculos?key=" + encodeURIComponent(_key) +
      "&telegram_id=" + tid +
      "&email=" + encodeURIComponent(email),
      {{ method: "DELETE" }}
    );
    if (resp.ok) location.reload();
    else alert("Error al eliminar el vínculo");
  }}

  async function agregarVinculo(e) {{
    e.preventDefault();
    const tid   = document.getElementById("inp-tid").value.trim();
    const email = document.getElementById("inp-email").value.trim();
    const resp  = await fetch(
      "/admin/vinculos?key=" + encodeURIComponent(_key),
      {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify({{ telegram_id: parseInt(tid), email }})
      }}
    );
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
    app.mount("/", StaticFiles(directory=_DIST, html=True), name="addin")
    logger.info("Add-in estático servido desde %s", _DIST)


# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
