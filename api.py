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
from fastapi import Depends, FastAPI, Header, HTTPException, Request
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
        await _ptb_app.bot.set_webhook(
            url=f"{_WEBHOOK_URL}/telegram/webhook",
            allowed_updates=Update.ALL_TYPES,
        )
        await _ptb_app.start()
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
    datos: list[list]
    pregunta: str


class PeticionAnalisis(BaseModel):
    datos: list[list]


class PeticionEdicion(BaseModel):
    datos: list[list]
    instruccion: str


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
