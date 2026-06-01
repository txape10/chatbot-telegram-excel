import base64
import json
import logging
import re
import time
import pandas as pd

from config import SYSTEM_PROMPT
from services.llm_provider import (  # noqa: F401 — re-exportados
    LLMError, LLMProvider, obtener_proveedor, obtener_proveedor_privado,
)

__all__ = ["LLMError", "obtener_proveedor_privado"]

logger = logging.getLogger(__name__)


def _estimar_tokens(texto: str) -> int:
    """Estimación rápida: ~4 caracteres por token."""
    return len(texto) // 4


def _chat_timed(nombre: str, messages: list, temperature: float = 0,
                max_tokens: int | None = None) -> str:
    """Llama al LLM activo y logea el tiempo empleado."""
    t0 = time.perf_counter()
    resultado = obtener_proveedor().chat(messages=messages, temperature=temperature,
                                         max_tokens=max_tokens)
    ms = (time.perf_counter() - t0) * 1000
    logger.debug("LLM %-25s → %6.0f ms", nombre, ms)
    return resultado


def _construir_mensajes(historial: list[dict], pregunta: str,
                         proveedor: LLMProvider | None = None,
                         system_override: str | None = None,
                         user_id: int | None = None) -> list[dict]:
    """Construye la lista de mensajes ajustando el historial si la petición
    supera el presupuesto de tokens del proveedor activo."""
    proveedor = proveedor or obtener_proveedor()
    prompt_sistema = system_override or SYSTEM_PROMPT

    # Inyectar few-shot del usuario si tiene ejemplos guardados (RAG)
    if user_id:
        try:
            from utils.rag import obtener_ejemplos
            ejemplos = obtener_ejemplos(user_id, limite=3)
            if ejemplos:
                partes = "\n\n".join(
                    f"Pregunta: {e['pregunta']}\nRespuesta: {e['respuesta']}"
                    for e in ejemplos
                )
                prompt_sistema = (
                    prompt_sistema
                    + "\n\nEjemplos de respuestas que este usuario valoró positivamente:\n\n"
                    + partes
                )
        except Exception:
            pass

    mensajes_sistema = [{"role": "system", "content": prompt_sistema}]
    tokens_fijos = _estimar_tokens(prompt_sistema) + _estimar_tokens(pregunta)
    presupuesto_historial = proveedor.max_tokens_peticion - tokens_fijos

    historial_valido = list(historial)
    # Calcular coste de cada par (usuario+modelo) una sola vez → O(n) en lugar de O(n²)
    tokens_por_par = [
        _estimar_tokens(str(historial_valido[i:i + 2]))
        for i in range(0, len(historial_valido), 2)
    ]
    tokens_hist = sum(tokens_por_par)
    pares_eliminados = 0
    while tokens_hist > presupuesto_historial and pares_eliminados < len(tokens_por_par):
        tokens_hist -= tokens_por_par[pares_eliminados]
        pares_eliminados += 1
    if pares_eliminados:
        historial_valido = historial_valido[pares_eliminados * 2:]
        logger.debug("Historial recortado a %d mensajes por límite de tokens", len(historial_valido))

    mensajes_historial = []
    for mensaje in historial_valido:
        rol = "assistant" if mensaje["role"] == "model" else "user"
        mensajes_historial.append({"role": rol, "content": mensaje["parts"][0]})

    return mensajes_sistema + mensajes_historial + [{"role": "user", "content": pregunta}]


def obtener_respuesta(historial: list[dict], pregunta: str,
                       proveedor: LLMProvider | None = None,
                       system_override: str | None = None,
                       user_id: int | None = None) -> str:
    """Envía la pregunta al LLM y devuelve la respuesta.

    Args:
        proveedor: si se indica, usa ese proveedor en lugar del activo.
        system_override: reemplaza el system prompt por defecto.
        user_id: si se indica, inyecta few-shot examples del historial del usuario.
    """
    proveedor_activo = proveedor or obtener_proveedor()
    mensajes = _construir_mensajes(historial, pregunta, proveedor_activo, system_override, user_id)
    tokens_estimados = _estimar_tokens(str(mensajes))
    logger.debug("Petición LLM — tokens estimados: %d", tokens_estimados)
    return proveedor_activo.chat(mensajes)


def _limpiar_json(texto: str) -> str:
    """Elimina bloques de código markdown que el LLM añade por error."""
    if "```" in texto:
        lineas = [l for l in texto.splitlines() if not l.startswith("```")]
        texto = "\n".join(lineas).strip()
    return texto


def extraer_estructura_excel(pregunta: str) -> dict | None:
    """Interpreta la descripción del usuario y devuelve la estructura JSON para crear el xlsx."""
    from prompts.excel import CREAR_EXCEL_SISTEMA, CREAR_EXCEL_USUARIO
    try:
        texto = obtener_proveedor().chat(
            messages=[
                {"role": "system", "content": CREAR_EXCEL_SISTEMA},
                {"role": "user",   "content": CREAR_EXCEL_USUARIO.format(pregunta=pregunta)},
            ],
            temperature=0,
            max_tokens=3000,
        )
        logger.debug("Estructura Excel del LLM: %s", texto)
        texto = texto.strip()
        if texto == "RESPUESTA_LIBRE":
            return None
        return json.loads(_limpiar_json(texto))
    except Exception as error:
        logger.warning("Error extrayendo estructura Excel: %s", error)
        return None


# El usuario pidió explícitamente ordenar (edición real, no solo "agrupar y mostrar")
_RE_ORDENAR_EXPLICITO = re.compile(
    r"\b(?:ordena|reordena|sort|clasifica)\b",
    re.IGNORECASE,
)


def _limpiar_ops_espurias(ops: list[dict], pregunta: str) -> list[dict]:
    """Elimina ops de edición que el LLM añadió como 'preparación' cuando la intención
    real del usuario era solo una consulta de lectura.

    Caso habitual: "Agrúpame por X y dame el top 3" → el LLM devuelve
    [{"op":"ordenar",...}, {"op":"query",...}]. El ordenar es espurio — el usuario
    nunca pidió modificar el archivo, solo ver el resultado agrupado.
    """
    tiene_query = any(op.get("op") == "query" for op in ops)
    if not tiene_query:
        return ops  # Sin query → no hay ambigüedad, dejar pasar

    # Si hay query Y ops de edición que el LLM suele mezclar por confusión, filtrar los espurios
    _OPS_CONFUNDIBLES = {"ordenar", "pivotear"}
    ordenar_explicito = bool(_RE_ORDENAR_EXPLICITO.search(pregunta))

    limpias = [
        op for op in ops
        if op.get("op") not in _OPS_CONFUNDIBLES or ordenar_explicito
    ]
    # Asegurar que queda al menos la query
    return limpias if limpias else ops


def extraer_operacion_edicion(
    df: pd.DataFrame,
    pregunta: str,
    macros_disponibles: list[str] | None = None,
) -> list[dict] | dict | None:
    """Llama al LLM con el prompt de edición y parsea el JSON.

    Devuelve:
      - list[dict]  → pipeline de operaciones DSL a ejecutar en orden
      - dict con aclaracion_necesaria=True → pedir aclaración al usuario
      - None        → RESPUESTA_LIBRE o error de parseo
    """
    from prompts.excel import EDITOR_DSL_SISTEMA, EDITOR_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {t}" for c, t in df.dtypes.items())
    muestra  = df.head(3).to_string(index=False)

    if macros_disponibles:
        macros_info = "Macros guardadas del usuario: " + ", ".join(macros_disponibles) + "\n"
    else:
        macros_info = ""

    texto = None
    try:
        texto = _chat_timed("editor_dsl", [
            {"role": "system", "content": EDITOR_DSL_SISTEMA},
            {"role": "user",   "content": EDITOR_DSL_USUARIO.format(
                columnas=columnas, tipos=tipos, muestra=muestra,
                macros_info=macros_info, pregunta=pregunta,
            )},
        ], temperature=0, max_tokens=600)
        logger.debug("Respuesta editor DSL del LLM: %s", texto)
        texto = texto.strip()
        if texto == "RESPUESTA_LIBRE":
            return None
        parsed = json.loads(_limpiar_json(texto))
        # Aclaración → dict plano (no es un pipeline)
        if isinstance(parsed, dict) and parsed.get("aclaracion_necesaria"):
            return parsed
        # Pipeline → siempre lista; normalizar dict suelto por si el LLM no siguió el formato
        if isinstance(parsed, dict):
            parsed = [parsed]
        if isinstance(parsed, list):
            return _limpiar_ops_espurias(parsed, pregunta)
        return None
    except Exception as error:
        logger.warning("Error extrayendo operación de edición: %s | respuesta LLM: %r", error, texto)
        return None


def extraer_query_dsl(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Llama al LLM con el prompt DSL y parsea el JSON resultante.

    Devuelve:
      - dict con la query DSL          → ejecutar consulta
      - {"aclaracion_necesaria": True, "pregunta": ..., "opciones": [...]}  → pedir aclaración
      - None                           → RESPUESTA_LIBRE o error de parseo
    """
    from prompts.excel import QUERY_DSL_SISTEMA, QUERY_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {t}" for c, t in df.dtypes.items())
    muestra  = df.head(3).to_string(index=False)

    texto = None
    try:
        texto = _chat_timed("query_dsl", [
            {"role": "system", "content": QUERY_DSL_SISTEMA},
            {"role": "user",   "content": QUERY_DSL_USUARIO.format(
                columnas=columnas, tipos=tipos, muestra=muestra, pregunta=pregunta,
            )},
        ], temperature=0, max_tokens=400)
        logger.debug("Respuesta DSL del LLM: %s", texto)
        texto = texto.strip()
        if texto == "RESPUESTA_LIBRE":
            return None
        parsed = json.loads(_limpiar_json(texto))
        # Propagar aclaración tal cual — el handler decide qué hacer
        return parsed
    except Exception as error:
        logger.warning("Error extrayendo query DSL: %s | respuesta LLM: %r", error, texto)
        return None


def extraer_operacion_combinar(df1: pd.DataFrame, df2: pd.DataFrame,
                               pregunta: str) -> dict:
    """Extrae la columna clave y tipo de join para combinar dos DataFrames."""
    from prompts.excel import COMBINAR_DSL_SISTEMA, COMBINAR_DSL_USUARIO

    cols_a       = ", ".join(f"'{c}'" for c in df1.columns)
    cols_b       = ", ".join(f"'{c}'" for c in df2.columns)
    cols_comunes = ", ".join(f"'{c}'" for c in df1.columns if c in df2.columns) or "ninguna"

    texto = None
    try:
        texto = _chat_timed("combinar_dsl", [
            {"role": "system", "content": COMBINAR_DSL_SISTEMA},
            {"role": "user",   "content": COMBINAR_DSL_USUARIO.format(
                cols_a=cols_a, cols_b=cols_b, cols_comunes=cols_comunes, pregunta=pregunta,
            )},
        ], temperature=0, max_tokens=100)
        logger.debug("Respuesta combinar DSL del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo operación combinar: %s | respuesta LLM: %r", error, texto)
        return {"col": None, "como": "inner"}


def extraer_peticion_grafico(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Extrae los parámetros del gráfico pedido en lenguaje natural."""
    from prompts.excel import GRAFICO_DSL_SISTEMA, GRAFICO_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {t}" for c, t in df.dtypes.items())

    texto = None
    try:
        texto = _chat_timed("grafico_dsl", [
            {"role": "system", "content": GRAFICO_DSL_SISTEMA},
            {"role": "user",   "content": GRAFICO_DSL_USUARIO.format(
                columnas=columnas, tipos=tipos, pregunta=pregunta,
            )},
        ], temperature=0, max_tokens=150)
        logger.debug("Respuesta gráfico DSL del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo petición de gráfico: %s | respuesta LLM: %r", error, texto)
        return None


def extraer_params_pivote(df: pd.DataFrame, pregunta: str) -> dict | None:
    """Extrae los parámetros de tabla dinámica (filas, columnas, valores, función)."""
    from prompts.excel import PIVOTE_DSL_SISTEMA, PIVOTE_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {t}" for c, t in df.dtypes.items())
    muestra  = df.head(3).to_string(index=False)

    texto = None
    try:
        texto = _chat_timed("pivote_dsl", [
            {"role": "system", "content": PIVOTE_DSL_SISTEMA},
            {"role": "user",   "content": PIVOTE_DSL_USUARIO.format(
                columnas=columnas, tipos=tipos, muestra=muestra, pregunta=pregunta,
            )},
        ], temperature=0, max_tokens=200)
        logger.debug("Params pivote del LLM: %s", texto)
        return json.loads(_limpiar_json(texto.strip()))
    except Exception as error:
        logger.warning("Error extrayendo params de tabla dinámica: %s | respuesta LLM: %r", error, texto)
        return None


def extraer_operaciones_macro(descripcion: str) -> list[dict] | None:
    """Convierte una descripción de macro en una lista de operaciones DSL."""
    from prompts.excel import MACRO_DSL_SISTEMA, MACRO_DSL_USUARIO

    texto = None
    try:
        texto = _chat_timed("macro_dsl", [
            {"role": "system", "content": MACRO_DSL_SISTEMA},
            {"role": "user",   "content": MACRO_DSL_USUARIO.format(descripcion=descripcion)},
        ], temperature=0, max_tokens=400)
        logger.debug("Operaciones macro del LLM: %s", texto)
        ops = json.loads(_limpiar_json(texto.strip()))
        return ops if isinstance(ops, list) else None
    except Exception as error:
        logger.warning("Error extrayendo operaciones de macro: %s | respuesta LLM: %r", error, texto)
        return None


def extraer_regla_formato(df: pd.DataFrame, instruccion: str) -> list[dict] | None:
    """Extrae las reglas de formato condicional desde lenguaje natural.

    Devuelve siempre una lista de reglas DSL o None si no se pudo interpretar.
    """
    from prompts.excel import FORMATO_DSL_SISTEMA, FORMATO_DSL_USUARIO

    columnas = ", ".join(f"'{c}'" for c in df.columns)
    tipos    = ", ".join(f"{c}: {t}" for c, t in df.dtypes.items())
    muestra  = df.head(3).to_string(index=False)

    texto = None
    try:
        texto = _chat_timed("formato_dsl", [
            {"role": "system", "content": FORMATO_DSL_SISTEMA},
            {"role": "user",   "content": FORMATO_DSL_USUARIO.format(
                columnas=columnas, tipos=tipos, muestra=muestra, pregunta=instruccion,
            )},
        ], temperature=0, max_tokens=400)
        logger.debug("Regla formato del LLM: %s", texto)
        parsed = json.loads(_limpiar_json(texto.strip()))
        # Normalizar: el LLM puede devolver un dict (regla única) o lista
        if isinstance(parsed, dict):
            return [parsed]
        if isinstance(parsed, list):
            return parsed
        return None
    except Exception as error:
        logger.warning("Error extrayendo regla de formato: %s | respuesta LLM: %r", error, texto)
        return None


def _col_letra(n: int) -> str:
    """Convierte índice 0-based a letra de columna Excel (A, B, ..., Z, AA, ...)."""
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def extraer_formula(df: pd.DataFrame, instruccion: str) -> dict | None:
    """Extrae la definición de una fórmula Excel desde lenguaje natural.

    Devuelve {"col_nueva": str, "formula": "=D{row}-C{row}"} o None si falla.
    El placeholder {row} se sustituye por el número de fila real al expandir.
    """
    from prompts.excel import FORMULA_DSL_SISTEMA, FORMULA_DSL_USUARIO

    columnas_info = "\n".join(
        f"  '{c}' → {_col_letra(i)}" for i, c in enumerate(df.columns)
    )
    nueva_col_letra = _col_letra(len(df.columns))
    muestra = df.head(3).to_string(index=False)

    texto = None
    try:
        texto = _chat_timed("formula_dsl", [
            {"role": "system", "content": FORMULA_DSL_SISTEMA},
            {"role": "user", "content": FORMULA_DSL_USUARIO.format(
                columnas_info=columnas_info,
                nueva_col_letra=nueva_col_letra,
                muestra=muestra,
                instruccion=instruccion,
            )},
        ], temperature=0, max_tokens=150)
        logger.debug("Respuesta formula DSL del LLM: %s", texto)
        parsed = json.loads(_limpiar_json(texto.strip()))
        if isinstance(parsed, dict) and "formula" in parsed and "col_nueva" in parsed:
            return parsed
        return None
    except Exception as error:
        logger.warning("Error extrayendo fórmula Excel: %s | respuesta LLM: %r", error, texto)
        return None


def transcribir_audio(audio_bytes: bytes, filename: str = "audio.ogg") -> str:
    """Transcribe un mensaje de voz usando el proveedor activo."""
    return obtener_proveedor().transcribir(audio_bytes, filename)


def analizar_imagen(imagen_bytes: bytes, pregunta: str = "") -> str:
    """Analiza una captura de pantalla de Excel usando visión IA."""
    imagen_b64 = base64.standard_b64encode(imagen_bytes).decode("utf-8")
    texto_usuario = (
        pregunta if pregunta
        else "Analiza esta captura de Excel y explica qué hace, qué fórmulas usa y cómo podría mejorarla."
    )
    mensajes = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{imagen_b64}"}},
                {"type": "text", "text": texto_usuario},
            ],
        },
    ]
    return obtener_proveedor().vision(mensajes)
