import asyncio
import io
import logging
import re
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import TelegramError
from services.llm import (obtener_respuesta, extraer_query_dsl,
                          extraer_operacion_edicion, extraer_estructura_excel,
                          extraer_operacion_combinar, extraer_peticion_grafico)
from services.tts import texto_a_audio
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.df_context import (obtener_df, guardar_df, obtener_df_secundario,
                               obtener_nombre_secundario, restaurar_undo, hay_undo)
from utils.file_meta import obtener_meta
from utils.auth import solo_autorizados
from utils.user_prefs import get_version, ya_fue_preguntado, marcar_preguntado, VERSIONES, get_modo_respuesta
from prompts.excel import EXPLICAR_FORMULA, PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO
from excel.query_engine import ejecutar_query, formatear_resultado, QueryError
from excel.editor import aplicar_edicion, exportar_xlsx, combinar_dataframes, EditorError
from excel.exporter import crear_tabla_dinamica, crear_desde_descripcion
from excel.analyzer import analisis_estadistico_completo, analisis_correlaciones, analisis_tendencia
from excel.charts import generar_grafico_personalizado, ChartError

# Detección de intención de edición: verbos que implican modificar el archivo
_RE_EDICION = re.compile(
    r"\b("
    r"a[ñn]ade[r]?\s|agrega[r]?\s|crea\s+(?:una\s+)?columna|nueva\s+columna|"
    r"calcula\s+(?:una\s+)?columna|"
    r"ordena[r]?\s|ordena\s+(?:el|los|por)|"
    r"elimina[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"quita[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"borra[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"elimina[r]?\s+(?:la\s+|el\s+)?columna|"
    r"quita[r]?\s+(?:la\s+|el\s+)?columna|"
    r"borra[r]?\s+(?:la\s+|el\s+)?columna|"
    r"rellena[r]?\s+(?:los\s+|las\s+)?(?:vac[ií]os?|nulos?|huecos?)|"
    r"completa[r]?\s+(?:los\s+|las\s+)?(?:vac[ií]os?|nulos?)|"
    r"renombra[r]?\s|cambia[r]?\s+el\s+nombre\s+de\s|"
    r"aplica[r]?\s+formato\s+condicional|aplica[r]?\s+color|"
    r"pinta[r]?\s+(?:en\s+)?(?:rojo|verde|amarillo|naranja|azul)|colorea[r]?\s|"
    r"normaliza[r]?\s+(?:el\s+)?(?:texto|datos)|"
    r"limpia[r]?\s+(?:los\s+)?espacios|quita[r]?\s+(?:los\s+)?espacios|"
    r"(?:unifica[r]?|convierte[r]?\s+a|pon[er]?\s+en)\s+(?:may[uú]sculas?|min[uú]sculas?)|"
    r"capitaliza[r]?\s|"
    r"(?:corrige[r]?|estandariza[r]?|formatea[r]?)\s+(?:las\s+)?fechas?|"
    r"(?:des)?pivotea[r]?|"
    r"convierte[r]?\s+(?:las\s+)?columnas?\s+en\s+filas?|"
    r"convierte[r]?\s+(?:las\s+)?filas?\s+en\s+columnas?|"
    r"(?:meses?|trimestres?|periodos?)\s+en\s+filas?"
    r")\b",
    re.IGNORECASE,
)

# Detección de combinación de dos archivos (B3)
_RE_COMBINAR = re.compile(
    r"\b("
    r"une[r]?\s|combina[r]?\s|junta[r]?\s|mezcla[r]?\s|fusiona[r]?\s|"
    r"cruza[r]?\s|cruce\s+(?:con|de)|"
    r"(?:los\s+)?dos\s+archivos|ambos\s+archivos|"
    r"merge[r]?\s|join\s"
    r")\b",
    re.IGNORECASE,
)

# Detección de creación de Excel desde descripción (B1/B2)
_RE_CREAR_EXCEL = re.compile(
    r"\b("
    r"crea[rm]?\s+(?:un[ao]?\s+)?(?:nuevo\s+)?(?:excel|tabla|hoja|archivo|libro)|"
    r"hazme\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja|archivo|plantilla)|"
    r"haz\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja|archivo)|"
    r"genera[rm]?\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja|archivo)|"
    r"necesito\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja)\s+(?:con|para|de)|"
    r"quiero\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja)\s+(?:con|para|de)"
    r")\b",
    re.IGNORECASE,
)

# Detección de "explícame este archivo" (E3)
_RE_EXPLICAR_ARCHIVO = re.compile(
    r"\b("
    r"expl[ií]came\s+(?:este\s+)?(?:archivo|excel|datos?|tabla)|"
    r"qu[eé]\s+(?:contiene|hay en|tiene)\s+(?:este\s+)?(?:archivo|excel)|"
    r"descr[ií]beme\s+(?:este\s+)?(?:archivo|excel|datos?)|"
    r"resumen\s+(?:del?\s+)?archivo|analiza\s+(?:este\s+)?archivo|"
    r"de\s+qu[eé]\s+(?:va|trata)\s+(?:este\s+)?(?:archivo|excel)"
    r")\b",
    re.IGNORECASE,
)

# Detección de exportar a CSV (E3)
_RE_EXPORTAR_CSV = re.compile(
    r"\b("
    r"exporta[r]?\s+(?:a\s+|como\s+|en\s+)?csv|"
    r"guarda[r]?\s+(?:como\s+|en\s+)?csv|"
    r"descarga[r]?\s+(?:en\s+|como\s+)?csv|"
    r"convierte[r]?\s+a\s+csv|"
    r"en\s+formato\s+csv|formato\s+csv"
    r")\b",
    re.IGNORECASE,
)

# Detección de deshacer (E2)
_RE_UNDO = re.compile(
    r"\b("
    r"deshaz|deshacer|desh[aá]cer|"
    r"vuelve\s+atr[aá]s|volver\s+atr[aá]s|"
    r"revertir|revierte|rev[eé]rtelo|"
    r"undo|ctrl\s*\+?\s*z|"
    r"cancela\s+(?:el\s+)?(?:[uú]ltimo\s+)?cambio|"
    r"recupera\s+(?:el\s+)?(?:archivo|excel|datos?)\s+anterior"
    r")\b",
    re.IGNORECASE,
)

# Detección de gráfico bajo demanda (E1)
_RE_GRAFICO = re.compile(
    r"\b("
    r"gr[aá]fico[s]?\b|chart\b|"
    r"dibuja[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"pinta[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"genera[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"hazme\s+(?:un\s+)?gr[aá]fico|"
    r"muestra[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"crea[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"visualiza[r]?\s|representa[r]?\s+gr[aá]ficamente|"
    r"gr[aá]fico\s+de\s+(?:barras?|l[ií]neas?|sectores?|tarta|pie|dispersi[oó]n|scatter)|"
    r"histograma\b"
    r")\b",
    re.IGNORECASE,
)

# Detección de análisis estadístico / correlaciones (C1/C2) / tendencia (C3)
_RE_STATS = re.compile(
    r"\b("
    r"estad[ií]stica[s]?|distribuc[ií][oó]n|correlac[ií][oó]n[es]?|"
    r"an[áa]lisis\s+(?:completo|estad[ií]stico|de\s+datos)|"
    r"media\s+y\s+(?:mediana|desviaci[oó]n)|resumen\s+estad[ií]stico|"
    r"desviaci[oó]n\s+(?:est[aá]ndar|t[ií]pica)|percentil[es]?|"
    r"qu[eé]\s+columnas\s+(?:est[aá]n\s+)?(?:m[aá]s\s+)?relacionadas|"
    r"c[oó]mo\s+se\s+relacionan|mapa\s+de\s+calor|heatmap|"
    r"tendencia[s]?|proyecci[oó]n|evoluci[oó]n\s+de\s+|"
    r"c[oó]mo\s+(?:evolucionan|van\s+(?:mis\s+)?|crecen|bajan)|"
    r"predicci[oó]n|pron[oó]stico"
    r")\b",
    re.IGNORECASE,
)

# Detección de tabla dinámica: basta con que mencione "tabla dinámica/s"
# salvo que sea una pregunta puramente informativa (qué es, cómo funciona, etc.)
_RE_TABLA_DINAMICA = re.compile(
    r"tabla[s]?\s+din[aá]mica[s]?",
    re.IGNORECASE,
)
_RE_SOLO_INFORMATIVA = re.compile(
    r"^[¿\s]*(qu[eé]\s+es\b|qu[eé]\s+son\b|c[oó]mo\s+(se\s+)?func|"
    r"para\s+qu[eé]\b|cu[aá]ndo\s+usar|por\s+qu[eé]\b|"
    r"explica[rm]?|qu[eé]\s+hace\b|qu[eé]\s+hacen\b|en\s+qu[eé]\s+consiste)",
    re.IGNORECASE,
)

logger = logging.getLogger(__name__)


async def _enviar_respuesta(update: Update, user_id: int,
                             mensaje_carga, texto: str,
                             parse_mode: str = None) -> None:
    """Envía la respuesta en texto o texto+voz según la preferencia del usuario.

    Siempre edita `mensaje_carga` con el texto (para que el usuario pueda
    releerlo). Si el modo es 'voz', además envía un mensaje de audio a
    continuación. Si el TTS falla, la respuesta de texto ya está enviada.
    """
    await mensaje_carga.edit_text(texto, parse_mode=parse_mode)

    if get_modo_respuesta(user_id) == "voz":
        try:
            audio_bytes = await texto_a_audio(texto)
            if audio_bytes:
                await update.message.reply_voice(voice=io.BytesIO(audio_bytes))
        except Exception as error:
            logger.warning("TTS falló para user_id %s: %s", user_id, error)


_TECLADO_VERSION = InlineKeyboardMarkup([
    [
        InlineKeyboardButton("Microsoft 365",      callback_data="version_365"),
        InlineKeyboardButton("Excel 2021",         callback_data="version_2021"),
    ],
    [
        InlineKeyboardButton("Excel 2019",         callback_data="version_2019"),
        InlineKeyboardButton("Excel 2016 o anterior", callback_data="version_2016"),
    ],
])


@solo_autorizados
async def responder_mensaje(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await procesar_pregunta(update, context, update.message.text.strip())


async def procesar_pregunta(update: Update, context: ContextTypes.DEFAULT_TYPE,
                             pregunta: str) -> None:
    """Núcleo de procesamiento de una pregunta de texto.

    Reutilizable desde cualquier handler (texto, voz, etc.) una vez que
    el texto ya ha sido extraído o transcrito.
    """
    user_id = update.effective_user.id

    # ── Explicador automático de fórmulas ─────────────────────────────────────
    if pregunta.startswith("="):
        await _explicar_formula(update, user_id, pregunta)
        return

    df_activo    = obtener_df(user_id)
    df_secundario = obtener_df_secundario(user_id)

    # ── Combinar dos archivos (B3) ────────────────────────────────────────────
    if df_activo is not None and df_secundario is not None and _RE_COMBINAR.search(pregunta):
        logger.info("Intención de combinación detectada para user_id %s: %r", user_id, pregunta)
        await _intentar_combinar(update, user_id, df_secundario, df_activo, pregunta)
        return

    # ── Deshacer última operación (E2) ───────────────────────────────────────
    if df_activo is not None and _RE_UNDO.search(pregunta):
        logger.info("Intención de undo para user_id %s", user_id)
        await _deshacer_operacion(update, user_id)
        return

    # ── Editor de archivo (si hay df activo y petición de modificación) ────────
    if df_activo is not None and _RE_EDICION.search(pregunta):
        logger.info("Intención de edición detectada para user_id %s: %r", user_id, pregunta)
        await _intentar_edicion(update, user_id, df_activo, pregunta, context)
        return

    # ── Creación de Excel desde descripción ──────────────────────────────────
    if _RE_CREAR_EXCEL.search(pregunta) and not _RE_EDICION.search(pregunta):
        logger.info("Intención crear Excel detectada para user_id %s: %r", user_id, pregunta)
        await _crear_excel_desde_descripcion(update, user_id, pregunta)
        return

    # ── Gráfico bajo demanda ─────────────────────────────────────────────────
    if df_activo is not None and _RE_GRAFICO.search(pregunta):
        logger.info("Intención gráfico bajo demanda para user_id %s: %r", user_id, pregunta)
        await _generar_grafico_bajo_demanda(update, user_id, df_activo, pregunta)
        return

    # ── Generador de tabla dinámica ───────────────────────────────────────────
    if _RE_TABLA_DINAMICA.search(pregunta) and not _RE_SOLO_INFORMATIVA.search(pregunta):
        logger.info("Intención tabla dinámica detectada para user_id %s: %r", user_id, pregunta)
        await _generar_tabla_dinamica(update, user_id)
        return

    # ── Análisis estadístico / correlaciones ──────────────────────────────────
    if df_activo is not None and _RE_STATS.search(pregunta):
        logger.info("Intención análisis estadístico detectada para user_id %s: %r", user_id, pregunta)
        await _analizar_estadisticas(update, user_id, df_activo, pregunta)
        return

    # ── Explícame este archivo ────────────────────────────────────────────────
    if df_activo is not None and _RE_EXPLICAR_ARCHIVO.search(pregunta):
        logger.info("Intención explicar archivo para user_id %s", user_id)
        await _explicar_archivo(update, user_id, df_activo)
        return

    # ── Exportar a CSV ────────────────────────────────────────────────────────
    if df_activo is not None and _RE_EXPORTAR_CSV.search(pregunta):
        logger.info("Intención exportar CSV para user_id %s", user_id)
        await _exportar_csv(update, user_id, df_activo)
        return

    # ── Flujo normal ──────────────────────────────────────────────────────────
    mensaje_carga = await update.message.reply_text("⏳ Pensando...")

    historial = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)

    # Inyectar versión de Excel si está configurada
    version = get_version(user_id)
    prefijo_version = (
        PREGUNTA_CON_VERSION.format(version=VERSIONES.get(version, version))
        if version else ""
    )

    if contexto_excel:
        pregunta_completa = prefijo_version + PREGUNTA_CON_CONTEXTO.format(
            contexto=contexto_excel, pregunta=pregunta
        )
    else:
        pregunta_completa = prefijo_version + pregunta

    try:
        # ── Intento DSL si hay DataFrame activo ───────────────────────────────
        if df_activo is not None and contexto_excel:
            respuesta = await _intentar_dsl(
                update, user_id, df_activo, pregunta, historial,
                pregunta_completa, mensaje_carga,
            )
            if respuesta is not None:
                return

        # ── Flujo LLM normal ──────────────────────────────────────────────────
        respuesta = obtener_respuesta(historial, pregunta_completa)
        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", respuesta)
        await _enviar_respuesta(update, user_id, mensaje_carga, respuesta)

        # Preguntar versión la primera vez (tras responder, no antes)
        if not ya_fue_preguntado(user_id):
            marcar_preguntado(user_id)
            await update.message.reply_text(
                "💡 Para darte respuestas más precisas, ¿qué versión de Excel usas?\n"
                "(Puedes cambiarlo siempre con /version)",
                reply_markup=_TECLADO_VERSION,
            )

    except TelegramError as error:
        logger.error("Error de Telegram para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ Error al enviar la respuesta. Inténtalo de nuevo.")
    except Exception as error:
        logger.error("Error inesperado para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text(
            "⚠️ El asistente no está disponible en este momento. Inténtalo en unos segundos."
        )


# Operaciones que requieren confirmación antes de ejecutarse
_OPS_DESTRUCTIVAS = {"eliminar_columna", "eliminar_duplicados", "filtrar_exportar"}


async def _intentar_edicion(update: Update, user_id: int, df, pregunta: str,
                             context=None) -> None:
    """Extrae la operación de edición, aplica el cambio y envía el archivo modificado."""
    mensaje_carga = await update.message.reply_text("⏳ Aplicando modificación...")
    try:
        op = await asyncio.to_thread(extraer_operacion_edicion, df, pregunta)

        if op is None:
            # El LLM dijo RESPUESTA_LIBRE → dejar que el flujo normal responda
            await mensaje_carga.delete()
            await _responder_con_llm(update, user_id, pregunta, mensaje_carga=None)
            return

        # Operaciones destructivas: pedir confirmación primero
        if op.get("op") in _OPS_DESTRUCTIVAS and context is not None:
            await mensaje_carga.delete()
            await _pedir_confirmacion(update, context, user_id, df, op)
            return

        df_mod, descripcion, extras = await asyncio.to_thread(aplicar_edicion, df, op)

        # Obtener el nombre del archivo original si está disponible
        meta = obtener_meta(user_id)
        nombre_base = meta["nombre"] if meta else "archivo"

        fmt_cond = extras if op.get("op") == "formato_condicional" else None
        buf, nombre_archivo = await asyncio.to_thread(
            exportar_xlsx, df_mod, nombre_base, descripcion, fmt_cond
        )

        # Actualizar el df en memoria con la versión modificada
        guardar_df(user_id, df_mod)

        caption = f"✅ {descripcion}\n\nEl archivo incluye los cambios aplicados."
        await update.message.reply_document(document=buf, filename=nombre_archivo, caption=caption)
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except EditorError as error:
        logger.warning("EditorError para user_id %s: %s", user_id, error)
        try:
            await mensaje_carga.edit_text(f"⚠️ No pude aplicar la modificación: {error}")
        except Exception:
            pass
    except Exception as error:
        logger.error("Error en edición para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo modificar el archivo. Inténtalo de nuevo.")
        except Exception:
            pass


async def _responder_con_llm(update: Update, user_id: int, pregunta: str,
                              mensaje_carga=None) -> None:
    """Responde usando el LLM normal (sin contexto de edición)."""
    if mensaje_carga is None:
        mensaje_carga = await update.message.reply_text("⏳ Pensando...")
    historial = obtener_historial(user_id)
    contexto_excel = obtener_contexto(user_id)
    version = get_version(user_id)
    prefijo = PREGUNTA_CON_VERSION.format(version=VERSIONES.get(version, version)) if version else ""
    pregunta_completa = prefijo + (
        PREGUNTA_CON_CONTEXTO.format(contexto=contexto_excel, pregunta=pregunta)
        if contexto_excel else pregunta
    )
    try:
        respuesta = obtener_respuesta(historial, pregunta_completa)
        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", respuesta)
        await _enviar_respuesta(update, user_id, mensaje_carga, respuesta)
    except Exception as error:
        logger.error("Error LLM fallback para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ Error al obtener respuesta. Inténtalo de nuevo.")


async def _generar_tabla_dinamica(update: Update, user_id: int) -> None:
    """Genera y envía un .xlsx con tabla dinámica usando los datos del usuario (si los tiene)."""
    mensaje_carga = await update.message.reply_text("⏳ Generando tabla dinámica...")
    try:
        df = obtener_df(user_id)
        usar_datos_usuario = df is not None and not df.empty
        logger.info("Generando tabla dinámica para user_id %s (datos propios: %s)",
                    user_id, usar_datos_usuario)

        buf, nombre = await asyncio.to_thread(crear_tabla_dinamica, df)

        if usar_datos_usuario:
            caption = (
                "📊 Archivo preparado para tabla dinámica interactiva\n\n"
                "El archivo tiene dos hojas:\n"
                "· Datos — tus datos como Excel Table (con filtros activos)\n"
                "· Resúmenes — agrupaciones y cruces precalculados\n\n"
                "Para crear la tabla dinámica interactiva en Excel:\n"
                "1. Abre el archivo y haz clic en cualquier celda de la hoja Datos\n"
                "2. Ve a Insertar → Tabla dinámica → Aceptar\n"
                "3. Arrastra los campos al panel de la derecha"
            )
        else:
            caption = (
                "📊 Ejemplo de tabla dinámica (datos de muestra)\n\n"
                "No tenía tus datos en memoria — puede que hayas reiniciado el bot "
                "o aún no hayas subido ningún archivo.\n\n"
                "Sube tu Excel y repite la petición para generarla con tus propios datos.\n\n"
                "Para crear la TD interactiva: clic en Datos → Insertar → Tabla dinámica → Aceptar"
            )

        # Enviar el archivo PRIMERO; borrar el mensaje de carga después
        await update.message.reply_document(document=buf, filename=nombre, caption=caption)
        try:
            await mensaje_carga.delete()
        except Exception:
            pass   # si ya no existe el mensaje de carga, no es crítico

    except Exception as error:
        logger.error("Error generando tabla dinámica para user_id %s: %s",
                     user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo generar el archivo. Inténtalo de nuevo.")
        except Exception:
            pass


async def _intentar_dsl(update, user_id, df, pregunta, historial,
                        pregunta_completa, mensaje_carga):
    """Intenta resolver la pregunta mediante el motor DSL.

    Devuelve la respuesta si tuvo éxito, o None si debe continuar por LLM normal.
    """
    try:
        query = await asyncio.to_thread(extraer_query_dsl, df, pregunta)
        if query is None:
            return None   # el LLM decidió RESPUESTA_LIBRE → flujo normal

        resultado, descripcion = await asyncio.to_thread(ejecutar_query, df, query)
        texto = formatear_resultado(resultado, descripcion)

        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", texto)
        await _enviar_respuesta(update, user_id, mensaje_carga, texto, parse_mode="Markdown")
        return texto

    except QueryError as error:
        logger.warning("QueryError para user_id %s: %s — usando LLM normal", user_id, error)
        return None   # error de datos → mejor responder con LLM normal
    except Exception as error:
        logger.warning("Error en DSL para user_id %s: %s — usando LLM normal", user_id, error)
        return None


async def _explicar_formula(update: Update, user_id: int, formula: str) -> None:
    """Explica paso a paso una fórmula de Excel."""
    mensaje_carga = await update.message.reply_text("⏳ Analizando la fórmula...")
    prompt = EXPLICAR_FORMULA.format(formula=formula)
    try:
        historial = obtener_historial(user_id)
        respuesta = obtener_respuesta(historial, prompt)
        agregar_mensaje(user_id, "user", formula)
        agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)
    except Exception as error:
        logger.error("Error explicando fórmula para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo analizar la fórmula. Inténtalo de nuevo.")


async def _intentar_combinar(update: Update, user_id: int,
                              df_a: "pd.DataFrame", df_b: "pd.DataFrame",
                              pregunta: str) -> None:
    """Combina los dos DataFrames en memoria según la petición del usuario."""
    nombre_a = obtener_nombre_secundario(user_id)
    meta_b   = obtener_meta(user_id)
    nombre_b = meta_b["nombre"] if meta_b else "archivo"

    mensaje_carga = await update.message.reply_text("⏳ Combinando archivos...")
    try:
        op = await asyncio.to_thread(extraer_operacion_combinar, df_a, df_b, pregunta)
        df_result, descripcion = await asyncio.to_thread(combinar_dataframes, df_a, df_b, op)

        # El resultado pasa a ser el df activo
        guardar_df(user_id, df_result)

        buf, nombre_archivo = await asyncio.to_thread(
            exportar_xlsx, df_result, "combinado", descripcion
        )

        caption = (
            f"✅ {descripcion}\n\n"
            f"Archivo A: *{nombre_a}*\n"
            f"Archivo B: *{nombre_b}*"
        )
        await update.message.reply_document(document=buf, filename=nombre_archivo,
                                             caption=caption, parse_mode="Markdown")
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except EditorError as error:
        logger.warning("EditorError en combinar para user_id %s: %s", user_id, error)
        try:
            await mensaje_carga.edit_text(f"⚠️ No pude combinar los archivos: {error}")
        except Exception:
            pass
    except Exception as error:
        logger.error("Error combinando archivos para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudieron combinar los archivos. Inténtalo de nuevo.")
        except Exception:
            pass


async def _crear_excel_desde_descripcion(update: Update, user_id: int, pregunta: str) -> None:
    """Genera un .xlsx con la estructura que el usuario describió en lenguaje natural."""
    mensaje_carga = await update.message.reply_text("⏳ Creando tu Excel...")
    try:
        estructura = await asyncio.to_thread(extraer_estructura_excel, pregunta)

        if estructura is None:
            await mensaje_carga.edit_text(
                "⚠️ No pude interpretar la estructura del Excel. "
                "Descríbela con más detalle (columnas, datos, etc.)."
            )
            return

        buf, nombre = await asyncio.to_thread(crear_desde_descripcion, estructura)

        titulo = estructura.get("titulo", "Excel")
        columnas = estructura.get("columnas", [])
        n_filas = len(estructura.get("datos", []))
        caption = (
            f"📄 *{titulo}*\n\n"
            f"Columnas: {', '.join(columnas)}\n"
            f"Filas de datos: {n_filas}\n\n"
            "Puedes abrir el archivo directamente en Excel."
        )

        await update.message.reply_document(document=buf, filename=nombre, caption=caption,
                                             parse_mode="Markdown")
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except Exception as error:
        logger.error("Error creando Excel desde descripción para user_id %s: %s",
                     user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo crear el archivo. Inténtalo de nuevo.")
        except Exception:
            pass


_NOMBRES_OP = {
    "eliminar_columna":    "Eliminar columna",
    "eliminar_duplicados": "Eliminar duplicados",
    "filtrar_exportar":    "Filtrar y exportar",
}


async def _pedir_confirmacion(update: Update, context, user_id: int,
                               df, op: dict) -> None:
    """Muestra un botón Sí/No antes de ejecutar una operación destructiva."""
    import json as _json
    nombre_op = _NOMBRES_OP.get(op.get("op", ""), op.get("op", "operación"))

    # Detalles legibles de la operación
    detalles = ""
    if op.get("op") == "eliminar_columna":
        cols = op.get("columnas", [])
        detalles = f"Columnas a eliminar: *{', '.join(cols)}*"
    elif op.get("op") == "eliminar_duplicados":
        cols = op.get("columnas")
        detalles = f"Comparando: *{'todas las columnas' if not cols else ', '.join(cols)}*"
    elif op.get("op") == "filtrar_exportar":
        filtros = op.get("filtros", [])
        detalles = "Filtros: " + ", ".join(
            f"{f.get('col')} {f.get('op')} {f.get('val')}" for f in filtros
        )

    # Guardar la operación pendiente en user_data para recuperarla en el callback
    context.user_data["op_pendiente"] = _json.dumps(op, ensure_ascii=False)

    teclado = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Sí, aplicar",  callback_data="confirmar_op_si"),
            InlineKeyboardButton("❌ Cancelar",      callback_data="confirmar_op_no"),
        ]
    ])
    await update.message.reply_text(
        f"⚠️ *{nombre_op}*\n{detalles}\n\n¿Confirmas la operación?",
        parse_mode="Markdown",
        reply_markup=teclado,
    )


async def callback_confirmacion(update, context) -> None:
    """Ejecuta o cancela la operación destructiva según la respuesta del usuario."""
    import json as _json
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    if query.data == "confirmar_op_no":
        context.user_data.pop("op_pendiente", None)
        await query.edit_message_text("❌ Operación cancelada.")
        return

    op_json = context.user_data.pop("op_pendiente", None)
    if not op_json:
        await query.edit_message_text("⚠️ No encontré la operación pendiente. Vuelve a intentarlo.")
        return

    op = _json.loads(op_json)
    df = obtener_df(user_id)
    if df is None:
        await query.edit_message_text("⚠️ No hay ningún archivo activo.")
        return

    await query.edit_message_text("⏳ Aplicando...")
    try:
        df_mod, descripcion, extras = await asyncio.to_thread(aplicar_edicion, df, op)
        meta = obtener_meta(user_id)
        nombre_base = meta["nombre"] if meta else "archivo"
        fmt_cond = extras if op.get("op") == "formato_condicional" else None
        buf, nombre_archivo = await asyncio.to_thread(
            exportar_xlsx, df_mod, nombre_base, descripcion, fmt_cond
        )
        guardar_df(user_id, df_mod)
        await query.message.reply_document(
            document=buf, filename=nombre_archivo,
            caption=f"✅ {descripcion}\n\nEl archivo incluye los cambios aplicados."
        )
        await query.edit_message_text(f"✅ {descripcion}")
    except EditorError as error:
        await query.edit_message_text(f"⚠️ No pude aplicar la modificación: {error}")
    except Exception as error:
        logger.error("Error en callback confirmación user_id %s: %s", user_id, error, exc_info=True)
        await query.edit_message_text("⚠️ No se pudo aplicar. Inténtalo de nuevo.")


async def _deshacer_operacion(update: Update, user_id: int) -> None:
    """Restaura el DataFrame al estado anterior a la última edición."""
    if not hay_undo(user_id):
        await update.message.reply_text(
            "↩️ No hay ninguna operación anterior que deshacer."
        )
        return

    df_restaurado = restaurar_undo(user_id)
    meta = obtener_meta(user_id)
    nombre_base = meta["nombre"] if meta else "archivo"

    buf, nombre_archivo = await asyncio.to_thread(
        exportar_xlsx, df_restaurado, nombre_base, "Versión anterior restaurada"
    )
    await update.message.reply_document(
        document=buf,
        filename=nombre_archivo,
        caption=(
            f"↩️ Operación deshecha.\n"
            f"Archivo restaurado: *{df_restaurado.shape[0]} filas × "
            f"{df_restaurado.shape[1]} columnas*"
        ),
        parse_mode="Markdown",
    )


async def _generar_grafico_bajo_demanda(update: Update, user_id: int,
                                        df, pregunta: str) -> None:
    """Extrae parámetros del gráfico con el LLM y genera la imagen PNG."""
    mensaje_carga = await update.message.reply_text("⏳ Generando gráfico...")
    try:
        params = await asyncio.to_thread(extraer_peticion_grafico, df, pregunta)

        if params is None:
            await mensaje_carga.edit_text(
                "⚠️ No entendí qué gráfico quieres. Intenta con algo como:\n"
                "• _Gráfico de barras de Ventas por Región_\n"
                "• _Hazme un gráfico de líneas de la columna Beneficio_",
                parse_mode="Markdown",
            )
            return

        col_y   = params.get("col_y")
        col_x   = params.get("col_x")
        tipo    = params.get("tipo", "barras")
        agregar = params.get("agregar")

        buf, titulo = await asyncio.to_thread(
            generar_grafico_personalizado, df, col_y, col_x, tipo, agregar
        )

        meta = obtener_meta(user_id)
        nombre = meta["nombre"] if meta else "archivo"
        caption = f"📊 {titulo} — {nombre}"

        await update.message.reply_photo(photo=buf, caption=caption)
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except ChartError as error:
        logger.warning("ChartError para user_id %s: %s", user_id, error)
        try:
            await mensaje_carga.edit_text(f"⚠️ {error}")
        except Exception:
            pass
    except Exception as error:
        logger.error("Error generando gráfico para user_id %s: %s",
                     user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo generar el gráfico. Inténtalo de nuevo.")
        except Exception:
            pass


async def _explicar_archivo(update: Update, user_id: int, df) -> None:
    """Genera una descripción en lenguaje natural del archivo activo usando el LLM."""
    mensaje_carga = await update.message.reply_text("⏳ Analizando el archivo...")
    try:
        from excel.analyzer import resumir, analizar_calidad

        resumen_tecnico = resumir(df)
        calidad         = analizar_calidad(df)

        cols_num  = df.select_dtypes(include="number").columns.tolist()
        cols_text = df.select_dtypes(include="object").columns.tolist()
        muestra   = df.head(3).to_string(index=False)

        prompt = (
            f"Tengo un archivo Excel con los siguientes datos:\n\n"
            f"Resumen: {resumen_tecnico}\n\n"
            f"Columnas numéricas: {', '.join(cols_num) or 'ninguna'}\n"
            f"Columnas de texto: {', '.join(cols_text) or 'ninguna'}\n\n"
            f"Primeras filas:\n{muestra}\n\n"
            f"Problemas de calidad detectados: {calidad if calidad else 'ninguno'}\n\n"
            "Explícame en español, de forma clara y concisa, qué contiene este archivo: "
            "qué tipo de datos son, para qué podría servir, qué columnas son clave, "
            "y si hay algún problema de calidad relevante que deba saber."
        )

        historial = obtener_historial(user_id)
        respuesta = await asyncio.to_thread(obtener_respuesta, historial, prompt)
        agregar_mensaje(user_id, "user", "Explícame este archivo")
        agregar_mensaje(user_id, "model", respuesta)
        await _enviar_respuesta(update, user_id, mensaje_carga, respuesta)

    except Exception as error:
        logger.error("Error explicando archivo para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No pude analizar el archivo. Inténtalo de nuevo.")
        except Exception:
            pass


async def _exportar_csv(update: Update, user_id: int, df) -> None:
    """Exporta el DataFrame activo como archivo CSV."""
    mensaje_carga = await update.message.reply_text("⏳ Generando CSV...")
    try:
        import io as _io

        def _crear_csv():
            buf = _io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")  # utf-8-sig para Excel en Windows
            buf.seek(0)
            return buf

        buf = await asyncio.to_thread(_crear_csv)
        meta = obtener_meta(user_id)
        nombre_base = meta["nombre"].rsplit(".", 1)[0] if meta else "archivo"
        nombre_csv  = f"{nombre_base}.csv"

        await update.message.reply_document(
            document=buf,
            filename=nombre_csv,
            caption=f"📄 *{nombre_csv}*\n{df.shape[0]:,} filas × {df.shape[1]} columnas",
            parse_mode="Markdown",
        )
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except Exception as error:
        logger.error("Error exportando CSV para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo generar el CSV. Inténtalo de nuevo.")
        except Exception:
            pass


_RE_CORR = re.compile(
    r"correlac[ií][oó]n|relacionad|mapa\s+de\s+calor|heatmap|c[oó]mo\s+se\s+relacionan",
    re.IGNORECASE,
)
_RE_TENDENCIA = re.compile(
    r"tendencia[s]?|proyecci[oó]n|evoluci[oó]n\s+de\s+|"
    r"c[oó]mo\s+(?:evolucionan|van\s+(?:mis\s+)?|crecen|bajan)|"
    r"predicci[oó]n|pron[oó]stico",
    re.IGNORECASE,
)


async def _analizar_estadisticas(update: Update, user_id: int, df, pregunta: str) -> None:
    """Devuelve estadísticas, correlaciones o tendencia según lo que pida el usuario."""
    mensaje_carga = await update.message.reply_text("⏳ Calculando...")
    try:
        if _RE_TENDENCIA.search(pregunta):
            texto, buf_img = await asyncio.to_thread(analisis_tendencia, df)
            await update.message.reply_text(texto, parse_mode="Markdown")
            if buf_img is not None:
                await update.message.reply_photo(photo=buf_img,
                                                 caption="📈 Gráfico de tendencia")
        elif _RE_CORR.search(pregunta):
            texto, buf_img = await asyncio.to_thread(analisis_correlaciones, df)
            await update.message.reply_text(texto, parse_mode="Markdown")
            if buf_img is not None:
                await update.message.reply_photo(photo=buf_img,
                                                 caption="🔥 Mapa de correlaciones")
        else:
            texto = await asyncio.to_thread(analisis_estadistico_completo, df)
            await update.message.reply_text(texto, parse_mode="Markdown")

        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except Exception as error:
        logger.error("Error en análisis estadístico para user_id %s: %s",
                     user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo completar el análisis. Inténtalo de nuevo.")
        except Exception:
            pass
