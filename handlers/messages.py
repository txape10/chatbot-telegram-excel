import asyncio
import logging
import re
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from telegram.error import TelegramError
from services.llm import (obtener_respuesta, extraer_query_dsl,
                          extraer_operacion_edicion, extraer_estructura_excel)
from utils.history import obtener_historial, agregar_mensaje
from utils.excel_context import obtener_contexto
from utils.df_context import obtener_df, guardar_df
from utils.file_meta import obtener_meta
from utils.auth import solo_autorizados
from utils.user_prefs import get_version, ya_fue_preguntado, marcar_preguntado, VERSIONES
from prompts.excel import EXPLICAR_FORMULA, PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO
from excel.query_engine import ejecutar_query, formatear_resultado, QueryError
from excel.editor import aplicar_edicion, exportar_xlsx, EditorError
from excel.exporter import crear_tabla_dinamica, crear_desde_descripcion
from excel.analyzer import analisis_estadistico_completo, analisis_correlaciones, analisis_tendencia

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
    user_id = update.effective_user.id
    pregunta = update.message.text.strip()

    # ── Explicador automático de fórmulas ─────────────────────────────────────
    if pregunta.startswith("="):
        await _explicar_formula(update, user_id, pregunta)
        return

    # ── Editor de archivo (si hay df activo y petición de modificación) ────────
    df_activo = obtener_df(user_id)
    if df_activo is not None and _RE_EDICION.search(pregunta):
        logger.info("Intención de edición detectada para user_id %s: %r", user_id, pregunta)
        await _intentar_edicion(update, user_id, df_activo, pregunta)
        return

    # ── Creación de Excel desde descripción ──────────────────────────────────
    if _RE_CREAR_EXCEL.search(pregunta) and not _RE_EDICION.search(pregunta):
        logger.info("Intención crear Excel detectada para user_id %s: %r", user_id, pregunta)
        await _crear_excel_desde_descripcion(update, user_id, pregunta)
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
        await mensaje_carga.edit_text(respuesta)

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


async def _intentar_edicion(update: Update, user_id: int, df, pregunta: str) -> None:
    """Extrae la operación de edición, aplica el cambio y envía el archivo modificado."""
    mensaje_carga = await update.message.reply_text("⏳ Aplicando modificación...")
    try:
        op = await asyncio.to_thread(extraer_operacion_edicion, df, pregunta)

        if op is None:
            # El LLM dijo RESPUESTA_LIBRE → dejar que el flujo normal responda
            await mensaje_carga.delete()
            # Re-lanzar como mensaje normal (sin edición)
            await _responder_con_llm(update, user_id, pregunta, mensaje_carga=None)
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
        await mensaje_carga.edit_text(respuesta)
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
        await mensaje_carga.edit_text(texto, parse_mode="Markdown")
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
