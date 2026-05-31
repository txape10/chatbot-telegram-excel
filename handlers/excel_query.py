import asyncio
import io
import logging
import re

from telegram import Update
from telegram.ext import ContextTypes

from services.llm import (LLMError, obtener_respuesta, obtener_proveedor_privado,
                           extraer_query_dsl, extraer_peticion_grafico,
                           extraer_estructura_excel)
from utils.history import obtener_historial, agregar_mensaje
from utils.df_context import obtener_df
from utils.file_meta import obtener_meta
from utils.excel_context import obtener_contexto
from utils.user_prefs import get_modo_privado, get_version, VERSIONES
from prompts.excel import EXPLICAR_FORMULA, PREGUNTA_CON_VERSION, PREGUNTA_CON_CONTEXTO
from excel.query_engine import ejecutar_query, formatear_resultado, QueryError
from excel.editor import exportar_xlsx
from excel.charts import generar_grafico_personalizado, ChartError
from excel.exporter import crear_tabla_dinamica, crear_desde_descripcion
from excel.analyzer import (analisis_estadistico_completo, analisis_correlaciones,
                             analisis_tendencia, comparar_dataframes as _comparar_dfs)
from handlers.intent_patterns import _CLAVE_ACLARACION, _RE_CORR, _RE_TENDENCIA
from handlers.bot_helpers import _pedir_aclaracion, _enviar_respuesta

logger = logging.getLogger(__name__)


async def _intentar_dsl(update, user_id, df, pregunta, historial,
                        pregunta_completa, mensaje_carga, context=None):
    """Intenta resolver la pregunta mediante el motor DSL.

    Devuelve la respuesta si tuvo éxito, None si debe continuar por LLM normal,
    o el sentinel "aclaracion" cuando el LLM pide datos adicionales.
    """
    try:
        query = await asyncio.to_thread(extraer_query_dsl, df, pregunta)
        if query is None:
            return None

        # Aclaración solo viene como dict plano (nunca como lista)
        if isinstance(query, dict) and query.get(_CLAVE_ACLARACION) and context is not None:
            logger.info("Aclaración solicitada (DSL) para user_id %s", user_id)
            await mensaje_carga.delete()
            await _pedir_aclaracion(update, context, user_id, "dsl", query)
            return "aclaracion"

        # Soporta objeto único o array de operaciones (multi-paso)
        ops = query if isinstance(query, list) else [query]
        partes: list[str] = []
        for q in ops:
            resultado, descripcion = await asyncio.to_thread(ejecutar_query, df, q)
            partes.append(formatear_resultado(resultado, descripcion))
        texto = "\n\n".join(partes)

        agregar_mensaje(user_id, "user", pregunta)
        agregar_mensaje(user_id, "model", texto)
        await _enviar_respuesta(update, user_id, mensaje_carga, texto, parse_mode="Markdown")
        return texto

    except QueryError as error:
        logger.warning("QueryError para user_id %s: %s — usando LLM normal", user_id, error)
        return None
    except Exception as error:
        logger.warning("Error en DSL para user_id %s: %s — usando LLM normal", user_id, error)
        return None


async def _explicar_formula(update: Update, user_id: int, formula: str) -> None:
    """Explica paso a paso una fórmula de Excel."""
    mensaje_carga = await update.message.reply_text("⏳ Analizando la fórmula...")
    prompt = EXPLICAR_FORMULA.format(formula=formula)
    try:
        historial = obtener_historial(user_id)
        _modo_privado = get_modo_privado(user_id)
        respuesta = obtener_respuesta(
            historial, prompt,
            obtener_proveedor_privado() if _modo_privado else None,
        )
        if not _modo_privado:
            agregar_mensaje(user_id, "user", formula)
            agregar_mensaje(user_id, "model", respuesta)
        await mensaje_carga.edit_text(respuesta)
    except Exception as error:
        logger.error("Error explicando fórmula para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo analizar la fórmula. Inténtalo de nuevo.")


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

        await update.message.reply_document(document=buf, filename=nombre, caption=caption)
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except Exception as error:
        logger.error("Error generando tabla dinámica para user_id %s: %s",
                     user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo generar el archivo. Inténtalo de nuevo.")
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

        await update.message.reply_document(
            document=buf, filename=nombre, caption=caption, parse_mode="Markdown"
        )
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

        buf, titulo = await asyncio.to_thread(
            generar_grafico_personalizado,
            df,
            params.get("col_y"),
            params.get("col_x"),
            params.get("tipo", "barras"),
            params.get("agregar"),
        )

        meta = obtener_meta(user_id)
        nombre = meta["nombre"] if meta else "archivo"
        await update.message.reply_photo(photo=buf, caption=f"📊 {titulo} — {nombre}")
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
        logger.error("Error generando gráfico para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo generar el gráfico. Inténtalo de nuevo.")
        except Exception:
            pass


async def _analizar_estadisticas(update: Update, user_id: int, df, pregunta: str) -> None:
    """Devuelve estadísticas, correlaciones o tendencia según lo que pida el usuario."""
    mensaje_carga = await update.message.reply_text("⏳ Calculando...")
    try:
        if _RE_TENDENCIA.search(pregunta):
            texto, buf_img = await asyncio.to_thread(analisis_tendencia, df)
            await update.message.reply_text(texto, parse_mode="Markdown")
            if buf_img is not None:
                await update.message.reply_photo(photo=buf_img, caption="📈 Gráfico de tendencia")
        elif _RE_CORR.search(pregunta):
            texto, buf_img = await asyncio.to_thread(analisis_correlaciones, df)
            await update.message.reply_text(texto, parse_mode="Markdown")
            if buf_img is not None:
                await update.message.reply_photo(photo=buf_img, caption="🔥 Mapa de correlaciones")
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


async def _previsualizar(update: Update, user_id: int, df, pregunta: str) -> None:
    """Muestra las primeras o últimas N filas del DataFrame en el chat."""
    ultimas = bool(re.search(r"\b[uú]ltimas?\b", pregunta, re.IGNORECASE))
    m = re.search(r"\b(\d+)\b", pregunta)
    n = int(m.group(1)) if m else 10
    n = min(n, 30)

    muestra = df.tail(n) if ultimas else df.head(n)
    pos_txt = f"últimas {n}" if ultimas else f"primeras {n}"

    tabla = muestra.to_string(index=False, max_colwidth=20)
    texto = (
        f"📋 *{pos_txt} filas* ({df.shape[0]:,} total × {df.shape[1]} columnas)\n\n"
        f"```\n{tabla}\n```"
    )
    await update.message.reply_text(texto, parse_mode="Markdown")


async def _valores_unicos(update: Update, user_id: int, df, pregunta: str) -> None:
    """Muestra los valores únicos de una columna (o de todas si no se especifica)."""
    col_encontrada = None
    pregunta_lower = pregunta.lower()
    for col in df.columns:
        if col.lower() in pregunta_lower:
            col_encontrada = col
            break

    if col_encontrada:
        unicos = df[col_encontrada].dropna().unique()
        n_unicos = len(unicos)
        muestra = sorted(unicos, key=str)[:30]
        lista = "\n".join(f"  • {v}" for v in muestra)
        resto = f"\n  _...y {n_unicos - 30} más_" if n_unicos > 30 else ""
        texto = (
            f"🔍 *Valores únicos de '{col_encontrada}'* ({n_unicos} distintos)\n\n"
            f"{lista}{resto}"
        )
    else:
        lineas = [f"🔍 *Valores únicos por columna* ({df.shape[0]:,} filas)\n"]
        for col in df.columns:
            n_unicos = df[col].nunique()
            if df[col].dtype == object and n_unicos <= 10:
                vals = ", ".join(str(v) for v in sorted(df[col].dropna().unique(), key=str))
                lineas.append(f"• *{col}* ({n_unicos}): {vals}")
            else:
                lineas.append(f"• *{col}*: {n_unicos:,} valores distintos")
        texto = "\n".join(lineas)

    await update.message.reply_text(texto, parse_mode="Markdown")


async def _comparar_archivos(update: Update, user_id: int, df_a, df_b) -> None:
    """Compara los dos DataFrames en memoria y envía el informe de diferencias."""
    mensaje_carga = await update.message.reply_text("⏳ Comparando archivos...")
    try:
        from utils.df_context import obtener_nombre_secundario
        from utils.file_meta import obtener_meta as _obtener_meta

        nombre_a = obtener_nombre_secundario(user_id)
        meta_b   = _obtener_meta(user_id)
        nombre_b = meta_b["nombre"] if meta_b else "archivo activo"

        resumen, df_diff = await asyncio.to_thread(
            _comparar_dfs, df_a, df_b, nombre_a, nombre_b
        )

        await mensaje_carga.delete()
        await update.message.reply_text(resumen, parse_mode="Markdown")

        if df_diff is not None and not df_diff.empty:
            buf, nombre_archivo = await asyncio.to_thread(
                exportar_xlsx, df_diff, "diferencias", "Filas que difieren entre archivos"
            )
            await update.message.reply_document(
                document=buf,
                filename=nombre_archivo,
                caption="📋 Filas que difieren entre los dos archivos",
            )

    except Exception as error:
        logger.error("Error comparando archivos para user_id %s: %s", user_id, error, exc_info=True)
        try:
            await mensaje_carga.edit_text("⚠️ No se pudo comparar los archivos. Inténtalo de nuevo.")
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
        _modo_privado = get_modo_privado(user_id)
        respuesta = await asyncio.to_thread(
            obtener_respuesta, historial, prompt,
            obtener_proveedor_privado() if _modo_privado else None,
        )
        if not _modo_privado:
            agregar_mensaje(user_id, "user", "Explícame este archivo")
            agregar_mensaje(user_id, "model", respuesta)
        await _enviar_respuesta(update, user_id, mensaje_carga, respuesta)

    except LLMError as error:
        logger.warning("Error LLM explicando archivo para user_id %s [%s]: %s",
                       user_id, error.tipo, error)
        try:
            await mensaje_carga.edit_text(error.mensaje_usuario)
        except Exception:
            pass
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
        def _crear_csv():
            buf = io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")
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
