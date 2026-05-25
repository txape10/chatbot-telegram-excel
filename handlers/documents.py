import asyncio
import logging
import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from utils.auth import solo_autorizados
from utils.excel_context import guardar_contexto
from utils.df_context import guardar_df, obtener_df, guardar_df_secundario
from utils.file_meta import obtener_meta
from utils.file_meta import guardar_meta
from utils.chart_context import guardar_datos_grafico, obtener_datos_grafico
from utils.sheet_context import guardar_hojas, obtener_hoja, listar_hojas
from excel.reader import leer_excel_hojas, leer_csv
from excel.analyzer import resumir, resumir_hojas, construir_contexto, detectar_errores_xlsx
from excel.charts import generar_grafico
from config import MAX_FILAS, MAX_COLUMNAS, MAX_HOJAS, TAMANIO_MAXIMO_MB

logger = logging.getLogger(__name__)
DIRECTORIO_TEMP   = os.path.join(os.path.dirname(__file__), "..", "data", "temp")

_EXTENSIONES_EXCEL = (".xlsx", ".xls")
_EXTENSIONES_CSV   = (".csv",)

# Magic bytes para validar tipo real de archivo
_MAGIC_XLSX = b"PK\x03\x04"          # ZIP — usado por .xlsx
_MAGIC_XLS  = b"\xd0\xcf\x11\xe0"   # OLE2 — usado por .xls


def _validar_extension_real(ruta: str, es_xlsx: bool, es_xls: bool) -> bool:
    """Comprueba los primeros bytes del archivo para confirmar que es realmente Excel."""
    try:
        with open(ruta, "rb") as f:
            cabecera = f.read(8)
        if es_xlsx and cabecera[:4] == _MAGIC_XLSX:
            return True
        if es_xls and cabecera[:4] == _MAGIC_XLS:
            return True
        if not es_xlsx and not es_xls:
            return True   # CSV: no hay magic bytes fiables, confiamos en la extensión
        return False
    except Exception:
        return False


@solo_autorizados
async def recibir_documento(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    documento = update.message.document
    user_id   = update.effective_user.id
    nombre    = os.path.basename(documento.file_name or "archivo")   # sanitizar ruta
    nombre_lower = nombre.lower()
    es_xlsx = nombre_lower.endswith(".xlsx")
    es_xls  = nombre_lower.endswith(".xls")
    es_csv  = nombre_lower.endswith(".csv")

    if not es_xlsx and not es_xls and not es_csv:
        await update.message.reply_text("⚠️ Solo acepto archivos Excel (.xlsx, .xls) o CSV (.csv).")
        return

    if documento.file_size > TAMANIO_MAXIMO_MB * 1024 * 1024:
        await update.message.reply_text(f"⚠️ El archivo supera el límite de {TAMANIO_MAXIMO_MB} MB.")
        return

    mensaje_carga = await update.message.reply_text("⏳ Leyendo el archivo...")
    ruta = None

    try:
        os.makedirs(DIRECTORIO_TEMP, exist_ok=True)
        ruta = os.path.join(DIRECTORIO_TEMP, f"{user_id}_{nombre}")
        archivo = await documento.get_file()
        await archivo.download_to_drive(ruta)

        # Validar tipo real por magic bytes
        if not _validar_extension_real(ruta, es_xlsx, es_xls):
            await mensaje_carga.edit_text("⚠️ El archivo no es un Excel válido aunque tenga la extensión correcta.")
            return

        if es_csv:
            df = await asyncio.to_thread(leer_csv, ruta)
            _validar_limites(df, nombre, n_hojas=1)
            await _procesar_dataframe(update, user_id, df, nombre, mensaje_carga, errores=[])
        else:
            sheets = await asyncio.to_thread(leer_excel_hojas, ruta)

            # Límite de hojas
            if len(sheets) > MAX_HOJAS:
                await mensaje_carga.edit_text(
                    f"⚠️ El archivo tiene {len(sheets)} hojas. Solo proceso hasta {MAX_HOJAS}."
                )
                sheets = dict(list(sheets.items())[:MAX_HOJAS])

            errores = await asyncio.to_thread(detectar_errores_xlsx, ruta)

            if len(sheets) > 1:
                guardar_hojas(user_id, sheets)
                resumen = resumir_hojas(sheets, nombre)
                nombres = list(sheets.keys())
                botones = [
                    InlineKeyboardButton(f"📋 {n[:20]}", callback_data=f"sheet_{i}")
                    for i, n in enumerate(nombres)
                ]
                filas = [botones[i:i+2] for i in range(0, len(botones), 2)]
                await mensaje_carga.edit_text(resumen, parse_mode="Markdown",
                                              reply_markup=InlineKeyboardMarkup(filas))
            else:
                df = list(sheets.values())[0]
                _validar_limites(df, nombre, n_hojas=1)
                await _procesar_dataframe(update, user_id, df, nombre, mensaje_carga, errores)

    except _LimiteExcedidoError as error:
        logger.warning("Límite superado en archivo de user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text(f"⚠️ {error}")
    except Exception as error:
        logger.error("Error procesando archivo para user_id %s: %s", user_id, error, exc_info=True)
        await mensaje_carga.edit_text("⚠️ No se pudo leer el archivo. Comprueba que es válido.")
    finally:
        if ruta and os.path.exists(ruta):
            os.remove(ruta)


class _LimiteExcedidoError(Exception):
    pass


def _validar_limites(df, nombre: str, n_hojas: int) -> None:
    if len(df) > MAX_FILAS:
        raise _LimiteExcedidoError(
            f"El archivo tiene {len(df):,} filas. El límite es {MAX_FILAS:,}."
        )
    if len(df.columns) > MAX_COLUMNAS:
        raise _LimiteExcedidoError(
            f"El archivo tiene {len(df.columns)} columnas. El límite es {MAX_COLUMNAS}."
        )


async def _procesar_dataframe(update, user_id, df, nombre, mensaje_carga, errores):
    """Envía resumen, guarda contexto y lanza el gráfico con botones de tipo."""
    # Si ya hay un df activo, guardarlo como secundario antes de reemplazarlo
    df_anterior = obtener_df(user_id)
    if df_anterior is not None:
        meta_anterior = obtener_meta(user_id)
        nombre_anterior = meta_anterior["nombre"] if meta_anterior else "archivo anterior"
        guardar_df_secundario(user_id, df_anterior, nombre_anterior)

    resumen  = resumir(df, nombre, errores)
    contexto = construir_contexto(df, nombre)
    guardar_contexto(user_id, contexto)
    guardar_df(user_id, df)
    guardar_meta(user_id, nombre)
    await mensaje_carga.edit_text(resumen, parse_mode="Markdown")

    # Avisar si ahora hay dos archivos combinables
    if df_anterior is not None:
        meta_anterior = obtener_meta(user_id)
        nombre_anterior = meta_anterior["nombre"] if meta_anterior else "archivo anterior"
        cols_comunes = [c for c in df_anterior.columns if c in df.columns]
        if cols_comunes:
            sugerencia = f"une por {cols_comunes[0]}"
            await update.message.reply_text(
                f"💡 Tienes dos archivos en memoria:\n"
                f"· *{nombre_anterior}* (anterior)\n"
                f"· *{nombre}* (nuevo)\n\n"
                f"Puedes combinarlos, por ejemplo: «{sugerencia}»",
                parse_mode="Markdown",
            )

    guardar_datos_grafico(user_id, df, nombre)

    try:
        buffer = await asyncio.to_thread(generar_grafico, df, nombre, "barras")
        if buffer:
            teclado = InlineKeyboardMarkup([[
                InlineKeyboardButton("📈 Líneas",   callback_data="chart_lineas"),
                InlineKeyboardButton("🥧 Sectores", callback_data="chart_sectores"),
            ]])
            await update.message.reply_photo(
                photo=buffer,
                caption="📊 Gráfico de barras. ¿Cambiar tipo?",
                reply_markup=teclado,
            )
    except Exception as error_grafico:
        logger.warning("No se pudo generar el gráfico para '%s': %s", nombre, error_grafico)


async def callback_sheet(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query   = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    indice  = int(query.data.replace("sheet_", ""))

    resultado = obtener_hoja(user_id, indice)
    if not resultado:
        await query.edit_message_text("⚠️ No se encontraron los datos de la hoja.")
        return

    nombre_hoja, df = resultado
    hojas = listar_hojas(user_id)

    resumen  = resumir(df, f"hoja '{nombre_hoja}'")
    contexto = construir_contexto(df, f"hoja '{nombre_hoja}'")
    guardar_contexto(user_id, contexto)
    guardar_df(user_id, df)
    guardar_datos_grafico(user_id, df, nombre_hoja)
    guardar_meta(user_id, nombre_hoja, hoja=nombre_hoja)

    await query.edit_message_text(resumen, parse_mode="Markdown")

    otras = [n for i, n in enumerate(hojas) if i != indice]
    if otras:
        botones = [
            InlineKeyboardButton(f"📋 {n[:20]}", callback_data=f"sheet_{hojas.index(n)}")
            for n in otras
        ]
        filas = [botones[i:i+2] for i in range(0, len(botones), 2)]
        await query.message.reply_text("¿Quieres cambiar a otra hoja?",
                                       reply_markup=InlineKeyboardMarkup(filas))

    try:
        buffer = await asyncio.to_thread(generar_grafico, df, nombre_hoja, "barras")
        if buffer:
            teclado = InlineKeyboardMarkup([[
                InlineKeyboardButton("📈 Líneas",   callback_data="chart_lineas"),
                InlineKeyboardButton("🥧 Sectores", callback_data="chart_sectores"),
            ]])
            await query.message.reply_photo(
                photo=buffer,
                caption=f"📊 Gráfico de *{nombre_hoja}*. ¿Cambiar tipo?",
                reply_markup=teclado,
                parse_mode="Markdown",
            )
    except Exception as error_grafico:
        logger.warning("Error generando gráfico de hoja '%s': %s", nombre_hoja, error_grafico)


async def callback_chart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query   = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    tipo    = query.data.replace("chart_", "")

    datos = obtener_datos_grafico(user_id)
    if not datos:
        await query.edit_message_caption(caption="⚠️ Ya no tengo los datos. Vuelve a subir el archivo.")
        return

    try:
        import pandas as pd
        df     = pd.DataFrame(datos["df"])
        buffer = await asyncio.to_thread(generar_grafico, df, datos["nombre"], tipo)
        if not buffer:
            await query.answer("No hay datos numéricos para graficar.", show_alert=True)
            return

        nombres_tipo = {"lineas": "líneas", "sectores": "sectores", "barras": "barras"}
        otros_tipos  = [t for t in ["barras", "lineas", "sectores"] if t != tipo]
        iconos       = {"barras": "📊", "lineas": "📈", "sectores": "🥧"}
        teclado = InlineKeyboardMarkup([[
            InlineKeyboardButton(f"{iconos[t]} {nombres_tipo[t].capitalize()}",
                                 callback_data=f"chart_{t}")
            for t in otros_tipos
        ]])
        await query.message.reply_photo(
            photo=buffer,
            caption=f"📊 Gráfico de {nombres_tipo.get(tipo, tipo)}. ¿Cambiar tipo?",
            reply_markup=teclado,
        )
        await query.edit_message_caption(caption=f"📊 Gráfico de {nombres_tipo.get(tipo, tipo)} generado.")
    except Exception as error:
        logger.error("Error regenerando gráfico tipo '%s': %s", tipo, error, exc_info=True)
        await query.answer("No se pudo generar el gráfico.", show_alert=True)
