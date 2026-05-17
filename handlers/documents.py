import logging
import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from utils.auth import solo_autorizados
from utils.excel_context import guardar_contexto
from utils.chart_context import guardar_datos_grafico, obtener_datos_grafico
from utils.sheet_context import guardar_hojas, obtener_hoja, listar_hojas
from excel.reader import leer_excel_hojas, leer_csv
from excel.analyzer import resumir, resumir_hojas, construir_contexto, detectar_errores_xlsx
from excel.charts import generar_grafico

logger = logging.getLogger(__name__)

TAMANIO_MAXIMO_MB = 5
DIRECTORIO_TEMP = os.path.join(os.path.dirname(__file__), "..", "data", "temp")

_EXTENSIONES_EXCEL = (".xlsx", ".xls")
_EXTENSIONES_CSV   = (".csv",)


@solo_autorizados
async def recibir_documento(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    documento = update.message.document
    user_id = update.effective_user.id
    nombre = documento.file_name or ""
    es_excel = nombre.lower().endswith(_EXTENSIONES_EXCEL)
    es_csv   = nombre.lower().endswith(_EXTENSIONES_CSV)

    if not es_excel and not es_csv:
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

        if es_csv:
            await _procesar_dataframe(update, user_id, leer_csv(ruta), nombre, mensaje_carga, errores=[])
        else:
            sheets = leer_excel_hojas(ruta)
            errores = detectar_errores_xlsx(ruta)

            if len(sheets) > 1:
                # Multi-hoja: guardar todas y mostrar selector
                guardar_hojas(user_id, sheets)
                resumen = resumir_hojas(sheets, nombre)
                nombres = list(sheets.keys())
                botones = [
                    InlineKeyboardButton(f"📋 {n[:20]}", callback_data=f"sheet_{i}")
                    for i, n in enumerate(nombres)
                ]
                # Filas de 2 botones
                filas = [botones[i:i+2] for i in range(0, len(botones), 2)]
                teclado = InlineKeyboardMarkup(filas)
                await mensaje_carga.edit_text(resumen, parse_mode="Markdown", reply_markup=teclado)
            else:
                # Una sola hoja: procesar directamente
                df = list(sheets.values())[0]
                await _procesar_dataframe(update, user_id, df, nombre, mensaje_carga, errores)

    except Exception as error:
        logger.error("Error procesando archivo para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo leer el archivo. Comprueba que es válido.")
    finally:
        if ruta and os.path.exists(ruta):
            os.remove(ruta)


async def _procesar_dataframe(update, user_id, df, nombre, mensaje_carga, errores):
    """Envía resumen, guarda contexto y lanza el gráfico con botones de tipo."""
    resumen = resumir(df, nombre, errores)
    contexto = construir_contexto(df, nombre)
    guardar_contexto(user_id, contexto)
    await mensaje_carga.edit_text(resumen, parse_mode="Markdown")

    # Guardar datos para poder regenerar el gráfico
    guardar_datos_grafico(user_id, df, nombre)

    # Generar gráfico de barras por defecto
    try:
        buffer = generar_grafico(df, nombre, tipo="barras")
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
    """Cambia la hoja activa cuando el usuario pulsa un botón de selección de hoja."""
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    indice = int(query.data.replace("sheet_", ""))

    resultado = obtener_hoja(user_id, indice)
    if not resultado:
        await query.edit_message_text("⚠️ No se encontraron los datos de la hoja.")
        return

    nombre_hoja, df = resultado
    hojas = listar_hojas(user_id)
    nombre_archivo = f"hoja '{nombre_hoja}'"

    resumen = resumir(df, nombre_archivo)
    contexto = construir_contexto(df, nombre_archivo)
    guardar_contexto(user_id, contexto)
    guardar_datos_grafico(user_id, df, nombre_hoja)

    await query.edit_message_text(resumen, parse_mode="Markdown")

    # Ofrecer cambiar a otra hoja si hay más
    otras = [n for i, n in enumerate(hojas) if i != indice]
    if otras:
        botones = [
            InlineKeyboardButton(f"📋 {n[:20]}", callback_data=f"sheet_{hojas.index(n)}")
            for n in otras
        ]
        filas = [botones[i:i+2] for i in range(0, len(botones), 2)]
        await query.message.reply_text(
            "¿Quieres cambiar a otra hoja?",
            reply_markup=InlineKeyboardMarkup(filas)
        )

    # Gráfico de la hoja seleccionada
    try:
        buffer = generar_grafico(df, nombre_hoja, tipo="barras")
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
    """Regenera el gráfico con el tipo de visualización elegido."""
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    tipo = query.data.replace("chart_", "")

    datos = obtener_datos_grafico(user_id)
    if not datos:
        await query.edit_message_caption(caption="⚠️ Ya no tengo los datos del archivo. Vuelve a subirlo.")
        return

    try:
        import pandas as pd
        df = pd.DataFrame(datos["df"])
        buffer = generar_grafico(df, datos["nombre"], tipo=tipo)
        if not buffer:
            await query.answer("No hay datos numéricos para graficar.", show_alert=True)
            return

        nombres_tipo = {"lineas": "líneas", "sectores": "sectores", "barras": "barras"}
        otros_tipos = [t for t in ["barras", "lineas", "sectores"] if t != tipo]
        teclado = InlineKeyboardMarkup([[
            InlineKeyboardButton(
                f"{'📊' if t == 'barras' else '📈' if t == 'lineas' else '🥧'} {nombres_tipo[t].capitalize()}",
                callback_data=f"chart_{t}"
            )
            for t in otros_tipos
        ]])
        await query.message.reply_photo(
            photo=buffer,
            caption=f"📊 Gráfico de {nombres_tipo.get(tipo, tipo)}. ¿Cambiar tipo?",
            reply_markup=teclado,
        )
        # Quitar los botones del mensaje original
        await query.edit_message_caption(
            caption=f"📊 Gráfico de {nombres_tipo.get(tipo, tipo)} generado."
        )
    except Exception as error:
        logger.error("Error regenerando gráfico tipo '%s': %s", tipo, error)
        await query.answer("No se pudo generar el gráfico.", show_alert=True)
