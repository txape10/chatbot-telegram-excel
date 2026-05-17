import logging
import os
from telegram import Update
from telegram.ext import ContextTypes
from utils.auth import solo_autorizados
from utils.excel_context import guardar_contexto
from excel.reader import leer_excel
from excel.analyzer import resumir, construir_contexto
from excel.charts import generar_grafico

logger = logging.getLogger(__name__)

TAMANIO_MAXIMO_MB = 5
DIRECTORIO_TEMP = os.path.join(os.path.dirname(__file__), "..", "data", "temp")


@solo_autorizados
async def recibir_documento(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    documento = update.message.document
    user_id = update.effective_user.id

    # Validar que sea un archivo Excel
    if not documento.file_name.endswith((".xlsx", ".xls")):
        await update.message.reply_text("⚠️ Solo acepto archivos Excel (.xlsx o .xls).")
        return

    # Validar tamaño
    if documento.file_size > TAMANIO_MAXIMO_MB * 1024 * 1024:
        await update.message.reply_text(f"⚠️ El archivo supera el límite de {TAMANIO_MAXIMO_MB} MB.")
        return

    mensaje_carga = await update.message.reply_text("⏳ Leyendo el archivo...")
    ruta = None

    try:
        os.makedirs(DIRECTORIO_TEMP, exist_ok=True)
        ruta = os.path.join(DIRECTORIO_TEMP, f"{user_id}_{documento.file_name}")

        archivo = await documento.get_file()
        await archivo.download_to_drive(ruta)

        df = leer_excel(ruta)
        resumen = resumir(df, documento.file_name)
        contexto = construir_contexto(df, documento.file_name)
        guardar_contexto(user_id, contexto)

        await mensaje_carga.edit_text(resumen, parse_mode="Markdown")

        # Enviar gráfico — try independiente para no afectar al resumen
        try:
            buffer_grafico = generar_grafico(df, documento.file_name)
            if buffer_grafico:
                await update.message.reply_photo(
                    photo=buffer_grafico,
                    caption="📊 Gráfico generado automáticamente con los datos del archivo."
                )
            else:
                logger.info("Sin columnas numéricas para graficar en '%s'", documento.file_name)
        except Exception as error_grafico:
            logger.warning("No se pudo generar el gráfico para '%s': %s", documento.file_name, error_grafico)

    except Exception as error:
        logger.error("Error procesando Excel para user_id %s: %s", user_id, error)
        await mensaje_carga.edit_text("⚠️ No se pudo leer el archivo. Comprueba que es un Excel válido.")
    finally:
        if ruta and os.path.exists(ruta):
            os.remove(ruta)
