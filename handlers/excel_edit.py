import asyncio
import json as _json
import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from services.llm import (LLMError, extraer_operacion_edicion, extraer_peticion_grafico,
                           extraer_formula, extraer_query_dsl, extraer_operacion_combinar,
                           extraer_operaciones_macro)
from utils.df_context import guardar_df, obtener_nombre_secundario, restaurar_undo, hay_undo
from utils.file_meta import obtener_meta
from utils.macros import (guardar_macro as _db_guardar_macro, obtener_macro,
                           listar_macros as _db_listar_macros, borrar_macro as _db_borrar_macro)
from excel.editor import aplicar_edicion, exportar_xlsx, combinar_dataframes, EditorError
from excel.query_engine import ejecutar_query, formatear_resultado
from excel.charts import generar_grafico_personalizado
from handlers.intent_patterns import _CLAVE_ACLARACION
from handlers.bot_helpers import _pedir_aclaracion, _responder_con_llm

logger = logging.getLogger(__name__)

_OPS_DESTRUCTIVAS = {"eliminar_columna", "eliminar_duplicados", "filtrar_exportar"}
_NOMBRES_OP = {
    "eliminar_columna":    "Eliminar columna",
    "eliminar_duplicados": "Eliminar duplicados",
    "filtrar_exportar":    "Filtrar y exportar",
}


async def _intentar_edicion(update: Update, user_id: int, df, pregunta: str,
                             context=None) -> None:
    """Extrae la pipeline de operaciones, las aplica en orden y envía el archivo modificado."""
    mensaje_carga = await update.message.reply_text("⏳ Aplicando modificación...")
    try:
        resultado = await asyncio.to_thread(extraer_operacion_edicion, df, pregunta)

        if resultado is None:
            await mensaje_carga.delete()
            await _responder_con_llm(update, user_id, pregunta, mensaje_carga=None)
            return

        if isinstance(resultado, dict) and resultado.get(_CLAVE_ACLARACION):
            logger.info("Aclaración solicitada (edición) para user_id %s", user_id)
            await mensaje_carga.delete()
            await _pedir_aclaracion(update, context, user_id, "edicion", resultado)
            return

        ops = resultado if isinstance(resultado, list) else [resultado]

        for op in ops:
            if isinstance(op, dict) and op.get(_CLAVE_ACLARACION):
                await mensaje_carga.delete()
                await _pedir_aclaracion(update, context, user_id, "edicion", op)
                return

        if (len(ops) == 1 and ops[0].get("op") in _OPS_DESTRUCTIVAS
                and context is not None):
            await mensaje_carga.delete()
            await _pedir_confirmacion(update, context, user_id, df, ops[0])
            return

        guardar_df(user_id, df)

        meta = obtener_meta(user_id)
        nombre_base = meta["nombre"] if meta else "archivo"

        df_actual = df.copy()
        descripciones: list[str] = []
        graficos_extra: list[tuple] = []
        resultados_query: list[str] = []

        for op in ops:
            nombre_op = op.get("op", "")

            if nombre_op == "grafico":
                try:
                    params = await asyncio.to_thread(extraer_peticion_grafico, df_actual, pregunta)
                    if params:
                        buf_img, titulo = await asyncio.to_thread(
                            generar_grafico_personalizado,
                            df_actual,
                            params.get("col_y"),
                            params.get("col_x"),
                            params.get("tipo", "barras"),
                            params.get("agregar"),
                        )
                        graficos_extra.append((buf_img, titulo))
                        descripciones.append(f"Gráfico: {titulo}")
                except Exception as e_graf:
                    logger.warning("Gráfico en pipeline falló para user_id %s: %s", user_id, e_graf)
                continue

            if nombre_op == "tabla_dinamica":
                descripciones.append("Tabla dinámica preparada")
                continue

            if nombre_op == "formula":
                try:
                    params = await asyncio.to_thread(extraer_formula, df_actual, pregunta)
                    if params:
                        formula_tmpl = params.get("formula", "")
                        col_nueva    = params.get("col_nueva", "Fórmula")
                        formulas_col = [
                            formula_tmpl.replace("{row}", str(row))
                            for row in range(2, len(df_actual) + 2)
                        ]
                        df_actual = df_actual.copy()
                        df_actual[col_nueva] = formulas_col
                        descripciones.append(f"Columna '{col_nueva}' con fórmula Excel")
                except Exception as e_form:
                    logger.warning("Fórmula en pipeline bot user_id %s: %s", user_id, e_form)
                continue

            if nombre_op == "query":
                pregunta_q = op.get("pregunta", "")
                if pregunta_q:
                    try:
                        dsl_q = await asyncio.to_thread(extraer_query_dsl, df_actual, pregunta_q)
                        es_aclaracion = isinstance(dsl_q, dict) and dsl_q.get("aclaracion_necesaria")
                        if dsl_q and not es_aclaracion:
                            ops_q = dsl_q if isinstance(dsl_q, list) else [dsl_q]
                            partes_q: list[str] = []
                            for q in ops_q:
                                resultado_q, descripcion_q = await asyncio.to_thread(
                                    ejecutar_query, df_actual, q
                                )
                                partes_q.append(formatear_resultado(resultado_q, descripcion_q))
                            texto_q = "\n\n".join(partes_q)
                            resultados_query.append(f"*{pregunta_q}*\n{texto_q}")
                    except Exception as e_q:
                        logger.warning("Query inline bot user_id %s: %s", user_id, e_q)
                continue

            try:
                df_actual, descripcion, extras = await asyncio.to_thread(
                    aplicar_edicion, df_actual, op, False
                )
                descripciones.append(descripcion)
            except EditorError as error:
                logger.warning("Op '%s' falló en pipeline bot user_id %s: %s",
                               nombre_op, user_id, error)

        guardar_df(user_id, df_actual)

        resumen = "; ".join(descripciones) if descripciones else None
        if resumen:
            buf, nombre_archivo = await asyncio.to_thread(
                exportar_xlsx, df_actual, nombre_base, resumen, None
            )
            caption = f"✅ {resumen}\n\nEl archivo incluye los cambios aplicados."
            await update.message.reply_document(document=buf, filename=nombre_archivo, caption=caption)

        if resultados_query:
            texto_queries = "\n\n".join(resultados_query)
            await update.message.reply_text(
                f"📊 *Resultados:*\n\n{texto_queries}", parse_mode="Markdown"
            )

        for buf_img, titulo in graficos_extra:
            await update.message.reply_photo(photo=buf_img, caption=f"📊 {titulo}")

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


async def _intentar_combinar(update: Update, user_id: int,
                              df_a, df_b, pregunta: str) -> None:
    """Combina los dos DataFrames en memoria según la petición del usuario."""
    nombre_a = obtener_nombre_secundario(user_id)
    meta_b   = obtener_meta(user_id)
    nombre_b = meta_b["nombre"] if meta_b else "archivo"

    mensaje_carga = await update.message.reply_text("⏳ Combinando archivos...")
    try:
        op = await asyncio.to_thread(extraer_operacion_combinar, df_a, df_b, pregunta)
        df_result, descripcion = await asyncio.to_thread(combinar_dataframes, df_a, df_b, op)

        guardar_df(user_id, df_result)

        buf, nombre_archivo = await asyncio.to_thread(
            exportar_xlsx, df_result, "combinado", descripcion
        )
        caption = (
            f"✅ {descripcion}\n\n"
            f"Archivo A: *{nombre_a}*\n"
            f"Archivo B: *{nombre_b}*"
        )
        await update.message.reply_document(
            document=buf, filename=nombre_archivo, caption=caption, parse_mode="Markdown"
        )
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


async def _deshacer_operacion(update: Update, user_id: int) -> None:
    """Restaura el DataFrame al estado anterior a la última edición."""
    if not hay_undo(user_id):
        await update.message.reply_text("↩️ No hay ninguna operación anterior que deshacer.")
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


async def _pedir_confirmacion(update: Update, context, user_id: int, df, op: dict) -> None:
    """Muestra un botón Sí/No antes de ejecutar una operación destructiva."""
    nombre_op = _NOMBRES_OP.get(op.get("op", ""), op.get("op", "operación"))

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

    context.user_data["op_pendiente"] = _json.dumps(op, ensure_ascii=False)

    teclado = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Sí, aplicar", callback_data="confirmar_op_si"),
            InlineKeyboardButton("❌ Cancelar",     callback_data="confirmar_op_no"),
        ]
    ])
    await update.message.reply_text(
        f"⚠️ *{nombre_op}*\n{detalles}\n\n¿Confirmas la operación?",
        parse_mode="Markdown",
        reply_markup=teclado,
    )


async def callback_confirmacion(update, context) -> None:
    """Ejecuta o cancela la operación destructiva según la respuesta del usuario."""
    from utils.df_context import obtener_df

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


async def _guardar_macro(update: Update, user_id: int, nombre: str, pregunta: str) -> None:
    """Interpreta la descripción y guarda la macro en SQLite."""
    mensaje_carga = await update.message.reply_text(f"⏳ Definiendo macro '{nombre}'...")
    try:
        ops = await asyncio.to_thread(extraer_operaciones_macro, pregunta)
        if not ops:
            await mensaje_carga.edit_text(
                f"⚠️ No pude interpretar las operaciones de la macro '{nombre}'.\n"
                "Describe qué pasos quieres que haga, por ejemplo:\n"
                "_«Guarda una macro llamada limpieza que normalice el texto, "
                "elimine duplicados y ordene por Fecha»_",
                parse_mode="Markdown",
            )
            return

        _db_guardar_macro(user_id, nombre, ops, descripcion=pregunta)
        pasos = "\n".join(f"  {i+1}. {op.get('op', '?')}" for i, op in enumerate(ops))
        await mensaje_carga.edit_text(
            f"✅ Macro *{nombre}* guardada con {len(ops)} paso(s):\n{pasos}\n\n"
            f"Úsala con: _«aplica la macro {nombre}»_",
            parse_mode="Markdown",
        )
    except Exception as error:
        logger.error("Error guardando macro para user_id %s: %s", user_id, error, exc_info=True)
        await mensaje_carga.edit_text("⚠️ No se pudo guardar la macro. Inténtalo de nuevo.")


async def _ejecutar_macro(update: Update, user_id: int, df, nombre: str) -> None:
    """Ejecuta todas las operaciones de una macro guardada."""
    ops = obtener_macro(user_id, nombre)
    if ops is None:
        macros = _db_listar_macros(user_id)
        if macros:
            nombres = ", ".join(f"*{m['nombre']}*" for m in macros)
            await update.message.reply_text(
                f"⚠️ No encontré la macro '{nombre}'.\n"
                f"Tus macros disponibles: {nombres}",
                parse_mode="Markdown",
            )
        else:
            await update.message.reply_text(
                f"⚠️ No encontré la macro '{nombre}' y no tienes ninguna guardada.\n"
                "Crea una con: _«Guarda una macro llamada X que haga Y»_",
                parse_mode="Markdown",
            )
        return

    mensaje_carga = await update.message.reply_text(
        f"⏳ Ejecutando macro *{nombre}* ({len(ops)} pasos)...", parse_mode="Markdown"
    )
    try:
        df_actual = df.copy()
        descripciones = []
        for op in ops:
            df_actual, desc, _ = await asyncio.to_thread(aplicar_edicion, df_actual, op)
            descripciones.append(desc)

        meta = obtener_meta(user_id)
        nombre_base = meta["nombre"] if meta else "archivo"
        buf, nombre_archivo = await asyncio.to_thread(
            exportar_xlsx, df_actual, nombre_base, f"Macro '{nombre}' aplicada"
        )
        guardar_df(user_id, df_actual)

        pasos_txt = "\n".join(f"  {i+1}. {d}" for i, d in enumerate(descripciones))
        await update.message.reply_document(
            document=buf, filename=nombre_archivo,
            caption=f"✅ Macro *{nombre}* completada:\n{pasos_txt}",
            parse_mode="Markdown",
        )
        try:
            await mensaje_carga.delete()
        except Exception:
            pass

    except EditorError as error:
        await mensaje_carga.edit_text(f"⚠️ Error en la macro '{nombre}': {error}")
    except Exception as error:
        logger.error("Error ejecutando macro para user_id %s: %s", user_id, error, exc_info=True)
        await mensaje_carga.edit_text("⚠️ No se pudo ejecutar la macro. Inténtalo de nuevo.")


async def _listar_macros(update: Update, user_id: int) -> None:
    macros = _db_listar_macros(user_id)
    if not macros:
        await update.message.reply_text(
            "No tienes ninguna macro guardada.\n"
            "Crea una con: _«Guarda una macro llamada limpieza que normalice el texto»_",
            parse_mode="Markdown",
        )
        return
    lineas = ["📋 *Tus macros guardadas:*\n"]
    for m in macros:
        ops = obtener_macro(user_id, m["nombre"])
        n_pasos = len(ops) if ops else 0
        lineas.append(f"• *{m['nombre']}* — {n_pasos} paso(s)")
    lineas.append("\n_Usa «aplica la macro X» para ejecutar una._")
    await update.message.reply_text("\n".join(lineas), parse_mode="Markdown")


async def _borrar_macro(update: Update, user_id: int, nombre: str) -> None:
    eliminada = _db_borrar_macro(user_id, nombre)
    if eliminada:
        await update.message.reply_text(f"🗑️ Macro *{nombre}* eliminada.", parse_mode="Markdown")
    else:
        await update.message.reply_text(
            f"⚠️ No encontré ninguna macro llamada '{nombre}'."
        )
