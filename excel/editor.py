"""Editor de DataFrames y exportador a .xlsx.

Aplica operaciones de modificación sobre el archivo del usuario y devuelve
el resultado como un nuevo .xlsx listo para descargar.

Operaciones: añadir_columna, ordenar, eliminar_duplicados, filtrar_exportar,
             rellenar_nulos, renombrar_columna, eliminar_columna, formato_condicional.
"""
import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

_COLORES_HEX = {
    "rojo":     "FF6B6B",
    "verde":    "70AD47",
    "amarillo": "FFD966",
    "naranja":  "F4B942",
    "azul":     "5B9BD5",
}

_OPERADORES_ARITMETICOS = {"+", "-", "*", "/"}

_COMPARADORES = {
    "<":  lambda a, b: a < b,
    ">":  lambda a, b: a > b,
    "<=": lambda a, b: a <= b,
    ">=": lambda a, b: a >= b,
    "==": lambda a, b: a == b,
    "!=": lambda a, b: a != b,
}


class EditorError(Exception):
    """Error controlado al aplicar una operación de edición."""


# ── Dispatcher principal ──────────────────────────────────────────────────────

def aplicar_edicion(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str, dict | None]:
    """Aplica una operación de edición al DataFrame.

    Devuelve (df_modificado, descripcion, extras).
    extras puede contener instrucciones de formato condicional para exportar_xlsx.
    """
    tipo = str(op.get("op", "")).lower().strip()
    df = df.copy()

    if tipo == "añadir_columna":
        df, desc = _añadir_columna(df, op)
        return df, desc, None

    elif tipo == "ordenar":
        df, desc = _ordenar(df, op)
        return df, desc, None

    elif tipo == "eliminar_duplicados":
        df, desc = _eliminar_duplicados(df, op)
        return df, desc, None

    elif tipo == "filtrar_exportar":
        df, desc = _filtrar_exportar(df, op)
        return df, desc, None

    elif tipo == "rellenar_nulos":
        df, desc = _rellenar_nulos(df, op)
        return df, desc, None

    elif tipo == "renombrar_columna":
        df, desc = _renombrar_columna(df, op)
        return df, desc, None

    elif tipo == "eliminar_columna":
        df, desc = _eliminar_columna(df, op)
        return df, desc, None

    elif tipo == "formato_condicional":
        desc = _describir_formato(op)
        return df, desc, op   # df sin cambios; extras lleva las instrucciones

    else:
        raise EditorError(f"Operación no reconocida: '{tipo}'")


# ── Operaciones ───────────────────────────────────────────────────────────────

def _añadir_columna(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    nombre = op.get("nombre")
    if not nombre:
        raise EditorError("Falta el nombre de la columna nueva ('nombre').")

    col1     = op.get("col1")
    col2     = op.get("col2")
    operador = op.get("operador")
    valor_fijo = op.get("valor_fijo")

    if not col1 or not operador:
        raise EditorError("Faltan 'col1' y 'operador' para añadir_columna.")
    _validar_col(df, col1)
    if operador not in _OPERADORES_ARITMETICOS:
        raise EditorError(f"Operador aritmético no permitido: '{operador}'. Usa: + - * /")

    s1 = pd.to_numeric(df[col1], errors="coerce")

    if col2:
        _validar_col(df, col2)
        s2 = pd.to_numeric(df[col2], errors="coerce")
    elif valor_fijo is not None:
        s2 = float(valor_fijo)
    else:
        raise EditorError("Falta 'col2' o 'valor_fijo' para la operación.")

    if operador == "+":
        df[nombre] = s1 + s2
    elif operador == "-":
        df[nombre] = s1 - s2
    elif operador == "*":
        df[nombre] = s1 * s2
    elif operador == "/":
        divisor = s2 if isinstance(s2, pd.Series) else s2
        df[nombre] = s1 / divisor

    ref = col2 if col2 else str(valor_fijo)
    return df, f"Columna '{nombre}' añadida ({col1} {operador} {ref})"


def _ordenar(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    col = op.get("col")
    _validar_col(df, col)
    asc = str(op.get("orden", "asc")).lower() == "asc"
    df = df.sort_values(col, ascending=asc).reset_index(drop=True)
    return df, f"Ordenado por '{col}' {'↑ ascendente' if asc else '↓ descendente'}"


def _eliminar_duplicados(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    columnas = op.get("columnas") or None
    if columnas:
        for c in columnas:
            _validar_col(df, c)
    antes = len(df)
    df = df.drop_duplicates(subset=columnas).reset_index(drop=True)
    eliminadas = antes - len(df)
    return df, f"{eliminadas} fila(s) duplicada(s) eliminada(s)"


def _filtrar_exportar(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    from excel.query_engine import _aplicar_filtros
    filtros = op.get("filtros", [])
    if not filtros:
        raise EditorError("Falta 'filtros' para filtrar_exportar.")
    df = _aplicar_filtros(df, filtros).reset_index(drop=True)
    return df, f"{len(df)} filas tras aplicar los filtros"


def _rellenar_nulos(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    col    = op.get("col")
    metodo = str(op.get("metodo", "media")).lower()
    cols   = [col] if col else [c for c in df.columns if df[c].isnull().any()]

    if col:
        _validar_col(df, col)

    rellenadas = 0
    for c in cols:
        nulos_antes = int(df[c].isnull().sum())
        if nulos_antes == 0:
            continue
        serie_num = pd.to_numeric(df[c], errors="coerce")
        if metodo == "media":
            df[c] = df[c].fillna(serie_num.mean())
        elif metodo == "mediana":
            df[c] = df[c].fillna(serie_num.median())
        elif metodo == "cero":
            df[c] = df[c].fillna(0)
        elif metodo == "anterior":
            df[c] = df[c].ffill()
        elif metodo == "siguiente":
            df[c] = df[c].bfill()
        elif metodo == "valor":
            df[c] = df[c].fillna(op.get("valor", 0))
        else:
            raise EditorError(f"Método de relleno no reconocido: '{metodo}'")
        rellenadas += nulos_antes

    return df, f"{rellenadas} celda(s) vacía(s) rellenada(s) (método: {metodo})"


def _renombrar_columna(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    mapeo = op.get("columnas", {})
    if not mapeo:
        raise EditorError("Falta 'columnas' {{nombre_actual: nombre_nuevo}}.")
    for c in mapeo:
        _validar_col(df, c)
    df = df.rename(columns=mapeo)
    cambios = ", ".join(f"'{v}' → '{n}'" for v, n in mapeo.items())
    return df, f"Columnas renombradas: {cambios}"


def _eliminar_columna(df: pd.DataFrame, op: dict) -> tuple[pd.DataFrame, str]:
    columnas = op.get("columnas", [])
    if not columnas:
        raise EditorError("Falta 'columnas' (lista de nombres a eliminar).")
    for c in columnas:
        _validar_col(df, c)
    df = df.drop(columns=columnas)
    return df, f"Eliminada(s): {', '.join(columnas)}"


def _describir_formato(op: dict) -> str:
    col      = op.get("col", "?")
    cond     = op.get("condicion", "?")
    valor    = op.get("valor", "?")
    color    = op.get("color", "rojo")
    return f"Formato condicional: '{col}' {cond} {valor} → {color}"


# ── Exportador ────────────────────────────────────────────────────────────────

def exportar_xlsx(df: pd.DataFrame, nombre_base: str,
                  descripcion: str = "",
                  formato_condicional: dict | None = None) -> tuple[io.BytesIO, str]:
    """Convierte el DataFrame modificado a .xlsx con cabeceras formateadas.

    Devuelve (buffer, nombre_archivo).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos"

    # Cabeceras
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=ci, value=str(col))
        c.font      = Font(bold=True, color="FFFFFF")
        c.fill      = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center")

    # Datos (con filas alternas ligeramente coloreadas)
    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, valor in enumerate(row, 1):
            val = None if (isinstance(valor, float) and pd.isna(valor)) else valor
            c = ws.cell(row=ri, column=ci, value=val)
            if ri % 2 == 0:
                c.fill = PatternFill("solid", fgColor="EBF3FB")

    # Formato condicional (coloreado de celdas)
    if formato_condicional:
        _aplicar_formato_condicional(ws, df, formato_condicional)

    # Ajustar anchos de columna
    for col_cells in ws.columns:
        ancho = max((len(str(c.value or "")) for c in col_cells), default=8) + 3
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(ancho, 40)

    # Hoja de información si hay descripción
    if descripcion:
        ws_info = wb.create_sheet("Modificación")
        ws_info["A1"] = "Operación aplicada"
        ws_info["A1"].font = Font(bold=True)
        ws_info["A2"] = descripcion
        ws_info.column_dimensions["A"].width = max(len(descripcion) + 4, 30)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    base = nombre_base.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
    return buf, f"{base}_modificado.xlsx"


def _aplicar_formato_condicional(ws, df: pd.DataFrame, op: dict) -> None:
    col      = op.get("col")
    condicion = op.get("condicion", "<")
    valor_ref = op.get("valor")
    color_key = str(op.get("color", "rojo")).lower()
    color_hex = _COLORES_HEX.get(color_key, "FF6B6B")

    if col not in df.columns or valor_ref is None:
        return

    fn_cond = _COMPARADORES.get(condicion)
    if not fn_cond:
        return

    col_idx = list(df.columns).index(col) + 1
    try:
        v_ref = float(valor_ref)
    except (TypeError, ValueError):
        return

    fill_color = PatternFill("solid", fgColor=color_hex)
    for ri in range(2, len(df) + 2):
        celda = ws.cell(row=ri, column=col_idx)
        try:
            if fn_cond(float(celda.value), v_ref):
                celda.fill = fill_color
        except (TypeError, ValueError):
            pass


# ── Validación ────────────────────────────────────────────────────────────────

def _validar_col(df: pd.DataFrame, col: str | None) -> None:
    if not col:
        raise EditorError("Falta especificar la columna.")
    if col not in df.columns:
        cols = ", ".join(f"'{c}'" for c in df.columns)
        raise EditorError(f"Columna '{col}' no existe. Disponibles: {cols}")
