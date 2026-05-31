"""Motor de consultas DSL sobre DataFrames.

Operaciones soportadas: filtrar, contar, suma, promedio, max, min, agrupar, ordenar, top_n.
No ejecuta código arbitrario — solo un conjunto cerrado de operaciones seguras.
"""
import pandas as pd
from typing import Any

import re

_OPS_FILTRO = {"==", "!=", ">", ">=", "<", "<=", "contiene", "no_contiene", "empieza_por"}

# Valores especiales en filtros numéricos: se calculan sobre la columna del df completo
_VALS_ESTADISTICOS = {"media": "mean", "mediana": "median", "max": "max", "min": "min"}

# Filas de resumen que genera añadir_fila_total u otras fuentes externas
_RE_FILA_RESUMEN = re.compile(
    r"^\s*(total|subtotal|grand\s+total|total\s+general)\b", re.IGNORECASE
)


def _limpiar_filas_resumen(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina filas de resumen/total de la primera columna para no contaminar estadísticas."""
    if df.empty:
        return df
    primera = df.columns[0]
    mask = df[primera].astype(str).str.match(_RE_FILA_RESUMEN, na=False)
    return df[~mask].copy() if mask.any() else df


class QueryError(Exception):
    """Error controlado al ejecutar una query DSL."""


def ejecutar_query(df: pd.DataFrame, query: dict) -> tuple[Any, str]:
    """Ejecuta una query DSL sobre el DataFrame.

    Devuelve (resultado, descripcion).
    resultado puede ser un DataFrame, un float o un int.
    """
    op = str(query.get("op", "")).strip().lower()
    filtros = query.get("filtros", [])
    df = _limpiar_filas_resumen(df)   # excluir filas Total/Subtotal antes de calcular
    df_t = _aplicar_filtros(df, filtros)

    if op == "filtrar":
        return df_t.reset_index(drop=True), f"{len(df_t)} filas encontradas"

    elif op == "contar":
        por = query.get("por")
        if por:
            _validar_col(df_t, por)
            res = df_t.groupby(por).size().reset_index(name="Recuento")
            res = res.sort_values("Recuento", ascending=False).reset_index(drop=True)
            return res, f"Recuento por '{por}'"
        return int(len(df_t)), "Total de filas"

    elif op in ("suma", "promedio", "max", "min"):
        col = query.get("col")
        _validar_col(df_t, col)
        por = query.get("por")
        fn_map = {"suma": "sum", "promedio": "mean", "max": "max", "min": "min"}
        fn = fn_map[op]
        etiqueta = {"suma": "Suma", "promedio": "Promedio", "max": "Máximo", "min": "Mínimo"}[op]
        if por:
            _validar_col(df_t, por)
            res = df_t.groupby(por)[col].agg(fn).reset_index()
            res.columns = [por, f"{etiqueta} de {col}"]
            res = res.sort_values(res.columns[-1], ascending=False).reset_index(drop=True)
            return res, f"{etiqueta} de '{col}' por '{por}'"
        else:
            val = float(getattr(df_t[col], fn)())
            return round(val, 2), f"{etiqueta} de '{col}'"

    elif op == "agrupar":
        por = query.get("por")
        col = query.get("col")
        agg = str(query.get("agg", "suma")).lower()
        _validar_col(df_t, por)
        fn_map = {"suma": "sum", "promedio": "mean", "contar": "count", "max": "max", "min": "min"}
        if agg not in fn_map:
            raise QueryError(f"Agregación no reconocida: '{agg}'. Usa: {', '.join(fn_map)}")
        if col:
            _validar_col(df_t, col)
            res = df_t.groupby(por)[col].agg(fn_map[agg]).reset_index()
            etiqueta = {"suma": "Suma", "promedio": "Promedio", "contar": "Recuento",
                        "max": "Máximo", "min": "Mínimo"}[agg]
            res.columns = [por, f"{etiqueta} de {col}"]
        else:
            res = df_t.groupby(por).size().reset_index(name="Recuento")
        res = res.sort_values(res.columns[-1], ascending=False).reset_index(drop=True)
        return res, f"Agrupado por '{por}'"

    elif op == "ordenar":
        col = query.get("col")
        _validar_col(df_t, col)
        asc = str(query.get("orden", "desc")).lower() == "asc"
        res = df_t.sort_values(col, ascending=asc).reset_index(drop=True)
        return res, f"Ordenado por '{col}' {'↑' if asc else '↓'}"

    elif op == "top_n":
        col = query.get("col")
        _validar_col(df_t, col)
        n = max(1, int(query.get("n", 5)))
        asc = str(query.get("orden", "desc")).lower() == "asc"
        res = df_t.sort_values(col, ascending=asc).head(n).reset_index(drop=True)
        label = "Últimos" if asc else "Top"
        return res, f"{label} {n} por '{col}'"

    elif op == "top_n_por_grupo":
        col = query.get("col")
        por = query.get("por")
        _validar_col(df_t, col)
        _validar_col(df_t, por)
        n = max(1, int(query.get("n", 3)))
        asc = str(query.get("orden", "desc")).lower() == "asc"
        # Excluir filas donde la columna de grupo es nula o vacía
        df_t = df_t[df_t[por].notna() & (df_t[por].astype(str).str.strip() != "")]
        serie_num = pd.to_numeric(df_t[col], errors="coerce")
        res = (df_t.assign(**{col: serie_num})
               .sort_values(col, ascending=asc)
               .groupby(por, sort=False)
               .head(n)
               .sort_values([por, col], ascending=[True, asc])
               .reset_index(drop=True))
        label = "Últimos" if asc else "Top"
        return res, f"{label} {n} de '{col}' por grupo '{por}'"

    else:
        raise QueryError(f"Operación no reconocida: '{op}'")


def _resolver_val_numerico(df: pd.DataFrame, col: str, val) -> float:
    """Convierte val a float. Si es un alias estadístico ('media', 'mediana', etc.)
    lo calcula sobre la columna del df completo."""
    if isinstance(val, str) and val.lower() in _VALS_ESTADISTICOS:
        fn = _VALS_ESTADISTICOS[val.lower()]
        return float(getattr(pd.to_numeric(df[col], errors="coerce"), fn)())
    return float(val)


def _aplicar_filtros(df: pd.DataFrame, filtros: list[dict]) -> pd.DataFrame:
    if not filtros:
        return df
    mascara = pd.Series([True] * len(df), index=df.index)
    for f in filtros:
        col = f.get("col")
        op  = f.get("op", "==")
        val = f.get("val")
        _validar_col(df, col)
        if op not in _OPS_FILTRO:
            raise QueryError(f"Operador no permitido: '{op}'. Válidos: {', '.join(_OPS_FILTRO)}")
        serie = df[col]
        if op == "==":
            mascara &= (serie == val)
        elif op == "!=":
            mascara &= (serie != val)
        elif op == ">":
            mascara &= (pd.to_numeric(serie, errors="coerce") > _resolver_val_numerico(df, col, val))
        elif op == ">=":
            mascara &= (pd.to_numeric(serie, errors="coerce") >= _resolver_val_numerico(df, col, val))
        elif op == "<":
            mascara &= (pd.to_numeric(serie, errors="coerce") < _resolver_val_numerico(df, col, val))
        elif op == "<=":
            mascara &= (pd.to_numeric(serie, errors="coerce") <= _resolver_val_numerico(df, col, val))
        elif op == "contiene":
            mascara &= serie.astype(str).str.contains(str(val), case=False, na=False)
        elif op == "no_contiene":
            mascara &= ~serie.astype(str).str.contains(str(val), case=False, na=False)
        elif op == "empieza_por":
            mascara &= serie.astype(str).str.startswith(str(val), na=False)
    return df[mascara].copy()


def _validar_col(df: pd.DataFrame, col: str | None) -> None:
    if not col:
        raise QueryError("Falta especificar 'col'.")
    if col not in df.columns:
        cols = ", ".join(f"'{c}'" for c in df.columns)
        raise QueryError(f"Columna '{col}' no existe. Disponibles: {cols}")


def formatear_resultado(resultado: Any, descripcion: str, max_filas: int = 15) -> str:
    """Convierte el resultado de una query a texto Markdown para Telegram."""
    if isinstance(resultado, pd.DataFrame):
        if resultado.empty:
            return f"*{descripcion}*\n\nSin resultados."
        total = len(resultado)
        muestra = resultado.head(max_filas)
        texto = f"*{descripcion}*"
        if total > max_filas:
            texto += f" (primeras {max_filas} de {total})"
        texto += f"\n```\n{muestra.to_string(index=False)}\n```"
        return texto
    else:
        return f"*{descripcion}*: `{resultado}`"
