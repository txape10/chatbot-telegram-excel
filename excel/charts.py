import io
import logging
import sys

import pandas as pd

logger = logging.getLogger(__name__)


def _importar_plt():
    """Importa matplotlib.pyplot de forma lazy.

    matplotlib.use("Agg") debe llamarse antes del primer import de pyplot.
    El guard sobre sys.modules evita llamarlo de nuevo una vez ya cargado
    (llamarlo después de pyplot emite un warning en versiones recientes).
    """
    if "matplotlib.pyplot" not in sys.modules:
        import matplotlib as _mpl
        _mpl.use("Agg")
    import matplotlib.pyplot as plt
    return plt


class ChartError(ValueError):
    """Error controlado al generar un gráfico personalizado."""


def generar_grafico(df: pd.DataFrame, nombre_archivo: str, tipo: str = "barras") -> io.BytesIO | None:
    """Genera un gráfico con las columnas numéricas del DataFrame.
    tipo: 'barras' | 'lineas' | 'sectores'
    Devuelve un buffer PNG o None si no hay datos numéricos suficientes."""
    plt = _importar_plt()

    # Intentar convertir columnas object a numérico (por si vienen como texto)
    df = df.copy()
    for col in df.select_dtypes(include="object").columns:
        convertida = pd.to_numeric(df[col], errors="coerce")
        if convertida.notna().sum() > len(df) * 0.5:
            df[col] = convertida

    columnas_numericas = df.select_dtypes(include="number").columns.tolist()
    logger.info("Columnas numéricas detectadas en '%s': %s", nombre_archivo, columnas_numericas)
    if not columnas_numericas:
        logger.info("Tipos de columnas: %s", df.dtypes.to_dict())
        return None

    columnas_texto = df.select_dtypes(exclude="number").columns.tolist()
    eje_x = columnas_texto[0] if columnas_texto else None
    col_y = columnas_numericas[0]
    muestra = df.head(15)

    fig, ax = plt.subplots(figsize=(10, 5))

    if tipo == "sectores":
        valores = muestra[col_y].dropna()
        etiquetas = muestra[eje_x].astype(str) if eje_x else [str(i) for i in range(len(valores))]
        # Agrupar los menores en "Otros" si hay más de 8 sectores
        if len(valores) > 8:
            top = valores.nlargest(7)
            otros = pd.Series([valores.sum() - top.sum()], index=["Otros"])
            valores = pd.concat([top, otros])
            etiquetas = list(muestra.loc[top.index, eje_x].astype(str)) + ["Otros"] if eje_x else list(range(8))
        ax.pie(valores, labels=etiquetas, autopct="%1.1f%%", startangle=140)
        ax.set_title(f"{nombre_archivo} — {col_y}")

    elif tipo == "lineas":
        if eje_x:
            ax.plot(muestra[eje_x].astype(str), muestra[col_y], marker="o", color="steelblue")
            ax.set_xlabel(eje_x)
        else:
            ax.plot(range(len(muestra)), muestra[col_y], marker="o", color="steelblue")
        ax.set_ylabel(col_y)
        ax.set_title(f"{nombre_archivo} — {col_y}")
        plt.xticks(rotation=45, ha="right")

    else:  # barras (por defecto)
        if eje_x:
            ax.bar(muestra[eje_x].astype(str), muestra[col_y], color="steelblue")
            ax.set_xlabel(eje_x)
        else:
            ax.bar(range(len(muestra)), muestra[col_y], color="steelblue")
        ax.set_ylabel(col_y)
        ax.set_title(f"{nombre_archivo} — {col_y}")
        plt.xticks(rotation=45, ha="right")

    plt.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=120)
    plt.close(fig)
    buffer.seek(0)
    return buffer


# ── Gráfico personalizado bajo demanda ───────────────────────────────────────

_AGGFUNCS = {
    "suma":     "sum",
    "promedio": "mean",
    "contar":   "count",
    "max":      "max",
    "min":      "min",
}


def generar_grafico_personalizado(
    df: pd.DataFrame,
    col_y: str,
    col_x: str | None = None,
    tipo: str = "barras",
    agregar: str | None = None,
) -> tuple[io.BytesIO, str]:
    """Genera un gráfico con parámetros explícitos indicados por el usuario.

    Parámetros
    ----------
    col_y    : columna numérica para el eje Y (requerida).
    col_x    : columna para el eje X / categorías. None → índice.
    tipo     : 'barras' | 'lineas' | 'sectores' | 'dispersion'.
    agregar  : 'suma' | 'promedio' | 'contar' | 'max' | 'min' | None.

    Devuelve (buffer_png, descripcion).
    Lanza ChartError si las columnas no existen o los datos no son válidos.
    """
    plt = _importar_plt()

    # ── Validación ────────────────────────────────────────────────────────────
    if col_y not in df.columns:
        raise ChartError(f"La columna '{col_y}' no existe en el archivo.")
    if col_x and col_x not in df.columns:
        raise ChartError(f"La columna '{col_x}' no existe en el archivo.")

    tipo = tipo.lower() if tipo else "barras"
    if tipo not in ("barras", "lineas", "sectores", "dispersion"):
        tipo = "barras"

    df = df.copy()

    # ── Agregación ────────────────────────────────────────────────────────────
    descripcion_agg = ""
    if agregar and col_x:
        func = _AGGFUNCS.get(agregar, "sum")
        try:
            df = df.groupby(col_x, as_index=False)[col_y].agg(func)
            descripcion_agg = f" ({agregar} de {col_y} por {col_x})"
        except Exception as error:
            raise ChartError(f"No se pudo agrupar: {error}") from error

    # Limitar a 20 puntos para legibilidad
    muestra = df.head(20)

    x_vals = muestra[col_x].astype(str) if col_x else range(len(muestra))
    y_vals = pd.to_numeric(muestra[col_y], errors="coerce")

    if y_vals.dropna().empty:
        raise ChartError(f"La columna '{col_y}' no contiene valores numéricos.")

    # ── Dibujar ───────────────────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(10, 5))
    titulo = f"{col_y}{descripcion_agg}"

    if tipo == "sectores":
        validos = y_vals.dropna()
        etiquetas = list(muestra.loc[validos.index, col_x].astype(str)) if col_x else list(range(len(validos)))
        if len(validos) > 8:
            top = validos.nlargest(7)
            otros_val = validos.sum() - top.sum()
            validos = pd.concat([top, pd.Series([otros_val], index=[-1])])
            etiquetas = [muestra.loc[i, col_x] if col_x else str(i) for i in top.index] + ["Otros"]
        ax.pie(validos, labels=etiquetas, autopct="%1.1f%%", startangle=140)

    elif tipo == "lineas":
        ax.plot(x_vals, y_vals, marker="o", color="steelblue")
        if col_x:
            ax.set_xlabel(col_x)
            plt.xticks(rotation=45, ha="right")
        ax.set_ylabel(col_y)

    elif tipo == "dispersion":
        if col_x:
            x_num = pd.to_numeric(muestra[col_x], errors="coerce")
            if x_num.dropna().empty:
                raise ChartError(
                    f"La columna '{col_x}' debe ser numérica para un gráfico de dispersión."
                )
            ax.scatter(x_num, y_vals, color="steelblue", alpha=0.7)
            ax.set_xlabel(col_x)
        else:
            ax.scatter(range(len(y_vals)), y_vals, color="steelblue", alpha=0.7)
        ax.set_ylabel(col_y)

    else:  # barras (por defecto)
        ax.bar(x_vals, y_vals, color="steelblue")
        if col_x:
            ax.set_xlabel(col_x)
            plt.xticks(rotation=45, ha="right")
        ax.set_ylabel(col_y)

    ax.set_title(titulo)
    plt.tight_layout()

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=120)
    plt.close(fig)
    buffer.seek(0)
    return buffer, titulo
