import io
import logging
import matplotlib
matplotlib.use("Agg")  # backend sin pantalla
import matplotlib.pyplot as plt
import pandas as pd

logger = logging.getLogger(__name__)


def generar_grafico(df: pd.DataFrame, nombre_archivo: str, tipo: str = "barras") -> io.BytesIO | None:
    """Genera un gráfico con las columnas numéricas del DataFrame.
    tipo: 'barras' | 'lineas' | 'sectores'
    Devuelve un buffer PNG o None si no hay datos numéricos suficientes."""

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
