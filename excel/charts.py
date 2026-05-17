import io
import matplotlib
matplotlib.use("Agg")  # backend sin pantalla
import matplotlib.pyplot as plt
import pandas as pd


def generar_grafico(df: pd.DataFrame, nombre_archivo: str) -> io.BytesIO | None:
    """Genera un gráfico de barras con las columnas numéricas del DataFrame.
    Devuelve un buffer PNG o None si no hay datos numéricos suficientes."""
    columnas_numericas = df.select_dtypes(include="number").columns.tolist()
    if not columnas_numericas:
        return None

    # Usar la primera columna no numérica como eje X (etiquetas), si existe
    columnas_texto = df.select_dtypes(exclude="number").columns.tolist()
    eje_x = columnas_texto[0] if columnas_texto else df.index
    col_y = columnas_numericas[0]

    fig, ax = plt.subplots(figsize=(10, 5))
    muestra = df.head(15)

    if columnas_texto:
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
