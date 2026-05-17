import pandas as pd


def resumir(df: pd.DataFrame, nombre_archivo: str) -> str:
    """Devuelve un resumen legible del DataFrame para mostrar al usuario."""
    nulos = {col: int(v) for col, v in df.isnull().sum().items() if v > 0}
    duplicados = int(df.duplicated().sum())

    lineas = [
        f"📊 *{nombre_archivo}*",
        f"• Filas: {len(df)} | Columnas: {len(df.columns)}",
        f"• Columnas: {', '.join(df.columns)}",
    ]
    if nulos:
        detalle = ", ".join(f"{col}: {n}" for col, n in nulos.items())
        lineas.append(f"• Celdas vacías: {detalle}")
    else:
        lineas.append("• Sin celdas vacías ✅")

    if duplicados:
        lineas.append(f"• Filas duplicadas: {duplicados} ⚠️")
    else:
        lineas.append("• Sin duplicados ✅")

    lineas.append("\nPuedes hacerme preguntas sobre este archivo.")
    return "\n".join(lineas)


def construir_contexto(df: pd.DataFrame, nombre_archivo: str) -> str:
    """Construye el contexto que se enviará al LLM junto con las preguntas del usuario."""
    nulos = {col: int(v) for col, v in df.isnull().sum().items() if v > 0}
    duplicados = int(df.duplicated().sum())

    texto = (
        f"El usuario ha subido un archivo Excel llamado '{nombre_archivo}'.\n"
        f"Filas: {len(df)} | Columnas: {len(df.columns)}\n"
        f"Nombres de columnas: {', '.join(df.columns)}\n"
    )
    if nulos:
        texto += f"Celdas vacías por columna: {nulos}\n"
    if duplicados:
        texto += f"Filas duplicadas: {duplicados}\n"

    texto += f"\nPrimeras filas del archivo:\n{df.head(10).to_string(index=False)}\n"
    return texto
