import openpyxl
import pandas as pd

_ERRORES_EXCEL = {"#REF!", "#VALOR!", "#N/A", "#DIV/0!", "#NOMBRE?", "#NULO!", "#NUM!", "#VALUE!", "#NAME?"}


def detectar_errores_xlsx(ruta: str) -> list[str]:
    """Escanea un .xlsx y devuelve lista de celdas con errores de fórmula."""
    errores = []
    try:
        wb = openpyxl.load_workbook(ruta, data_only=True)
        for hoja in wb.worksheets:
            for fila in hoja.iter_rows():
                for celda in fila:
                    val = str(celda.value) if celda.value is not None else ""
                    if any(e in val for e in _ERRORES_EXCEL):
                        errores.append(f"{hoja.title}!{celda.coordinate}: `{celda.value}`")
                        if len(errores) >= 15:
                            return errores
    except Exception:
        pass
    return errores


def analizar_calidad(df: pd.DataFrame) -> list[str]:
    """Detecta problemas de calidad de datos en el DataFrame.
    Devuelve una lista de avisos legibles para el usuario."""
    avisos = []
    total = len(df)
    if total == 0:
        return avisos

    for col in df.columns:
        serie = df[col]
        pct_nulos = serie.isnull().mean()

        # Columna casi vacía (>80 % nulos)
        if pct_nulos > 0.8:
            avisos.append(f"• `{col}`: {pct_nulos:.0%} de valores vacíos")
            continue   # si está casi vacía, las demás comprobaciones no aportan

        no_nulos = serie.dropna()

        # Columna constante
        if len(no_nulos) > 0 and no_nulos.nunique() == 1:
            avisos.append(f"• `{col}`: columna constante (valor siempre '{no_nulos.iloc[0]}')")

        # Mezcla texto/número
        if serie.dtype == object:
            convertidos = pd.to_numeric(no_nulos, errors="coerce")
            tiene_nums  = convertidos.notna().any()
            tiene_texto = convertidos.isna().any()
            if tiene_nums and tiene_texto:
                avisos.append(f"• `{col}`: mezcla de texto y números en la misma columna")

        # Outliers con IQR (solo numéricas)
        if pd.api.types.is_numeric_dtype(serie) and len(no_nulos) >= 4:
            q1, q3 = no_nulos.quantile(0.25), no_nulos.quantile(0.75)
            iqr = q3 - q1
            if iqr > 0:
                n_outliers = int(((no_nulos < q1 - 1.5 * iqr) | (no_nulos > q3 + 1.5 * iqr)).sum())
                if n_outliers > 0:
                    avisos.append(f"• `{col}`: {n_outliers} posibles valor(es) atípico(s) (IQR)")

        # Fechas inválidas (columnas cuyo nombre sugiere fecha)
        _PALABRAS_FECHA = ("fecha", "date", "dia", "día", "mes", "año", "year", "month")
        if serie.dtype == object and any(p in str(col).lower() for p in _PALABRAS_FECHA):
            convertidas = pd.to_datetime(no_nulos, errors="coerce")
            n_invalidas = int(convertidas.isna().sum())
            if n_invalidas > 0:
                avisos.append(f"• `{col}`: {n_invalidas} fecha(s) con formato no reconocido")

    return avisos


def resumir(df: pd.DataFrame, nombre_archivo: str, errores: list[str] | None = None) -> str:
    """Devuelve un resumen legible del DataFrame para mostrar al usuario."""
    nulos = {col: int(v) for col, v in df.isnull().sum().items() if v > 0}
    duplicados = int(df.duplicated().sum())

    lineas = [
        f"📊 *{nombre_archivo}*",
        f"• Filas: {len(df)} | Columnas: {len(df.columns)}",
        f"• Columnas: {', '.join(str(c) for c in df.columns)}",
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

    if errores:
        lineas.append(f"\n⚠️ *Errores de fórmula ({len(errores)}):*")
        for e in errores:
            lineas.append(f"  • {e}")

    avisos_calidad = analizar_calidad(df)
    if avisos_calidad:
        lineas.append(f"\n🔍 *Calidad de datos ({len(avisos_calidad)} aviso(s)):*")
        for aviso in avisos_calidad:
            lineas.append(f"  {aviso}")

    lineas.append("\nPuedes hacerme preguntas sobre este archivo.")
    return "\n".join(lineas)


def resumir_hojas(sheets: dict[str, pd.DataFrame], nombre_archivo: str) -> str:
    """Devuelve un resumen cuando el Excel tiene varias hojas."""
    lineas = [f"📊 *{nombre_archivo}* — {len(sheets)} hojas detectadas:\n"]
    for nombre, df in sheets.items():
        lineas.append(f"• *{nombre}*: {len(df)} filas × {len(df.columns)} columnas")
    lineas.append("\nSelecciona la hoja que quieres analizar:")
    return "\n".join(lineas)


def construir_contexto(df: pd.DataFrame, nombre_archivo: str) -> str:
    """Construye el contexto que se enviará al LLM junto con las preguntas del usuario."""
    nulos = {col: int(v) for col, v in df.isnull().sum().items() if v > 0}
    duplicados = int(df.duplicated().sum())

    texto = (
        f"El usuario ha subido un archivo llamado '{nombre_archivo}'.\n"
        f"Filas: {len(df)} | Columnas: {len(df.columns)}\n"
        f"Nombres de columnas: {', '.join(str(c) for c in df.columns)}\n"
    )
    if nulos:
        texto += f"Celdas vacías por columna: {nulos}\n"
    if duplicados:
        texto += f"Filas duplicadas: {duplicados}\n"

    texto += f"\nPrimeras filas del archivo:\n{df.head(10).to_string(index=False)}\n"
    return texto
