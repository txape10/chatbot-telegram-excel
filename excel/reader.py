import pandas as pd


def leer_excel(ruta: str) -> pd.DataFrame:
    """Lee la primera hoja de un Excel."""
    return pd.read_excel(ruta, engine="openpyxl")


def leer_excel_hojas(ruta: str) -> dict[str, pd.DataFrame]:
    """Lee todas las hojas de un Excel. Devuelve {nombre_hoja: DataFrame}."""
    return pd.read_excel(ruta, sheet_name=None, engine="openpyxl")


def leer_csv(ruta: str) -> pd.DataFrame:
    """Lee un CSV probando separadores comunes y encodings."""
    for sep in (",", ";", "\t"):
        for enc in ("utf-8", "latin-1"):
            try:
                df = pd.read_csv(ruta, sep=sep, encoding=enc)
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
    # Último recurso: pandas decide el separador
    return pd.read_csv(ruta)
