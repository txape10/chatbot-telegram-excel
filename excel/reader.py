import pandas as pd


def leer_excel(ruta: str) -> pd.DataFrame:
    return pd.read_excel(ruta, engine="openpyxl")
