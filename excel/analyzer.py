import io
import openpyxl
import numpy as np
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


# ── Análisis estadístico avanzado (C1) ───────────────────────────────────────

def analisis_estadistico_completo(df: pd.DataFrame) -> str:
    """Genera un informe estadístico detallado de todas las columnas numéricas."""
    cols_num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not cols_num:
        return "⚠️ El archivo no tiene columnas numéricas para analizar."

    lineas = [f"📊 *Análisis estadístico — {len(cols_num)} columna(s) numéricas*\n"]

    for col in cols_num:
        serie = df[col].dropna()
        if serie.empty:
            continue
        nulos = int(df[col].isnull().sum())
        lineas.append(f"*{col}*")
        lineas.append(f"  • Registros: {len(serie)} ({nulos} vacíos)")
        lineas.append(f"  • Media:     {serie.mean():.2f}")
        lineas.append(f"  • Mediana:   {serie.median():.2f}")
        lineas.append(f"  • Mín / Máx: {serie.min():.2f} / {serie.max():.2f}")
        lineas.append(f"  • Desv. std: {serie.std():.2f}")
        q1, q3 = serie.quantile(0.25), serie.quantile(0.75)
        lineas.append(f"  • P25 / P75: {q1:.2f} / {q3:.2f}")
        # Detectar asimetría
        skew = serie.skew()
        if abs(skew) > 1:
            dir_skew = "derecha (valores altos extremos)" if skew > 0 else "izquierda (valores bajos extremos)"
            lineas.append(f"  • Distribución sesgada hacia la {dir_skew}")
        lineas.append("")

    return "\n".join(lineas)


# ── Análisis de correlaciones (C2) ────────────────────────────────────────────

def analisis_correlaciones(df: pd.DataFrame) -> tuple[str, io.BytesIO | None]:
    """Calcula la matriz de correlaciones y genera un heatmap PNG.

    Devuelve (texto_resumen, buffer_imagen | None).
    """
    cols_num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if len(cols_num) < 2:
        return "⚠️ Se necesitan al menos 2 columnas numéricas para calcular correlaciones.", None

    corr = df[cols_num].corr()

    # Texto: top correlaciones (excluyendo diagonal)
    pares = []
    for i, c1 in enumerate(cols_num):
        for c2 in cols_num[i + 1:]:
            pares.append((abs(corr.loc[c1, c2]), corr.loc[c1, c2], c1, c2))
    pares.sort(reverse=True)

    lineas = ["📈 *Análisis de correlaciones*\n"]
    lineas.append("*Correlaciones más fuertes:*")
    for _, valor, c1, c2 in pares[:8]:
        if abs(valor) < 0.1:
            continue
        fuerza = (
            "muy fuerte" if abs(valor) > 0.8 else
            "fuerte"     if abs(valor) > 0.6 else
            "moderada"   if abs(valor) > 0.4 else "débil"
        )
        signo = "positiva" if valor > 0 else "negativa"
        lineas.append(f"  • *{c1}* ↔ *{c2}*: {valor:.2f} ({fuerza}, {signo})")

    if not any(abs(p[0]) >= 0.1 for p in pares):
        lineas.append("  No se detectan correlaciones relevantes entre columnas.")

    lineas.append(f"\n*Matriz completa:*\n```\n{corr.round(2).to_string()}\n```")
    texto = "\n".join(lineas)

    # Heatmap
    buf_img = _generar_heatmap(corr, cols_num)
    return texto, buf_img


def _generar_heatmap(corr: pd.DataFrame, cols: list[str]) -> io.BytesIO | None:
    """Genera un heatmap de correlaciones con matplotlib."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.colors as mcolors

        n = len(cols)
        fig, ax = plt.subplots(figsize=(max(6, n * 1.2), max(5, n * 1.0)))

        cmap = plt.cm.RdYlGn
        im = ax.imshow(corr.values, cmap=cmap, vmin=-1, vmax=1, aspect="auto")
        plt.colorbar(im, ax=ax, shrink=0.8)

        ax.set_xticks(range(n))
        ax.set_yticks(range(n))
        ax.set_xticklabels(cols, rotation=45, ha="right", fontsize=9)
        ax.set_yticklabels(cols, fontsize=9)

        for i in range(n):
            for j in range(n):
                val = corr.values[i, j]
                color = "white" if abs(val) > 0.6 else "black"
                ax.text(j, i, f"{val:.2f}", ha="center", va="center",
                        fontsize=8, color=color, fontweight="bold")

        ax.set_title("Mapa de correlaciones", fontsize=12, fontweight="bold", pad=12)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=130, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf
    except Exception:
        return None


# ── Análisis de tendencia (C3) ────────────────────────────────────────────────

_PALABRAS_FECHA = ("fecha", "date", "dia", "día", "mes", "año", "year", "month",
                   "periodo", "semana", "trimestre", "quarter")


def analisis_tendencia(df: pd.DataFrame) -> tuple[str, io.BytesIO | None]:
    """Detecta tendencias en columnas numéricas usando regresión lineal.

    Devuelve (texto_resumen, buffer_grafico | None).
    """
    cols_num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not cols_num:
        return "⚠️ No hay columnas numéricas para analizar la tendencia.", None

    # Detectar columna temporal (para usar como eje X)
    col_fecha = next(
        (c for c in df.columns if any(p in str(c).lower() for p in _PALABRAS_FECHA)),
        None,
    )

    # Construir eje X numérico
    x_label = "Período"
    if col_fecha:
        fechas = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        if fechas.notna().sum() >= 3:
            x_raw = fechas.ffill().map(lambda d: d.toordinal()).values.astype(float)
            x_label = str(col_fecha)
        else:
            col_fecha = None
    if col_fecha is None:
        x_raw = np.arange(len(df), dtype=float)

    # Normalizar X al rango [0, N-1] para que la pendiente sea interpretable
    x = x_raw - x_raw[0]
    x_norm_total = x[-1] if x[-1] != 0 else 1

    lineas = ["📈 *Análisis de tendencia*\n"]
    resultados = []   # para el gráfico

    for col in cols_num:
        serie = df[col].dropna()
        if len(serie) < 3:
            continue

        # Alinear longitudes
        n = min(len(x), len(serie))
        xi, yi = x[:n], serie.values[:n]

        try:
            coef, intercept = np.polyfit(xi, yi, 1)
        except Exception:
            continue

        # R²
        y_pred = coef * xi + intercept
        ss_res = float(((yi - y_pred) ** 2).sum())
        ss_tot = float(((yi - yi.mean()) ** 2).sum())
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else 0.0

        # Cambio total estimado en el período
        cambio_abs = coef * x_norm_total
        pct_cambio = (cambio_abs / abs(yi[0])) * 100 if yi[0] != 0 else 0.0

        # Interpretación
        if r2 < 0.25:
            tendencia = "sin tendencia clara (datos muy dispersos)"
        elif coef > 0:
            tendencia = f"📈 creciente  ({pct_cambio:+.1f}% en el período)"
        else:
            tendencia = f"📉 decreciente ({pct_cambio:+.1f}% en el período)"

        fiabilidad = "alta" if r2 > 0.7 else "moderada" if r2 > 0.4 else "baja"

        lineas.append(f"*{col}*")
        lineas.append(f"  • Tendencia: {tendencia}")
        lineas.append(f"  • R²: {r2:.2f} (ajuste {fiabilidad})")
        lineas.append(f"  • Inicio: {yi[0]:.2f}  →  Final: {yi[-1]:.2f}  (Δ {yi[-1]-yi[0]:+.2f})")
        lineas.append("")

        resultados.append((col, xi, yi, coef, intercept, r2))

    if not resultados:
        return "⚠️ No hay suficientes datos para calcular tendencias (mínimo 3 filas).", None

    buf = _generar_grafico_tendencia(resultados, x_label, df, col_fecha)
    return "\n".join(lineas), buf


def _generar_grafico_tendencia(resultados, x_label, df, col_fecha) -> io.BytesIO | None:
    """Genera un gráfico con los valores reales y la línea de tendencia."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        n_cols = len(resultados)
        fig, axes = plt.subplots(n_cols, 1, figsize=(9, 3.5 * n_cols), squeeze=False)

        for idx, (col, xi, yi, coef, intercept, r2) in enumerate(resultados):
            ax = axes[idx][0]

            # Etiquetas del eje X
            if col_fecha and col_fecha in df.columns:
                fechas = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
                x_ticks = fechas.dropna().values[:len(xi)]
                ax.plot(x_ticks, yi, "o-", color="#2E75B6", linewidth=1.5,
                        markersize=5, label="Valor real", zorder=3)
                y_trend = coef * xi + intercept
                ax.plot(x_ticks, y_trend, "--", color="#FF6B6B", linewidth=2,
                        label=f"Tendencia (R²={r2:.2f})", zorder=2)
                fig.autofmt_xdate()
            else:
                ax.plot(yi, "o-", color="#2E75B6", linewidth=1.5,
                        markersize=5, label="Valor real", zorder=3)
                y_trend = coef * xi + intercept
                ax.plot(y_trend, "--", color="#FF6B6B", linewidth=2,
                        label=f"Tendencia (R²={r2:.2f})", zorder=2)
                ax.set_xlabel(x_label, fontsize=9)

            ax.set_title(col, fontsize=11, fontweight="bold")
            ax.set_ylabel("Valor", fontsize=9)
            ax.legend(fontsize=8)
            ax.grid(True, alpha=0.3)

        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=130, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf
    except Exception:
        return None


# ── Comparar dos DataFrames (F3) ─────────────────────────────────────────────

def comparar_dataframes(df_a: pd.DataFrame, df_b: pd.DataFrame,
                        nombre_a: str = "Archivo A",
                        nombre_b: str = "Archivo B") -> tuple[str, pd.DataFrame | None]:
    """Compara dos DataFrames y devuelve (texto_resumen, df_diferencias | None).

    Detecta diferencias estructurales, filas exclusivas de cada archivo
    y el número de filas idénticas en ambos.
    """
    lineas = [f"🔄 *Comparación: {nombre_a} vs {nombre_b}*\n"]

    cols_a = set(df_a.columns)
    cols_b = set(df_b.columns)

    # ── Diferencias estructurales ─────────────────────────────────────────────
    solo_en_a = cols_a - cols_b
    solo_en_b = cols_b - cols_a
    if solo_en_a:
        lineas.append(f"📌 Columnas solo en *{nombre_a}*: {', '.join(sorted(solo_en_a))}")
    if solo_en_b:
        lineas.append(f"📌 Columnas solo en *{nombre_b}*: {', '.join(sorted(solo_en_b))}")

    lineas.append(
        f"\n📊 Filas: *{nombre_a}* = {len(df_a):,}  |  *{nombre_b}* = {len(df_b):,}"
    )

    cols_comunes = sorted(cols_a & cols_b)
    if not cols_comunes:
        lineas.append("\n⚠️ No hay columnas en común — no es posible comparar filas.")
        return "\n".join(lineas), None

    # ── Comparación de filas (usando columnas comunes como cadenas) ───────────
    a_str = df_a[cols_comunes].astype(str)
    b_str = df_b[cols_comunes].astype(str)

    merge = a_str.merge(b_str, on=cols_comunes, how="outer", indicator=True)
    solo_a = merge[merge["_merge"] == "left_only"].drop(columns="_merge")
    solo_b = merge[merge["_merge"] == "right_only"].drop(columns="_merge")
    n_comunes = int((merge["_merge"] == "both").sum())

    lineas.append(f"\n➕ Filas solo en *{nombre_a}*: {len(solo_a):,}")
    lineas.append(f"➖ Filas solo en *{nombre_b}*: {len(solo_b):,}")
    lineas.append(f"✅ Filas idénticas en ambos: {n_comunes:,}")

    # ── DataFrame de diferencias para exportar ────────────────────────────────
    df_diff = None
    if len(solo_a) > 0 or len(solo_b) > 0:
        partes = []
        if len(solo_a) > 0:
            tmp = solo_a.copy()
            tmp.insert(0, "_origen_", f"Solo en {nombre_a}")
            partes.append(tmp)
        if len(solo_b) > 0:
            tmp = solo_b.copy()
            tmp.insert(0, "_origen_", f"Solo en {nombre_b}")
            partes.append(tmp)
        df_diff = pd.concat(partes, ignore_index=True)

    if len(solo_a) == 0 and len(solo_b) == 0:
        lineas.append(
            "\n🎉 Los archivos tienen exactamente el mismo contenido en las columnas comunes."
        )
    else:
        lineas.append(
            f"\nSe adjunta un archivo con las {len(solo_a) + len(solo_b):,} "
            "filas que difieren entre ambos archivos."
        )

    return "\n".join(lineas), df_diff
