import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def _estilo_cabecera(celda):
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor="2E75B6")
    celda.alignment = Alignment(horizontal="center")


def _estilo_formula(celda):
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor="70AD47")
    celda.alignment = Alignment(horizontal="left")


def _ajustar_columnas(hoja):
    for col in hoja.columns:
        ancho = max(len(str(c.value or "")) for c in col) + 4
        hoja.column_dimensions[get_column_letter(col[0].column)].width = min(ancho, 40)


def crear_ejemplo(funcion: str) -> tuple[io.BytesIO, str]:
    """Devuelve (buffer_xlsx, nombre_archivo). Usa ejemplo genérico si la función no está definida."""
    funcion_upper = funcion.upper().replace(" ", ".")
    creador = EJEMPLOS.get(funcion_upper, _ejemplo_generico)
    wb, nombre = creador(funcion_upper)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, nombre


# ── Ejemplos concretos ──────────────────────────────────────────────────────

def _ejemplo_buscarv(funcion: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BUSCARV"

    cabeceras = ["ID", "Producto", "Precio"]
    datos = [(1, "Teclado", 29.99), (2, "Ratón", 19.99), (3, "Monitor", 199.99),
             (4, "Auriculares", 49.99), (5, "Webcam", 39.99)]

    for i, cab in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=i, value=cab)
        _estilo_cabecera(c)
    for fila, dato in enumerate(datos, 2):
        for col, valor in enumerate(dato, 1):
            ws.cell(row=fila, column=col, value=valor)

    ws["E1"] = "ID a buscar"
    _estilo_cabecera(ws["E1"])
    ws["E2"] = 3

    ws["F1"] = "Producto encontrado"
    _estilo_cabecera(ws["F1"])
    c = ws["F2"]
    c.value = '=BUSCARV(E2,$A$2:$C$6,2,FALSO)'
    _estilo_formula(c)

    ws["G1"] = "Precio encontrado"
    _estilo_cabecera(ws["G1"])
    c = ws["G2"]
    c.value = '=BUSCARV(E2,$A$2:$C$6,3,FALSO)'
    _estilo_formula(c)

    _ajustar_columnas(ws)
    return wb, "ejemplo_BUSCARV.xlsx"


def _ejemplo_buscarx(funcion: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BUSCARX"

    cabeceras = ["Código", "Artículo", "Stock", "Precio"]
    datos = [("A001", "Silla", 12, 89.90), ("A002", "Mesa", 5, 149.90),
             ("A003", "Lámpara", 30, 24.50), ("A004", "Estantería", 8, 59.99)]

    for i, cab in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=i, value=cab)
        _estilo_cabecera(c)
    for fila, dato in enumerate(datos, 2):
        for col, valor in enumerate(dato, 1):
            ws.cell(row=fila, column=col, value=valor)

    ws["F1"] = "Código a buscar"
    _estilo_cabecera(ws["F1"])
    ws["F2"] = "A003"

    ws["G1"] = "Resultado BUSCARX"
    _estilo_cabecera(ws["G1"])
    c = ws["G2"]
    c.value = '=BUSCARX(F2,$A$2:$A$5,$C$2:$C$5,"No encontrado")'
    _estilo_formula(c)

    _ajustar_columnas(ws)
    return wb, "ejemplo_BUSCARX.xlsx"


def _ejemplo_sumar_si(funcion: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SUMAR.SI"

    cabeceras = ["Mes", "Vendedor", "Ventas"]
    datos = [("Enero", "Ana", 1200), ("Enero", "Luis", 950), ("Febrero", "Ana", 1400),
             ("Febrero", "Luis", 1100), ("Marzo", "Ana", 1600), ("Marzo", "Luis", 800)]

    for i, cab in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=i, value=cab)
        _estilo_cabecera(c)
    for fila, dato in enumerate(datos, 2):
        for col, valor in enumerate(dato, 1):
            ws.cell(row=fila, column=col, value=valor)

    ws["E1"] = "Vendedor"
    _estilo_cabecera(ws["E1"])
    ws["F1"] = "Total ventas"
    _estilo_cabecera(ws["F1"])

    ws["E2"] = "Ana"
    c = ws["F2"]
    c.value = '=SUMAR.SI($B$2:$B$7,E2,$C$2:$C$7)'
    _estilo_formula(c)

    ws["E3"] = "Luis"
    c = ws["F3"]
    c.value = '=SUMAR.SI($B$2:$B$7,E3,$C$2:$C$7)'
    _estilo_formula(c)

    _ajustar_columnas(ws)
    return wb, "ejemplo_SUMAR.SI.xlsx"


def _ejemplo_contar_si(funcion: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONTAR.SI"

    cabeceras = ["Empleado", "Departamento", "Nota"]
    datos = [("Ana", "Ventas", "Aprobado"), ("Luis", "IT", "Pendiente"),
             ("Sara", "Ventas", "Aprobado"), ("Pedro", "IT", "Aprobado"),
             ("Marta", "Ventas", "Pendiente"), ("Carlos", "IT", "Aprobado")]

    for i, cab in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=i, value=cab)
        _estilo_cabecera(c)
    for fila, dato in enumerate(datos, 2):
        for col, valor in enumerate(dato, 1):
            ws.cell(row=fila, column=col, value=valor)

    ws["E1"] = "Estado"
    _estilo_cabecera(ws["E1"])
    ws["F1"] = "Cantidad"
    _estilo_cabecera(ws["F1"])

    ws["E2"] = "Aprobado"
    c = ws["F2"]
    c.value = '=CONTAR.SI($C$2:$C$7,E2)'
    _estilo_formula(c)

    ws["E3"] = "Pendiente"
    c = ws["F3"]
    c.value = '=CONTAR.SI($C$2:$C$7,E3)'
    _estilo_formula(c)

    _ajustar_columnas(ws)
    return wb, "ejemplo_CONTAR.SI.xlsx"


def _ejemplo_si(funcion: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SI"

    cabeceras = ["Alumno", "Nota", "Resultado", "Calificación"]
    datos = [("Ana", 7.5), ("Luis", 4.2), ("Sara", 9.1), ("Pedro", 5.0), ("Marta", 3.8)]

    for i, cab in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=i, value=cab)
        _estilo_cabecera(c)
    for fila, dato in enumerate(datos, 2):
        ws.cell(row=fila, column=1, value=dato[0])
        ws.cell(row=fila, column=2, value=dato[1])
        c = ws.cell(row=fila, column=3)
        c.value = f'=SI(B{fila}>=5,"Aprobado","Suspenso")'
        _estilo_formula(c)
        c2 = ws.cell(row=fila, column=4)
        c2.value = f'=SI(B{fila}>=9,"Sobresaliente",SI(B{fila}>=7,"Notable",SI(B{fila}>=5,"Aprobado","Suspenso")))'
        _estilo_formula(c2)

    _ajustar_columnas(ws)
    return wb, "ejemplo_SI.xlsx"


def _ejemplo_generico(funcion: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = funcion[:30]

    cabeceras = ["ID", "Categoría", "Valor A", "Valor B", "Resultado"]
    datos = [(1, "Norte", 100, 200), (2, "Sur", 150, 80),
             (3, "Este", 90, 310), (4, "Oeste", 200, 120), (5, "Norte", 75, 95)]

    for i, cab in enumerate(cabeceras, 1):
        c = ws.cell(row=1, column=i, value=cab)
        _estilo_cabecera(c)
    for fila, dato in enumerate(datos, 2):
        for col, valor in enumerate(dato, 1):
            ws.cell(row=fila, column=col, value=valor)
        c = ws.cell(row=fila, column=5)
        c.value = f"Aplica aquí {funcion}"
        _estilo_formula(c)

    ws["G1"] = f"Función: {funcion}"
    ws["G2"] = "Modifica la columna Resultado con tu fórmula"

    _ajustar_columnas(ws)
    return wb, f"ejemplo_{funcion}.xlsx"


# Registro de ejemplos (debe estar después de las funciones)
EJEMPLOS = {
    "BUSCARV": _ejemplo_buscarv,
    "BUSCARX": _ejemplo_buscarx,
    "SUMAR.SI": _ejemplo_sumar_si,
    "CONTAR.SI": _ejemplo_contar_si,
    "SI": _ejemplo_si,
}


# ── Plantillas de uso ────────────────────────────────────────────────────────

def crear_plantilla(nombre: str) -> tuple[io.BytesIO, str]:
    """Genera una plantilla lista para usar. nombre: presupuesto|gastos|kpis|inventario."""
    creadores = {
        "presupuesto": _plantilla_presupuesto,
        "gastos":      _plantilla_gastos,
        "kpis":        _plantilla_kpis,
        "inventario":  _plantilla_inventario,
    }
    creador = creadores.get(nombre, _plantilla_presupuesto)
    wb, nombre_archivo = creador()
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, nombre_archivo


def _estilo_total(celda):
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor="ED7D31")
    celda.alignment = Alignment(horizontal="right")


def _plantilla_presupuesto():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Presupuesto"

    cabeceras = ["Categoría", "Presupuesto", "Real", "Diferencia", "% Cumplido"]
    for i, cab in enumerate(cabeceras, 1):
        _estilo_cabecera(ws.cell(row=1, column=i, value=cab))

    categorias = ["Vivienda", "Alimentación", "Transporte", "Salud", "Ocio",
                  "Ropa", "Educación", "Ahorro", "Otros"]
    for fila, cat in enumerate(categorias, 2):
        ws.cell(row=fila, column=1, value=cat)
        ws.cell(row=fila, column=2, value=0)   # Presupuesto — el usuario lo rellena
        ws.cell(row=fila, column=3, value=0)   # Real — el usuario lo rellena
        c_dif = ws.cell(row=fila, column=4)
        c_dif.value = f"=B{fila}-C{fila}"
        _estilo_formula(c_dif)
        c_pct = ws.cell(row=fila, column=5)
        c_pct.value = f'=SI(B{fila}=0,"—",C{fila}/B{fila})'
        c_pct.number_format = "0.0%"
        _estilo_formula(c_pct)

    fila_total = len(categorias) + 2
    ws.cell(row=fila_total, column=1, value="TOTAL")
    letras = {2: "B", 3: "C", 4: "D"}
    for col in range(2, 5):
        c = ws.cell(row=fila_total, column=col)
        c.value = f"=SUMA({letras[col]}2:{letras[col]}{fila_total - 1})"
        _estilo_total(c)

    _ajustar_columnas(ws)
    return wb, "plantilla_presupuesto.xlsx"


def _plantilla_gastos():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gastos"

    cabeceras = ["Fecha", "Descripción", "Categoría", "Importe", "Acumulado"]
    for i, cab in enumerate(cabeceras, 1):
        _estilo_cabecera(ws.cell(row=1, column=i, value=cab))

    import datetime
    hoy = datetime.date.today()
    ejemplos = [
        (hoy, "Supermercado", "Alimentación", 65.30),
        (hoy, "Gasolina", "Transporte", 48.00),
        (hoy, "Netflix", "Ocio", 15.99),
    ]
    for fila, (fecha, desc, cat, imp) in enumerate(ejemplos, 2):
        ws.cell(row=fila, column=1, value=fecha).number_format = "dd/mm/yyyy"
        ws.cell(row=fila, column=2, value=desc)
        ws.cell(row=fila, column=3, value=cat)
        ws.cell(row=fila, column=4, value=imp)
        c = ws.cell(row=fila, column=5)
        c.value = f"=SUMA($D$2:D{fila})"
        _estilo_formula(c)

    # Hoja de resumen por categoría
    ws2 = wb.create_sheet("Resumen")
    cabeceras2 = ["Categoría", "Total gastado"]
    for i, cab in enumerate(cabeceras2, 1):
        _estilo_cabecera(ws2.cell(row=1, column=i, value=cab))
    cats = ["Alimentación", "Transporte", "Ocio", "Salud", "Vivienda", "Otros"]
    for fila, cat in enumerate(cats, 2):
        ws2.cell(row=fila, column=1, value=cat)
        c = ws2.cell(row=fila, column=2)
        c.value = f'=SUMAR.SI(Gastos!$C:$C,A{fila},Gastos!$D:$D)'
        _estilo_formula(c)

    _ajustar_columnas(ws)
    _ajustar_columnas(ws2)
    return wb, "plantilla_gastos.xlsx"


def _plantilla_kpis():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KPIs"

    cabeceras = ["KPI", "Objetivo", "Real", "Cumplimiento", "Estado"]
    for i, cab in enumerate(cabeceras, 1):
        _estilo_cabecera(ws.cell(row=1, column=i, value=cab))

    kpis = [
        ("Ventas mensuales (€)", 50000, 0),
        ("Nuevos clientes", 20, 0),
        ("Tasa de conversión (%)", 15, 0),
        ("Satisfacción cliente (1-10)", 8, 0),
        ("Tickets resueltos", 100, 0),
        ("Coste por adquisición (€)", 30, 0),
    ]
    for fila, (kpi, obj, real) in enumerate(kpis, 2):
        ws.cell(row=fila, column=1, value=kpi)
        ws.cell(row=fila, column=2, value=obj)
        ws.cell(row=fila, column=3, value=real)
        c_pct = ws.cell(row=fila, column=4)
        c_pct.value = f'=SI(B{fila}=0,"—",C{fila}/B{fila})'
        c_pct.number_format = "0.0%"
        _estilo_formula(c_pct)
        c_est = ws.cell(row=fila, column=5)
        c_est.value = f'=SI(C{fila}=0,"Sin datos",SI(C{fila}>=B{fila},"✅ OK","⚠️ Por debajo"))'
        _estilo_formula(c_est)

    _ajustar_columnas(ws)
    return wb, "plantilla_kpis.xlsx"


def _plantilla_inventario():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventario"

    cabeceras = ["Código", "Producto", "Cantidad", "Precio unit.", "Valor total", "Stock mín.", "Alerta"]
    for i, cab in enumerate(cabeceras, 1):
        _estilo_cabecera(ws.cell(row=1, column=i, value=cab))

    productos = [
        ("P001", "Teclado mecánico", 15, 45.99, 5),
        ("P002", "Ratón inalámbrico", 8, 22.50, 10),
        ("P003", "Monitor 24\"", 3, 189.00, 2),
        ("P004", "Auriculares USB", 20, 35.00, 5),
        ("P005", "Webcam HD", 6, 49.99, 4),
    ]
    for fila, (cod, prod, qty, precio, stock_min) in enumerate(productos, 2):
        ws.cell(row=fila, column=1, value=cod)
        ws.cell(row=fila, column=2, value=prod)
        ws.cell(row=fila, column=3, value=qty)
        ws.cell(row=fila, column=4, value=precio)
        c_val = ws.cell(row=fila, column=5)
        c_val.value = f"=C{fila}*D{fila}"
        _estilo_formula(c_val)
        ws.cell(row=fila, column=6, value=stock_min)
        c_alert = ws.cell(row=fila, column=7)
        c_alert.value = f'=SI(C{fila}<=F{fila},"🔴 Reponer","✅ OK")'
        _estilo_formula(c_alert)

    fila_total = len(productos) + 2
    ws.cell(row=fila_total, column=2, value="VALOR TOTAL STOCK")
    c_tot = ws.cell(row=fila_total, column=5)
    c_tot.value = f"=SUMA(E2:E{fila_total - 1})"
    _estilo_total(c_tot)

    _ajustar_columnas(ws)
    return wb, "plantilla_inventario.xlsx"
