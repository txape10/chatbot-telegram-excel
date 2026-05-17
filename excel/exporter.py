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
