# VBA Básico en Excel — Macros

## Qué es VBA y cuándo usarlo

VBA (Visual Basic for Applications) permite automatizar tareas repetitivas en Excel. Úsalo cuando:
- Repites los mismos pasos manualmente cada día/semana
- Necesitas procesar muchas filas con lógica compleja
- Quieres crear botones que ejecuten acciones
- Las fórmulas no son suficientes

---

## Acceder al editor VBA

- **Alt + F11** → abre el editor
- **Alt + F8** → lista de macros disponibles
- Pestaña Desarrollador → Editor de Visual Basic (activar en: Archivo → Opciones → Personalizar cinta)

---

## Estructura básica de una macro

```vba
Sub NombreDeLaMacro()
    ' Esto es un comentario
    ' El código va aquí
    MsgBox "¡Hola desde VBA!"
End Sub
```

---

## Operaciones más comunes

### Trabajar con celdas
```vba
' Leer y escribir valores
Range("A1").Value = "Hola"
Cells(1, 1).Value = "Hola"        ' Cells(fila, columna)

' Leer valor de una celda
Dim valor As String
valor = Range("B2").Value

' Formatear celda
Range("A1").Font.Bold = True
Range("A1").Interior.Color = RGB(255, 255, 0)   ' fondo amarillo
Range("A1").Font.Color = RGB(255, 0, 0)          ' texto rojo
Range("A1").NumberFormat = "#,##0.00"             ' formato número
```

### Última fila con datos
```vba
Dim ultimaFila As Long
ultimaFila = Cells(Rows.Count, "A").End(xlUp).Row
' Equivalente a Ctrl+Fin en columna A
```

### Bucle por filas
```vba
Dim i As Long
For i = 2 To ultimaFila
    If Cells(i, 1).Value > 100 Then
        Cells(i, 2).Value = "Alto"
    Else
        Cells(i, 2).Value = "Bajo"
    End If
Next i
```

### Copiar y pegar
```vba
Range("A1:C10").Copy Destination:=Range("E1")
Range("A1:C10").Copy
Range("E1").PasteSpecial Paste:=xlPasteValues   ' solo valores, sin fórmulas
Application.CutCopyMode = False                  ' quitar el borde parpadeante
```

---

## Macros más útiles y frecuentes

### Eliminar filas en blanco
```vba
Sub EliminarFilasEnBlanco()
    Dim i As Long
    For i = Cells(Rows.Count, "A").End(xlUp).Row To 1 Step -1
        If Cells(i, 1).Value = "" Then
            Rows(i).Delete
        End If
    Next i
End Sub
```

### Copiar datos a otra hoja
```vba
Sub CopiarAOtraHoja()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Set wsOrigen = Sheets("Datos")
    Set wsDestino = Sheets("Resumen")
    
    wsOrigen.Range("A1:D100").Copy wsDestino.Range("A1")
End Sub
```

### Guardar como PDF
```vba
Sub GuardarComoPDF()
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:="C:\Informes\informe.pdf", _
        Quality:=xlQualityStandard
    MsgBox "PDF guardado correctamente"
End Sub
```

### Resaltar duplicados
```vba
Sub ResaltarDuplicados()
    Dim rng As Range
    Dim celda As Range
    Set rng = Range("A2:A100")
    
    For Each celda In rng
        If celda.Value <> "" Then
            If WorksheetFunction.CountIf(rng, celda.Value) > 1 Then
                celda.Interior.Color = RGB(255, 200, 200)
            End If
        End If
    Next celda
End Sub
```

---

## Cuadros de mensaje e input

```vba
' Mostrar mensaje
MsgBox "Proceso completado"
MsgBox "¿Continuar?", vbYesNo + vbQuestion, "Confirmar"

' Pedir datos al usuario
Dim nombre As String
nombre = InputBox("¿Cuál es tu nombre?", "Nombre", "Por defecto")

' Capturar respuesta Yes/No
If MsgBox("¿Eliminar datos?", vbYesNo) = vbYes Then
    ' código si el usuario dice Sí
End If
```

---

## Manejo de errores

```vba
Sub MacroSegura()
    On Error GoTo ManejadorError
    
    ' código principal
    Range("A1").Value = 1 / 0   ' causará error
    
    Exit Sub

ManejadorError:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
```

---

## Consejos importantes

- **Grabar macro primero**: Desarrollador → Grabar macro → haz los pasos manualmente → Detener. Luego edita el código generado.
- **Activar macros**: Excel puede bloquear macros por seguridad. En archivos propios, guardar como `.xlsm` (con macros).
- **No usar `.Select` ni `.Activate`**: código grabado lleno de `.Select` es lento y frágil. Referenciar directamente: `Range("A1").Value` en vez de `Range("A1").Select: Selection.Value`.
- **Velocidad**: para macros lentas, añadir al inicio: `Application.ScreenUpdating = False` y al final: `Application.ScreenUpdating = True`.
