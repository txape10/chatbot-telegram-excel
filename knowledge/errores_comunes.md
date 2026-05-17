# Errores Comunes en Excel y Cómo Resolverlos

## Tabla resumen rápida

| Error | Significa | Causa más frecuente |
|---|---|---|
| `#¡VALOR!` | Tipo de dato incorrecto | Texto donde se espera número |
| `#N/A` | Valor no encontrado | BUSCARV no encuentra el dato |
| `#¡REF!` | Referencia inválida | Se eliminó una celda referenciada |
| `#¡DIV/0!` | División por cero | El divisor es 0 o está vacío |
| `#¿NOMBRE?` | Nombre no reconocido | Error tipográfico en la fórmula |
| `#¡NUM!` | Número inválido | Raíz de negativo, número demasiado grande |
| `#¡NULO!` | Intersección vacía | Espacio en vez de `:` o `,` en el rango |
| `######` | Columna muy estrecha | Ampliar el ancho de la columna |

---

## #¡VALOR! — Tipo de dato incorrecto

**Causas**:
- Sumar celdas que contienen texto
- Espacios invisibles en celdas numéricas
- Fechas almacenadas como texto

**Soluciones**:
```
=SUMAR.SI(A1:A10,"<>"&"")          → ignora texto al sumar
=ESPACIOS(A1)                       → elimina espacios extra
=VALOR(A1)                          → convierte texto a número
=SI.ERROR(A1+B1, "Error de datos") → captura el error
```

**Detectar si una celda es realmente número**:
```
=ESNUMERO(A1)    → devuelve VERDADERO si es número real
```

---

## #N/A — Valor no encontrado

**Causas**:
- BUSCARV no encuentra el valor buscado
- El valor existe pero con espacios o formato diferente
- Buscar número en columna de texto o viceversa

**Soluciones**:
```
=SI.ERROR(BUSCARV(A1,B:C,2,0), "No encontrado")
=BUSCARX(A1, B:B, C:C, "No encontrado")       → más limpio en Excel 365
```

**Diagnóstico**:
```
=ESPACIOS(A1)=A1                    → FALSO si hay espacios ocultos
=COINCIDIR(A1, B:B, 0)             → si devuelve #N/A, confirma que no existe
```

---

## #¡REF! — Referencia inválida

**Causas**:
- Se eliminó una fila, columna o hoja que estaba en la fórmula
- DESREF apunta fuera del rango válido
- Copiar fórmula que sale del borde de la hoja

**Solución**: revisar la fórmula y reconstruir la referencia. No hay atajo; hay que identificar qué se eliminó. Usar Ctrl+Z si fue reciente.

---

## #¡DIV/0! — División por cero

**Soluciones**:
```
=SI(B1=0, 0, A1/B1)                → devuelve 0 si el divisor es 0
=SI(B1="", "", A1/B1)              → devuelve vacío si B1 está vacío
=SI.ERROR(A1/B1, 0)                → captura cualquier error
=SIERROR(A1/B1,"Sin datos")        → versión antigua (Excel 2007+)
```

---

## #¿NOMBRE? — Nombre no reconocido

**Causas**:
- Error tipográfico en el nombre de la función (ej: `=SUMAS(A1:A10)`)
- Texto sin comillas dentro de la fórmula (ej: `=SI(A1=Sí,...)` en vez de `=SI(A1="Sí",...)`)
- Función de Excel 365 usada en versión anterior
- Nombre de rango inexistente

**Solución**: revisar ortografía, añadir comillas al texto, verificar compatibilidad de versión.

---

## #¡NUM! — Número inválido

**Causas**:
- `=RAIZ(-1)` — raíz de número negativo
- Número demasiado grande para Excel (>9.99×10^307)
- IRR/TIR sin solución convergente

**Solución**:
```
=SI(A1<0, "Negativo", RAIZ(A1))   → validar antes de calcular
```

---

## Errores silenciosos (los más peligrosos)

Son errores que no muestran mensaje pero dan resultados incorrectos:

**Fechas como texto**: parecen fechas pero Excel las trata como texto.
```
=ESNUMERO(A1)    → una fecha real devuelve VERDADERO
```

**Números como texto**: columna de números que no suma correctamente.
```
Seleccionar rango → Datos → Texto en columnas → Finalizar  (fuerza conversión)
```

**BUSCARV con coincidencia aproximada por defecto**:
```
=BUSCARV(A1,B:C,2)       ← PELIGROSO: último argumento omitido = VERDADERO (aproximado)
=BUSCARV(A1,B:C,2,FALSO) ← CORRECTO: siempre especificar FALSO para coincidencia exacta
```

---

## SI.ERROR vs SIERROR

```
=SI.ERROR(fórmula, valor_si_error)   → Excel 2007+ — captura CUALQUIER error
=SIERROR(fórmula, valor_si_error)    → idéntico, nombre alternativo
=SI.ND(fórmula, valor_si_nd)         → solo captura errores #N/A (más específico)
```

**Recomendación**: usar `SI.ND` en BUSCARV para no ocultar errores inesperados:
```
=SI.ND(BUSCARV(A1,B:C,2,FALSO), "No existe")
```
