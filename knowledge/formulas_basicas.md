# Fórmulas Básicas de Excel

## SUMA y variantes
```
=SUMA(A1:A10)                          → suma un rango
=SUMAR.SI(A1:A10,">=100")             → suma si cumple condición
=SUMAR.SI.CONJUNTO(C1:C10,A1:A10,"Ventas",B1:B10,"Norte")  → suma con múltiples condiciones
```
**Ejemplo práctico**: Tienes ventas por departamento en columna A y montos en B. Para sumar solo las de "Ventas": `=SUMAR.SI(A2:A100,"Ventas",B2:B100)`

---

## SI y variantes
```
=SI(A1>100,"Alto","Bajo")                          → condición simple
=SI(Y(A1>0,B1>0),"Ambos positivos","No")          → con Y (AND)
=SI(O(A1="Sí",B1="Sí"),"Al menos uno","Ninguno") → con O (OR)
=SI.ERROR(BUSCARV(...),"No encontrado")            → capturar errores
=IFS(A1<5,"Bajo",A1<8,"Medio",A1>=8,"Alto")      → múltiples condiciones (Excel 2019+)
```

---

## BUSCARV y BUSCARH
```
=BUSCARV(valor_buscado, rango_tabla, num_columna, [exacto])
=BUSCARV(A2, $E$2:$G$100, 2, FALSO)   → busca A2 en tabla E:G, devuelve col 2
```
**Limitación importante**: solo busca hacia la derecha. Para buscar en cualquier dirección usar ÍNDICE/COINCIDIR.

**Ejemplo**: Buscar el precio de un producto por su código:
```
=BUSCARV(B2, Productos!$A:$C, 3, FALSO)
```

---

## CONTAR y variantes
```
=CONTAR(A1:A10)          → cuenta celdas con números
=CONTARA(A1:A10)         → cuenta celdas no vacías
=CONTAR.SI(A1:A10,">0") → cuenta si cumple condición
=CONTAR.SI.CONJUNTO(A1:A10,"Sí",B1:B10,"Norte")  → múltiples condiciones
=CONTAR.BLANCO(A1:A10)  → cuenta celdas vacías
```

---

## Texto
```
=CONCATENAR(A1," ",B1)          → unir texto (mejor usar &)
=A1&" "&B1                      → forma moderna de concatenar
=IZQUIERDA(A1,5)               → primeros 5 caracteres
=DERECHA(A1,3)                 → últimos 3 caracteres
=EXTRAE(A1,3,4)                → desde posición 3, toma 4 caracteres
=LARGO(A1)                     → longitud del texto
=MAYUSC(A1) / MINUSC(A1)      → convertir mayúsculas/minúsculas
=NOMPROPIO(A1)                 → Primera Letra En Mayúscula
=ESPACIOS(A1)                  → eliminar espacios extra
=SUSTITUIR(A1,"viejo","nuevo") → reemplazar texto
=ENCONTRAR("x",A1)             → posición de un carácter
=TEXTO(A1,"dd/mm/yyyy")        → formatear número/fecha como texto
```

---

## Fechas
```
=HOY()                         → fecha actual
=AHORA()                       → fecha y hora actual
=DIA(A1) / MES(A1) / AÑO(A1) → extraer partes de fecha
=FECHA(2024,12,31)             → construir fecha
=DIAS(B1,A1)                   → días entre dos fechas
=DIAS.LAB(A1,B1)               → días laborables entre dos fechas
=FIN.MES(A1,0)                 → último día del mes de A1
=FIN.MES(A1,1)                 → último día del mes siguiente
=DIA.SEM(A1,2)                 → día de la semana (2=lunes=1)
```

---

## Matemáticas y estadística
```
=REDONDEAR(A1,2)               → redondear a 2 decimales
=REDONDEAR.MAS(A1,0)          → siempre redondea hacia arriba
=REDONDEAR.MENOS(A1,0)        → siempre redondea hacia abajo
=ENTERO(A1)                    → parte entera
=RESIDUO(A1,B1)               → resto de la división
=POTENCIA(A1,2)               → elevar al cuadrado
=RAIZ(A1)                     → raíz cuadrada
=ABS(A1)                      → valor absoluto
=PROMEDIO(A1:A10)             → media aritmética
=MEDIANA(A1:A10)              → mediana
=MODA(A1:A10)                 → valor más frecuente
=MAX(A1:A10) / MIN(A1:A10)   → máximo y mínimo
=K.ESIMO.MAYOR(A1:A10,2)     → segundo mayor valor
=DESVEST(A1:A10)              → desviación estándar
```

---

## Búsqueda y referencia
```
=INDIRECTO("A"&B1)            → referencia dinámica construida con texto
=DESREF(A1,2,3)               → celda desplazada 2 filas y 3 columnas
=FILA() / COLUMNA()           → número de fila/columna actual
=FILAS(A1:A10) / COLUMNAS()  → contar filas/columnas de un rango
=COINCIDIR(A1,B1:B10,0)      → posición de un valor en un rango
=ELEGIR(2,"Uno","Dos","Tres") → devuelve el elemento en posición 2
```
