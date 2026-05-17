# Fórmulas Avanzadas de Excel

## ÍNDICE + COINCIDIR (alternativa superior a BUSCARV)

La combinación más poderosa de Excel. Busca en cualquier dirección.

```
=ÍNDICE(rango_resultado, COINCIDIR(valor_buscado, rango_búsqueda, 0))
```

**Ejemplo**: Buscar el salario de una persona por su nombre, donde el nombre está en columna B y el salario en columna A (BUSCARV no podría hacer esto):
```
=ÍNDICE(A2:A100, COINCIDIR("Juan", B2:B100, 0))
```

**Con doble COINCIDIR (búsqueda en matriz)**:
```
=ÍNDICE(B2:E10, COINCIDIR("Juan",A2:A10,0), COINCIDIR("Marzo",B1:E1,0))
```

---

## LAMBDA (Excel 365 — crear funciones propias)

Permite crear funciones reutilizables sin VBA.

```
=LAMBDA(x, x*x)               → función que eleva al cuadrado
=LAMBDA(base,altura, base*altura/2)  → área de triángulo
```

**Uso real**: definir en Administrador de nombres como "AreaTriangulo" y llamar con:
```
=AreaTriangulo(B2, C2)
```

---

## LET (asignar nombres a cálculos intermedios)

Evita repetir fórmulas largas y mejora la legibilidad.

```
=LET(
    ventas, SUMA(B2:B100),
    objetivo, 50000,
    porcentaje, ventas/objetivo,
    TEXTO(porcentaje,"0.0%")
)
```

---

## XLOOKUP / BUSCARX (Excel 365)

Reemplaza a BUSCARV, BUSCARH e ÍNDICE/COINCIDIR con una sola función.

```
=BUSCARX(valor, rango_búsqueda, rango_resultado, [si_no_encontrado], [modo])
=BUSCARX("Juan", A2:A100, C2:C100, "No encontrado")
```

**Ventajas sobre BUSCARV**:
- Busca en cualquier dirección
- Devuelve múltiples columnas a la vez
- Tiene valor por defecto si no encuentra

---

## FILTER / FILTRAR (Excel 365)

Devuelve un rango filtrado dinámicamente (¡sin tablas dinámicas!).

```
=FILTRAR(A2:C100, B2:B100="Ventas")           → filas donde B="Ventas"
=FILTRAR(A2:C100, (B2:B100="Ventas")*(C2:C100>1000))  → múltiples condiciones
=FILTRAR(A2:C100, B2:B100="Ventas", "Sin datos")       → con mensaje si vacío
```

---

## UNIQUE / ÚNICOS (Excel 365)

Devuelve valores únicos de un rango.

```
=ÚNICOS(A2:A100)               → lista sin duplicados
=ÚNICOS(A2:C100, FALSO, FALSO) → filas únicas
```

---

## SORT / ORDENAR (Excel 365)

Ordena un rango dinámicamente.

```
=ORDENAR(A2:C100)              → ordena por primera columna, ascendente
=ORDENAR(A2:C100, 3, -1)      → ordena por columna 3, descendente
```

---

## SEQUENCE / SECUENCIA (Excel 365)

Genera una secuencia de números.

```
=SECUENCIA(10)                 → números del 1 al 10
=SECUENCIA(5,3)               → matriz 5 filas x 3 columnas
=SECUENCIA(12,1,1,1)          → meses del 1 al 12
```

---

## Fórmulas matriciales clásicas (Ctrl+Shift+Enter)

Para versiones antiguas sin las funciones 365.

**Contar con múltiples condiciones** (alternativa a CONTAR.SI.CONJUNTO):
```
{=SUMA((A2:A100="Ventas")*(B2:B100="Norte"))}
```

**Sumar los N mayores valores**:
```
{=SUMA(K.ESIMO.MAYOR(A2:A100,FILA(INDIRECTO("1:5"))))}   → suma top 5
```

---

## SUMAPRODUCTO (versátil y sin Ctrl+Shift+Enter)

El "todo terreno" de Excel. Multiplica arrays y suma el resultado.

```
=SUMAPRODUCTO((A2:A100="Ventas")*(B2:B100="Norte")*C2:C100)
→ suma C donde A="Ventas" Y B="Norte"

=SUMAPRODUCTO((A2:A100="Ventas")/CONTAR.SI(A2:A100,A2:A100))
→ contar valores únicos que sean "Ventas"
```

---

## Fórmulas de texto avanzadas

```
=UNIRCADENAS(", ",VERDADERO,A2:A10)   → une con separador, ignora vacíos
=DIVIDIRTEXTO(A1,",")                  → divide texto en celdas (Excel 365)
=TEXTOANTES(A1,"@")                    → texto antes de "@" (Excel 365)
=TEXTODESPUES(A1,"@")                  → texto después de "@" (Excel 365)
```

---

## Validación de datos con fórmulas

En Datos → Validación → Personalizada:
```
=Y(A1>0, A1<1000)              → solo números entre 0 y 1000
=ESNUMERO(A1)                  → solo permite números
=LARGO(A1)<=10                 → máximo 10 caracteres
=CONTAR.SI($A$1:$A$100,A1)=1  → no permite duplicados
```
