# Formato Condicional en Excel

## Qué es

Permite aplicar formato automáticamente (color, negrita, iconos...) a celdas según su valor o una fórmula. Se actualiza en tiempo real cuando cambian los datos.

**Ruta**: Inicio → Estilos → Formato condicional

---

## Reglas predefinidas (las más rápidas)

### Resaltar reglas de celdas
- Mayor que / Menor que / Entre
- Igual a / Texto que contiene
- Fecha que corresponde a (hoy, ayer, esta semana...)
- Valores duplicados / únicos

### Reglas superiores/inferiores
- 10 valores superiores/inferiores
- 10% superiores/inferiores
- Por encima/debajo del promedio

### Barras de datos
Muestra una barra proporcional al valor dentro de cada celda. Muy útil para comparar rápidamente.

### Escalas de color
Degradado de color entre mínimo y máximo (ej: rojo → amarillo → verde).

### Conjuntos de iconos
Flechas, semáforos, estrellas, indicadores según rangos de valor.

---

## Reglas con fórmulas personalizadas (las más potentes)

Formato condicional → Nueva regla → "Usar una fórmula..."

**Clave**: la fórmula debe devolver VERDADERO o FALSO. Usar referencias mixtas correctamente.

### Ejemplos prácticos

**Resaltar fila completa según valor de una columna**:
```
=$C2="Completado"     → aplica a $A:$Z, resalta fila entera si col C = "Completado"
```

**Filas alternadas (sin tabla)**:
```
=RESIDUO(FILA(),2)=0  → colorea filas pares
```

**Resaltar duplicados en un rango**:
```
=CONTAR.SI($A$2:$A$100,A2)>1
```

**Resaltar la celda máxima de un rango**:
```
=A2=MAX($A$2:$A$100)
```

**Resaltar fines de semana en un calendario**:
```
=DIA.SEM(A1,2)>=6     → sábado=6, domingo=7
```

**Resaltar fechas vencidas**:
```
=Y(A2<HOY(), B2<>"Completado")    → vencida y no completada
```

**Resaltar si la celda contiene texto específico**:
```
=ESNUMERO(ENCONTRAR("urgente",MINUSC(A2)))
```

---

## Gestión de reglas

Formato condicional → Administrar reglas:
- Ver todas las reglas de la hoja o selección
- Cambiar el orden de prioridad (la primera regla que se cumple gana)
- Marcar "Detener si es verdad" para que no evalúe reglas siguientes

---

## Consejos importantes

- **Referencias relativas vs absolutas**: en fórmulas de FC, `$C2` fija la columna pero no la fila (para aplicar a filas completas). `$C$2` fijaría ambas (solo útil para comparar contra una celda fija).
- **Rendimiento**: muchas reglas complejas sobre rangos grandes pueden ralentizar el archivo. Limitar el rango al mínimo necesario.
- **Copiar formato condicional**: usar "Copiar formato" (brocha) o Pegado especial → Solo formato.
- **Borrar reglas**: Formato condicional → Borrar reglas → De las celdas seleccionadas / De toda la hoja.
