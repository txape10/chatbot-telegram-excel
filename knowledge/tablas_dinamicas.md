# Tablas Dinámicas en Excel

## Qué son y cuándo usarlas

Las tablas dinámicas permiten resumir, analizar y explorar grandes volúmenes de datos en segundos, sin escribir fórmulas. Son ideales cuando tienes una tabla de datos plana y quieres ver totales, medias o recuentos agrupados por categorías.

---

## Crear una tabla dinámica — Pasos

1. Haz clic en cualquier celda de tus datos
2. **Insertar → Tabla dinámica**
3. Elige dónde colocarla (nueva hoja recomendado)
4. En el panel derecho, arrastra campos a las 4 zonas:
   - **Filas**: las categorías por las que agrupar (ej: Departamento)
   - **Columnas**: segunda dimensión (ej: Mes)
   - **Valores**: qué calcular (ej: Suma de Ventas)
   - **Filtros**: filtro global (ej: Año)

---

## Tipos de cálculo en Valores

Clic derecho sobre un valor → "Resumir valores por":
- Suma, Promedio, Recuento, Máx, Mín, Producto
- Desviación estándar, Varianza

"Mostrar valores como":
- % del total general
- % del total de fila / columna
- Diferencia respecto a...
- Acumulado
- Clasificación de mayor a menor

---

## Trucos esenciales

**Actualizar datos**: clic derecho en la tabla → Actualizar (o Alt+F5)

**Agrupar fechas automáticamente**: arrastra una fecha a Filas → Excel agrupa por año/trimestre/mes automáticamente. Clic derecho → Agrupar para personalizar.

**Segmentadores de datos**: Herramientas de tabla dinámica → Analizar → Insertar Segmentación. Crea botones visuales para filtrar con un clic.

**Escala de tiempo**: igual pero específico para fechas. Permite filtrar por año, trimestre, mes o día deslizando una barra.

**Campo calculado**: Analizar → Campos, elementos y conjuntos → Campo calculado. Crea métricas nuevas: `= Ventas / Unidades` para obtener precio medio.

---

## Diseño y formato

- **Diseño de informe**: compacto (por defecto) / esquema / tabular
- **Totales**: puedes activar/desactivar totales de fila y columna
- **Subtotales**: mostrar u ocultar por nivel de agrupación
- **Filas en blanco**: insertar entre grupos para mejor legibilidad
- **Estilos**: pestaña Diseño → galería de estilos predefinidos

---

## Gráfico dinámico

Analizar → Gráfico dinámico. Se conecta directamente a la tabla y se actualiza con ella. Admite segmentadores compartidos.

---

## Tabla dinámica desde múltiples tablas (Power Pivot)

Si tienes datos en varias hojas relacionadas:
1. Convierte cada hoja en Tabla (Ctrl+T)
2. Datos → Obtener y transformar → Desde tabla
3. En Power Pivot, define relaciones entre tablas
4. Crea tabla dinámica usando el Modelo de datos

---

## Errores frecuentes y soluciones

| Problema | Causa | Solución |
|---|---|---|
| Los números se cuentan en vez de sumar | Hay celdas con texto en la columna | Limpiar datos, convertir a número |
| Fechas no se agrupan | Hay celdas con texto en vez de fecha real | Usar FECHANUMERO() para convertir |
| No aparecen todos los datos | Filtros activos | Revisar zona de Filtros y limpiar |
| Valores duplicados en filas | Espacios invisibles en los datos | Usar ESPACIOS() para limpiar |
