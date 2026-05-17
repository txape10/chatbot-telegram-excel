# Power Query en Excel

## Qué es y cuándo usarlo

Power Query es el motor de transformación de datos de Excel. Úsalo cuando:
- Importas datos desde fuentes externas (CSV, web, bases de datos, otras hojas)
- Necesitas limpiar datos repetidamente (mismo proceso cada semana/mes)
- Quieres combinar datos de varias fuentes
- Las fórmulas de Excel son demasiado complejas para el tratamiento que necesitas

**Ruta**: Datos → Obtener y transformar datos

---

## Fuentes de datos soportadas

- Archivos: Excel, CSV, TXT, JSON, XML, PDF
- Bases de datos: SQL Server, Access, MySQL, PostgreSQL, Oracle
- Online: SharePoint, OneDrive, Web (scraping básico)
- Otras: carpeta entera de archivos, API REST

---

## Transformaciones más usadas

### Limpiar datos
- **Quitar filas en blanco**: Inicio → Quitar filas → Quitar filas en blanco
- **Quitar duplicados**: Inicio → Quitar filas → Quitar duplicados
- **Recortar espacios**: clic derecho en columna → Transformar → Recortar
- **Limpiar caracteres no imprimibles**: clic derecho → Transformar → Limpiar
- **Cambiar tipo de dato**: clic en icono junto al nombre de columna

### Columnas
- **Dividir columna**: por delimitador, por número de caracteres, por posición
- **Columna desde ejemplos**: escribes el resultado deseado y Power Query deduce la transformación
- **Combinar columnas**: seleccionar varias → clic derecho → Combinar columnas
- **Columna condicional**: Agregar columna → Columna condicional (equivale a SI anidados)
- **Columna personalizada**: fórmulas M propias

### Filas y estructura
- **Filtrar filas**: flecha en encabezado de columna
- **Ordenar**: flecha en encabezado
- **Transponer**: Transformar → Transponer
- **Anular dinamización (Unpivot)**: convierte columnas en filas. Clave para datos en formato matriz → formato tabla
- **Dinamizar (Pivot)**: lo contrario, convierte valores únicos de una columna en columnas separadas

### Agrupar y agregar
Inicio → Agrupar por:
- Agrupar por una o varias columnas
- Calcular: suma, promedio, recuento, mín, máx, todos los valores

---

## Combinar consultas

### Combinar (Merge) — equivale a JOIN/BUSCARV
Inicio → Combinar → Combinar consultas:
- Combinación externa izquierda (la más común — como BUSCARV)
- Combinación interna (solo filas con coincidencia en ambas)
- Combinación externa completa (todas las filas de ambas)

### Anexar (Append) — apilar tablas verticalmente
Inicio → Combinar → Anexar consultas. Útil para unir datos de varios meses/años que tienen la misma estructura.

---

## Combinar todos los archivos de una carpeta

1. Datos → Obtener datos → Desde archivo → Desde carpeta
2. Selecciona la carpeta con todos los CSV/Excel
3. Power Query combina automáticamente todos los archivos
4. Añade una columna con el nombre del archivo de origen

---

## Lenguaje M (fórmulas de Power Query)

Cada transformación genera código M en segundo plano. Vista previa:
- Vista → Editor avanzado

Ejemplos de código M útil:

```m
// Filtrar filas donde columna > 100
= Table.SelectRows(#"Paso anterior", each [Ventas] > 100)

// Añadir columna condicional
= Table.AddColumn(tabla, "Categoría", each if [Ventas] > 1000 then "Alto" else "Bajo")

// Reemplazar valores
= Table.ReplaceValue(tabla, "Sí", "Yes", Replacer.ReplaceText, {"Columna1"})

// Cambiar tipo de dato
= Table.TransformColumnTypes(tabla, {{"Fecha", type date}, {"Ventas", type number}})
```

---

## Actualizar datos

- **Manual**: Datos → Actualizar todo (o clic derecho en la tabla → Actualizar)
- **Automático**: Propiedades de consulta → Actualizar cada X minutos
- **Al abrir el archivo**: Propiedades → Actualizar datos al abrir el archivo

---

## Consejos clave

- **Cada paso queda registrado**: en el panel "Pasos aplicados" puedes ver, editar o eliminar cualquier transformación
- **No modificar datos origen**: Power Query crea una copia, los datos originales no se tocan
- **Deshabilitar carga**: si solo usas la consulta como intermedia, clic derecho → Deshabilitar carga (no crea tabla en la hoja pero sí está disponible para otras consultas)
- **Renombrar pasos**: doble clic en cada paso para darle nombre descriptivo — mucho más fácil de mantener
