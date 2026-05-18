# 📖 Guía de usuario — Asistente de Excel en Telegram

Esta guía explica cómo sacar el máximo partido al bot. No hace falta saber programar: todo se hace en lenguaje natural desde Telegram.

---

## Índice

1. [Primeros pasos](#1-primeros-pasos)
2. [Configuración inicial](#2-configuración-inicial)
3. [Hacer preguntas sobre Excel](#3-hacer-preguntas-sobre-excel)
4. [Subir y analizar archivos](#4-subir-y-analizar-archivos)
5. [Consultar datos en lenguaje natural](#5-consultar-datos-en-lenguaje-natural)
6. [Modificar archivos](#6-modificar-archivos)
7. [Crear archivos Excel desde cero](#7-crear-archivos-excel-desde-cero)
8. [Análisis estadístico y gráficos](#8-análisis-estadístico-y-gráficos)
9. [Combinar y comparar archivos](#9-combinar-y-comparar-archivos)
10. [Macros personales](#10-macros-personales)
11. [Entrada y respuesta por voz](#11-entrada-y-respuesta-por-voz)
12. [Comandos de referencia](#12-comandos-de-referencia)
13. [Preguntas frecuentes](#13-preguntas-frecuentes)

---

## 1. Primeros pasos

Abre Telegram, busca el bot y envía el comando `/start`. Recibirás un mensaje de bienvenida con un resumen de todo lo que puedes hacer.

> 💡 El bot solo responde a usuarios autorizados. Si no recibes respuesta, contacta con el administrador para que añada tu user ID a la lista.

---

## 2. Configuración inicial

### Versión de Excel

El bot adapta sus respuestas a tu versión de Excel. Configúrala con:

```
/version
```

Aparecerán cuatro botones: **Microsoft 365**, **Excel 2021**, **Excel 2019** y **Excel 2016 o anterior**. Pulsa el tuyo.

> ⚠️ Algunas funciones como `BUSCARX`, `UNIRCADENAS` o matrices dinámicas solo están disponibles en versiones recientes. Si no configuras tu versión, el bot no podrá ajustar las fórmulas que te sugiere.

### Modo de respuesta

Elige si quieres que el bot te responda solo en texto o también con audio:

```
/modo
```

Pulsa **🔊 Por voz** o **💬 Solo texto**.

---

## 3. Hacer preguntas sobre Excel

Escribe cualquier duda directamente, como si hablaras con un compañero experto:

```
¿Cómo busco un valor en una tabla y devuelvo el dato de otra columna?
```

```
¿Cuál es la diferencia entre BUSCARV y BUSCARX?
```

```
¿Cómo congelo la primera fila en Excel?
```

### Explicar una fórmula paso a paso

Si pegas una fórmula que empieza por `=`, el bot la desglosa automáticamente:

```
=SI(Y(A2>100,B2="Madrid"),A2*0.1,0)
```

### Obtener un ejemplo concreto

```
/ejemplo BUSCARV
```

```
/ejemplo SUMAR.SI.CONJUNTO
```

Si no escribes ninguna función, el bot elige una al azar.

### Generar un archivo de ejemplo

```
/generar BUSCARV
```

Recibirás un `.xlsx` con datos de muestra y la función ya aplicada, listo para abrir en Excel.

---

## 4. Subir y analizar archivos

### Subir un archivo

Adjunta cualquier archivo `.xlsx`, `.xls` o `.csv` directamente en el chat (como si enviaras una foto). El bot responde automáticamente con:

- **Resumen**: número de filas, columnas, valores nulos y filas duplicadas
- **Calidad de datos**: detecta columnas casi vacías, mezcla de tipos, outliers estadísticos y fechas con formato incorrecto
- **Gráfico automático**: genera una imagen PNG con la visualización más adecuada para tus datos

### Archivos con varias hojas

Si el Excel tiene más de una hoja, el bot muestra botones para que elijas cuál analizar.

### Analizar una captura de pantalla

Envía una **imagen** de tu pantalla de Excel y el bot la analizará con visión IA:

> *"Analiza esta captura de Excel y explica qué hace, qué fórmulas usa y cómo podría mejorarla."*

También puedes añadir una pregunta concreta junto a la imagen:

> *"¿Por qué esta fórmula devuelve #¡VALOR!?"*

### Ver el estado de tu sesión

```
/estado
```

Muestra el archivo activo (filas × columnas), si hay un archivo secundario cargado, cuántos mensajes tiene el historial, el modo de respuesta y la versión de Excel configurada.

---

## 5. Consultar datos en lenguaje natural

Con un archivo activo, puedes hacer preguntas directamente sobre los datos:

| Lo que escribes | Lo que hace el bot |
|---|---|
| "¿Cuánto suma Ventas por Región?" | Agrupa y suma |
| "Muéstrame el top 5 por Importe" | Ordena y muestra los 5 primeros |
| "¿Cuántos pedidos hay con estado Pendiente?" | Filtra y cuenta |
| "¿Cuál es el promedio de Precio por Categoría?" | Agrupa y promedia |
| "¿Cuál es el máximo de Ventas en enero?" | Filtra por mes y calcula máximo |

El bot usa un motor interno estructurado, **no ejecuta código libre**, lo que garantiza que no se modifica nada sin tu permiso.

### Previsualizar filas

```
Muéstrame las primeras 10 filas
```

```
Dame las últimas 5 filas
```

### Ver valores únicos de una columna

```
¿Qué valores únicos hay en la columna Categoría?
```

```
Lista los productos distintos que aparecen
```

---

## 6. Modificar archivos

Pide cambios sobre el archivo que tienes activo. El bot aplica la operación, te envía el `.xlsx` actualizado y lo deja listo para seguir trabajando.

### Operaciones disponibles

#### Añadir una columna calculada
```
Añade una columna Margen que sea Precio multiplicado por 0,3
```
```
Crea una columna IVA con el 21% del campo Importe
```

#### Ordenar
```
Ordena por Fecha de forma descendente
```
```
Ordena primero por Región y luego por Ventas descendente
```

#### Rellenar valores vacíos
```
Rellena los vacíos de la columna Categoría con "Sin categoría"
```
```
Rellena los nulos de Precio con el valor 0
```

#### Renombrar columna
```
Renombra la columna "Impt" a "Importe"
```

#### Formato condicional
```
Colorea en rojo las celdas de Ventas menores de 100
```
```
Marca en verde los valores de Margen mayores de 500
```

#### Buscar y reemplazar
```
Busca "Ene" y reemplázalo por "Enero" en la columna Mes
```
```
Reemplaza todos los "N/A" por vacío en todo el archivo
```

#### Dividir una columna
```
Divide la columna Nombre completo por el espacio en Nombre y Apellido
```
```
Separa la columna Dirección por la coma en Calle y Ciudad
```

#### Concatenar columnas
```
Une Nombre y Apellido separados por un espacio en una columna NombreCompleto
```
```
Concatena Calle, Ciudad y País con ", " en una columna Dirección
```

#### Normalizar texto
```
Pon en mayúsculas toda la columna Categoría
```
```
Limpia espacios y pon en formato título la columna Nombre
```

#### Estandarizar fechas
```
Corrige el formato de las fechas de la columna Fecha
```
```
Estandariza todas las fechas al formato DD/MM/YYYY
```

#### Despivotar (columnas → filas)
```
Convierte las columnas de meses en filas
```
```
Transforma las columnas Enero, Febrero, Marzo en filas con columnas Mes y Valor
```

#### Pivotar (filas → columnas)
```
Crea una tabla con Vendedor en filas, Mes en columnas y suma de Ventas como valores
```

#### Eliminar duplicados ⚠️
```
Elimina los duplicados
```

#### Eliminar columna ⚠️
```
Elimina la columna Notas
```

#### Filtrar y exportar ⚠️
```
Filtra los pedidos de Madrid y expórtalos
```

> ⚠️ Las operaciones marcadas son **destructivas** (no se pueden deshacer fácilmente). El bot pedirá confirmación antes de ejecutarlas.

### Deshacer la última operación

Si el resultado no es el esperado:

```
Deshacer
```

```
Revertir el último cambio
```

El bot restaura el archivo al estado anterior y te lo envía.

---

## 7. Crear archivos Excel desde cero

Describe en lenguaje natural la tabla que necesitas:

```
Hazme un Excel con columnas Fecha, Concepto, Importe y Categoría para llevar mis gastos
```

```
Crea una tabla de inventario con Referencia, Producto, Stock mínimo, Stock actual y Precio unitario
```

```
Necesito una hoja de seguimiento de proyectos con Proyecto, Responsable, Fecha inicio, Fecha fin y Estado
```

### Plantillas predefinidas

```
/plantilla
```

Elige entre cuatro plantillas listas para usar: **Presupuesto personal**, **Control de gastos**, **KPIs de negocio** e **Inventario**.

### Tabla dinámica

```
/pivote
```

Genera un `.xlsx` con tu archivo activo formateado como **Excel Table** (con filtros activados) y una segunda hoja con resúmenes estáticos. Para crear la tabla dinámica interactiva en Excel: `Insertar → Tabla dinámica → Aceptar`.

---

## 8. Análisis estadístico y gráficos

### Estadísticas completas

```
Dame las estadísticas del archivo
```

```
Analiza numéricamente las columnas
```

Para cada columna numérica obtendrás: media, mediana, mínimo, máximo, desviación estándar, percentiles P25/P75 y detección de sesgo (si los datos están sesgados hacia valores altos o bajos).

### Mapa de correlaciones

```
Muéstrame las correlaciones entre columnas
```

```
Haz un mapa de calor
```

Recibirás un ranking de los pares de columnas más correlacionados y una imagen heatmap.

### Análisis de tendencia

```
¿Qué tendencia tienen mis ventas?
```

```
Analiza la evolución del campo Importe
```

Para cada columna numérica: regresión lineal, coeficiente R² (fiabilidad de la tendencia), variación porcentual entre el primer y último valor, e interpretación visual (📈 / 📉 / sin tendencia clara). Incluye gráfico con la línea de tendencia.

### Gráficos personalizados

```
Hazme un gráfico de barras de Ventas por Región
```

```
Quiero un gráfico de líneas de Ingresos por Mes
```

```
Muéstrame un gráfico de sectores con la distribución de Categoría
```

```
Dibuja un gráfico de dispersión de Precio vs Ventas
```

Tipos disponibles: **barras**, **líneas**, **sectores** (tarta) y **dispersión**.

### Análisis narrativo

```
Explícame el archivo
```

El bot genera un resumen en lenguaje natural del contenido, la calidad de los datos y los puntos de atención más relevantes.

---

## 9. Combinar y comparar archivos

### Combinar dos archivos

1. Sube el **primer archivo** → el bot lo analiza y lo deja como archivo activo
2. Sube el **segundo archivo** → el bot detecta que ya hay uno activo, mueve el primero a un slot secundario y te avisa
3. Pide la combinación:

```
Une los dos archivos por la columna ID
```

```
Combina con todos los clientes aunque no tengan pedidos (left join)
```

```
Fusiona incluyendo todos los registros de ambos archivos
```

Tipos de combinación disponibles:
- **inner** (por defecto): solo filas que existen en ambos archivos
- **left**: todas las filas del primer archivo, aunque no tengan correspondencia en el segundo
- **right**: todas las filas del segundo archivo
- **outer**: todas las filas de ambos archivos

Si hay columnas con el mismo nombre en los dos archivos, se añaden sufijos `_A` y `_B` automáticamente.

### Comparar dos archivos

Con dos archivos cargados:

```
Compara los dos archivos
```

```
¿Qué diferencias hay entre los dos Excel?
```

Recibirás un informe de texto con:
- Columnas presentes en uno y no en el otro
- Número de filas únicas de cada archivo
- Número de filas compartidas

Si hay diferencias, también recibirás un `.xlsx` con todas las filas etiquetadas por origen.

---

## 10. Macros personales

Las macros te permiten guardar una secuencia de operaciones con un nombre y ejecutarla cuando quieras, sin tener que repetir cada paso.

### Guardar una macro

Después de describir lo que quieres hacer (no es necesario haberlo ejecutado antes):

```
Guarda esta macro como LimpiarDatos: normaliza el texto de todas las columnas, estandariza las fechas y elimina duplicados
```

```
Crea una macro llamada PrepararInforme que ordene por Fecha descendente y filtre los registros de Madrid
```

### Ejecutar una macro

```
Ejecuta la macro LimpiarDatos
```

```
Aplica la macro PrepararInforme
```

### Ver tus macros guardadas

```
Lista mis macros
```

### Borrar una macro

```
Borra la macro LimpiarDatos
```

---

## 11. Entrada y respuesta por voz

### Enviar mensajes por voz

Pulsa el micrófono en Telegram y envía un mensaje de voz. El bot transcribirá el audio automáticamente y procesará tu petición igual que si la hubieras escrito.

> 💬 *"¿Cuánto suma la columna Ventas agrupado por Región?"*

Después de tu primer mensaje de voz, el bot te preguntará si prefieres también recibir las respuestas en audio.

### Configurar el modo de respuesta

En cualquier momento:

```
/modo
```

- **🔊 Por voz**: el bot envía texto + audio en cada respuesta
- **💬 Solo texto**: solo texto (más rápido, mejor para fórmulas y tablas)

> 💡 El modo voz siempre incluye también el texto, para que puedas leer fórmulas y bloques de código aunque estés escuchando.

---

## 12. Comandos de referencia

| Comando | Descripción |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/ayuda` | Menú de ayuda por categorías |
| `/ejemplo [FUNCIÓN]` | Explicación detallada de una función |
| `/generar [FUNCIÓN]` | Genera un `.xlsx` de ejemplo con esa función |
| `/plantilla` | Plantillas listas para usar |
| `/pivote` | Genera archivo para tabla dinámica |
| `/version` | Configura tu versión de Excel |
| `/modo` | Elige respuestas por voz o solo texto |
| `/estado` | Ver estado actual de la sesión |
| `/privado` | Activar/desactivar modo privado |
| `/limpiar` | Borrar historial y archivos activos |

### Modo privado

```
/privado
```

Cuando el modo privado está **activo**, el bot no guarda ningún mensaje en el historial. Las conversaciones no se almacenan en la base de datos. Para desactivarlo, vuelve a ejecutar `/privado`.

### Limpiar la sesión

```
/limpiar
```

Borra el historial de conversación y libera los archivos activos (activo, secundario y undo). Útil para empezar con un archivo nuevo sin que el bot confunda contextos.

---

## 13. Preguntas frecuentes

**¿El bot guarda mis archivos?**
No de forma permanente. Los archivos se procesan en memoria y se eliminan al limpiar la sesión o al reiniciar el bot. Solo se guardan en SQLite las preferencias de usuario, el historial de conversación y las macros que tú guardes explícitamente. Si activas el modo privado (`/privado`), ni siquiera el historial se guarda.

**¿Qué pasa si subo un archivo muy grande?**
El bot tiene límites configurables (por defecto: 500.000 filas, 200 columnas, 10 hojas). Si tu archivo los supera, recibirás un mensaje de aviso.

**¿El bot puede ejecutar código en mi PC?**
No. Todas las operaciones sobre datos usan un lenguaje de instrucciones estructurado (DSL). El LLM extrae la intención y la convierte en una operación concreta, pero nunca se ejecuta código arbitrario.

**¿Funciona si el PC está apagado?**
No. El bot corre en local. Para que esté disponible 24/7 habría que desplegarlo en un servidor (está en el roadmap).

**¿El bot entiende español con acento?**
Sí, tanto el texto como el audio. La transcripción de voz usa Groq Whisper con idioma configurado en español.

**¿Puedo usarlo desde el móvil?**
Sí. Al estar el bot en Telegram, funciona desde cualquier dispositivo (móvil, tablet, otro PC) donde tengas Telegram instalado, siempre que el PC donde corre el bot esté encendido.

**El bot tardó mucho en responder, ¿es normal?**
Las operaciones sobre archivos grandes (estadísticas, correlaciones, tendencias) pueden tardar unos segundos. El bot siempre envía un mensaje de "⏳ Procesando…" mientras trabaja. Las preguntas simples de texto suelen responderse en menos de 2 segundos.

**¿Cómo sé qué archivo está activo?**
Usa `/estado` para ver el nombre del archivo activo, sus dimensiones y si hay un archivo secundario o undo disponibles.
