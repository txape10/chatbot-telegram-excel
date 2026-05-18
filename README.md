# 🤖 Chatbot de Telegram — Asistente Personal de Excel

Asistente personal de Excel en Telegram con IA. Responde preguntas, analiza archivos, modifica datos y genera documentos, todo desde el chat — incluso por voz.

Funciona en local (PC propio) y es accesible desde cualquier dispositivo con Telegram mientras el PC esté encendido.

---

## Funcionalidades principales

### Conversación y consultas
- Responde cualquier pregunta sobre Excel en español con ejemplos prácticos
- Mantiene el hilo de la conversación por usuario (historial en SQLite)
- Se adapta a la versión de Excel del usuario (365, 2021, 2019, 2016)
- Escribe `=FORMULA(...)` y el bot la explica paso a paso
- **Voz**: envía un mensaje de voz → transcripción con Groq Whisper → respuesta en texto o también por voz (configurable con `/modo`)

### Análisis de archivos
- Sube un `.xlsx`, `.xls` o `.csv` y el bot responde:
  - Resumen: filas, columnas, nulos, duplicados
  - Calidad de datos: outliers, mezcla de tipos, fechas inválidas, columnas constantes
  - Gráfico automático (barras, líneas o sectores según el contenido)
  - Soporte multi-hoja con selector inline
- Sube una **captura de pantalla** de Excel → análisis con visión IA

### Consultas sobre datos en lenguaje natural
Con un archivo activo puedes preguntar directamente:
- "¿Cuánto suma Ventas por Región?"
- "Muéstrame el top 5 por Importe"
- "¿Cuántos pedidos hay con estado Pendiente?"

El bot usa un motor DSL interno (sin ejecutar código arbitrario) que soporta: filtrar, contar, sumar, promediar, agrupar, ordenar y top N, con filtros encadenables.

### Modificación de archivos
Pide cambios en lenguaje natural sobre el archivo que tienes activo:

| Petición de ejemplo | Operación |
|---|---|
| "Añade una columna Margen que sea Precio × 0,3" | Nueva columna calculada |
| "Ordena por Fecha descendente" | Ordenar |
| "Elimina los duplicados" | Eliminar duplicados (con confirmación) |
| "Rellena los vacíos de Categoría con 'Sin categoría'" | Rellenar nulos |
| "Elimina la columna Notas" | Eliminar columna (con confirmación) |
| "Renombra 'Impt' a 'Importe'" | Renombrar columna |
| "Colorea en rojo las ventas menores de 100" | Formato condicional |
| "Filtra los pedidos de Madrid y expórtalos" | Filtrar y exportar (con confirmación) |
| "Busca 'Enero' y reemplaza por 'January'" | Buscar y reemplazar |
| "Divide Nombre por espacio en Nombre y Apellido" | Dividir columna |
| "Concatena Nombre y Apellido en NombreCompleto" | Concatenar columnas |

El bot envía el `.xlsx` modificado y actualiza el archivo activo en memoria. Las operaciones destructivas piden confirmación antes de ejecutarse.

### Deshacer
- "Deshacer" / "revertir" → restaura el DataFrame al estado anterior a la última edición

### Creación de archivos
- "Hazme un Excel con columnas Fecha, Concepto, Importe y Categoría" → genera y envía el `.xlsx`
- "Crea una tabla de inventario con Referencia, Producto, Stock, Precio unitario y Total" → ídem

### Análisis estadístico y tendencias
- "Dame estadísticas del archivo" → media, mediana, mín/máx, desviación estándar, percentiles P25/P75 y detección de sesgo por cada columna numérica
- "Muéstrame las correlaciones" → ranking de pares más correlacionados + imagen heatmap
- "Analiza la tendencia" → regresión lineal por columna numérica, R², variación porcentual y gráfico con línea de tendencia

### Gráficos personalizados
- "Hazme un gráfico de barras de Ventas por Mes" → gráfico a medida (barras, líneas, sectores o dispersión) con agrupación si se pide

### Combinación y comparación de archivos
Sube dos archivos y únelos o compáralos:
- "Une los dos archivos por ID" → inner join
- "Combina con todos los clientes" → left join
- "Compara los dos archivos" → informe de diferencias (columnas, filas únicas/compartidas) + `.xlsx` de diff

El bot detecta automáticamente la columna común y gestiona columnas duplicadas con sufijos `_A` / `_B`.

### Modificaciones avanzadas

| Petición de ejemplo | Operación |
|---|---|
| "Limpia los textos de Categoría (mayúsculas)" | Normalizar texto (strip/upper/lower/title) |
| "Convierte la columna Fecha a formato fecha" | Estandarizar fechas (detección automática de formato) |
| "Convierte las columnas de meses en filas" | Despivotear (melt / unpivot) |
| "Agrupa Vendedor en columnas con suma de Ventas" | Pivotear (pivot_table) |

### Exploración rápida
- "Muéstrame las primeras 10 filas" → previsualización en bloque de código
- "Qué valores únicos hay en Categoría" → lista de valores únicos de la columna
- "Explícame el archivo" → análisis narrativo completo con el LLM
- "Exporta como CSV" → envía el DataFrame activo como `.csv` UTF-8

### Macros personales
Guarda secuencias de operaciones con nombre y ejecútalas cuando quieras:
- "Guarda esta macro como LimpiarFechas" → el LLM convierte la descripción en operaciones DSL y las guarda en SQLite
- "Ejecuta la macro LimpiarFechas" → aplica las operaciones en secuencia
- "Lista mis macros" / "Borra la macro LimpiarFechas"

### Tabla dinámica
- "Tabla dinámica" → genera un `.xlsx` con la hoja de datos como **Excel Table** (con filtros activos) y una segunda hoja con resúmenes estáticos por agrupación. Para la TD interactiva: `Insertar → Tabla dinámica → Aceptar` en Excel.

---

## Comandos

| Comando | Descripción |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/ayuda` | Menú de categorías con botones |
| `/ejemplo BUSCARV` | Explica una función de Excel (aleatorio si no se especifica) |
| `/generar BUSCARV` | Genera un `.xlsx` de ejemplo para esa función |
| `/plantilla` | Plantillas listas: presupuesto, gastos, KPIs, inventario |
| `/pivote` | Genera archivo preparado para tabla dinámica |
| `/version` | Configura tu versión de Excel |
| `/modo` | Elige si las respuestas son por voz o solo texto |
| `/estado` | Muestra el estado actual de la sesión (archivo activo, historial, modo…) |
| `/privado` | Activa/desactiva el modo privado (sin historial en SQLite) |
| `/limpiar` | Borra historial y contexto del archivo activo |

---

## Stack técnico

| Componente | Tecnología |
|---|---|
| Lenguaje | Python 3.11+ |
| Bot framework | python-telegram-bot v21 |
| LLM texto | Groq — llama-3.3-70b-versatile |
| LLM visión | Groq — meta-llama/llama-4-scout-17b-16e-instruct |
| STT (voz a texto) | Groq — whisper-large-v3-turbo |
| TTS (texto a voz) | edge-tts — es-ES-ElviraNeural (sin API key) |
| Excel / CSV | pandas + openpyxl |
| Gráficos | matplotlib |
| Persistencia | SQLite (historial, preferencias, metadatos, macros) |
| Coste | 0 € — Groq free tier |

---

## Instalación y puesta en marcha

### Requisitos previos
- Python 3.11 o superior
- Cuenta en [Groq Console](https://console.groq.com) (gratuita, sin tarjeta)
- Bot de Telegram creado con [@BotFather](https://t.me/BotFather)

### 1. Clonar el repositorio

```bash
git clone https://github.com/txape10/chatbot-telegram-excel.git
cd chatbot-telegram-excel
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 3. Configurar variables de entorno

```bash
cp .env.example .env
```

Edita `.env` con tus credenciales:

```env
TELEGRAM_TOKEN=token_obtenido_de_botfather
GROQ_API_KEY=clave_de_groq_console
AUTHORIZED_USERS=tu_user_id_de_telegram
```

Para obtener tu user ID de Telegram, envía un mensaje a [@userinfobot](https://t.me/userinfobot).

### 4. Arrancar el bot

```bash
python bot.py
```

El bot queda escuchando. Desde Telegram, envía `/start` para comenzar.

---

## Estructura del proyecto

```
├── bot.py                  ← Punto de entrada
├── config.py               ← Variables de entorno y límites de seguridad
├── handlers/
│   ├── messages.py         ← Detección de intención por regex + LLM; procesar_pregunta()
│   ├── commands.py         ← Comandos /cmd y callbacks InlineKeyboard
│   ├── documents.py        ← Procesamiento de archivos Excel/CSV subidos
│   ├── images.py           ← Análisis de capturas de pantalla con visión IA
│   └── audio.py            ← Mensajes de voz y audio: transcripción Whisper + preferencia de modo
├── excel/
│   ├── reader.py           ← Lectura de .xlsx y .csv
│   ├── analyzer.py         ← Resumen, calidad, estadísticas, correlaciones, tendencias, comparar
│   ├── query_engine.py     ← Motor DSL de consultas (9 operaciones)
│   ├── editor.py           ← Motor de edición (15 operaciones + exportar_xlsx)
│   ├── charts.py           ← Gráficos PNG automáticos y personalizados
│   └── exporter.py         ← Ejemplos, plantillas, crear desde descripción, tabla dinámica
├── services/
│   ├── llm.py              ← Groq: texto, visión, STT, DSLs de consulta/edición/gráfico/macro
│   └── tts.py              ← Síntesis de voz con edge-tts
├── utils/
│   ├── history.py          ← Historial de conversación (SQLite)
│   ├── df_context.py       ← DataFrame activo + secundario + undo por usuario
│   ├── excel_context.py    ← Contexto textual del archivo para el LLM
│   ├── file_meta.py        ← Metadatos del último archivo subido (SQLite)
│   ├── user_prefs.py       ← Preferencias: versión Excel, modo respuesta, modo privado (SQLite)
│   ├── macros.py           ← Macros personales por usuario (SQLite)
│   └── auth.py             ← Whitelist de acceso por user_id
├── prompts/
│   └── excel.py            ← Todas las plantillas de texto enviadas al LLM
├── tests/                  ← 136 tests unitarios (pytest)
├── knowledge/              ← Base de conocimiento en Markdown
└── data/
    ├── historial.db        ← SQLite
    └── logs/bot.log        ← Log rotativo (5 MB × 3 backups)
```

---

## Tests

```bash
pytest
```

136 tests unitarios que cubren: lectura de archivos, análisis de calidad, motor DSL, editor de archivos (incl. normalización, fechas, pivot/unpivot), combinación de archivos, tendencias, generación de Excel y plantillas.

---

## Notas de seguridad

- Acceso restringido por whitelist de `user_id` (variable `AUTHORIZED_USERS`)
- Validación de tipo de archivo por magic bytes (no solo por extensión)
- Límites de tamaño: MAX_FILAS, MAX_COLUMNAS, MAX_HOJAS configurables en `config.py`
- Las operaciones de edición y consulta usan DSL cerrada: el LLM no ejecuta código Python arbitrario
- Las operaciones destructivas piden confirmación antes de ejecutarse
- Modo privado: el historial no se persiste en SQLite cuando está activo
- El `.env` nunca se sube al repositorio (está en `.gitignore`)

---

## Roadmap

- [x] Conversación y historial
- [x] Análisis de archivos Excel/CSV
- [x] Motor de consultas DSL en lenguaje natural
- [x] Editor de archivos (15 operaciones, incl. buscar/reemplazar, dividir, concatenar)
- [x] Creación de Excel desde descripción
- [x] Análisis estadístico, correlaciones y tendencias
- [x] Tabla dinámica (Excel Table + resúmenes)
- [x] Normalización de texto y estandarización de fechas
- [x] Pivot / unpivot de tablas
- [x] Combinación de dos archivos (inner/left/right/outer join)
- [x] Comparación de dos archivos con informe de diferencias
- [x] Entrada por voz (Groq Whisper STT)
- [x] Respuestas por voz (edge-tts TTS, configurable)
- [x] Gráficos personalizados bajo demanda
- [x] Deshacer la última operación
- [x] Confirmaciones para operaciones destructivas
- [x] Previsualización de filas y valores únicos
- [x] Explícame archivo + Exportar CSV
- [x] Macros personales (guardar, ejecutar, listar, borrar)
- [x] Modo privado (sin historial)
- [x] `/estado`, `/modo`, `/privado`
- [ ] Tablas dinámicas interactivas nativas (evaluando xlwings vs XML injection)
- [ ] Despliegue en Railway/Render para disponibilidad 24/7
