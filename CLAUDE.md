# 🤖 Chatbot de Telegram — Asistente Personal de Excel

## WHY — Por qué existe este proyecto

Crear un asistente personal en Telegram que resuelva dudas sobre Microsoft Excel de forma inmediata, con ejemplos prácticos y explicaciones claras, disponible desde cualquier dispositivo (móvil, tablet, otro PC) siempre que el PC principal esté encendido.

## WHAT — Qué hace

Un bot de Telegram que:
- Responde preguntas sobre Excel en español con ejemplos prácticos
- Cubre fórmulas, tablas dinámicas, formato condicional, gráficos, macros/VBA, Power Query, etc.
- Mantiene contexto de la conversación (historial por usuario en SQLite)
- Analiza archivos Excel y CSV subidos por el usuario (resumen, calidad, gráfico automático)
- Consulta datos con lenguaje natural (motor DSL: filtrar, agrupar, sumar, ordenar…)
- Modifica archivos Excel a petición (añadir columnas, ordenar, buscar/reemplazar, dividir/concatenar columnas…)
- Crea archivos Excel desde descripción en lenguaje natural
- Genera resúmenes estadísticos, mapas de correlaciones y gráficos de tendencia
- Genera archivos de tabla dinámica (Excel Table + resúmenes estáticos)
- Genera gráficos personalizados bajo demanda (barras, líneas, sectores, dispersión)
- Explica fórmulas paso a paso si el mensaje empieza por `=`
- Analiza capturas de pantalla de Excel con visión IA
- Entiende mensajes de voz (Groq Whisper) y puede responder también por voz (edge-tts)
- Compara dos archivos Excel y devuelve un informe de diferencias
- Permite guardar y ejecutar macros personales (secuencias de operaciones con nombre)
- Es accesible desde cualquier dispositivo con Telegram instalado

## Principios Karpathy

- No asumo ni complico sin preguntar
- Código mínimo que resuelve el problema (sin abstracciones especulativas)
- Cambios quirúrgicos: solo toco lo solicitado, respeto código existente
- Define éxito con tests/verificación antes de empezar

## HOW — Stack técnico y decisiones tomadas

### Lenguaje
- **Python 3.11+**

### LLM
- **Groq** con modelo `llama-3.3-70b-versatile` — gratuito, sin tarjeta de crédito
- Análisis de imágenes: `meta-llama/llama-4-scout-17b-16e-instruct` (visión)
- Límite TPM gestionado dinámicamente: system prompt ~1.200 tokens, historial truncado automáticamente si se acerca a 9.000 tokens

### Librería de Telegram
- `python-telegram-bot` v21+

### Ejecución
- **Local en PC del usuario** (sin servidor externo de momento)
- El bot es accesible desde cualquier dispositivo con Telegram mientras el PC esté encendido
- Migración futura posible a Railway o Render (gratis, sin cambiar código)

### Dependencias principales
```
python-telegram-bot==21.10
groq==0.13.1
python-dotenv==1.1.0
pandas==2.2.3
openpyxl==3.1.5
matplotlib==3.9.4
pillow==11.2.1
edge-tts==6.1.12
```

## Estructura del proyecto

```
3 - Chatbot de Telegram para Excel/
├── CLAUDE.md               ← este archivo
├── .env                    ← tokens y claves (NO subir a git)
├── .env.example            ← plantilla sin valores reales
├── .gitignore
├── requirements.txt
├── logging_config.py       ← logging a consola + fichero rotativo (data/logs/bot.log)
├── bot.py                  ← punto de entrada principal
├── config.py               ← configuración, variables de entorno, whitelist y límites de seguridad
├── handlers/
│   ├── __init__.py
│   ├── messages.py         ← núcleo de detección de intención: procesar_pregunta() + todos los flujos
│   ├── commands.py         ← /start, /ayuda, /ejemplo, /generar, /plantilla, /version, /modo, /estado, /privado, /limpiar + callbacks InlineKeyboard
│   ├── documents.py        ← archivos .xlsx/.xls/.csv: validación, límites, multi-hoja, slot secundario, asyncio.to_thread
│   ├── images.py           ← análisis de capturas de pantalla con LLM de visión
│   └── audio.py            ← mensajes de voz y audio: transcripción Whisper → procesar_pregunta; pregunta preferencia de modo
├── services/
│   ├── __init__.py
│   ├── llm.py              ← Groq: texto (Llama 3.3), visión (Llama 4 Scout), STT (Whisper), DSLs de consulta/edición/gráfico/macro
│   │                          Gestión dinámica de tokens: truncado automático del historial
│   └── tts.py              ← síntesis de voz con edge-tts (es-ES-ElviraNeural); limpieza de Markdown antes de síntesis
├── excel/
│   ├── __init__.py
│   ├── reader.py           ← leer .xlsx (multi-hoja) y .csv con detección de separador
│   ├── analyzer.py         ← resumen, calidad, estadísticas, correlaciones, tendencias, comparar_dataframes
│   ├── query_engine.py     ← motor DSL: filtrar/contar/suma/promedio/max/min/agrupar/ordenar/top_n
│   ├── editor.py           ← 15 operaciones de edición + exportar_xlsx; buscar/reemplazar, dividir, concatenar, pivot/unpivot, normalizar, fechas, combinar
│   ├── charts.py           ← gráficos PNG automáticos y personalizados (barras/líneas/sectores/dispersión)
│   └── exporter.py         ← ejemplos, plantillas, crear_desde_descripcion(), crear_tabla_dinamica()
├── utils/
│   ├── __init__.py
│   ├── history.py          ← historial de conversación en SQLite
│   ├── excel_context.py    ← contexto del archivo subido (en memoria por sesión)
│   ├── chart_context.py    ← datos para regenerar gráficos (en memoria por sesión)
│   ├── sheet_context.py    ← hojas de Excel para selector multi-hoja (en memoria)
│   ├── user_prefs.py       ← preferencias: versión Excel, modo respuesta (texto/voz), modo privado (SQLite)
│   ├── auth.py             ← decorador @solo_autorizados (whitelist por user_id)
│   ├── knowledge.py        ← carga de knowledge/*.md; solo ejemplos_respuestas.md va al system prompt
│   ├── file_meta.py        ← metadata del último archivo subido por usuario (SQLite)
│   ├── df_context.py       ← DataFrame activo + secundario + undo por usuario (en memoria)
│   └── macros.py           ← macros personales por usuario: CRUD en SQLite (user_macros), operaciones como lista DSL
├── prompts/
│   ├── __init__.py
│   └── excel.py            ← todas las plantillas de texto del bot (SYSTEM_BASE, DSL, editor, creación)
├── tests/
│   ├── test_analyzer.py    ← 12 tests para resumir, construir_contexto, analizar_calidad
│   ├── test_exporter.py    ← 11 tests para crear_ejemplo y crear_plantilla
│   ├── test_reader.py      ← 8 tests para leer_excel, leer_excel_hojas, leer_csv
│   ├── test_query_engine.py ← 24 tests para el motor DSL (todas las operaciones y errores)
│   ├── test_editor.py      ← 23 tests para el editor (8 operaciones, EditorError, exportar_xlsx)
│   ├── test_b1_c1.py       ← 16 tests para crear_desde_descripcion, análisis estadístico y correlaciones
│   ├── test_c3_c4_c5.py    ← 26 tests para tendencia, normalización, estandarizar fechas, pivot/unpivot
│   └── test_b3.py          ← 13 tests para combinar_dataframes (inner/left/right/outer, errores, integridad)
├── knowledge/              ← base de conocimiento en Markdown (8 archivos)
└── data/
    ├── historial.db        ← SQLite: historial + preferencias de usuario
    ├── logs/bot.log        ← log rotativo (máx. 5 MB × 3 backups)
    └── temp/               ← archivos temporales durante el procesamiento
```

## Variables de entorno (.env)

```
TELEGRAM_TOKEN=token_obtenido_de_botfather
GROQ_API_KEY=clave_de_groq_console
AUTHORIZED_USERS=id1,id2
```

## Decisiones técnicas importantes

| Decisión | Motivo |
|---|---|
| Solo `ejemplos_respuestas.md` en el system prompt | El resto de knowledge sobrepasaba el límite de 12.000 TPM del tier gratuito de Groq |
| Truncado dinámico del historial en `llm.py` | Evita errores 413 en conversaciones largas sin cortar el contexto del usuario bruscamente |
| `asyncio.to_thread()` para pandas y matplotlib | Evita bloquear el event loop async de python-telegram-bot en operaciones pesadas |
| Magic bytes para validar Excel | La extensión puede ser falsa; se leen los primeros bytes para confirmar el tipo real |
| Límites MAX_FILAS / MAX_COLUMNAS / MAX_HOJAS | Un Excel de 5 MB puede tener 500k filas y colapsar el proceso |
| DSL cerrada para consultas y ediciones | El LLM extrae un JSON estructurado (no código Python libre) → sin riesgo de inyección |
| `RESPUESTA_LIBRE` como centinela del LLM | Si la pregunta no es una operación de datos, el LLM responde exactamente ese literal y el bot cae al flujo normal |
| Excel Table en lugar de TD nativa | openpyxl no puede crear PivotTable interactivas; la Excel Table permite que el usuario cree la TD en 2 clicks desde Excel. Pendiente: evaluar xlwings para TD nativas (solo local, no válido en cloud) |
| TTS con `edge-tts` (sin API key) | Microsoft Neural TTS gratuito, voz `es-ES-ElviraNeural`; Markdown limpiado antes de síntesis; cap de 600 chars con frase de cierre |
| `_enviar_respuesta()` siempre envía texto | El modo voz añade audio, pero el texto se envía siempre para que fórmulas/código sean legibles |
| `context.user_data["op_pendiente"]` para confirmaciones | Persiste el JSON de operación destructiva entre mensaje y callback, evitando el límite de 64 bytes de callback_data |
| Undo con slot `_undo` en `df_context` | `guardar_df()` auto-guarda el df actual antes de sustituirlo; `restaurar_undo()` hace swap activo↔undo — llamarlo dos veces equivale a redo sin estado extra |
| Macros como lista DSL en SQLite | El LLM convierte la descripción en lenguaje natural a lista de operaciones DSL; se almacena como JSON en `user_macros`; sin código arbitrario |
| Modo privado omite `agregar_mensaje()` | Cuando está activo, el historial no se escribe en SQLite; el resto del flujo es idéntico |
| Orden de detección en `procesar_pregunta()` | Fórmulas → Comparar → Combinar → Deshacer → Edición → Crear Excel → Gráfico bajo demanda → Tabla dinámica → Stats/Tendencia → Previsualizar → Valores únicos → Explícame → Exportar CSV → Macros → DSL → LLM. Cada check es un regex O(1) antes de llamar a la API |

## Comandos del bot

| Comando | Descripción |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/ayuda` | Categorías con botones InlineKeyboard |
| `/ejemplo` | Explicación de una función: `/ejemplo BUSCARV` (aleatorio si no se especifica) |
| `/generar` | Genera un .xlsx de ejemplo: `/generar BUSCARV` |
| `/plantilla` | Plantillas .xlsx listas: presupuesto, gastos, KPIs, inventario |
| `/pivote` | Genera archivo con Excel Table + resúmenes estáticos para crear tabla dinámica |
| `/version` | Configura la versión de Excel del usuario (365, 2021, 2019, 2016) |
| `/modo` | Elige si las respuestas son por voz o solo texto |
| `/estado` | Muestra el estado actual de la sesión (archivo activo, historial, modo, versión) |
| `/privado` | Activa/desactiva el modo privado (sin historial en SQLite) |
| `/limpiar` | Borra historial de conversación y contexto Excel (activo + secundario + undo) |

## Funcionalidades sin comando (detección automática por intención)

| Qué escribe/envía el usuario | Qué hace el bot |
|---|---|
| Sube un `.xlsx` / `.csv` | Resumen, calidad de datos, gráfico automático, selector multi-hoja |
| Sube una captura de pantalla | Analiza el Excel con visión IA |
| Envía un mensaje de voz | Transcribe con Whisper → procesa como texto → responde por voz si el modo está activo |
| `=FORMULA(...)` | Explica la fórmula paso a paso |
| "añade una columna Margen que sea Precio×0.3" | Modifica el archivo y envía el .xlsx actualizado |
| "ordena por Fecha descendente" | Ordena el archivo y lo envía |
| "elimina duplicados / rellena vacíos / renombra columna…" | Aplica la operación (con confirmación si es destructiva) y envía el .xlsx |
| "busca 'Enero' y reemplaza por 'January'" | Buscar y reemplazar en una columna o en todo el archivo |
| "divide la columna Nombre por espacio en Nombre y Apellido" | Divide columna de texto en dos |
| "concatena Nombre y Apellido separados por espacio en NombreCompleto" | Concatena columnas de texto |
| "colorea en rojo los valores de Ventas menores de 100" | Aplica formato condicional y envía el .xlsx |
| "cuánto suma Ventas por Región" | Consulta DSL sobre el archivo activo, responde en texto |
| "muéstrame el top 5 por Importe" | Consulta DSL, responde en texto |
| "hazme un Excel con columnas Fecha, Concepto, Importe" | Genera y envía el .xlsx desde descripción |
| "dame estadísticas del archivo" | Análisis estadístico completo (media, std, percentiles, sesgo) |
| "correlaciones" / "mapa de calor" | Análisis de correlaciones + imagen heatmap PNG |
| "¿qué tendencia tienen mis ventas?" | Regresión lineal por columna, R², gráfico de tendencia |
| "normaliza el texto / pon en mayúsculas" | Limpia espacios y unifica capitalización en columnas de texto |
| "corrige las fechas / estandariza el formato" | Parsea formatos mixtos y unifica a DD/MM/YYYY |
| "convierte las columnas de meses en filas" | Despivotea (melt): columnas → filas |
| "pivotea por Producto y Mes" | Pivota (pivot_table): filas → columnas con agregación |
| "une por ID" / "combina los dos archivos" | Combina el archivo activo con el anterior (inner/left/right/outer) |
| "compara los dos archivos" | Informe de diferencias (columnas, filas únicas, filas compartidas) + .xlsx de diferencias |
| "tabla dinámica" | Genera Excel Table + resúmenes estáticos (ver nota en Decisiones técnicas) |
| "hazme un gráfico de barras de Ventas por Mes" | Gráfico personalizado bajo demanda (barras/líneas/sectores/dispersión) |
| "muéstrame las primeras 10 filas" | Previsualización de N filas (primeras o últimas) en bloque de código |
| "qué valores únicos hay en Categoría" | Lista de valores únicos de la columna indicada |
| "explícame el archivo" | Análisis narrativo completo del contenido con el LLM |
| "exporta como CSV" | Exporta el DataFrame activo como .csv UTF-8 con BOM |
| "guarda esta macro como LimpiarFechas" | Guarda la última descripción de operaciones como macro con nombre |
| "ejecuta la macro LimpiarFechas" | Carga la macro y aplica sus operaciones DSL en secuencia |
| "deshacer" / "revertir" | Restaura el DataFrame al estado anterior a la última edición |

## Roadmap

### Fase 1 — MVP ✅
- [x] Conversación básica Telegram ↔ Groq (Llama 3.3)
- [x] Historial por usuario (SQLite)
- [x] Whitelist de acceso por user_id

### Fase 2 — Robustez ✅
- [x] Mensajes de carga, manejo de errores, comandos registrados en Telegram

### Fase 3 — Excel real ✅
- [x] Subida de archivos .xlsx, resumen automático, contexto para preguntas de seguimiento

### Fase 4 — Generación y visualización ✅
- [x] /generar (ejemplos .xlsx), gráficos automáticos, InlineKeyboard en /ayuda, análisis de imágenes

### Fase 5 — Enriquecimiento ✅
- [x] CSV, multi-hoja, tipos de gráfico, /plantilla, explicador de fórmulas, /version, base de conocimiento

### Fase 6 — Calidad y robustez (Sprint 1) ✅
- [x] Logging a fichero rotativo (`data/logs/bot.log`)
- [x] Security hardening: límites de filas/columnas/hojas, validación de tipo real por magic bytes, sanitización de nombres de archivo
- [x] `asyncio.to_thread()` en todas las operaciones bloqueantes (pandas, matplotlib, openpyxl)
- [x] Fix error 413: system prompt reducido, historial limitado a 6, truncado dinámico de tokens

### Fase 6 — Calidad y robustez (Sprint 2) ✅
- [x] Módulo `prompts/excel.py`: todas las plantillas de texto del bot centralizadas
- [x] `utils/file_meta.py`: metadata del último archivo subido por usuario en SQLite (nombre, hoja activa); se limpia con /limpiar
- [x] `excel/analyzer.py`: función `analizar_calidad()` detecta columnas casi vacías, constantes, mezcla texto/número, outliers IQR y fechas inválidas
- [x] 34 tests unitarios en `tests/`: `test_reader`, `test_analyzer`, `test_exporter` — todos en verde

### Fase 6 — Calidad y robustez (Sprint 3) ✅
- [x] `excel/query_engine.py`: DSL cerrada con 9 operaciones (filtrar, contar, suma, promedio, max, min, agrupar, ordenar, top_n) y filtros encadenables (== != > >= < <= contiene no_contiene empieza_por)
- [x] `utils/df_context.py`: DataFrame activo en memoria por usuario
- [x] `services/llm.py`: `extraer_query_dsl()` — interpreta la pregunta y devuelve JSON o RESPUESTA_LIBRE
- [x] `handlers/messages.py`: intenta DSL antes del LLM normal si hay Excel activo; fallback transparente
- [x] 24 tests nuevos en `tests/test_query_engine.py` — 58/58 en verde

### Sprint A — Editor de archivos Excel ✅
- [x] `excel/editor.py`: 8 operaciones (añadir_columna, ordenar, eliminar_duplicados, filtrar_exportar, rellenar_nulos, renombrar_columna, eliminar_columna, formato_condicional)
- [x] `excel/editor.py`: `exportar_xlsx()` con cabeceras azules, filas alternas y hoja de info
- [x] `services/llm.py`: `extraer_operacion_edicion()` — extrae JSON de operación o RESPUESTA_LIBRE
- [x] `handlers/messages.py`: detección por regex `_RE_EDICION` → `_intentar_edicion()` → actualiza df en memoria
- [x] 23 tests en `tests/test_editor.py` — 81/81 en verde

### Sprint B1/B2 — Crear Excel desde descripción ✅
- [x] `prompts/excel.py`: CREAR_EXCEL_SISTEMA + CREAR_EXCEL_USUARIO
- [x] `services/llm.py`: `extraer_estructura_excel()` — extrae JSON de estructura (título, columnas, datos, totales)
- [x] `excel/exporter.py`: `crear_desde_descripcion()` — xlsx con cabeceras, datos y fila SUMA opcional
- [x] `handlers/messages.py`: regex `_RE_CREAR_EXCEL` → `_crear_excel_desde_descripcion()`

### Sprint C1/C2 — Análisis estadístico y correlaciones ✅
- [x] `excel/analyzer.py`: `analisis_estadistico_completo()` — media, mediana, min, max, std, P25/P75, sesgo por columna numérica
- [x] `excel/analyzer.py`: `analisis_correlaciones()` — texto con ranking de pares + heatmap PNG (matplotlib)
- [x] `handlers/messages.py`: rama en `_analizar_estadisticas()` para correlaciones vs estadísticas generales
- [x] 16 tests en `tests/test_b1_c1.py` — 97/97 en verde

### Sprint C3+C4+C5 — Tendencia, normalización y pivot/unpivot ✅
- [x] `excel/analyzer.py`: `analisis_tendencia()` — regresión lineal por columna numérica, R², cambio porcentual, interpretación (📈/📉/sin tendencia), gráfico PNG con línea de tendencia
- [x] `excel/editor.py`: `normalizar_texto()` — strip/upper/lower/title/todas sobre una columna o todas las de texto
- [x] `excel/editor.py`: `estandarizar_fechas()` — parseo de formatos mixtos (DD/MM, ISO, con guiones) y unificación
- [x] `excel/editor.py`: `despivotear()` — columnas → filas (pd.melt), nombres de variable y valor configurables
- [x] `excel/editor.py`: `pivotear()` — filas → columnas (pd.pivot_table), 5 funciones de agregación
- [x] `handlers/messages.py`: `_RE_TENDENCIA` + rama en `_analizar_estadisticas()`; `_RE_EDICION` y `_RE_STATS` ampliados
- [x] `prompts/excel.py`: EDITOR_DSL_SISTEMA actualizado con las 4 nuevas operaciones
- [x] 26 tests en `tests/test_c3_c4_c5.py` — 123/123 en verde

### Sprint B3 — Combinar dos archivos Excel ✅
- [x] `utils/df_context.py`: slot secundario por usuario (guardar/obtener/borrar_df_secundario, borrar_todo)
- [x] `handlers/documents.py`: al procesar un nuevo archivo con df activo, el anterior pasa a slot secundario y se avisa al usuario con la columna sugerida
- [x] `excel/editor.py`: `combinar_dataframes()` — merge por columna clave, 4 tipos de join (inner/left/right/outer), sufijos _A/_B para columnas duplicadas, autodetección de columna común
- [x] `prompts/excel.py`: COMBINAR_DSL_SISTEMA + COMBINAR_DSL_USUARIO
- [x] `services/llm.py`: `extraer_operacion_combinar()` — extrae col y tipo de join del LLM
- [x] `handlers/messages.py`: `_RE_COMBINAR` + `_intentar_combinar()` — resultado pasa a ser el df activo
- [x] `handlers/commands.py`: /limpiar usa `borrar_todo()` → limpia activo + secundario
- [x] 13 tests en `tests/test_b3.py` — 136/136 en verde

### Sprint D1 — Entrada por voz (STT) ✅
- [x] `services/llm.py`: `transcribir_audio()` — Groq Whisper (`whisper-large-v3-turbo`), OGG/Opus, idioma español
- [x] `handlers/audio.py`: `recibir_voz()` y `recibir_audio()` — descarga, transcribe, muestra "🎤 Te escuché: …", llama a `procesar_pregunta()`
- [x] `handlers/messages.py`: `procesar_pregunta()` extraído como función pública; `responder_mensaje()` lo llama directamente
- [x] `bot.py`: handlers para `filters.VOICE` y `filters.AUDIO`

### Sprint D2 — Respuestas por voz (TTS) + preferencia de modo ✅
- [x] `services/tts.py`: `texto_a_audio()` con edge-tts (`es-ES-ElviraNeural`), `_limpiar_markdown()`, cap 600 chars
- [x] `utils/user_prefs.py`: migraciones `modo_respuesta`, `preguntado_modo`; `get/set_modo_respuesta()`, `ya_fue_preguntado_modo()`, `marcar_preguntado_modo()`
- [x] `handlers/messages.py`: `_enviar_respuesta()` — siempre envía texto, añade audio si modo='voz'
- [x] `handlers/audio.py`: tras primera respuesta, pregunta preferencia con InlineKeyboard si aún no se ha preguntado
- [x] `handlers/commands.py`: `/modo` + `callback_modo()` — cambia preferencia en cualquier momento
- [x] `bot.py`: callback `^modo_` registrado

### Sprint E1 — Gráficos bajo demanda ✅
- [x] `excel/charts.py`: `generar_grafico_personalizado()` — barras/líneas/sectores/dispersión, groupby con 5 funciones de agregación, límite 20 categorías
- [x] `prompts/excel.py`: GRAFICO_DSL_SISTEMA + GRAFICO_DSL_USUARIO
- [x] `services/llm.py`: `extraer_peticion_grafico()` — extrae {col_x, col_y, tipo, agregar}
- [x] `handlers/messages.py`: `_RE_GRAFICO` + `_generar_grafico_bajo_demanda()`

### Sprint E2 — Deshacer operaciones ✅
- [x] `utils/df_context.py`: slot `_dataframes_undo`; `guardar_df()` auto-guarda antes de sustituir; `restaurar_undo()` (swap activo↔undo), `hay_undo()`
- [x] `handlers/messages.py`: `_RE_UNDO` + `_deshacer_operacion()` — exporta y envía el df restaurado
- [x] `handlers/commands.py`: `/estado` muestra si hay undo disponible

### Sprint E3 — Explícame archivo + Exportar CSV ✅
- [x] `handlers/messages.py`: `_RE_EXPLICAR_ARCHIVO` + `_explicar_archivo()` — construye prompt con `resumir()` + `analizar_calidad()` y llama al LLM
- [x] `handlers/messages.py`: `_RE_EXPORTAR_CSV` + `_exportar_csv()` — exporta df activo como UTF-8 con BOM

### Sprint F1 — Confirmaciones para operaciones destructivas ✅
- [x] `handlers/messages.py`: `_OPS_DESTRUCTIVAS` set; `_intentar_edicion()` intercepta operaciones destructivas
- [x] `handlers/messages.py`: `_pedir_confirmacion()` — guarda op en `context.user_data["op_pendiente"]`, muestra teclado Sí/No
- [x] `handlers/messages.py`: `callback_confirmacion()` — lee `op_pendiente`, ejecuta via `aplicar_edicion()`
- [x] `bot.py`: callback `^confirmar_op_` registrado

### Sprint F2 — Previsualizar filas + Valores únicos ✅
- [x] `handlers/messages.py`: `_RE_PREVIEW` + `_previsualizar()` — primeras/últimas N filas en bloque de código
- [x] `handlers/messages.py`: `_RE_VALORES_UNICOS` + `_valores_unicos()` — lista de valores únicos por columna o resumen completo

### Sprint F3 — Comparar dos archivos ✅
- [x] `excel/analyzer.py`: `comparar_dataframes()` — outer merge con indicator, informe de columnas/filas únicas/compartidas, devuelve df_diff con columna `_origen_`
- [x] `handlers/messages.py`: `_RE_COMPARAR` + `_comparar_archivos()` — envía informe texto + .xlsx de diferencias si las hay

### Sprint F4 — Macros personales + Modo privado + Buscar/Dividir/Concatenar ✅
- [x] `utils/macros.py`: tabla `user_macros` SQLite; `guardar_macro()`, `obtener_macro()`, `listar_macros()`, `borrar_macro()`
- [x] `prompts/excel.py`: MACRO_DSL_SISTEMA + MACRO_DSL_USUARIO
- [x] `services/llm.py`: `extraer_operaciones_macro()` — convierte descripción en lista de ops DSL
- [x] `handlers/messages.py`: `_RE_GUARDAR_MACRO`, `_RE_EJECUTAR_MACRO`, `_RE_LISTAR_MACROS`, `_RE_BORRAR_MACRO` + handlers correspondientes
- [x] `utils/user_prefs.py`: migración `modo_privado`; `get_modo_privado()`, `toggle_modo_privado()`
- [x] `handlers/commands.py`: `/privado` alterna modo; `/estado` muestra estado privado
- [x] `handlers/messages.py`: `procesar_pregunta()` omite `agregar_mensaje()` si modo privado activo
- [x] `excel/editor.py`: `_buscar_reemplazar()`, `_dividir_columna()`, `_concatenar_columnas()` añadidas al dispatcher
- [x] `prompts/excel.py`: EDITOR_DSL_SISTEMA actualizado con las 3 nuevas operaciones

### Pendiente — decisión futura
- [ ] Tablas dinámicas interactivas nativas: evaluar xlwings (requiere Excel en la máquina, no válido en cloud) vs XML injection con openpyxl (válido en cloud, complejo)

### Fase 7 — Despliegue
- [ ] Despliegue en Railway o Render para disponibilidad 24/7

## Convenciones de código

- Código en español (variables, comentarios, mensajes al usuario)
- Un archivo por responsabilidad
- Variables de entorno siempre desde `.env`
- Manejo de excepciones en todas las llamadas a APIs externas
- Operaciones de I/O y CPU-bound siempre en `asyncio.to_thread()`

## Control de versiones

- **Usuario GitHub**: txape10
- **Repositorio**: github.com/txape10/chatbot-telegram-excel
- **Rama principal**: `main`
- **Archivos que NUNCA deben subirse**: `.env`, `__pycache__/`, `*.pyc`, `data/`

## Notas importantes

- Para obtener el token de Telegram: `@BotFather` → `/newbot`
- Para obtener la clave de Groq: [console.groq.com](https://console.groq.com) (gratuito, sin tarjeta)
- Coste actual: 0 € — Groq free tier ofrece ~14.400 peticiones/día
