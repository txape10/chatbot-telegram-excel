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
- Modifica archivos Excel a petición (añadir columnas, ordenar, eliminar duplicados, colorear…)
- Crea archivos Excel desde descripción en lenguaje natural
- Genera resúmenes estadísticos y mapas de correlaciones
- Genera archivos de tabla dinámica (Excel Table + resúmenes estáticos)
- Explica fórmulas paso a paso si el mensaje empieza por `=`
- Analiza capturas de pantalla de Excel con visión IA
- Es accesible desde cualquier dispositivo con Telegram instalado

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
│   ├── messages.py         ← respuesta a texto; detección automática de fórmulas (=)
│   ├── commands.py         ← /start, /ayuda, /ejemplo, /generar, /plantilla, /version, /limpiar + callbacks InlineKeyboard
│   ├── documents.py        ← archivos .xlsx/.xls/.csv: validación, límites, multi-hoja, gráficos, asyncio.to_thread
│   └── images.py           ← análisis de capturas de pantalla con LLM de visión
├── services/
│   ├── __init__.py
│   └── llm.py              ← Groq: obtener_respuesta (Llama 3.3) + analizar_imagen (Llama 4 Scout)
│                              Gestión dinámica de tokens: truncado automático del historial
├── excel/
│   ├── __init__.py
│   ├── reader.py           ← leer .xlsx (multi-hoja) y .csv con detección de separador
│   ├── analyzer.py         ← resumen, errores de fórmula, analizar_calidad(), análisis estadístico, correlaciones
│   ├── query_engine.py     ← motor DSL: filtrar/contar/suma/promedio/max/min/agrupar/ordenar/top_n
│   ├── editor.py           ← modificación de archivos: 8 operaciones + exportar_xlsx con estilos
│   ├── charts.py           ← gráficos PNG (barras / líneas / sectores) con matplotlib
│   └── exporter.py         ← ejemplos, plantillas, crear_desde_descripcion(), crear_tabla_dinamica()
├── utils/
│   ├── __init__.py
│   ├── history.py          ← historial de conversación en SQLite
│   ├── excel_context.py    ← contexto del archivo subido (en memoria por sesión)
│   ├── chart_context.py    ← datos para regenerar gráficos (en memoria por sesión)
│   ├── sheet_context.py    ← hojas de Excel para selector multi-hoja (en memoria)
│   ├── user_prefs.py       ← preferencias de usuario: versión Excel (SQLite)
│   ├── auth.py             ← decorador @solo_autorizados (whitelist por user_id)
│   ├── knowledge.py        ← carga de knowledge/*.md; solo ejemplos_respuestas.md va al system prompt
│   ├── file_meta.py        ← metadata del último archivo subido por usuario (SQLite)
│   └── df_context.py       ← DataFrame activo por usuario (en memoria, para motor DSL y editor)
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
│   └── test_c3_c4_c5.py    ← 26 tests para tendencia, normalización, estandarizar fechas, pivot/unpivot
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
| Orden de detección en messages.py | Fórmulas → Edición → Crear Excel → Tabla dinámica → Estadísticas → DSL → LLM. Cada check es un regex O(1) antes de llamar a la API |

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
| `/limpiar` | Borra historial de conversación y contexto Excel |

## Funcionalidades sin comando (detección automática por intención)

| Qué escribe el usuario | Qué hace el bot |
|---|---|
| Sube un `.xlsx` / `.csv` | Resumen, calidad de datos, gráfico automático, selector multi-hoja |
| `=FORMULA(...)` | Explica la fórmula paso a paso |
| Sube una captura de pantalla | Analiza el Excel con visión IA |
| "añade una columna Margen que sea Precio×0.3" | Modifica el archivo y envía el .xlsx actualizado |
| "ordena por Fecha descendente" | Ordena el archivo y lo envía |
| "elimina duplicados / rellena vacíos / renombra columna…" | Aplica la operación y envía el .xlsx |
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
| "tabla dinámica" | Genera Excel Table + resúmenes estáticos (ver nota en Decisiones técnicas) |

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

### Sprint B3 — Combinar dos archivos Excel ⏳ pendiente
- [ ] Gestión de segundo df en memoria (`df_context.py`: añadir slot secundario por usuario)
- [ ] Detección por regex de intención de combinación ("une por ID", "combina los dos archivos")
- [ ] Operación `combinar` en editor: merge por columna clave (inner/left/outer), descripción del resultado
- [ ] Flujo de subida: si el usuario ya tiene un df activo y sube otro archivo, ofrecer opción de combinar

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
