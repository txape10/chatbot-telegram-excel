# 🤖 Chatbot de Telegram — Asistente Personal de Excel

## WHY — Por qué existe este proyecto

Crear un asistente personal en Telegram que resuelva dudas sobre Microsoft Excel de forma inmediata, con ejemplos prácticos y explicaciones claras, disponible desde cualquier dispositivo (móvil, tablet, otro PC) siempre que el PC principal esté encendido.

## WHAT — Qué hace

Un bot de Telegram que:
- Responde preguntas sobre Excel en español
- Proporciona ejemplos prácticos con datos reales
- Cubre fórmulas, tablas dinámicas, formato condicional, gráficos, macros/VBA, Power Query, etc.
- Mantiene contexto de la conversación (historial)
- Analiza archivos Excel y CSV subidos por el usuario
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
│   ├── analyzer.py         ← resumen, detección de errores de fórmula, multi-hoja
│   ├── charts.py           ← gráficos PNG (barras / líneas / sectores) con matplotlib
│   └── exporter.py         ← ejemplos de funciones + 4 plantillas listas (presupuesto, gastos, KPIs, inventario)
├── utils/
│   ├── __init__.py
│   ├── history.py          ← historial de conversación en SQLite
│   ├── excel_context.py    ← contexto del archivo subido (en memoria por sesión)
│   ├── chart_context.py    ← datos para regenerar gráficos (en memoria por sesión)
│   ├── sheet_context.py    ← hojas de Excel para selector multi-hoja (en memoria)
│   ├── user_prefs.py       ← preferencias de usuario: versión Excel (SQLite)
│   ├── auth.py             ← decorador @solo_autorizados (whitelist por user_id)
│   └── knowledge.py        ← carga de knowledge/*.md; solo ejemplos_respuestas.md va al system prompt
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

## Comandos del bot

| Comando | Descripción |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/ayuda` | Categorías con botones InlineKeyboard |
| `/ejemplo` | Explicación de una función: `/ejemplo BUSCARV` (aleatorio si no se especifica) |
| `/generar` | Genera un .xlsx de ejemplo: `/generar BUSCARV` |
| `/plantilla` | Plantillas .xlsx listas: presupuesto, gastos, KPIs, inventario |
| `/version` | Configura la versión de Excel del usuario (365, 2021, 2019, 2016) |
| `/limpiar` | Borra historial de conversación y contexto Excel |

## Funcionalidades especiales (sin comando)

- **Subir Excel/CSV**: resumen automático, detección de errores, gráfico, selector multi-hoja
- **Escribir `=FORMULA(...)`**: el bot la desglosa paso a paso automáticamente
- **Subir una captura de pantalla**: el bot la analiza con visión IA

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

### Fase 6 — Calidad y robustez (pendiente)
- [ ] Sprint 2: tests, analista automático ampliado, metadata en SQLite, prompts a módulo
- [ ] Sprint 3: engine de queries pandas con DSL cerrada (filtrar, agrupar, ordenar, top N...)

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
