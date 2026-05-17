# 🤖 Chatbot de Telegram — Asistente Personal de Excel

## WHY — Por qué existe este proyecto

Crear un asistente personal en Telegram que resuelva dudas sobre Microsoft Excel de forma inmediata, con ejemplos prácticos y explicaciones claras, disponible desde cualquier dispositivo (móvil, tablet, otro PC) siempre que el PC principal esté encendido.

## WHAT — Qué hace

Un bot de Telegram que:
- Responde preguntas sobre Excel en español
- Proporciona ejemplos prácticos con datos reales
- Cubre fórmulas, tablas dinámicas, formato condicional, gráficos, macros/VBA, Power Query, etc.
- Mantiene contexto de la conversación (historial)
- Es accesible desde cualquier dispositivo con Telegram instalado

## HOW — Stack técnico y decisiones tomadas

### Lenguaje
- **Python**

### LLM (gratuito)
- **Groq** con modelo `llama-3.3-70b-versatile` — gratuito, sin tarjeta de crédito

### Librería de Telegram
- `python-telegram-bot`

### Ejecución
- **Local en PC del usuario** (sin servidor externo de momento)
- El bot es accesible desde cualquier dispositivo con Telegram mientras el PC esté encendido
- Migración futura posible a Railway o Render (gratis, sin cambiar código)

### Dependencias principales
```
python-telegram-bot
groq
python-dotenv
pandas
openpyxl
numpy
matplotlib
pillow
```

## Estructura del proyecto (objetivo final)

```
3 - Chatbot de Telegram para Excel/
├── CLAUDE.md               ← este archivo
├── .env                    ← tokens y claves (NO subir a git)
├── .env.example            ← plantilla sin valores reales
├── .gitignore
├── requirements.txt
├── bot.py                  ← punto de entrada principal
├── config.py               ← configuración, variables de entorno y whitelist
├── handlers/
│   ├── __init__.py
│   ├── messages.py         ← lógica de respuesta a mensajes de texto
│   ├── commands.py         ← comandos /start, /ayuda, /ejemplo, /limpiar
│   └── documents.py        ← recepción y procesamiento de archivos .xlsx
├── services/
│   ├── __init__.py
│   └── llm.py              ← integración con la API de Groq (Llama 3.3)
├── excel/
│   ├── __init__.py
│   ├── reader.py           ← leer archivos .xlsx con pandas
│   ├── analyzer.py         ← análisis, resumen y detección de errores
│   └── exporter.py         ← generación de archivos .xlsx con ejemplos
├── utils/
│   ├── __init__.py
│   └── history.py          ← gestión del historial de conversación
└── knowledge/              ← base de conocimiento en Markdown
```

## Variables de entorno (.env)

```
TELEGRAM_TOKEN=token_obtenido_de_botfather
GROQ_API_KEY=clave_de_groq_console
AUTHORIZED_USERS=id1,id2
```

## System prompt del asistente

El bot debe comportarse como un experto en Microsoft Excel con más de 20 años de experiencia. Ante cada pregunta debe:
1. Dar una explicación clara y concisa
2. Incluir un ejemplo práctico con datos reales
3. Mostrar la fórmula o los pasos exactos
4. Añadir consejos o variantes útiles si los hay

Responde siempre en español.

## Comandos del bot

| Comando | Descripción |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/ayuda` | Categorías de temas disponibles |
| `/ejemplo` | Ejemplo aleatorio o de una función concreta: `/ejemplo BUSCARV` |
| `/limpiar` | Borrar el historial de conversación |

## Funcionalidades — Roadmap

### Fase 1 — MVP (conversación básica) ✅
- [x] Conexión básica bot Telegram ↔ Groq (Llama 3.3)
- [x] Respuestas a preguntas de texto sobre Excel
- [x] Historial de conversación en memoria por usuario (últimos 10 mensajes)
- [x] Comandos básicos: /start, /ayuda, /limpiar
- [x] Control de acceso: whitelist de `user_id` autorizados en `.env`

### Fase 2 — Robustez ✅
- [x] Mensajes de carga mientras el LLM procesa ("⏳ Pensando...")
- [x] Manejo de errores con mensajes claros al usuario
- [x] Persistencia del historial en SQLite (sobrevive reinicios)
- [x] Comando /ejemplo con función aleatoria o concreta (/ejemplo BUSCARV)
- [x] Comandos registrados en Telegram (menú al escribir "/")

### Fase 3 — Excel real con pandas ✅
- [x] Carpeta `excel/` con `reader.py` y `analyzer.py`
- [x] `handlers/documents.py`: recibir archivos `.xlsx` por Telegram (máx. 5 MB)
- [x] Resumen automático del Excel subido (filas, columnas, nulos, duplicados)
- [x] Responder preguntas sobre el archivo subido usando el LLM

### Fase 4 — Generación y visualización
- [ ] `excel/exporter.py`: generar archivos `.xlsx` con ejemplos prácticos
- [ ] Gráficos automáticos con `matplotlib` enviados como imagen
- [ ] Menús con botones (InlineKeyboard) por categorías de temas
- [ ] Análisis de imágenes: captura de pantalla de Excel interpretada por el LLM

### Fase 5 — Despliegue y producción
- [ ] Base de conocimiento desde PDFs propios
- [ ] Despliegue en Railway o Render para disponibilidad 24/7

## Documentación de referencia técnica

- **python-telegram-bot**: https://docs.python-telegram-bot.org (usar v21+)
- **Groq API**: https://console.groq.com/docs
- **openpyxl** (fase 3): https://openpyxl.readthedocs.io
- **Python objetivo**: 3.11+

## Base de conocimiento de Excel

La carpeta `knowledge/` contiene archivos Markdown con contenido estructurado sobre Excel. Claude Code debe usarlos para:
1. Construir un system prompt rico y detallado para el bot
2. Cargarlos como contexto adicional en las respuestas cuando sea relevante

```
knowledge/
├── formulas_basicas.md        ← SUMA, SI, BUSCARV, CONTAR.SI...
├── formulas_avanzadas.md      ← LAMBDA, LET, ÍNDICE/COINCIDIR...
├── tablas_dinamicas.md        ← pasos, opciones, trucos
├── formato_condicional.md     ← reglas, fórmulas personalizadas
├── power_query.md             ← transformaciones más usadas
├── vba_basico.md              ← macros más útiles y frecuentes
├── errores_comunes.md         ← #¡VALOR!, #N/A, #¡REF! y cómo resolverlos
└── ejemplos_respuestas.md     ← tono y formato exacto que debe usar el bot
```

## Convenciones de código

- Código en español (variables, comentarios, mensajes al usuario)
- Un archivo por responsabilidad (no mezclar lógica de Telegram con llamadas a la API)
- Variables de entorno siempre desde `.env`, nunca hardcodeadas
- Manejo de excepciones en todas las llamadas a APIs externas
- El historial de conversación se limita a los últimos 10 mensajes para no superar el contexto

## Control de versiones

- **Usuario GitHub**: txape10
- **Repositorio**: github.com/txape10/chatbot-telegram-excel
- **Rama principal**: `main`
- **Flujo de trabajo**: commits frecuentes con mensajes descriptivos en español
- **Archivos que NUNCA deben subirse a GitHub**: `.env`, `__pycache__/`, `*.pyc`

## Notas importantes

- El archivo `.env` nunca debe subirse a git (añadirlo al `.gitignore`)
- Para obtener el token de Telegram: buscar `@BotFather` en Telegram → `/newbot`
- Para obtener la clave de Groq: registrarse en [console.groq.com](https://console.groq.com) (gratuito, sin tarjeta)
- Coste: 0€ — Groq ofrece ~14.400 peticiones/día gratis en el tier gratuito
