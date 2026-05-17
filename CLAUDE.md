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
- **Google Gemini API** (1.500 peticiones/día gratis) — opción principal
- Alternativa: **Groq** con Llama 3.3 (14.400 peticiones/día gratis)

### Librería de Telegram
- `python-telegram-bot`

### Ejecución
- **Local en PC del usuario** (sin servidor externo de momento)
- El bot es accesible desde cualquier dispositivo con Telegram mientras el PC esté encendido
- Migración futura posible a Railway o Render (gratis, sin cambiar código)

### Dependencias principales
```
python-telegram-bot
google-generativeai
python-dotenv
```

## Estructura del proyecto

```
3 - Chatbot de Telegram para Excel/
├── CLAUDE.md               ← este archivo
├── .env                    ← tokens y claves (NO subir a git)
├── .env.example            ← plantilla sin valores reales
├── .gitignore
├── requirements.txt
├── bot.py                  ← punto de entrada principal
├── config.py               ← configuración y variables de entorno
├── handlers/
│   ├── __init__.py
│   ├── messages.py         ← lógica de respuesta a mensajes
│   └── commands.py         ← comandos /start, /ayuda, /ejemplo
├── services/
│   ├── __init__.py
│   └── gemini.py           ← integración con la API de Gemini
└── utils/
    ├── __init__.py
    └── history.py          ← gestión del historial de conversación
```

## Variables de entorno (.env)

```
TELEGRAM_TOKEN=token_obtenido_de_botfather
GEMINI_API_KEY=clave_de_google_ai_studio
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
| `/ejemplo` | Ejemplo aleatorio de función útil de Excel |
| `/limpiar` | Borrar el historial de conversación |

## Funcionalidades — Roadmap

### Fase 1 — MVP ✅ (implementar primero)
- [ ] Conexión básica bot Telegram ↔ Gemini API
- [ ] Respuestas a preguntas de texto sobre Excel
- [ ] Historial de conversación por usuario (últimos N mensajes)
- [ ] Comandos básicos: /start, /ayuda, /limpiar

### Fase 2 — Mejoras medias 🚀
- [ ] Menús con botones (InlineKeyboard) por categorías: Fórmulas / Tablas dinámicas / VBA / Gráficos / Power Query
- [ ] Comando /ejemplo con función aleatoria útil
- [ ] Persistencia del historial en SQLite (sobrevive reinicios)
- [ ] Manejo de errores y mensajes de carga ("Pensando...")

### Fase 3 — Funcionalidades avanzadas 💡
- [ ] Análisis de imágenes: el usuario envía captura de su Excel y el bot la interpreta (Gemini tiene visión)
- [ ] Generación de archivos .xlsx con ejemplos usando `openpyxl`
- [ ] Base de conocimiento propia (PDF de fórmulas favoritas)
- [ ] Despliegue en Railway o Render para disponibilidad 24/7

## Documentación de referencia técnica

- **python-telegram-bot**: https://docs.python-telegram-bot.org (usar v21+)
- **Google Gemini API**: https://ai.google.dev/docs
- **google-generativeai SDK**: https://github.com/google/generative-ai-python
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

### Crear el repositorio (pendiente)
El repositorio aún no está creado. Claude Code debe crearlo con:
```bash
git init
git remote add origin https://github.com/txape10/chatbot-telegram-excel.git
```
Y asegurarse de que el `.gitignore` esté configurado antes del primer commit.

## Notas importantes

- El archivo `.env` nunca debe subirse a git (añadirlo al `.gitignore`)
- Para obtener el token de Telegram: buscar `@BotFather` en Telegram → `/newbot`
- Para obtener la clave de Gemini: registrarse en [Google AI Studio](https://aistudio.google.com)
- Coste estimado: prácticamente 0€ para uso personal con los tiers gratuitos
