# 🤖 Asistente Excel con IA — Bot Telegram + Add-in Excel

## WHY — Por qué existe este proyecto

Asistente de inteligencia artificial para Microsoft Excel accesible desde dos canales:
- **Bot de Telegram**: disponible desde cualquier dispositivo, incluso por voz
- **Add-in de Excel**: panel lateral dentro de Excel sin salir de la aplicación

Ambos canales comparten el mismo motor de IA y pueden instalarse juntos o por separado.

---

## HOW — Stack técnico

### Lenguaje
- **Python 3.11+** (backend) + **JavaScript/webpack** (Add-in frontend)

### Proveedor de IA — configurable
- `LLM_PROVIDER` en `.env`: `groq` | `ollama` | `gemini` | `mistral` | `openai` | `azure`
- Por defecto: **Groq** (gratuito, llama-3.3-70b-versatile)
- Abstracción en `services/llm_provider.py` — cambiar de proveedor no requiere cambios de código
- Límite TPM gestionado dinámicamente por `max_tokens_peticion` de cada proveedor

### Modos de despliegue
| Modo | Cómo arranca | Bot | API Add-in |
|---|---|---|---|
| Personal Windows | `scripts\arrancar_personal.bat` → `bot.py` | Polling | No |
| Render (cloud) | `python api.py` + WEBHOOK_URL | Webhook | Sí |
| Empresa (servidor Linux) | `python api.py` sin WEBHOOK_URL | Polling | Sí |

`api.py` detecta el modo automáticamente según `WEBHOOK_URL` y `TELEGRAM_TOKEN`.

### Módulos activables
```env
ENABLE_TELEGRAM=true   # Bot de Telegram
ENABLE_ADDIN=true      # API REST + ficheros estáticos del Add-in
```

### Dependencias principales
```
python-telegram-bot==21.10
groq==0.13.1
openai>=1.30.0          ← también para Ollama, Gemini, Mistral (API compatible)
python-dotenv==1.1.0
pandas==2.2.3
openpyxl==3.1.5
matplotlib==3.9.4
pillow==11.2.1
edge-tts==6.1.12
fastapi==0.115.12
uvicorn==0.34.3
aiofiles==24.1.0
```

---

## Estructura del proyecto

```
├── instalar.py             ← Instalador interactivo (genera .env con checklist)
├── bot.py                  ← Arranque personal (polling, sin API)
├── api.py                  ← API REST + bot integrado (empresa/cloud)
├── telegram_app.py         ← Configuración PTB compartida entre bot.py y api.py
├── config.py               ← Variables de entorno, flags ENABLE_*, límites seguridad
├── logging_config.py       ← Logging consola + fichero rotativo (data/logs/bot.log)
├── render.yaml             ← Configuración Render (cloud gratuito)
├── .env.example            ← Plantilla de variables (sin valores reales)
├── requirements.txt
│
├── handlers/
│   ├── messages.py         ← Núcleo: procesar_pregunta() + detección de intención
│   ├── commands.py         ← /start /ayuda /ejemplo /generar /plantilla /version /modo /estado /privado /limpiar
│   ├── documents.py        ← Archivos .xlsx/.xls/.csv: validación, multi-hoja, slot secundario
│   ├── images.py           ← Análisis de capturas con visión IA
│   └── audio.py            ← Voz → Whisper → procesar_pregunta; pregunta preferencia de modo
│
├── services/
│   ├── llm_provider.py     ← Abstracción IA: GroqProvider, OllamaProvider, GeminiProvider,
│   │                          MistralProvider, OpenAIProvider, AzureOpenAIProvider
│   │                          Factory singleton: obtener_proveedor()
│   ├── llm.py              ← Funciones de IA (usa llm_provider): chat, DSLs, Whisper, visión
│   └── tts.py              ← edge-tts (es-ES-ElviraNeural); limpieza Markdown; cap 600 chars
│
├── excel/
│   ├── reader.py           ← Leer .xlsx (multi-hoja) y .csv con detección de separador
│   ├── analyzer.py         ← Resumen, calidad, estadísticas, correlaciones, tendencias, comparar
│   ├── query_engine.py     ← DSL consultas: filtrar/contar/suma/promedio/max/min/agrupar/ordenar/top_n
│   ├── editor.py           ← 15 operaciones de edición + exportar_xlsx
│   ├── charts.py           ← Gráficos PNG automáticos y personalizados
│   └── exporter.py         ← Ejemplos, plantillas, crear_desde_descripcion(), tabla_dinámica
│
├── utils/
│   ├── history.py          ← Historial SQLite
│   ├── df_context.py       ← DataFrame activo + secundario + undo por usuario (en memoria)
│   ├── excel_context.py    ← Contexto textual del archivo para el LLM
│   ├── chart_context.py    ← Datos para regenerar gráficos
│   ├── sheet_context.py    ← Hojas para selector multi-hoja
│   ├── user_prefs.py       ← Preferencias: versión Excel, modo voz, modo privado (SQLite)
│   ├── auth.py             ← Decorador @solo_autorizados (whitelist por user_id)
│   ├── knowledge.py        ← Carga knowledge/*.md → system prompt
│   ├── file_meta.py        ← Metadata último archivo subido (SQLite)
│   ├── macros.py           ← Macros personales CRUD en SQLite
│   └── sheet_context.py    ← Hojas de Excel para selector multi-hoja
│
├── prompts/excel.py        ← Todas las plantillas de texto enviadas al LLM
│
├── excel-addin/            ← Add-in de Excel (Office.js + webpack)
│   ├── src/taskpane/
│   │   ├── taskpane.html/js/css ← UI principal
│   │   ├── themes.js       ← Sistema de temas (default, empresa, zelda Easter egg)
│   │   ├── auth.whitelist.js ← Autenticación por dominio/correo (sin Azure AD)
│   │   └── auth.sso.js     ← Stub SSO para futura integración Azure AD
│   ├── webpack.config.js   ← DefinePlugin inyecta ALLOWED_DOMAINS/EMAILS en build
│   ├── manifest.xml        ← Manifiesto del Add-in para Office 365
│   └── .env                ← Variables del Add-in (no sube a git)
│
├── scripts/
│   ├── instalar_personal.bat   ← Primera instalación Windows
│   ├── arrancar_personal.bat   ← Arranque diario Windows (polling)
│   ├── instalar_empresa.sh     ← Primera instalación Linux servidor
│   ├── arrancar_empresa.sh     ← Arranque manual Linux (API + polling)
│   └── asistente-excel.service ← Fichero systemd para servidor empresa
│
├── tests/                  ← 352 tests unitarios (pytest)
├── knowledge/              ← Base de conocimiento Markdown (8 archivos)
├── docs/                   ← Documentación interna (NO sube a GitHub)
└── data/                   ← SQLite + logs (NO sube a GitHub)
```

---

## Variables de entorno (.env)

```env
# Módulos
ENABLE_TELEGRAM=true
ENABLE_ADDIN=true

# Proveedor de IA (groq|ollama|gemini|mistral|openai|azure)
LLM_PROVIDER=groq
LLM_MODEL=llama-3.3-70b-versatile
GROQ_API_KEY=...

# Bot de Telegram
TELEGRAM_TOKEN=...
AUTHORIZED_USERS=id1,id2

# Add-in
API_KEY=...
ALLOWED_DOMAINS=empresa.eu,empresa.com
ALLOWED_EMAILS=usuario@gmail.com

# Despliegue cloud (vacío en local)
WEBHOOK_URL=
ADDIN_URL=
```

---

## Decisiones técnicas importantes

| Decisión | Motivo |
|---|---|
| `services/llm_provider.py` como abstracción | Cambiar de Groq a Ollama/Azure/etc. sin tocar el resto del código |
| `obtener_proveedor()` como singleton | Una sola instancia por proceso; el proveedor se inicializa al primer uso |
| `max_tokens_peticion` por proveedor | Groq free tiene 9.000 tokens de margen; otros proveedores permiten más |
| `api.py` detecta polling vs webhook automáticamente | Si WEBHOOK_URL está vacía → polling; si está definida → webhook |
| `ENABLE_TELEGRAM` / `ENABLE_ADDIN` | La empresa puede desplegar solo el Add-in sin el bot |
| `instalar.py` genera el `.env` | Evita errores en la configuración manual; muestra URLs de obtención de claves |
| `auth.whitelist.js` + webpack DefinePlugin | Los dominios/correos autorizados se inyectan en el bundle en build time; nunca en el repo |
| `auth.sso.js` como stub | Misma interfaz que whitelist; listo para implementar Azure AD sin cambiar el resto |
| Solo `ejemplos_respuestas.md` en el system prompt | El resto de knowledge sobrepasaba el límite TPM del tier gratuito de Groq |
| DSL cerrada para consultas y ediciones | El LLM extrae JSON estructurado; nunca ejecuta código Python arbitrario |
| `RESPUESTA_LIBRE` como centinela del LLM | Si la pregunta no es operación de datos, el LLM responde ese literal y el bot cae al flujo normal |
| Magic bytes para validar Excel | La extensión puede ser falsa; se leen los primeros bytes para confirmar el tipo real |
| `asyncio.to_thread()` para pandas y matplotlib | Evita bloquear el event loop async de python-telegram-bot |
| Undo con slot `_undo` en `df_context` | `guardar_df()` auto-guarda antes de sustituir; `restaurar_undo()` hace swap activo↔undo |
| `context.user_data["op_pendiente"]` para confirmaciones | Persiste el JSON de operación destructiva entre mensaje y callback (límite 64 bytes de callback_data) |
| Macros como lista DSL en SQLite | Sin código arbitrario; el LLM convierte descripción → lista de operaciones JSON |
| Excel Table en lugar de TD nativa | openpyxl no puede crear PivotTable interactivas; la Excel Table permite crearla en 2 clicks desde Excel |
| TTS con `edge-tts` (sin API key) | Microsoft Neural TTS gratuito; voz `es-ES-ElviraNeural`; cap 600 chars |
| Temas visuales con CSS Custom Properties | Cambio de tema sin recargar la página; localStorage persiste la preferencia |
| Easter egg Zelda con Web Audio API | Jingle sintetizado con OscillatorNode (onda cuadrada NES); no requiere archivos de audio |

---

## Estado del proyecto

### ✅ Completado

- Bot Telegram con todas las funcionalidades (Fases 1-6 + Sprints A-G)
- Add-in Excel (temas, autenticación por dominio, Easter egg Zelda)
- Despliegue en Render (cloud gratuito, 24/7) — betatester activo
- Abstracción de proveedor IA (6 proveedores)
- Módulos activables independientemente
- Instalador interactivo (`instalar.py`)
- Scripts por modo de despliegue (`scripts/`)
- README completo
- Documentación para reunión con administrador (`docs/`)
- **Sprint G**: aclaración inteligente con InlineKeyboard cuando la petición es ambigua
- **Sprint G**: detección automática de idioma del usuario (multiidioma sin coste)
- **Fix Add-in**: API_KEY inyectada por webpack DefinePlugin (ya no hardcodeada)
- **Seguridad**: eliminadas todas las referencias a empresa/dominios corporativos de archivos públicos
- **Instalador Add-in**: `scripts/instalar_addin.bat` + `instalar_addin.ps1` — un doble clic instala el complemento, comparte carpeta SMB y registra el catálogo en el Centro de confianza de Excel automáticamente (con opción manual si el antivirus lo bloquea)
- **Errores LLM**: `LLMError` con mensajes descriptivos por tipo (saturación, timeout, conexión, autenticación, límite de tokens) en todos los handlers
- **Add-in UI**: barra de archivo activo (libro, hoja, rango en tiempo real) + historial de conversación colapsable con persistencia en localStorage
- **Tests sprints D2/E1/E2/F3/F4**: 105 tests — TTS markdown, gráficos, undo/secundario, comparar archivos, macros CRUD, buscar/dividir/concatenar, preferencias
- **Tests sprints G/E3/F1/F2**: 111 tests — detección de intenciones por regex (73 patrones), lógica pura de preview/valores únicos/CSV/LLM utils. Suite total: 352/352 ✅
- **Fix editor**: `pd.to_numeric(errors='ignore')` deprecado sustituido por try/except explícito
- **Sprint H1 — Enviar al bot desde el Add-in**: botón "📤 Enviar al bot" en el panel del Add-in; `/vincular email` y `/desvincular` en Telegram; endpoint `POST /enviar-al-bot`; tabla `user_links` en SQLite
- **Sprint H2 — Panel de administración**: `GET /admin` (HTML con gráfico + tabla de usuarios) y `GET /admin/stats` (JSON); protegido por `ADMIN_KEY`; estadísticas: mensajes totales/hoy/por usuario, actividad 7 días, Add-ins vinculados
- **Sprint H2+**: análisis de sesiones con funciones de ventana SQLite (LAG/SUM OVER); `utils/device_emails.py` mapea device_id → email + versión Excel; Add-in envía `user_email` y `excel_version` en cada petición; modo respuesta con iconos explícitos ("💬 Texto" / "🔊 Voz" + "🔒 Priv")
- **Panel admin — rediseño con tabs**: 4 pestañas (Resumen / Usuarios / IA / Sistema); gráfico de actividad corregido con alturas en píxeles; sección Usuarios con tops (más activos, más sesiones, inactivos >30 días) + análisis de sesiones
- **Notificaciones — spam corregido**: eliminado el fallback automático a `AUTHORIZED_USERS[0]`; ahora `ALERT_TELEGRAM_ID` debe estar explícitamente en `.env` para recibir alertas; cooldown de 300 s por tipo de alerta
- **Sprint C — Formato condicional**: 7 tipos (valor, top/bottom, escala de color, barra de datos, iconos, texto, fórmula); DSL en `prompts/excel.py`, `extraer_regla_formato()` en `services/llm.py`, `POST /format` en `api.py`; routing automático en Add-in por keywords; `_aplicarFormatoCondicional()` con Office.js
- **Add-in — preservación de formato** (Opción A + B): `copyFrom(formats)` antes de escribir valores cuando hay rango fuente; `_inferirFormatos()` aplica `#,##0.00` / `#,##0` en tablas creadas desde cero

### 🐛 Bugs conocidos (en investigación)

- **Notificaciones del sistema — sección pendiente de rehacer**: la sección de suscriptores se eliminó del panel admin temporalmente; rediseñar con configuración por tipo de alerta y gestión desde la UI antes de reactivar.

### ⏳ Pendiente (bloqueado por reunión con admin)

- Despliegue en servidor empresa (Linux + cloudflared)
- Distribución del `manifest.xml` al catálogo Office 365

### 🔮 Futuro

- **UptimeRobot** (o similar): monitorizar que la URL de Render responde cada 5 min y avisar por Telegram/email si cae. Configuración de 2 min en uptimerobot.com — free tier ilimitado.
- **Notificaciones del sistema — rediseño**: sección de suscriptores eliminada temporalmente del panel; rediseñar con configuración por tipo de alerta (RAM, errores LLM, etc.), activar/desactivar por tipo, gestión desde UI antes de reactivar. `ALERT_TELEGRAM_ID` en `.env` como única forma de recibir alertas por ahora.
- **Aprendizaje de usuario (RAG sobre historial)**: el Add-in podría aprender de peticiones aceptadas; RAG sobre el historial SQLite para incluir pares pregunta→respuesta exitosos como few-shot en el system prompt. Requiere: mecanismo de "pulgar arriba" en la UI, endpoint `/context` que recupere ejemplos por usuario, integración en `_construir_mensajes()`.
- Autenticación SSO con Azure Active Directory (`auth.sso.js` ya preparado)
- Tablas dinámicas interactivas nativas con **xlwings** (gratuito, requiere Windows + Excel en el servidor). Activación automática: al arrancar se detecta `platform.system() == "Windows"` + `xlwings` disponible → flag `PIVOT_NATIVO_DISPONIBLE`. En Linux (Render) sigue usando la alternativa actual de openpyxl. Implementar cuando esté disponible el servidor Windows de empresa.

---

## Instancias en producción

| Instancia | URL/Acceso | Usuarios | Proveedor IA |
|---|---|---|---|
| Personal/Beta | Render — asistente-excel.onrender.com | Roberto + betatester | Groq |
| Empresa | Pendiente servidor admin | Empleados | Por decidir (Mistral/Azure) |

---

## Convenciones de código

- Código en español (variables, comentarios, mensajes al usuario)
- Un archivo por responsabilidad
- Variables de entorno siempre desde `.env`
- Manejo de excepciones en todas las llamadas a APIs externas
- Operaciones I/O y CPU-bound siempre en `asyncio.to_thread()`
- Documentación interna en `docs/` (gitignored)

---

## Add-in — flujo de cambios (IMPORTANTE)

El Add-in se **compila en local** y el `dist/` compilado va al repositorio.
**Render NO ejecuta npm** — solo `pip install -r requirements.txt`.

### Por qué
`npm install` consume ~400 MB de RAM en el contenedor de Render (free tier: 512 MB).
Pre-compilar en local elimina ese problema y hace los deploys más rápidos.

### Flujo obligatorio al tocar el Add-in
```
1. Editar archivos en excel-addin/src/
2. cd excel-addin && npm run build
3. git add excel-addin/dist/
4. git commit + git push
```

> **NUNCA** añadir `npm install` ni `npm run build` al `buildCommand` de `render.yaml`.
> El `dist/` está excluido de `.gitignore` a propósito — debe subir al repo.

---

## Control de versiones

- **Usuario GitHub**: txape10
- **Repositorio**: github.com/txape10/chatbot-telegram-excel
- **Rama principal**: `main`
- **Nunca subir**: `.env`, `data/`, `docs/`, `__pycache__/`, `*.pyc`
- **Render** redespliega automáticamente en cada `git push` a `main`
