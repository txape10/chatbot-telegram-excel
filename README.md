# Asistente Excel con IA

Asistente de inteligencia artificial para Microsoft Excel accesible desde dos canales:

- **Bot de Telegram** — escribe en lenguaje natural desde cualquier dispositivo
- **Add-in de Excel** — panel lateral dentro de Excel, sin salir de la aplicación

Ambos canales comparten el mismo motor de IA y pueden instalarse juntos o por separado.

---

## Instalación rápida

```bash
git clone https://github.com/txape10/chatbot-telegram-excel.git
cd chatbot-telegram-excel
pip install -r requirements.txt
python instalar.py
```

El instalador te guía paso a paso: elige los módulos, el proveedor de IA, introduce tus claves y genera el `.env` automáticamente.

---

## Modos de despliegue

### 🖥️ Personal (Windows, uso local)
El bot corre en tu PC mientras la ventana esté abierta.
```bash
scripts\instalar_personal.bat   # primera vez
scripts\arrancar_personal.bat   # arranque diario
```

### ☁️ Render (cloud, 24/7 gratuito)
El bot corre en la nube y responde aunque el PC esté apagado. Se redesploya automáticamente con cada `git push`.
```bash
# Conecta el repo en render.com y configura las variables de entorno
# WEBHOOK_URL=https://tu-servicio.onrender.com
# ADDIN_URL=https://tu-servicio.onrender.com/
```
> El tier gratuito puede tardar ~50 segundos tras un período de inactividad.

### 🏢 Servidor empresa (Linux interno)
Bot + API REST en un servidor interno. El Add-in se sirve a través de Cloudflare Tunnel sin abrir puertos en el firewall.
```bash
scripts/instalar_empresa.sh    # primera vez
scripts/arrancar_empresa.sh    # arranque manual (en producción usa systemd)
# sudo cp scripts/asistente-excel.service /etc/systemd/system/
```

---

## Proveedores de IA

Configura el proveedor con `LLM_PROVIDER` en el `.env`. Puedes cambiarlo en cualquier momento sin tocar el código.

| Proveedor | Coste | Datos | Caso de uso |
|---|---|---|---|
| `groq` | **0$** | EE.UU. | Personal / beta — rápido, generoso |
| `ollama` | **0$** | Local (ninguno sale) | Máxima privacidad, servidor propio |
| `gemini` | **0$** | EE.UU. | Alternativa a Groq, Gemini 1.5 Flash |
| `mistral` | **0$** | UE 🇪🇺 | Empresa sin coste, empresa europea |
| `openai` | 💲 | EE.UU. | GPT-4o-mini, máxima capacidad |
| `azure` | 💲 | UE 🇪🇺 | Empresa, cumple RGPD, garantía contractual |

Cualquier proveedor con API compatible con OpenAI también funciona añadiendo `OPENAI_BASE_URL`.

---

## Módulos activables

```env
ENABLE_TELEGRAM=true   # Bot de Telegram
ENABLE_ADDIN=true      # Add-in de Excel (API REST + ficheros estáticos)
```

La empresa puede optar por desplegar solo el Add-in sin el bot de Telegram.

---

## Funcionalidades del bot

### Conversación y consultas
- Responde preguntas sobre Excel en el idioma del usuario (multiidioma automático)
- Mantiene historial de conversación por usuario (SQLite)
- Se adapta a la versión de Excel del usuario (`/version`: 365, 2021, 2019, 2016)
- Escribe `=FORMULA(...)` → explicación paso a paso
- Mensajes de voz → transcripción Whisper → respuesta en texto o voz (`/modo`)
- Cuando una petición es ambigua, pregunta con botones en lugar de fallar

### Análisis de archivos
- Sube `.xlsx`, `.xls` o `.csv` → resumen, calidad de datos, gráfico automático
- Soporte multi-hoja con selector inline
- Sube una captura de pantalla de Excel → análisis con visión IA

### Consultas en lenguaje natural
Con un archivo activo puedes preguntar directamente:
- "¿Cuánto suma Ventas por Región?"
- "Muéstrame el top 5 por Importe"
- "¿Cuántos pedidos hay con estado Pendiente?"

Motor DSL interno (sin ejecución de código arbitrario): filtrar, contar, sumar, promediar, agrupar, ordenar, top N.

### Modificación de archivos

| Petición | Operación |
|---|---|
| "Añade una columna Margen que sea Precio × 0,3" | Nueva columna calculada |
| "Ordena por Fecha descendente" | Ordenar |
| "Elimina los duplicados" | Eliminar duplicados |
| "Rellena los vacíos de Categoría con 'Sin categoría'" | Rellenar nulos |
| "Colorea en rojo las ventas menores de 100" | Formato condicional |
| "Busca 'Enero' y reemplaza por 'January'" | Buscar y reemplazar |
| "Divide Nombre por espacio en Nombre y Apellido" | Dividir columna |
| "Concatena Nombre y Apellido en NombreCompleto" | Concatenar columnas |
| "Normaliza el texto de Categoría" | Mayúsculas/minúsculas/título |
| "Estandariza las fechas" | Detecta y unifica formatos mixtos |
| "Convierte los meses en filas" | Unpivot (melt) |
| "Agrupa por Vendedor con suma de Ventas" | Pivot table |

Las operaciones destructivas piden confirmación. "Deshacer" restaura el estado anterior.

### Análisis avanzado
- Estadísticas completas: media, mediana, std, percentiles, sesgo
- Correlaciones + heatmap PNG
- Tendencia: regresión lineal, R², gráfico con línea de tendencia
- Combinar dos archivos: inner/left/right/outer join con autodetección de columna común
- Comparar dos archivos: informe de diferencias + `.xlsx` de diff

### Creación y exportación
- "Hazme un Excel con columnas Fecha, Concepto, Importe" → genera y envía el `.xlsx`
- Plantillas listas: presupuesto, gastos, KPIs, inventario
- Tabla dinámica: Excel Table + resúmenes estáticos
- Gráficos personalizados: barras, líneas, sectores, dispersión
- Exportar como CSV

### Macros personales
- "Guarda esta macro como LimpiarFechas" → almacena la secuencia de operaciones
- "Ejecuta la macro LimpiarFechas" → aplica las operaciones en orden

---

## Comandos

| Comando | Descripción |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/ayuda` | Menú de categorías con botones |
| `/ejemplo BUSCARV` | Explica una función de Excel |
| `/generar BUSCARV` | Genera un `.xlsx` de ejemplo |
| `/plantilla` | Plantillas listas para usar |
| `/pivote` | Genera archivo para tabla dinámica |
| `/version` | Configura tu versión de Excel |
| `/modo` | Respuestas por texto o voz |
| `/estado` | Estado de la sesión actual |
| `/privado` | Activa/desactiva historial en SQLite |
| `/limpiar` | Borra historial y contexto de archivo |

---

## Seguridad

- **Bot Telegram**: whitelist de IDs numéricos (`AUTHORIZED_USERS`) — no falsificables
- **Add-in**: whitelist de dominio corporativo + correos individuales, verificada por servidor
- **API REST**: clave de API en header `X-API-Key` obligatoria en todas las rutas
- **Archivos**: validación por magic bytes, límites de tamaño/filas/columnas/hojas, sin ejecución de macros
- **DSL cerrada**: el LLM extrae JSON estructurado, nunca ejecuta código arbitrario
- **Comunicaciones**: HTTPS (TLS 1.3) en todas las conexiones externas
- **Servidor**: solo conexiones salientes, sin puertos entrantes abiertos
- **`.env`**: nunca se sube al repositorio (en `.gitignore`)

---

## Stack técnico

| Componente | Tecnología |
|---|---|
| Lenguaje | Python 3.11+ |
| Bot framework | python-telegram-bot v21 |
| API REST | FastAPI + uvicorn |
| Add-in frontend | Office.js + webpack |
| Motor IA | Configurable: Groq / Ollama / Gemini / Mistral / OpenAI / Azure |
| STT (voz a texto) | Groq Whisper (o proveedor activo) |
| TTS (texto a voz) | edge-tts — es-ES-ElviraNeural (sin API key) |
| Excel / CSV | pandas + openpyxl |
| Gráficos | matplotlib |
| Persistencia | SQLite (historial, preferencias, macros, metadatos) |
| Coste base | 0 € |

---

## Estructura del proyecto

```
├── instalar.py             ← Instalador interactivo (genera .env)
├── bot.py                  ← Arranque modo personal (polling)
├── api.py                  ← API REST + bot integrado (empresa/cloud)
├── telegram_app.py         ← Configuración PTB compartida
├── config.py               ← Variables de entorno, flags de módulos
├── handlers/               ← messages, commands, documents, images, audio
├── excel/                  ← reader, analyzer, query_engine, editor, charts, exporter
├── services/
│   ├── llm_provider.py     ← Abstracción de proveedores de IA
│   ├── llm.py              ← Funciones de IA (usa llm_provider)
│   └── tts.py              ← Síntesis de voz
├── utils/                  ← history, df_context, auth, macros, user_prefs…
├── prompts/excel.py        ← Plantillas de texto para el LLM
├── excel-addin/            ← Add-in de Excel (Office.js + webpack)
├── scripts/                ← Scripts de instalación y arranque por modo
├── tests/                  ← 352 tests unitarios (pytest)
├── knowledge/              ← Base de conocimiento en Markdown
├── docs/                   ← Documentación interna (no sube a GitHub)
└── data/                   ← SQLite + logs (no sube a GitHub)
```

---

## Tests

```bash
pytest
```

352 tests que cubren: lectura de archivos, análisis de calidad, motor DSL, editor (15 operaciones), combinación y comparación de archivos, tendencias, creación de Excel, gráficos, undo, macros, preferencias, detección de intenciones por regex y lógica pura de handlers (preview, valores únicos, exportar CSV).

---

## Roadmap

- [x] Conversación con historial por usuario
- [x] Análisis de archivos Excel/CSV (calidad, estadísticas, correlaciones, tendencias)
- [x] Motor de consultas DSL en lenguaje natural
- [x] Editor de archivos (15 operaciones)
- [x] Creación de Excel desde descripción
- [x] Tabla dinámica, gráficos personalizados
- [x] Combinación y comparación de dos archivos
- [x] Entrada por voz (Whisper STT) y respuesta por voz (edge-tts TTS)
- [x] Macros personales, modo privado, deshacer
- [x] Add-in de Excel (panel lateral, temas visuales, Easter egg)
- [x] Autenticación por dominio corporativo (sin Azure AD)
- [x] Aclaración inteligente con InlineKeyboard cuando la petición es ambigua
- [x] Detección automática del idioma del usuario (multiidioma sin coste adicional)
- [x] API_KEY del Add-in inyectada en build time (sin valores hardcodeados)
- [x] Despliegue en Render (cloud gratuito, 24/7)
- [x] Abstracción de proveedor de IA (Groq/Ollama/Gemini/Mistral/OpenAI/Azure)
- [x] Módulos activables independientemente (Telegram / Add-in)
- [x] Instalador interactivo con checklist
- [x] Instalador one-click del Add-in (`.bat` auto-elevado, SMB share, registro Centro de confianza)
- [x] Mensajes de error descriptivos cuando el proveedor de IA falla (saturación, timeout, auth…)
- [x] Add-in UI: barra de archivo activo (libro/hoja/rango en tiempo real) e historial colapsable
- [ ] Add-in: subir archivo Excel directamente al bot desde el panel lateral
- [x] Tests completos para todos los sprints (352/352 ✅)
- [ ] Autenticación SSO con Azure Active Directory
- [ ] Despliegue en servidor de empresa con Cloudflare Tunnel
- [ ] Tablas dinámicas interactivas nativas (Windows + Excel en servidor)
- [ ] Panel de administración (estadísticas de uso, gestión de usuarios)
