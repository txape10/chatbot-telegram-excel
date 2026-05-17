"""Prompts del asistente de Excel. Centraliza todos los textos enviados al LLM."""

# ── Prompt de sistema base ───────────────────────────────────────────────────

SYSTEM_BASE = (
    "Eres un experto en Microsoft Excel con más de 20 años de experiencia. "
    "Ante cada pregunta debes:\n"
    "1. Dar una explicación clara y concisa\n"
    "2. Incluir un ejemplo práctico con datos reales\n"
    "3. Mostrar la fórmula o los pasos exactos\n"
    "4. Añadir consejos o variantes útiles si los hay\n\n"
    "Responde siempre en español."
)

# ── Prompts de comandos ──────────────────────────────────────────────────────

EJEMPLO_FUNCION = (
    "Explícame la función {funcion} de Microsoft Excel con un ejemplo práctico. "
    "Usa datos reales y concretos, muestra la fórmula exacta y añade un consejo útil."
)

# ── Prompts de mensajes ──────────────────────────────────────────────────────

EXPLICAR_FORMULA = (
    "Explica paso a paso esta fórmula de Excel:\n\n{formula}\n\n"
    "Desglosa cada argumento o parte, explica qué hace cada uno y muestra "
    "un ejemplo práctico con datos reales de cuándo y cómo usarla."
)

PREGUNTA_CON_VERSION = (
    "[El usuario usa {version}. Adapta la respuesta a las funciones "
    "disponibles en esa versión.]\n\n"
)

PREGUNTA_CON_CONTEXTO = "{contexto}\n\nPregunta del usuario: {pregunta}"

# ── Prompts de análisis de imágenes ─────────────────────────────────────────

IMAGEN_SIN_CAPTION = (
    "Analiza esta captura de Excel y explica qué hace, "
    "qué fórmulas usa y cómo podría mejorarla."
)
