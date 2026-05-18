"""Síntesis de voz con edge-tts (Microsoft Neural TTS, sin API key).

Voces recomendadas en español:
  es-ES-ElviraNeural  — femenina, España (por defecto)
  es-ES-AlvaroNeural  — masculina, España
  es-MX-DaliaNeural   — femenina, México
  es-MX-JorgeNeural   — masculina, México
"""
import re
import logging
import edge_tts

logger = logging.getLogger(__name__)

VOZ_ES    = "es-ES-ElviraNeural"
MAX_CHARS = 600   # ~45 segundos de audio; respuestas más largas se recortan


# ── Limpieza de Markdown ─────────────────────────────────────────────────────

def _limpiar_markdown(texto: str) -> str:
    """Elimina marcado Markdown para que el TTS lea de forma natural."""
    # Negrita e itálica
    texto = re.sub(r'\*{1,2}(.+?)\*{1,2}', r'\1', texto)
    texto = re.sub(r'_{1,2}(.+?)_{1,2}',   r'\1', texto)
    # Bloques y spans de código
    texto = re.sub(r'```[\s\S]+?```', '', texto)
    texto = re.sub(r'`(.+?)`',        r'\1', texto)
    # Cabeceras Markdown
    texto = re.sub(r'^#{1,4}\s+', '', texto, flags=re.MULTILINE)
    # Viñetas
    texto = re.sub(r'^\s*[•·\-\*]\s+', '', texto, flags=re.MULTILINE)
    # Filas de tabla
    texto = re.sub(r'\|[^\n]+\|', '', texto)
    # Saltos múltiples → uno solo
    texto = re.sub(r'\n{3,}', '\n\n', texto)
    # Emojis de puntos de lista al inicio de línea
    texto = re.sub(r'^[📊📈📉🔢✅⚠️💡•·]\s*', '', texto, flags=re.MULTILINE)
    return texto.strip()


# ── API pública ───────────────────────────────────────────────────────────────

async def texto_a_audio(texto: str, voz: str = VOZ_ES) -> bytes:
    """Convierte texto a audio MP3 (bytes).

    Si el texto supera MAX_CHARS se recorta al último espacio antes del límite
    y se añade una frase de cierre. Devuelve bytes vacíos si falla.
    """
    texto_limpio = _limpiar_markdown(texto)

    if not texto_limpio:
        return b""

    if len(texto_limpio) > MAX_CHARS:
        recortado = texto_limpio[:MAX_CHARS].rsplit(" ", 1)[0]
        texto_limpio = recortado + ". Puedes leer la respuesta completa en el chat."

    try:
        communicate = edge_tts.Communicate(texto_limpio, voz)
        audio = b""
        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                audio += chunk["data"]
        return audio
    except Exception as error:
        logger.warning("Error generando audio TTS: %s", error)
        return b""
