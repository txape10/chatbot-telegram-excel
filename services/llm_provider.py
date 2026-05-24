"""Abstracción del proveedor de IA.

Configura el proveedor activo con LLM_PROVIDER en .env:

  groq    → Groq (llama-3.3-70b-versatile) — gratuito, datos en EE.UU.
  ollama  → Ollama local — gratuito, sin datos fuera de la red
  openai  → OpenAI (gpt-4o-mini) — de pago, datos en EE.UU.

Variables de entorno relevantes:
  LLM_PROVIDER        groq | ollama | openai   (por defecto: groq)
  LLM_MODEL           nombre del modelo de chat (por defecto: según proveedor)
  LLM_MODEL_VISION    modelo con visión         (solo groq/openai)
  LLM_MODEL_AUDIO     modelo de transcripción   (solo groq/openai)
  GROQ_API_KEY        clave de Groq
  OPENAI_API_KEY      clave de OpenAI
  OLLAMA_URL          URL de Ollama             (por defecto: http://localhost:11434)
"""

import logging
import os
from abc import ABC, abstractmethod

logger = logging.getLogger(__name__)

# Presupuesto de tokens por petición.
# Groq free tier: ~12 000 TPM → margen de 9 000.
# Otros proveedores: límite mucho más alto.
_MAX_TOKENS_GROQ   = 9_000
_MAX_TOKENS_OTROS  = 32_000


class LLMProvider(ABC):
    """Interfaz común para todos los proveedores de IA."""

    # Subclases deben definir estos atributos
    modelo: str = ""
    max_tokens_peticion: int = _MAX_TOKENS_OTROS

    @abstractmethod
    def chat(self, messages: list[dict], temperature: float = 0.7,
             max_tokens: int | None = None) -> str:
        """Llamada de chat estándar. Devuelve el texto de la respuesta."""

    def transcribir(self, audio_bytes: bytes, filename: str = "audio.ogg") -> str:
        """Transcripción de audio a texto. No todos los proveedores lo soportan."""
        raise NotImplementedError(
            f"{self.__class__.__name__} no soporta transcripción de audio. "
            "Usa LLM_PROVIDER=groq o LLM_PROVIDER=openai para esta función."
        )

    def vision(self, messages: list[dict]) -> str:
        """Análisis de imágenes. No todos los proveedores lo soportan."""
        raise NotImplementedError(
            f"{self.__class__.__name__} no soporta análisis de imágenes. "
            "Usa LLM_PROVIDER=groq o LLM_PROVIDER=openai para esta función."
        )


# ---------------------------------------------------------------------------
# Groq
# ---------------------------------------------------------------------------

class GroqProvider(LLMProvider):
    """Proveedor Groq — gratuito, rápido, datos en EE.UU."""

    max_tokens_peticion = _MAX_TOKENS_GROQ

    def __init__(self):
        from groq import Groq
        self._cliente = Groq(api_key=os.getenv("GROQ_API_KEY", ""), timeout=60.0)
        self.modelo         = os.getenv("LLM_MODEL",        "llama-3.3-70b-versatile")
        self.modelo_vision  = os.getenv("LLM_MODEL_VISION", "meta-llama/llama-4-scout-17b-16e-instruct")
        self.modelo_audio   = os.getenv("LLM_MODEL_AUDIO",  "whisper-large-v3-turbo")

    def chat(self, messages, temperature=0.7, max_tokens=None):
        kwargs = {"model": self.modelo, "messages": messages, "temperature": temperature}
        if max_tokens:
            kwargs["max_tokens"] = max_tokens
        respuesta = self._cliente.chat.completions.create(**kwargs)
        return respuesta.choices[0].message.content

    def transcribir(self, audio_bytes, filename="audio.ogg"):
        transcripcion = self._cliente.audio.transcriptions.create(
            file=(filename, audio_bytes),
            model=self.modelo_audio,
            language="es",
        )
        return transcripcion.text.strip()

    def vision(self, messages):
        respuesta = self._cliente.chat.completions.create(
            model=self.modelo_vision,
            messages=messages,
        )
        return respuesta.choices[0].message.content


# ---------------------------------------------------------------------------
# Ollama (local, sin datos fuera de la red)
# ---------------------------------------------------------------------------

class OllamaProvider(LLMProvider):
    """Proveedor Ollama — modelo local, gratuito, máxima privacidad.

    Usa la API compatible con OpenAI que expone Ollama en /v1.
    Requisito: Ollama instalado y corriendo en OLLAMA_URL (por defecto localhost:11434).
    Modelos recomendados: llama3.2, mistral, qwen2.5
    """

    max_tokens_peticion = _MAX_TOKENS_OTROS

    def __init__(self):
        from openai import OpenAI
        base_url = os.getenv("OLLAMA_URL", "http://localhost:11434").rstrip("/") + "/v1"
        self._cliente = OpenAI(api_key="ollama", base_url=base_url, timeout=120.0)
        self.modelo = os.getenv("LLM_MODEL", "llama3.2")

    def chat(self, messages, temperature=0.7, max_tokens=None):
        kwargs = {"model": self.modelo, "messages": messages, "temperature": temperature}
        if max_tokens:
            kwargs["max_tokens"] = max_tokens
        respuesta = self._cliente.chat.completions.create(**kwargs)
        return respuesta.choices[0].message.content


# ---------------------------------------------------------------------------
# OpenAI
# ---------------------------------------------------------------------------

class OpenAIProvider(LLMProvider):
    """Proveedor OpenAI — de pago, datos en EE.UU.

    Compatible con Azure OpenAI cambiando OPENAI_API_KEY y OPENAI_BASE_URL.
    """

    max_tokens_peticion = _MAX_TOKENS_OTROS

    def __init__(self):
        from openai import OpenAI
        kwargs = {"api_key": os.getenv("OPENAI_API_KEY", ""), "timeout": 60.0}
        base_url = os.getenv("OPENAI_BASE_URL", "")
        if base_url:
            kwargs["base_url"] = base_url
        self._cliente = OpenAI(**kwargs)
        self.modelo        = os.getenv("LLM_MODEL",        "gpt-4o-mini")
        self.modelo_vision = os.getenv("LLM_MODEL_VISION", "gpt-4o-mini")
        self.modelo_audio  = os.getenv("LLM_MODEL_AUDIO",  "whisper-1")

    def chat(self, messages, temperature=0.7, max_tokens=None):
        kwargs = {"model": self.modelo, "messages": messages, "temperature": temperature}
        if max_tokens:
            kwargs["max_tokens"] = max_tokens
        respuesta = self._cliente.chat.completions.create(**kwargs)
        return respuesta.choices[0].message.content

    def transcribir(self, audio_bytes, filename="audio.ogg"):
        transcripcion = self._cliente.audio.transcriptions.create(
            file=(filename, audio_bytes),
            model=self.modelo_audio,
            language="es",
        )
        return transcripcion.text.strip()

    def vision(self, messages):
        respuesta = self._cliente.chat.completions.create(
            model=self.modelo_vision,
            messages=messages,
        )
        return respuesta.choices[0].message.content


# ---------------------------------------------------------------------------
# Factory / singleton
# ---------------------------------------------------------------------------

_proveedor: LLMProvider | None = None


def obtener_proveedor() -> LLMProvider:
    """Devuelve la instancia singleton del proveedor activo."""
    global _proveedor
    if _proveedor is None:
        nombre = os.getenv("LLM_PROVIDER", "groq").lower()
        proveedores = {
            "groq":   GroqProvider,
            "ollama": OllamaProvider,
            "openai": OpenAIProvider,
        }
        clase = proveedores.get(nombre)
        if clase is None:
            logger.warning("LLM_PROVIDER '%s' desconocido — usando groq", nombre)
            clase = GroqProvider
        _proveedor = clase()
        logger.info("Proveedor IA: %s | Modelo: %s", nombre, _proveedor.modelo)
    return _proveedor
