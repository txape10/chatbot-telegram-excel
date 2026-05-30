"""Abstracción del proveedor de IA.

Configura el proveedor activo con LLM_PROVIDER en .env:

  groq    → Groq (llama-3.3-70b-versatile)   0$  gratuito, datos en EE.UU.
  ollama  → Ollama local                      0$  sin datos fuera de la red
  gemini  → Google Gemini 1.5 Flash           0$  free tier generoso, datos en EE.UU.
  mistral → Mistral (Mistral Small)           0$  free tier, empresa europea
  openai  → OpenAI (gpt-4o-mini)             💲  de pago, datos en EE.UU.
  azure   → Azure OpenAI                     💲  de pago, datos en UE, cumple RGPD

Proveedor de respaldo (opcional):
  Si LLM_FALLBACK_PROVIDER está definido, los errores transitorios del proveedor
  primario (rate limit, timeout, conexión) se reintentan automáticamente con el
  proveedor de respaldo, sin que el usuario note nada.

  Configuración recomendada:
    LLM_PROVIDER=groq
    LLM_FALLBACK_PROVIDER=mistral    ← datos en la UE, gratis sin tarjeta
    MISTRAL_API_KEY=...

Variables de entorno relevantes:
  LLM_PROVIDER          groq|ollama|gemini|mistral|openai|azure  (por defecto: groq)
  LLM_MODEL             nombre del modelo de chat  (por defecto: según proveedor)
  LLM_FALLBACK_PROVIDER groq|ollama|gemini|mistral|openai|azure  (vacío = sin respaldo)
  LLM_FALLBACK_MODEL    modelo del proveedor de respaldo  (vacío = default del proveedor)
  LLM_MODEL_VISION      modelo con visión          (groq/openai/azure/gemini)
  LLM_MODEL_AUDIO       modelo de transcripción    (groq/openai/azure)
  GROQ_API_KEY          clave de Groq
  GEMINI_API_KEY        clave de Google AI Studio  (aistudio.google.com — gratuito)
  MISTRAL_API_KEY       clave de Mistral           (console.mistral.ai — gratuito)
  OPENAI_API_KEY        clave de OpenAI
  AZURE_OPENAI_KEY      clave de Azure OpenAI
  AZURE_OPENAI_URL      endpoint Azure             (https://<recurso>.openai.azure.com)
  AZURE_API_VERSION     versión API Azure          (por defecto: 2024-02-01)
  OLLAMA_URL            URL de Ollama              (por defecto: http://localhost:11434)
"""

import logging
import os
from abc import ABC, abstractmethod

logger = logging.getLogger(__name__)


def _stat(tipo: str, proveedor: str, ok: bool,
          fallback: bool = False, error_tipo: str | None = None) -> None:
    """Registra una llamada al LLM en la BD de estadísticas. Falla silenciosamente."""
    try:
        from utils.llm_stats import registrar
        registrar(proveedor, tipo, ok, fallback, error_tipo)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Excepción propia — errores controlados del proveedor de IA
# ---------------------------------------------------------------------------

class LLMError(Exception):
    """Error del proveedor de IA con mensaje listo para mostrar al usuario."""

    # Mensajes por tipo de error
    _MENSAJES = {
        "rate_limit":  "⏳ El servicio de IA está saturado ahora mismo. Espera unos segundos e inténtalo de nuevo.",
        "timeout":     "⏰ La IA tardó demasiado en responder. Inténtalo de nuevo.",
        "conexion":    "🌐 No se pudo conectar con el servicio de IA. Comprueba tu conexión a Internet.",
        "auth":        "🔑 Error de autenticación con el servicio de IA. Contacta con el administrador.",
        "limite":      "📏 El texto es demasiado largo para procesarlo de una vez. Prueba con una pregunta más corta o sube un archivo más pequeño.",
        "generico":    "⚠️ El servicio de IA no está disponible en este momento. Inténtalo de nuevo en unos segundos.",
    }

    def __init__(self, tipo: str = "generico", detalle: str = ""):
        self.tipo = tipo
        self.mensaje_usuario = self._MENSAJES.get(tipo, self._MENSAJES["generico"])
        super().__init__(detalle or self.mensaje_usuario)


def _manejar_error_groq(error: Exception) -> None:
    """Convierte errores del SDK de Groq en LLMError con mensaje descriptivo."""
    nombre = type(error).__name__
    texto  = str(error).lower()

    if nombre == "RateLimitError":
        raise LLMError("rate_limit", str(error)) from error
    if nombre == "APITimeoutError":
        raise LLMError("timeout", str(error)) from error
    if nombre == "APIConnectionError":
        raise LLMError("conexion", str(error)) from error
    if nombre == "AuthenticationError":
        raise LLMError("auth", str(error)) from error
    if nombre == "BadRequestError" and ("413" in texto or "too large" in texto or "tokens" in texto):
        raise LLMError("limite", str(error)) from error
    raise LLMError("generico", str(error)) from error

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
        # max_retries=0: ante un 429 falla inmediatamente → el FallbackProvider
        # lo captura y reintenta con el proveedor de respaldo sin esperar 17 s.
        self._cliente = Groq(api_key=os.getenv("GROQ_API_KEY", ""), timeout=60.0, max_retries=0)
        self.modelo         = os.getenv("LLM_MODEL",        "llama-3.3-70b-versatile")
        self.modelo_vision  = os.getenv("LLM_MODEL_VISION", "meta-llama/llama-4-scout-17b-16e-instruct")
        self.modelo_audio   = os.getenv("LLM_MODEL_AUDIO",  "whisper-large-v3-turbo")

    def chat(self, messages, temperature=0.7, max_tokens=None):
        try:
            kwargs = {"model": self.modelo, "messages": messages, "temperature": temperature}
            if max_tokens:
                kwargs["max_tokens"] = max_tokens
            respuesta = self._cliente.chat.completions.create(**kwargs)
            return respuesta.choices[0].message.content
        except LLMError:
            raise
        except Exception as error:
            logger.warning("Groq chat error: %s", error)
            _manejar_error_groq(error)

    def transcribir(self, audio_bytes, filename="audio.ogg"):
        try:
            transcripcion = self._cliente.audio.transcriptions.create(
                file=(filename, audio_bytes),
                model=self.modelo_audio,
                language="es",
            )
            return transcripcion.text.strip()
        except LLMError:
            raise
        except Exception as error:
            logger.warning("Groq transcripción error: %s", error)
            _manejar_error_groq(error)

    def vision(self, messages):
        try:
            respuesta = self._cliente.chat.completions.create(
                model=self.modelo_vision,
                messages=messages,
            )
            return respuesta.choices[0].message.content
        except LLMError:
            raise
        except Exception as error:
            logger.warning("Groq vision error: %s", error)
            _manejar_error_groq(error)


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
# Google Gemini (API compatible con OpenAI)
# ---------------------------------------------------------------------------

class GeminiProvider(LLMProvider):
    """Google Gemini — free tier generoso, datos en EE.UU.

    Obtén tu clave gratuita en: https://aistudio.google.com/app/apikey
    """

    max_tokens_peticion = _MAX_TOKENS_OTROS

    def __init__(self):
        from openai import OpenAI
        self._cliente = OpenAI(
            api_key=os.getenv("GEMINI_API_KEY", ""),
            base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
            timeout=60.0,
        )
        self.modelo        = os.getenv("LLM_MODEL",        "gemini-1.5-flash")
        self.modelo_vision = os.getenv("LLM_MODEL_VISION", "gemini-1.5-flash")

    def chat(self, messages, temperature=0.7, max_tokens=None):
        kwargs = {"model": self.modelo, "messages": messages, "temperature": temperature}
        if max_tokens:
            kwargs["max_tokens"] = max_tokens
        respuesta = self._cliente.chat.completions.create(**kwargs)
        return respuesta.choices[0].message.content

    def vision(self, messages):
        respuesta = self._cliente.chat.completions.create(
            model=self.modelo_vision, messages=messages,
        )
        return respuesta.choices[0].message.content


# ---------------------------------------------------------------------------
# Mistral (API compatible con OpenAI)
# ---------------------------------------------------------------------------

class MistralProvider(LLMProvider):
    """Mistral — free tier disponible, empresa europea.

    Obtén tu clave en: https://console.mistral.ai  (plan free disponible)
    """

    max_tokens_peticion = _MAX_TOKENS_OTROS

    def __init__(self):
        from openai import OpenAI
        self._cliente = OpenAI(
            api_key=os.getenv("MISTRAL_API_KEY", ""),
            base_url="https://api.mistral.ai/v1",
            timeout=60.0,
        )
        self.modelo = os.getenv("LLM_MODEL", "mistral-small-latest")

    def chat(self, messages, temperature=0.7, max_tokens=None):
        kwargs = {"model": self.modelo, "messages": messages, "temperature": temperature}
        if max_tokens:
            kwargs["max_tokens"] = max_tokens
        respuesta = self._cliente.chat.completions.create(**kwargs)
        return respuesta.choices[0].message.content


# ---------------------------------------------------------------------------
# Azure OpenAI (datos en UE, cumple RGPD)
# ---------------------------------------------------------------------------

class AzureOpenAIProvider(LLMProvider):
    """Azure OpenAI — de pago, datos en la UE, cumple RGPD.

    Requiere: cuenta Azure + recurso Azure OpenAI desplegado.
    Variables: AZURE_OPENAI_KEY, AZURE_OPENAI_URL, AZURE_API_VERSION
    """

    max_tokens_peticion = _MAX_TOKENS_OTROS

    def __init__(self):
        from openai import AzureOpenAI
        self._cliente = AzureOpenAI(
            api_key=os.getenv("AZURE_OPENAI_KEY", ""),
            azure_endpoint=os.getenv("AZURE_OPENAI_URL", ""),
            api_version=os.getenv("AZURE_API_VERSION", "2024-02-01"),
            timeout=60.0,
        )
        self.modelo        = os.getenv("LLM_MODEL",        "gpt-4o-mini")
        self.modelo_vision = os.getenv("LLM_MODEL_VISION", "gpt-4o-mini")
        self.modelo_audio  = os.getenv("LLM_MODEL_AUDIO",  "whisper")

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
            model=self.modelo_vision, messages=messages,
        )
        return respuesta.choices[0].message.content


# ---------------------------------------------------------------------------
# Fallback automático entre proveedores
# ---------------------------------------------------------------------------

# Errores transitorios que justifican reintentar con el proveedor de respaldo.
# auth / limite / generico indican un problema real que el respaldo no resolvería.
_ERRORES_RECUPERABLES = {"rate_limit", "timeout", "conexion"}


class FallbackProvider(LLMProvider):
    """Wrapper transparente: usa el primario y cae al secundario en errores transitorios.

    Solo aplica fallback a chat(). Audio y visión no tienen respaldo porque
    los proveedores de respaldo habituales (Mistral) no los soportan.
    """

    def __init__(self, primario: LLMProvider, secundario: LLMProvider,
                 nombre_primario: str = "primary", nombre_secundario: str = "secondary"):
        self._primario            = primario
        self._secundario          = secundario
        self._nombre_primario     = nombre_primario
        self._nombre_secundario   = nombre_secundario
        self.modelo               = primario.modelo
        self.max_tokens_peticion  = primario.max_tokens_peticion

    def chat(self, messages, temperature=0.7, max_tokens=None):
        try:
            resultado = self._primario.chat(messages, temperature, max_tokens)
            _stat("chat", self._nombre_primario, ok=True)
            return resultado
        except LLMError as error:
            _stat("chat", self._nombre_primario, ok=False, error_tipo=error.tipo)
            if error.tipo not in _ERRORES_RECUPERABLES:
                raise
            logger.warning(
                "Proveedor primario falló (%s) — reintentando con respaldo", error.tipo
            )
            try:
                resultado = self._secundario.chat(messages, temperature, max_tokens)
                _stat("chat", self._nombre_secundario, ok=True, fallback=True)
                return resultado
            except LLMError as error2:
                _stat("chat", self._nombre_secundario, ok=False, fallback=True, error_tipo=error2.tipo)
                raise

    def transcribir(self, audio_bytes, filename="audio.ogg"):
        try:
            r = self._primario.transcribir(audio_bytes, filename)
            _stat("audio", self._nombre_primario, ok=True)
            return r
        except Exception as exc:
            _stat("audio", self._nombre_primario, ok=False)
            raise exc

    def vision(self, messages):
        try:
            r = self._primario.vision(messages)
            _stat("vision", self._nombre_primario, ok=True)
            return r
        except Exception as exc:
            _stat("vision", self._nombre_primario, ok=False)
            raise exc


# ---------------------------------------------------------------------------
# Wrapper estadístico para proveedores directos (sin FallbackProvider)
# ---------------------------------------------------------------------------

class _ProveedorInstrumentado(LLMProvider):
    """Envuelve un proveedor directo añadiendo registro de estadísticas."""

    def __init__(self, inner: LLMProvider, nombre: str):
        self._inner              = inner
        self._nombre             = nombre
        self.modelo              = inner.modelo
        self.max_tokens_peticion = inner.max_tokens_peticion

    def chat(self, messages, temperature=0.7, max_tokens=None):
        try:
            r = self._inner.chat(messages, temperature, max_tokens)
            _stat("chat", self._nombre, ok=True)
            return r
        except LLMError as exc:
            _stat("chat", self._nombre, ok=False, error_tipo=exc.tipo)
            raise

    def transcribir(self, audio_bytes, filename="audio.ogg"):
        try:
            r = self._inner.transcribir(audio_bytes, filename)
            _stat("audio", self._nombre, ok=True)
            return r
        except Exception as exc:
            _stat("audio", self._nombre, ok=False)
            raise exc

    def vision(self, messages):
        try:
            r = self._inner.vision(messages)
            _stat("vision", self._nombre, ok=True)
            return r
        except Exception as exc:
            _stat("vision", self._nombre, ok=False)
            raise exc


# ---------------------------------------------------------------------------
# Factory / singleton
# ---------------------------------------------------------------------------

_proveedor: LLMProvider | None = None

_PROVEEDORES = {
    "groq":    GroqProvider,
    "ollama":  OllamaProvider,
    "gemini":  GeminiProvider,
    "mistral": MistralProvider,
    "openai":  OpenAIProvider,
    "azure":   AzureOpenAIProvider,
}


def _instanciar_proveedor(nombre: str, modelo_override: str = "") -> LLMProvider:
    """Crea una instancia del proveedor indicado, usando modelo_override si se indica."""
    clase = _PROVEEDORES.get(nombre)
    if clase is None:
        logger.warning("Proveedor '%s' desconocido — usando groq", nombre)
        clase = GroqProvider

    if not modelo_override:
        return clase()

    # Aplicar el modelo de override temporalmente para esta instanciación
    modelo_guardado = os.environ.get("LLM_MODEL")
    os.environ["LLM_MODEL"] = modelo_override
    instancia = clase()
    if modelo_guardado is not None:
        os.environ["LLM_MODEL"] = modelo_guardado
    else:
        os.environ.pop("LLM_MODEL", None)
    return instancia


def obtener_proveedor() -> LLMProvider:
    """Devuelve la instancia singleton del proveedor activo."""
    global _proveedor
    if _proveedor is None:
        nombre          = os.getenv("LLM_PROVIDER",          "groq").lower()
        nombre_fallback = os.getenv("LLM_FALLBACK_PROVIDER", "").lower()
        modelo_fallback = os.getenv("LLM_FALLBACK_MODEL",    "")

        primario = _instanciar_proveedor(nombre)

        if nombre_fallback and nombre_fallback != nombre and nombre_fallback in _PROVEEDORES:
            # Si no se define LLM_FALLBACK_MODEL, limpiar LLM_MODEL para que el
            # proveedor de respaldo use su propio modelo por defecto (evita que
            # Mistral intente cargar llama-3.3-70b-versatile, por ejemplo).
            override = modelo_fallback or ""
            if not override:
                modelo_principal = os.environ.get("LLM_MODEL")
                if modelo_principal:
                    os.environ.pop("LLM_MODEL", None)
                    secundario = _PROVEEDORES[nombre_fallback]()
                    os.environ["LLM_MODEL"] = modelo_principal
                else:
                    secundario = _PROVEEDORES[nombre_fallback]()
            else:
                secundario = _instanciar_proveedor(nombre_fallback, override)

            _proveedor = FallbackProvider(primario, secundario, nombre, nombre_fallback)
            logger.info(
                "Proveedor IA: %s (%s) | Respaldo: %s (%s)",
                nombre, primario.modelo, nombre_fallback, secundario.modelo,
            )
        else:
            _proveedor = _ProveedorInstrumentado(primario, nombre)
            logger.info("Proveedor IA: %s | Modelo: %s", nombre, primario.modelo)

    return _proveedor


# ---------------------------------------------------------------------------
# Proveedor para modo privado (datos en la UE, RGPD)
# ---------------------------------------------------------------------------

_proveedor_privado: LLMProvider | None = None


def obtener_proveedor_privado() -> LLMProvider:
    """Devuelve el proveedor para el modo privado (/privado).

    Las peticiones en modo privado se enrutan a este proveedor, que por
    defecto es Mistral (empresa francesa, datos procesados en la UE,
    cumple RGPD). Configurable con LLM_PRIVADO_PROVIDER en .env.

    Variables de entorno:
      LLM_PRIVADO_PROVIDER  nombre del proveedor  (por defecto: mistral)
      LLM_PRIVADO_MODEL     modelo                (vacío = default del proveedor)
    """
    global _proveedor_privado
    if _proveedor_privado is None:
        nombre = os.getenv("LLM_PRIVADO_PROVIDER", "mistral").lower()
        modelo = os.getenv("LLM_PRIVADO_MODEL", "")
        _proveedor_privado = _instanciar_proveedor(nombre, modelo)
        logger.info(
            "Proveedor privado (EU/RGPD): %s (%s)", nombre, _proveedor_privado.modelo
        )
    return _proveedor_privado
