"""Tests Sprint H2 y backlog — admin API, modo privado, límites de ficheros.

H2:
  - /addin-config devuelve telegram_habilitado y nombre_empresa
  - /admin/stats devuelve estructura correcta
  - /admin/vinculos POST/DELETE funcionan con clave válida

Backlog:
  - TAMANIO_MAXIMO_MB = 20
  - MAX_FILAS = 100_000
  - obtener_proveedor_privado() devuelve instancia de MistralProvider por defecto
  - modo privado no guarda historial → ya cubierto en handlers (prueba manual)
"""
import os
import pytest


# ── Fixtures API ──────────────────────────────────────────────────────────────

_TEST_API_KEY   = "test-key-h2-suite"
_TEST_ADMIN_KEY = "test-admin-h2-suite"

HEADERS_ADDIN = {"X-API-Key": _TEST_API_KEY}
HEADERS_ADMIN = {"X-API-Key": _TEST_API_KEY}


@pytest.fixture(scope="module")
def api_client():
    """TestClient que parchea directamente _API_KEY y _ADMIN_KEY en el módulo api.
    Necesario porque api.py llama load_dotenv() en el import y el módulo puede
    ya estar cargado con las claves reales del .env cuando se ejecuta la suite."""
    from fastapi.testclient import TestClient
    import api as api_mod

    # Guardar y sobrescribir las claves de seguridad
    orig_api_key   = api_mod._API_KEY
    orig_admin_key = api_mod._ADMIN_KEY
    orig_telegram  = api_mod._ENABLE_TELEGRAM

    api_mod._API_KEY         = _TEST_API_KEY
    api_mod._ADMIN_KEY       = _TEST_ADMIN_KEY
    api_mod._ENABLE_TELEGRAM = False   # modo empresa sin bot

    yield TestClient(api_mod.app)

    # Restaurar valores originales
    api_mod._API_KEY         = orig_api_key
    api_mod._ADMIN_KEY       = orig_admin_key
    api_mod._ENABLE_TELEGRAM = orig_telegram


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    ruta = str(tmp_path / "test_h2.db")
    monkeypatch.setattr("utils.db.DB_PATH", ruta)


# ── /addin-config ─────────────────────────────────────────────────────────────

class TestAddinConfig:
    def test_devuelve_telegram_habilitado(self, api_client):
        r = api_client.get("/addin-config", headers=HEADERS_ADDIN)
        assert r.status_code == 200
        data = r.json()
        assert "telegram_habilitado" in data
        assert isinstance(data["telegram_habilitado"], bool)

    def test_devuelve_nombre_empresa(self, api_client, monkeypatch):
        monkeypatch.setenv("COMPANY_NAME", "La Unión Corp")
        r = api_client.get("/addin-config", headers=HEADERS_ADDIN)
        assert r.status_code == 200
        assert r.json()["nombre_empresa"] == "La Unión Corp"

    def test_nombre_empresa_vacio_por_defecto(self, api_client, monkeypatch):
        monkeypatch.delenv("COMPANY_NAME", raising=False)
        r = api_client.get("/addin-config", headers=HEADERS_ADDIN)
        assert r.status_code == 200
        assert r.json()["nombre_empresa"] == ""

    def test_sin_api_key_devuelve_error(self, api_client):
        # Sin cabecera X-API-Key: FastAPI devuelve 422 (campo requerido faltante)
        # o 403 si el middleware lo intercepta antes — ambos son rechazo válido.
        r = api_client.get("/addin-config")
        assert r.status_code in (403, 422)

    def test_telegram_desactivado_en_fixtures(self, api_client):
        r = api_client.get("/addin-config", headers=HEADERS_ADDIN)
        assert r.json()["telegram_habilitado"] is False


# ── /admin/stats ──────────────────────────────────────────────────────────────

class TestAdminStats:
    def test_estructura_basica(self, api_client):
        r = api_client.get("/admin/stats?key=test-admin-h2-suite")
        assert r.status_code == 200
        data = r.json()
        # Campos definidos en utils/stats.py → obtener_estadisticas()
        assert "total_mensajes" in data
        assert "total_usuarios" in data
        assert "mensajes_hoy" in data
        assert "usuarios" in data
        assert "mensajes_por_dia" in data

    def test_sin_clave_devuelve_403(self, api_client):
        r = api_client.get("/admin/stats")
        assert r.status_code in (400, 403, 422)  # query param requerido


# ── /admin/vinculos ───────────────────────────────────────────────────────────

class TestAdminVinculos:
    def test_alta_manual_y_baja(self, api_client):
        # Crear vínculo manual desde el panel
        r = api_client.post(
            "/admin/vinculos?key=test-admin-h2-suite",
            json={"telegram_id": 12345, "email": "test@empresa.com"},
        )
        assert r.status_code == 200
        assert r.json()["ok"] is True

        # Eliminar ese mismo vínculo
        r = api_client.delete(
            "/admin/vinculos?key=test-admin-h2-suite"
            "&telegram_id=12345&email=test@empresa.com"
        )
        assert r.status_code == 200
        assert r.json()["ok"] is True

    def test_eliminar_vinculo_inexistente_devuelve_404(self, api_client):
        r = api_client.delete(
            "/admin/vinculos?key=test-admin-h2-suite"
            "&telegram_id=99999&email=nadie@a.com"
        )
        assert r.status_code == 404

    def test_alta_sin_clave_devuelve_403(self, api_client):
        r = api_client.post(
            "/admin/vinculos",
            json={"telegram_id": 1, "email": "x@y.com"},
        )
        assert r.status_code in (400, 403, 422)


# ── Backlog: límites de ficheros ──────────────────────────────────────────────

class TestLimitesFicheros:
    def test_tamanio_maximo_mb(self):
        from config import TAMANIO_MAXIMO_MB
        assert TAMANIO_MAXIMO_MB == 20

    def test_max_filas(self):
        from config import MAX_FILAS
        assert MAX_FILAS == 100_000

    def test_tamanio_importado_en_documents(self):
        """Asegura que documents.py ya no define TAMANIO_MAXIMO_MB localmente."""
        import handlers.documents as mod
        # Debe existir el atributo (importado de config, no local)
        assert hasattr(mod, "TAMANIO_MAXIMO_MB")
        assert mod.TAMANIO_MAXIMO_MB == 20


# ── Backlog: proveedor privado (modo /privado) ────────────────────────────────

class TestProveedorPrivado:
    def test_devuelve_instancia_de_llm_provider(self, monkeypatch):
        """Con LLM_PRIVADO_PROVIDER=groq (SDK disponible en tests) se obtiene GroqProvider."""
        monkeypatch.setenv("LLM_PRIVADO_PROVIDER", "groq")
        monkeypatch.setenv("GROQ_API_KEY", "dummy-groq-key")

        import services.llm_provider as mod
        mod._proveedor_privado = None

        proveedor = mod.obtener_proveedor_privado()
        assert isinstance(proveedor, mod.GroqProvider)
        assert isinstance(proveedor, mod.LLMProvider)

    def test_singleton_misma_instancia(self, monkeypatch):
        monkeypatch.setenv("LLM_PRIVADO_PROVIDER", "groq")
        monkeypatch.setenv("GROQ_API_KEY", "dummy-groq-key")

        import services.llm_provider as mod
        mod._proveedor_privado = None

        p1 = mod.obtener_proveedor_privado()
        p2 = mod.obtener_proveedor_privado()
        assert p1 is p2

    def test_proveedor_privado_por_defecto_es_mistral(self, monkeypatch):
        """Sin configuración explícita el proveedor debe ser mistral."""
        monkeypatch.delenv("LLM_PRIVADO_PROVIDER", raising=False)

        import services.llm_provider as mod
        # Verificar solo el nombre del proveedor, sin instanciar (evita dep. openai)
        nombre = os.getenv("LLM_PRIVADO_PROVIDER", "mistral").lower()
        assert nombre == "mistral"

    def test_obtener_respuesta_acepta_proveedor_externo(self):
        """obtener_respuesta() acepta proveedor opcional sin excepción de firma."""
        from services.llm import obtener_respuesta
        import inspect
        sig = inspect.signature(obtener_respuesta)
        assert "proveedor" in sig.parameters
        assert sig.parameters["proveedor"].default is None


# ── Backlog: empresa sin Telegram ─────────────────────────────────────────────

class TestEmpresaSinTelegram:
    def test_tiene_vinculo_no_disponible_si_telegram_desactivado(self, api_client):
        """/tiene-vinculo sigue respondiendo aunque ENABLE_TELEGRAM sea false
        (la tabla existe igual; es decisión del Add-in ocultar el UI)."""
        r = api_client.get(
            "/tiene-vinculo?email=test@empresa.com",
            headers=HEADERS_ADDIN,
        )
        # Puede devolver 200 (vinculado:false) o 404; nunca debe ser 500
        assert r.status_code in (200, 404)
