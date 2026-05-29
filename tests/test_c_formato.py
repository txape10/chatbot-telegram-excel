"""Tests Sprint C — Formato condicional.

Cubre:
  - _describir_una_regla() para los 7 tipos de regla
  - _describir_reglas_formato() para múltiples reglas
  - POST /format: respuesta con regla válida (devuelve reglas=[...])
  - POST /format: fallback cuando el LLM no puede interpretar
  - POST /format: sin datos devuelve error
"""
import pytest
from unittest.mock import patch


# ── Fixtures ──────────────────────────────────────────────────────────────────

_TEST_API_KEY = "test-key-formato-suite"
HEADERS = {"X-API-Key": _TEST_API_KEY}

_DATOS_VENTAS = [
    ["Producto", "Ventas", "Estado"],
    ["A", 500, "activo"],
    ["B", 150, "inactivo"],
    ["C", 800, "activo"],
    ["D", 300, "inactivo"],
]


@pytest.fixture(scope="module")
def api_client():
    from fastapi.testclient import TestClient
    import api as api_mod

    orig_key      = api_mod._API_KEY
    orig_telegram = api_mod._ENABLE_TELEGRAM

    api_mod._API_KEY         = _TEST_API_KEY
    api_mod._ENABLE_TELEGRAM = False

    yield TestClient(api_mod.app)

    api_mod._API_KEY         = orig_key
    api_mod._ENABLE_TELEGRAM = orig_telegram


@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    monkeypatch.setattr("utils.db.DB_PATH", str(tmp_path / "test_fmt.db"))


# ── _describir_una_regla ──────────────────────────────────────────────────────

class TestDescribirRegla:
    def _describir(self, regla):
        from api import _describir_una_regla
        return _describir_una_regla(regla)

    def test_tipo_valor(self):
        desc = self._describir({"tipo": "valor", "col": "Ventas", "op": ">", "valor": 500})
        assert "Ventas" in desc
        assert ">" in desc
        assert "500" in desc

    def test_tipo_top_bottom_top(self):
        desc = self._describir({
            "tipo": "top_bottom", "col": "Ventas",
            "direccion": "top", "n": 5, "porcentaje": False,
        })
        assert "Ventas" in desc
        assert "superiores" in desc.lower()
        assert "5" in desc

    def test_tipo_top_bottom_bottom_pct(self):
        desc = self._describir({
            "tipo": "top_bottom", "col": "Ventas",
            "direccion": "bottom", "n": 10, "porcentaje": True,
        })
        assert "inferiores" in desc.lower()
        assert "10" in desc
        assert "%" in desc

    def test_tipo_escala(self):
        desc = self._describir({"tipo": "escala", "col": "Ventas"})
        assert "Ventas" in desc
        assert "color" in desc.lower() or "escala" in desc.lower()

    def test_tipo_barra(self):
        desc = self._describir({"tipo": "barra", "col": "Ventas"})
        assert "Ventas" in desc
        assert "barra" in desc.lower()

    def test_tipo_icono(self):
        desc = self._describir({"tipo": "icono", "col": "Ventas", "estilo": "semaforo"})
        assert "Ventas" in desc
        assert "semaforo" in desc or "icono" in desc.lower()

    def test_tipo_texto(self):
        desc = self._describir({
            "tipo": "texto", "col": "Estado",
            "op": "contiene", "valor": "activo",
        })
        assert "Estado" in desc
        assert "activo" in desc

    def test_tipo_formula(self):
        desc = self._describir({"tipo": "formula", "formula": "=A1>100", "color": "rojo"})
        assert "formula" in desc.lower() or "fórmula" in desc.lower()

    def test_tipo_desconocido_no_falla(self):
        desc = self._describir({"tipo": "xyz"})
        assert isinstance(desc, str)

    def test_tipo_formula_con_col(self):
        desc = self._describir({"tipo": "formula", "col": "Ventas", "formula": "=A1>0"})
        assert "Ventas" in desc


# ── _describir_reglas_formato (múltiples reglas) ──────────────────────────────

class TestDescribirReglas:
    def _describir(self, reglas):
        from api import _describir_reglas_formato
        return _describir_reglas_formato(reglas)

    def test_regla_unica(self):
        desc = self._describir([{"tipo": "valor", "col": "Ventas", "op": ">", "valor": 500}])
        assert "Ventas" in desc
        # No muestra conteo cuando es una sola regla
        assert "2 reglas" not in desc

    def test_multiples_reglas(self):
        reglas = [
            {"tipo": "valor", "col": "Estado", "op": "==", "valor": "Rechazado", "color": "rojo"},
            {"tipo": "valor", "col": "Estado", "op": "==", "valor": "Aprobado",  "color": "verde"},
            {"tipo": "valor", "col": "Estado", "op": "==", "valor": "Pendiente", "color": "amarillo"},
        ]
        desc = self._describir(reglas)
        assert "3 reglas" in desc
        assert "Estado" in desc


# ── POST /format ──────────────────────────────────────────────────────────────

class TestEndpointFormat:
    def test_regla_valor_devuelve_formato(self, api_client):
        regla_mock = [
            {"tipo": "valor", "col": "Ventas", "op": ">", "valor": 400, "color": "verde"}
        ]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "colorea en verde ventas > 400"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "formato"
        assert data["reglas"] == regla_mock
        assert "descripcion" in data
        assert "Ventas" in data["descripcion"]

    def test_regla_escala_devuelve_formato(self, api_client):
        regla_mock = [{"tipo": "escala", "col": "Ventas", "colores": ["rojo", "verde"]}]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "escala de color en ventas"},
            )
        assert r.status_code == 200
        assert r.json()["tipo"] == "formato"

    def test_regla_top_bottom(self, api_client):
        regla_mock = [{
            "tipo": "top_bottom", "col": "Ventas",
            "direccion": "top", "n": 3, "porcentaje": False, "color": "verde",
        }]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "top 3 ventas en verde"},
            )
        assert r.status_code == 200
        assert "3" in r.json()["descripcion"]

    def test_regla_icono(self, api_client):
        regla_mock = [{"tipo": "icono", "col": "Ventas", "estilo": "semaforo"}]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "semáforo en ventas"},
            )
        assert r.status_code == 200
        assert r.json()["tipo"] == "formato"

    def test_llm_no_interpreta_devuelve_texto(self, api_client):
        with patch("api.extraer_regla_formato", return_value=None):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "haz algo bonito"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "texto"
        assert "interpretar" in data["respuesta"].lower() or "formato" in data["respuesta"].lower()

    def test_sin_datos_devuelve_400(self, api_client):
        r = api_client.post(
            "/format",
            headers=HEADERS,
            json={"datos": [["Producto"]], "instruccion": "colorea"},
        )
        assert r.status_code == 400

    def test_sin_api_key_devuelve_error(self, api_client):
        r = api_client.post(
            "/format",
            json={"datos": _DATOS_VENTAS, "instruccion": "colorea"},
        )
        assert r.status_code in (403, 422)

    def test_regla_texto(self, api_client):
        regla_mock = [{
            "tipo": "texto", "col": "Estado",
            "op": "contiene", "valor": "activo", "color": "verde",
        }]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "verde si estado contiene activo"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "formato"
        assert "activo" in data["descripcion"]

    def test_regla_formula(self, api_client):
        regla_mock = [{"tipo": "formula", "formula": "=$B2>500", "color": "azul"}]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "formula para ventas mayores"},
            )
        assert r.status_code == 200
        assert r.json()["tipo"] == "formato"

    def test_regla_barra(self, api_client):
        regla_mock = [{"tipo": "barra", "col": "Ventas", "color": "azul"}]
        with patch("api.extraer_regla_formato", return_value=regla_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS, "instruccion": "barras de datos en ventas"},
            )
        assert r.status_code == 200
        assert r.json()["tipo"] == "formato"

    def test_multiples_reglas_mismo_endpoint(self, api_client):
        """Tres reglas en una sola petición (caso del bug reportado)."""
        reglas_mock = [
            {"tipo": "valor", "col": "Estado", "op": "==", "valor": "Rechazado", "color": "rojo"},
            {"tipo": "valor", "col": "Estado", "op": "==", "valor": "Aprobado",  "color": "verde"},
            {"tipo": "valor", "col": "Estado", "op": "==", "valor": "Pendiente", "color": "amarillo"},
        ]
        with patch("api.extraer_regla_formato", return_value=reglas_mock):
            r = api_client.post(
                "/format",
                headers=HEADERS,
                json={"datos": _DATOS_VENTAS,
                      "instruccion": "rojo=Rechazado, verde=Aprobado, amarillo=Pendiente"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "formato"
        assert len(data["reglas"]) == 3
        assert "3 reglas" in data["descripcion"]
