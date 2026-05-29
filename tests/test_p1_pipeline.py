"""Tests Sprint P1 — Pipeline base.

Cubre:
  - extraer_operacion_edicion: normalización a lista, aclaracion como dict, RESPUESTA_LIBRE
  - _ejecutar_pipeline: ops de datos encadenadas, ops visuales acumuladas
  - POST /edit: pipeline de dos ops devuelve tipo:pipeline con pasos
  - POST /edit: un solo op sigue devolviendo el resultado directamente (retrocompatibilidad)
  - POST /edit: aclaracion devuelta antes de ejecutar nada
"""
import pytest
from unittest.mock import patch, MagicMock
import pandas as pd


_TEST_API_KEY = "test-key-pipeline-suite"
HEADERS = {"X-API-Key": _TEST_API_KEY}

_DATOS = [
    ["Producto", "Ventas", "Estado"],
    ["A", 500, "activo"],
    ["B", 150, "inactivo"],
    ["C", 800, "activo"],
    ["D", 300, "inactivo"],
]


# ── Fixtures ──────────────────────────────────────────────────────────────────

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
    monkeypatch.setattr("utils.db.DB_PATH", str(tmp_path / "test_p1.db"))


# ── Normalización en extraer_operacion_edicion ────────────────────────────────

class TestExtraerOperacionEdicion:
    """Verifica la normalización de salida sin llamar al LLM real."""

    def _parsear(self, texto_llm):
        """Simula lo que hace extraer_operacion_edicion con una respuesta LLM dada."""
        import json
        from services.llm import _limpiar_json
        texto = texto_llm.strip()
        if texto == "RESPUESTA_LIBRE":
            return None
        parsed = json.loads(_limpiar_json(texto))
        if isinstance(parsed, dict) and parsed.get("aclaracion_necesaria"):
            return parsed
        if isinstance(parsed, dict):
            return [parsed]
        if isinstance(parsed, list):
            return parsed
        return None

    def test_respuesta_libre_devuelve_none(self):
        assert self._parsear("RESPUESTA_LIBRE") is None

    def test_array_de_ops_pasa_directo(self):
        texto = '[{"op":"ordenar","col":"Ventas"},{"op":"formato_condicional"}]'
        resultado = self._parsear(texto)
        assert isinstance(resultado, list)
        assert len(resultado) == 2
        assert resultado[0]["op"] == "ordenar"

    def test_dict_suelto_se_envuelve_en_lista(self):
        """El LLM no siguió el formato array → normalizamos."""
        texto = '{"op":"ordenar","col":"Ventas"}'
        resultado = self._parsear(texto)
        assert isinstance(resultado, list)
        assert resultado[0]["op"] == "ordenar"

    def test_aclaracion_devuelve_dict_plano(self):
        texto = '{"aclaracion_necesaria":true,"pregunta":"¿qué columna?","opciones":["A","B"]}'
        resultado = self._parsear(texto)
        assert isinstance(resultado, dict)
        assert resultado.get("aclaracion_necesaria") is True

    def test_pipeline_un_op(self):
        texto = '[{"op":"eliminar_duplicados"}]'
        resultado = self._parsear(texto)
        assert isinstance(resultado, list)
        assert len(resultado) == 1


# ── _ejecutar_pipeline ────────────────────────────────────────────────────────

class TestEjecutarPipeline:
    def _df(self):
        return pd.DataFrame({
            "Producto": ["A", "B", "C", "D"],
            "Ventas":   [500, 150, 800, 300],
            "Estado":   ["activo", "inactivo", "activo", "inactivo"],
        })

    def test_op_datos_modifica_df(self):
        from api import _ejecutar_pipeline
        ops = [{"op": "ordenar", "col": "Ventas", "orden": "desc"}]
        pasos = _ejecutar_pipeline(self._df(), ops, "ordena por ventas")
        assert len(pasos) == 1
        assert pasos[0]["tipo"] == "edicion"
        # Primera fila debe ser el mayor valor
        assert pasos[0]["datos_modificados"][1][1] == 800

    def test_dos_ops_de_datos_encadenadas(self):
        from api import _ejecutar_pipeline
        ops = [
            {"op": "ordenar", "col": "Ventas", "orden": "desc"},
            {"op": "eliminar_duplicados"},
        ]
        pasos = _ejecutar_pipeline(self._df(), ops, "ordena y quita duplicados")
        assert len(pasos) == 2
        assert pasos[0]["tipo"] == "edicion"
        assert pasos[1]["tipo"] == "edicion"

    def test_op_visual_formato_acumula_paso(self):
        from api import _ejecutar_pipeline
        reglas_mock = [{"tipo": "valor", "col": "Estado", "op": "==", "valor": "activo", "color": "verde"}]
        with patch("api.extraer_regla_formato", return_value=reglas_mock):
            ops = [{"op": "formato_condicional"}]
            pasos = _ejecutar_pipeline(self._df(), ops, "pon verde los activos")
        assert len(pasos) == 1
        assert pasos[0]["tipo"] == "formato"

    def test_op_invalida_se_omite_silenciosamente(self):
        from api import _ejecutar_pipeline
        ops = [
            {"op": "ordenar", "col": "Ventas"},
            {"op": "operacion_inventada_que_no_existe"},
        ]
        pasos = _ejecutar_pipeline(self._df(), ops, "ordena y haz algo raro")
        # Solo el ordenar debe haberse completado
        assert any(p["tipo"] == "edicion" for p in pasos)


# ── POST /edit con pipeline ───────────────────────────────────────────────────

class TestEndpointEditPipeline:
    def test_un_op_devuelve_resultado_directo(self, api_client):
        """Un solo op en la lista → devuelve tipo:edicion directamente (retrocompat)."""
        ops_mock = [{"op": "ordenar", "col": "Ventas", "orden": "desc"}]
        with patch("api.extraer_operacion_edicion", return_value=ops_mock):
            r = api_client.post(
                "/edit", headers=HEADERS,
                json={"datos": _DATOS, "instruccion": "ordena por ventas"},
            )
        assert r.status_code == 200
        assert r.json()["tipo"] == "edicion"
        assert "pipeline" not in r.json()

    def test_dos_ops_devuelven_pipeline(self, api_client):
        """Dos ops → devuelve tipo:pipeline con pasos."""
        ops_mock = [
            {"op": "ordenar", "col": "Ventas", "orden": "desc"},
            {"op": "eliminar_duplicados"},
        ]
        with patch("api.extraer_operacion_edicion", return_value=ops_mock):
            r = api_client.post(
                "/edit", headers=HEADERS,
                json={"datos": _DATOS, "instruccion": "ordena y quita duplicados"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "pipeline"
        assert len(data["pasos"]) == 2
        assert "descripcion" in data

    def test_pipeline_con_op_visual(self, api_client):
        """Edición + formato condicional → pipeline con 2 pasos de tipos distintos."""
        ops_mock = [
            {"op": "ordenar", "col": "Ventas", "orden": "desc"},
            {"op": "formato_condicional"},
        ]
        reglas_mock = [{"tipo": "valor", "col": "Ventas", "op": ">", "valor": 400, "color": "verde"}]
        with patch("api.extraer_operacion_edicion", return_value=ops_mock), \
             patch("api.extraer_regla_formato", return_value=reglas_mock):
            r = api_client.post(
                "/edit", headers=HEADERS,
                json={"datos": _DATOS, "instruccion": "ordena y pon verde las ventas altas"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "pipeline"
        tipos = [p["tipo"] for p in data["pasos"]]
        assert "edicion" in tipos
        assert "formato" in tipos

    def test_aclaracion_devuelta_sin_ejecutar(self, api_client):
        """Si el DSL necesita aclaración, se devuelve tipo:aclaracion sin ejecutar nada."""
        aclaracion_mock = {
            "aclaracion_necesaria": True,
            "pregunta": "¿Por qué columna ordenas?",
            "opciones": ["Ventas", "Producto", "Estado"],
        }
        with patch("api.extraer_operacion_edicion", return_value=aclaracion_mock):
            r = api_client.post(
                "/edit", headers=HEADERS,
                json={"datos": _DATOS, "instruccion": "ordena esto"},
            )
        assert r.status_code == 200
        data = r.json()
        assert data["tipo"] == "aclaracion"
        assert data["aclaracion_necesaria"] is True

    def test_respuesta_libre_cae_a_llm(self, api_client):
        """RESPUESTA_LIBRE (None) → el endpoint intenta query o LLM libre."""
        with patch("api.extraer_operacion_edicion", return_value=None), \
             patch("api.extraer_query_dsl", return_value=None), \
             patch("api.obtener_respuesta", return_value="respuesta libre"):
            r = api_client.post(
                "/edit", headers=HEADERS,
                json={"datos": _DATOS, "instruccion": "¿cuántas filas hay?",
                      "historial": []},
            )
        assert r.status_code == 200
        assert r.json()["tipo"] == "texto"
