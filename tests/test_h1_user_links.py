"""Tests Sprint H1 — vinculación Telegram ↔ email (user_links.py).

Cubre:
  - vincular: alta, reasignación, idempotencia
  - obtener_telegram_id: email conocido / desconocido
  - obtener_emails: 0, 1 y N emails por usuario
  - desvincular: uno / todos
  - obtener_todos_los_vinculos: orden y estructura
"""
import pytest
from unittest.mock import patch


# Reapunta DB_PATH a un archivo temporal para aislar los tests
@pytest.fixture(autouse=True)
def db_temporal(tmp_path, monkeypatch):
    ruta = str(tmp_path / "test_links.db")
    monkeypatch.setattr("utils.user_links.DB_PATH", ruta)


# ── vincular ──────────────────────────────────────────────────────────────────

class TestVincular:
    def test_alta_basica(self):
        from utils.user_links import vincular, obtener_telegram_id
        vincular(111, "usuario@empresa.com")
        assert obtener_telegram_id("usuario@empresa.com") == 111

    def test_normaliza_email_a_minusculas(self):
        from utils.user_links import vincular, obtener_telegram_id
        vincular(222, "Usuario@Empresa.COM")
        assert obtener_telegram_id("usuario@empresa.com") == 222

    def test_idempotente(self):
        """Vincular el mismo email dos veces no duplica ni falla."""
        from utils.user_links import vincular, obtener_emails
        vincular(333, "a@b.com")
        vincular(333, "a@b.com")
        assert obtener_emails(333) == ["a@b.com"]

    def test_reasignar_email_a_otro_telegram(self):
        """Un email ya vinculado a otro usuario se reasigna al nuevo."""
        from utils.user_links import vincular, obtener_telegram_id
        vincular(100, "shared@b.com")
        vincular(200, "shared@b.com")
        assert obtener_telegram_id("shared@b.com") == 200

    def test_multiples_emails_mismo_usuario(self):
        """Un usuario puede tener varios emails vinculados."""
        from utils.user_links import vincular, obtener_emails
        vincular(400, "personal@gmail.com")
        vincular(400, "trabajo@empresa.eu")
        emails = obtener_emails(400)
        assert len(emails) == 2
        assert "personal@gmail.com" in emails
        assert "trabajo@empresa.eu" in emails


# ── obtener_telegram_id ───────────────────────────────────────────────────────

class TestObtenerTelegramId:
    def test_email_existente(self):
        from utils.user_links import vincular, obtener_telegram_id
        vincular(500, "existe@dom.com")
        assert obtener_telegram_id("existe@dom.com") == 500

    def test_email_inexistente_devuelve_none(self):
        from utils.user_links import obtener_telegram_id
        assert obtener_telegram_id("noexiste@dom.com") is None


# ── desvincular ───────────────────────────────────────────────────────────────

class TestDesvincular:
    def test_desvincular_email_concreto(self):
        from utils.user_links import vincular, desvincular, obtener_emails
        vincular(600, "a@b.com")
        vincular(600, "c@d.com")
        eliminados = desvincular(600, "a@b.com")
        assert eliminados == 1
        assert obtener_emails(600) == ["c@d.com"]

    def test_desvincular_todos(self):
        from utils.user_links import vincular, desvincular, obtener_emails
        vincular(700, "x@y.com")
        vincular(700, "z@w.com")
        eliminados = desvincular(700)
        assert eliminados == 2
        assert obtener_emails(700) == []

    def test_desvincular_inexistente_devuelve_cero(self):
        from utils.user_links import desvincular
        assert desvincular(999, "nada@nada.com") == 0


# ── obtener_todos_los_vinculos ────────────────────────────────────────────────

class TestObtenerTodosLosVinculos:
    def test_estructura_de_filas(self):
        from utils.user_links import vincular, obtener_todos_los_vinculos
        vincular(800, "uno@a.com")
        vincular(801, "dos@a.com")
        vinculos = obtener_todos_los_vinculos()
        assert len(vinculos) == 2
        for v in vinculos:
            assert "telegram_id" in v
            assert "email" in v
            assert "creado_en" in v

    def test_contiene_ambos_registros(self):
        """Ambos vínculos aparecen en la lista (no dependemos del orden exacto
        porque CURRENT_TIMESTAMP tiene precisión de segundos y pueden coincidir)."""
        from utils.user_links import vincular, obtener_todos_los_vinculos
        vincular(900, "primero@a.com")
        vincular(901, "segundo@a.com")
        emails = {v["email"] for v in obtener_todos_los_vinculos()}
        assert "primero@a.com" in emails
        assert "segundo@a.com" in emails

    def test_sin_vinculos_devuelve_lista_vacia(self):
        from utils.user_links import obtener_todos_los_vinculos
        assert obtener_todos_los_vinculos() == []
