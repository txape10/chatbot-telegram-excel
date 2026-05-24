#!/usr/bin/env python3
"""Instalador interactivo — Asistente Excel.

Genera el fichero .env con todas las variables necesarias
según los módulos y proveedor de IA elegidos.

Uso:
  python instalar.py
"""

import os
import secrets
import sys

# ── Colores de consola (funciona en Windows 10+ y Linux) ─────────────────────
class C:
    VERDE   = "\033[92m"
    AMARILLO= "\033[93m"
    ROJO    = "\033[91m"
    AZUL    = "\033[94m"
    NEGRITA = "\033[1m"
    RESET   = "\033[0m"

def ok(msg):   print(f"  {C.VERDE}✔{C.RESET}  {msg}")
def info(msg): print(f"  {C.AZUL}ℹ{C.RESET}  {msg}")
def warn(msg): print(f"  {C.AMARILLO}⚠{C.RESET}  {msg}")

def titulo(msg):
    print(f"\n{C.NEGRITA}{C.AZUL}{'─'*55}{C.RESET}")
    print(f"{C.NEGRITA}{C.AZUL}  {msg}{C.RESET}")
    print(f"{C.NEGRITA}{C.AZUL}{'─'*55}{C.RESET}\n")

def preguntar(prompt, por_defecto=""):
    sufijo = f" [{por_defecto}]" if por_defecto else ""
    valor = input(f"  {prompt}{sufijo}: ").strip()
    return valor if valor else por_defecto

def elegir(prompt, opciones: list[tuple[str, str]], por_defecto=1) -> int:
    """Muestra un menú numerado y devuelve el índice elegido (0-based)."""
    for i, (clave, desc) in enumerate(opciones, 1):
        marca = f"{C.VERDE}►{C.RESET}" if i == por_defecto else " "
        print(f"  {marca} [{i}] {C.NEGRITA}{clave}{C.RESET} — {desc}")
    while True:
        raw = input(f"\n  Elige una opción [{por_defecto}]: ").strip()
        if not raw:
            return por_defecto - 1
        if raw.isdigit() and 1 <= int(raw) <= len(opciones):
            return int(raw) - 1
        print(f"  {C.ROJO}Opción no válida.{C.RESET}")

def elegir_multiple(prompt, opciones: list[tuple[str, str]], por_defecto="1,2") -> list[int]:
    """Permite seleccionar varias opciones separadas por coma."""
    for i, (clave, desc) in enumerate(opciones, 1):
        print(f"    [{i}] {C.NEGRITA}{clave}{C.RESET} — {desc}")
    while True:
        raw = input(f"\n  Elige (separados por coma) [{por_defecto}]: ").strip()
        if not raw:
            raw = por_defecto
        indices = []
        try:
            for parte in raw.split(","):
                n = int(parte.strip())
                if 1 <= n <= len(opciones):
                    indices.append(n - 1)
        except ValueError:
            pass
        if indices:
            return indices
        print(f"  {C.ROJO}Selección no válida.{C.RESET}")


# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Activar colores en Windows
    if sys.platform == "win32":
        os.system("color")

    print(f"\n{C.NEGRITA}{'='*55}")
    print("  Instalador — Asistente Excel con IA")
    print(f"{'='*55}{C.RESET}\n")

    if os.path.exists(".env"):
        warn("Ya existe un fichero .env.")
        resp = preguntar("¿Sobreescribir? (s/N)", "N")
        if resp.lower() != "s":
            print("\n  Instalación cancelada.\n")
            sys.exit(0)

    config = {}

    # ── 1. Módulos ────────────────────────────────────────────────────────────
    titulo("1. ¿Qué módulos quieres instalar?")
    modulos_opts = [
        ("Bot Telegram",      "Chatbot accesible desde cualquier dispositivo con Telegram"),
        ("Add-in Excel",      "Panel lateral dentro de Excel, sin salir de la aplicación"),
    ]
    modulos_sel = elegir_multiple("Módulos", modulos_opts, por_defecto="1,2")
    enable_telegram = 0 in modulos_sel
    enable_addin    = 1 in modulos_sel
    config["ENABLE_TELEGRAM"] = "true" if enable_telegram else "false"
    config["ENABLE_ADDIN"]    = "true" if enable_addin    else "false"

    ok(f"Bot Telegram: {'activo' if enable_telegram else 'desactivado'}")
    ok(f"Add-in Excel: {'activo' if enable_addin    else 'desactivado'}")

    # ── 2. Proveedor de IA ────────────────────────────────────────────────────
    titulo("2. ¿Qué proveedor de IA quieres usar?")
    ia_opts = [
        ("Groq",    "Gratuito · Rápido · Datos enviados a EE.UU. · Recomendado para empezar"),
        ("Ollama",  "Gratuito · Local · Sin datos fuera de la red · Requiere GPU o CPU potente"),
        ("OpenAI",  "De pago · GPT-4o · Datos en EE.UU."),
    ]
    ia_idx = elegir("Proveedor", ia_opts, por_defecto=1)
    proveedor = ["groq", "ollama", "openai"][ia_idx]
    config["LLM_PROVIDER"] = proveedor
    ok(f"Proveedor: {proveedor}")

    # ── 3. Configuración del proveedor ────────────────────────────────────────
    titulo(f"3. Configuración de {proveedor.upper()}")

    if proveedor == "groq":
        info("Obtén tu clave gratuita en: https://console.groq.com")
        config["GROQ_API_KEY"] = preguntar("GROQ_API_KEY")
        config["LLM_MODEL"]    = preguntar("Modelo", "llama-3.3-70b-versatile")

    elif proveedor == "ollama":
        info("Ollama debe estar instalado y corriendo en tu servidor.")
        info("Modelos recomendados: llama3.2, mistral, qwen2.5")
        config["OLLAMA_URL"] = preguntar("URL de Ollama", "http://localhost:11434")
        config["LLM_MODEL"]  = preguntar("Modelo", "llama3.2")

    elif proveedor == "openai":
        info("Obtén tu clave en: https://platform.openai.com/api-keys")
        config["OPENAI_API_KEY"] = preguntar("OPENAI_API_KEY")
        config["LLM_MODEL"]      = preguntar("Modelo", "gpt-4o-mini")
        azure = preguntar("¿Usar Azure OpenAI? (s/N)", "N")
        if azure.lower() == "s":
            config["OPENAI_BASE_URL"] = preguntar("Azure endpoint URL")

    # ── 4. Bot de Telegram ────────────────────────────────────────────────────
    if enable_telegram:
        titulo("4. Bot de Telegram")
        info("Obtén el token en Telegram buscando @BotFather → /newbot")
        config["TELEGRAM_TOKEN"] = preguntar("TELEGRAM_TOKEN")
        info("IDs numéricos de Telegram. Búscalos con @userinfobot")
        config["AUTHORIZED_USERS"] = preguntar("IDs autorizados (separados por coma)")

    # ── 5. Add-in de Excel ────────────────────────────────────────────────────
    if enable_addin:
        titulo("5. Add-in de Excel")
        api_key = secrets.token_hex(32)
        info(f"API Key generada automáticamente: {api_key[:16]}…")
        config["API_KEY"] = api_key

        info("Dominios de correo corporativo con acceso (ej: empresa.com,empresa.eu)")
        config["ALLOWED_DOMAINS"] = preguntar("ALLOWED_DOMAINS (vacío si no aplica)", "")
        info("Correos individuales adicionales con acceso")
        config["ALLOWED_EMAILS"] = preguntar("ALLOWED_EMAILS (vacío si no aplica)", "")

    # ── 6. Despliegue ─────────────────────────────────────────────────────────
    titulo("6. ¿Cómo vas a desplegar?")
    deploy_opts = [
        ("Local Windows",      "PC personal, bot corriendo mientras la ventana esté abierta"),
        ("Render (cloud)",     "Gratuito, 24/7, redespliega con cada git push"),
        ("Servidor empresa",   "Servidor Linux interno con cloudflared para el Add-in"),
    ]
    deploy_idx = elegir("Despliegue", deploy_opts, por_defecto=1)

    config["WEBHOOK_URL"] = ""
    config["ADDIN_URL"]   = ""

    if deploy_idx == 1:  # Render
        info("La URL de Render es: https://<nombre-servicio>.onrender.com")
        url = preguntar("URL de Render (vacío si aún no la tienes)", "")
        if url:
            url = url.rstrip("/")
            config["WEBHOOK_URL"] = url
            config["ADDIN_URL"]   = url + "/"
    elif deploy_idx == 2:  # Servidor empresa
        info("Una vez creado el túnel Cloudflare, introduce la URL aquí")
        url = preguntar("URL del túnel Cloudflare (vacío si aún no la tienes)", "")
        if url:
            config["ADDIN_URL"] = url.rstrip("/") + "/"

    # ── Escribir .env ─────────────────────────────────────────────────────────
    titulo("Generando .env…")

    lineas = [
        "# Generado por instalar.py — NO subir a GitHub\n",
        "\n# ── Módulos ──────────────────────────────────────────\n",
        f"ENABLE_TELEGRAM={config.get('ENABLE_TELEGRAM', 'true')}\n",
        f"ENABLE_ADDIN={config.get('ENABLE_ADDIN', 'true')}\n",
        "\n# ── Proveedor de IA ───────────────────────────────────\n",
        f"LLM_PROVIDER={config.get('LLM_PROVIDER', 'groq')}\n",
        f"LLM_MODEL={config.get('LLM_MODEL', '')}\n",
    ]

    if proveedor == "groq":
        lineas += [f"GROQ_API_KEY={config.get('GROQ_API_KEY', '')}\n"]
    elif proveedor == "ollama":
        lineas += [f"OLLAMA_URL={config.get('OLLAMA_URL', 'http://localhost:11434')}\n"]
    elif proveedor == "openai":
        lineas += [f"OPENAI_API_KEY={config.get('OPENAI_API_KEY', '')}\n"]
        if config.get("OPENAI_BASE_URL"):
            lineas += [f"OPENAI_BASE_URL={config['OPENAI_BASE_URL']}\n"]

    if enable_telegram:
        lineas += [
            "\n# ── Bot de Telegram ───────────────────────────────────\n",
            f"TELEGRAM_TOKEN={config.get('TELEGRAM_TOKEN', '')}\n",
            f"AUTHORIZED_USERS={config.get('AUTHORIZED_USERS', '')}\n",
        ]

    if enable_addin:
        lineas += [
            "\n# ── Add-in Excel ──────────────────────────────────────\n",
            f"API_KEY={config.get('API_KEY', '')}\n",
            f"ALLOWED_DOMAINS={config.get('ALLOWED_DOMAINS', '')}\n",
            f"ALLOWED_EMAILS={config.get('ALLOWED_EMAILS', '')}\n",
        ]

    lineas += [
        "\n# ── URLs de despliegue ────────────────────────────────\n",
        f"WEBHOOK_URL={config.get('WEBHOOK_URL', '')}\n",
        f"ADDIN_URL={config.get('ADDIN_URL', '')}\n",
    ]

    with open(".env", "w", encoding="utf-8") as f:
        f.writelines(lineas)

    ok(".env generado correctamente")

    # ── Resumen final ─────────────────────────────────────────────────────────
    print(f"\n{C.NEGRITA}{'='*55}")
    print("  ✅  Instalación completada")
    print(f"{'='*55}{C.RESET}\n")

    deploy_nombres = ["local Windows", "Render", "servidor empresa"]
    print(f"  Módulos:    {'Bot Telegram' if enable_telegram else ''}"
          f"{'  +  ' if enable_telegram and enable_addin else ''}"
          f"{'Add-in Excel' if enable_addin else ''}")
    print(f"  IA:         {proveedor}")
    print(f"  Despliegue: {deploy_nombres[deploy_idx]}\n")

    if deploy_idx == 0:
        print(f"  Siguiente paso: {C.VERDE}scripts\\arrancar_personal.bat{C.RESET}\n")
    elif deploy_idx == 1:
        if not config.get("WEBHOOK_URL"):
            warn("Recuerda añadir WEBHOOK_URL y ADDIN_URL en el dashboard de Render")
            warn("una vez tengas la URL del servicio.")
        print(f"  Siguiente paso: {C.VERDE}git push{C.RESET} → Render redespliega automáticamente\n")
    elif deploy_idx == 2:
        print(f"  Siguiente paso: {C.VERDE}scripts/instalar_empresa.sh{C.RESET}\n")
        if not config.get("ADDIN_URL"):
            warn("Recuerda añadir ADDIN_URL en .env una vez tengas la URL del túnel Cloudflare")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n  Instalación cancelada.\n")
        sys.exit(0)
