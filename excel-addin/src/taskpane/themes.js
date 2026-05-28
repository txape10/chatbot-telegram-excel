/* global localStorage, document, Office, __COMPANY_NAME__ */

// Nombre de empresa inyectado por webpack DefinePlugin en tiempo de build.
// En desarrollo (sin COMPANY_NAME en .env) usa "Empresa" como valor por defecto.
const _NOMBRE_EMPRESA = typeof __COMPANY_NAME__ !== "undefined" ? __COMPANY_NAME__ : "Empresa";

/**
 * Definición de temas del Add-in.
 *
 * activo: true  → visible en el selector
 * activo: false → oculto (puede desbloquearse como Easter egg)
 *
 * Propiedades:
 *   id        — identificador único (guardado en localStorage)
 *   nombre    — texto en el selector
 *   activo    — visibilidad por defecto
 *   logo      — ruta a assets/ o null
 *   logoTexto — emoji/texto alternativo cuando no hay imagen
 *   subtitulo — texto pequeño bajo el título (opcional)
 *   vars      — variables CSS que se aplican al :root
 */
export const TEMAS = [
  {
    id: "default",
    nombre: "Excel",
    activo: true,
    logo: null,
    logoTexto: "📊",
    vars: {
      "--color-primario":          "#217346",
      "--color-acento":            "#107C41",
      "--color-fondo":             "#FFFFFF",
      "--color-fondo-panel":       "#F2F2F2",
      "--color-texto":             "#212121",
      "--color-texto-suave":       "#737373",
      "--color-borde":             "#C8C8C8",
      "--color-btn-texto":         "#FFFFFF",
      "--color-btn-copiar-texto":  "#FFFFFF",
    },
  },
  {
    id: "zelda",
    nombre: "Zelda",
    activo: false,   // oculto — se desbloquea con el Easter egg
    logo: "assets/logo-theme/logo_zelda.png",
    vars: {
      "--color-primario":          "#1B4D1B",
      "--color-acento":            "#C8A84B",
      "--color-fondo":             "#FFFEF0",
      "--color-fondo-panel":       "#F5E6C8",
      "--color-texto":             "#2C1A0E",
      "--color-texto-suave":       "#6B4F2A",
      "--color-borde":             "#C8A84B",
      "--color-btn-texto":         "#FFFEF0",
      "--color-btn-copiar-texto":  "#2C1A0E",
    },
  },
  {
    id: "empresa",
    nombre: _NOMBRE_EMPRESA,
    activo: true,
    logo: "assets/logo-theme/logo_union.png",
    subtitulo: _NOMBRE_EMPRESA,
    vars: {
      "--color-primario":          "#D42B2B",
      "--color-acento":            "#F4941D",
      "--color-fondo":             "#FFFFFF",
      "--color-fondo-panel":       "#F7F7F7",
      "--color-texto":             "#1A1A1A",
      "--color-texto-suave":       "#666666",
      "--color-borde":             "#D42B2B",
      "--color-btn-texto":         "#FFFFFF",
      "--color-btn-copiar-texto":  "#1A1A1A",
    },
  },
];

const CLAVE_STORAGE        = "asistente-excel-tema";
const CLAVE_ZELDA          = "asistente-excel-zelda-desbloqueado";
const TEMA_DEFECTO         = "default";
const DOMINIOS_EMPRESA     = ["empresa.eu", "empresa.com"];

// ── Easter egg ────────────────────────────────────────────────────────────────

export function estaZeldaDesbloqueado() {
  return localStorage.getItem(CLAVE_ZELDA) === "1";
}

export function desbloquearZelda() {
  localStorage.setItem(CLAVE_ZELDA, "1");
}

// ── Detección de dominio ──────────────────────────────────────────────────────

function _obtenerDominio() {
  try {
    const email = (Office?.context?.userProfile?.email || "").toLowerCase();
    const partes = email.split("@");
    return partes.length > 1 ? partes[1] : "";
  } catch {
    return "";
  }
}

function _temaDefectoPorDominio() {
  const dominio = _obtenerDominio();
  return DOMINIOS_EMPRESA.includes(dominio) ? "empresa" : TEMA_DEFECTO;
}

// ── API pública ───────────────────────────────────────────────────────────────

export function cargarTema() {
  const guardado = localStorage.getItem(CLAVE_STORAGE);
  const id = guardado || _temaDefectoPorDominio();
  const tema = TEMAS.find((t) => t.id === id) || TEMAS[0];
  aplicarTema(tema);
  return tema;
}

export function aplicarTema(tema) {
  const raiz = document.documentElement;
  Object.entries(tema.vars).forEach(([prop, valor]) => {
    raiz.style.setProperty(prop, valor);
  });

  // Logo imagen vs texto
  const logoImg  = document.getElementById("logo-tema");
  const logoTxt  = document.getElementById("logo-texto");
  if (tema.logo) {
    logoImg.src          = tema.logo;
    logoImg.style.display = "block";
    if (logoTxt) logoTxt.style.display = "none";
  } else if (tema.logoTexto) {
    logoImg.style.display = "none";
    if (logoTxt) {
      logoTxt.textContent  = tema.logoTexto;
      logoTxt.style.display = "inline";
    }
  } else {
    logoImg.style.display = "none";
    if (logoTxt) logoTxt.style.display = "none";
  }

  // Subtítulo
  const subtituloEl = document.getElementById("subtitulo-tema");
  if (subtituloEl) {
    subtituloEl.textContent  = tema.subtitulo || "";
    subtituloEl.style.display = tema.subtitulo ? "block" : "none";
  }

  localStorage.setItem(CLAVE_STORAGE, tema.id);

  // Sincronizar select de configuración
  const sel = document.getElementById("select-tema");
  if (sel) sel.value = tema.id;
}

export function aplicarTemaId(id) {
  const tema = TEMAS.find((t) => t.id === id);
  if (tema) aplicarTema(tema);
}

export function temasVisibles() {
  return TEMAS.filter((t) => t.activo || (t.id === "zelda" && estaZeldaDesbloqueado()));
}
