/* global Office, __ALLOWED_DOMAINS__, __ALLOWED_EMAILS__ */

/**
 * Control de acceso del Add-in — modo whitelist.
 *
 * Los dominios y correos se leen de variables de entorno definidas en .env
 * e inyectadas en el bundle por webpack (DefinePlugin).
 * El fichero .env está en .gitignore y NUNCA se sube al repositorio.
 *
 * Variables necesarias en .env:
 *   ALLOWED_DOMAINS=launioncorp.eu,launioncorp.com
 *   ALLOWED_EMAILS=usuario@ejemplo.com,otro@ejemplo.com
 */

// Arrays inyectados por webpack DefinePlugin en tiempo de build
const DOMINIOS_PERMITIDOS = __ALLOWED_DOMAINS__;
const CORREOS_PERMITIDOS  = __ALLOWED_EMAILS__;

// ── API ───────────────────────────────────────────────────────────────────────

export function obtenerEmailUsuario() {
  try {
    return (Office?.context?.userProfile?.email || "").toLowerCase().trim();
  } catch {
    return "";
  }
}

export function obtenerNombreUsuario() {
  try {
    return Office?.context?.userProfile?.displayName || "";
  } catch {
    return "";
  }
}

/**
 * Devuelve true si el usuario tiene acceso.
 * Si Office no proporciona el email, se permite (no penalizar entornos sin cuenta).
 */
export function estaAutorizado() {
  const email = obtenerEmailUsuario();
  if (!email) return true;

  const dominio = email.split("@")[1] || "";

  return (
    DOMINIOS_PERMITIDOS.includes(dominio) ||
    CORREOS_PERMITIDOS.includes(email)
  );
}
