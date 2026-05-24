/* global Office */

/**
 * Autenticación SSO con Office 365 / Azure AD.
 *
 * Requiere:
 *   1. App registrada en Azure AD (lo hace el administrador)
 *   2. CLIENT_ID en .env (el Application ID de Azure)
 *   3. Permiso "User.Read" concedido en Azure AD
 *
 * Misma interfaz que auth.whitelist.js — taskpane.js no cambia.
 *
 * Pendiente de implementar cuando el administrador registre la app en Azure.
 */

// ── Configuración ─────────────────────────────────────────────────────────────

// Dominios / correos autorizados (igual que en whitelist, como segundo filtro
// por si el token SSO se obtiene pero queremos restringir aún más)
const DOMINIOS_PERMITIDOS = [
  "launioncorp.eu",
  "launioncorp.com",
];

const CORREOS_PERMITIDOS = [
  "roberto.chapado@gmail.com",   // desarrollo / testing
];

// ── Estado interno ────────────────────────────────────────────────────────────

let _emailCache = null;

// ── API pública ───────────────────────────────────────────────────────────────

export function obtenerEmailUsuario() {
  if (_emailCache !== null) return _emailCache;
  try {
    _emailCache = (Office?.context?.userProfile?.email || "").toLowerCase().trim();
    return _emailCache;
  } catch {
    _emailCache = "";
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
 * Verificación SSO.
 *
 * Flujo completo (pendiente de implementar con el admin):
 *   1. Office.auth.getAccessToken() → token de Bootstrap
 *   2. Intercambio On-Behalf-Of en el backend → token de Microsoft Graph
 *   3. GET https://graph.microsoft.com/v1.0/me → email verificado
 *   4. Comprobar email contra DOMINIOS/CORREOS_PERMITIDOS
 *
 * Por ahora hace fallback a la misma lógica que auth.whitelist.js
 * para no bloquear el desarrollo hasta tener el admin disponible.
 */
export function estaAutorizado() {
  const email = obtenerEmailUsuario();

  // Sin email disponible → permitir (no penalizar entornos sin cuenta)
  if (!email) return true;

  const dominio = email.split("@")[1] || "";

  return (
    DOMINIOS_PERMITIDOS.includes(dominio) ||
    CORREOS_PERMITIDOS.includes(email)
  );

  /*
   * TODO — implementar cuando el admin registre la app en Azure AD:
   *
   * const token = await Office.auth.getAccessToken({ allowSignInPrompt: true });
   * const resp  = await fetch("/api/auth/verify", {
   *   headers: { Authorization: "Bearer " + token }
   * });
   * const { email, autorizado } = await resp.json();
   * return autorizado;
   */
}
