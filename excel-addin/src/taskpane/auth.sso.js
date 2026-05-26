/* global Office, __ALLOWED_DOMAINS__, __ALLOWED_EMAILS__ */

/**
 * Autenticación SSO con Office 365 / Azure AD.
 *
 * Requiere (para el flujo completo con token verificado):
 *   1. App registrada en Azure AD (lo hace el administrador)
 *   2. CLIENT_ID en excel-addin/.env (Application ID de Azure)
 *   3. Permiso "User.Read" concedido en Azure AD
 *
 * Misma interfaz que auth.whitelist.js — taskpane.js no cambia.
 *
 * Mientras no esté registrada la app en Azure AD, usa whitelist de
 * dominios/correos inyectada por webpack (mismo comportamiento que whitelist.js).
 */

// ── Configuración — inyectada por webpack DefinePlugin en tiempo de build ──────
// Mismas variables que auth.whitelist.js; NUNCA hardcodear aquí.
const DOMINIOS_PERMITIDOS = __ALLOWED_DOMAINS__;
const CORREOS_PERMITIDOS  = __ALLOWED_EMAILS__;

// ── API pública ───────────────────────────────────────────────────────────────

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
