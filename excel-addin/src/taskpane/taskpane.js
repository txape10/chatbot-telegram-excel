/* global Office, Excel, fetch, document, localStorage, navigator */

import {
  cargarTema, aplicarTema, temasVisibles,
  TEMAS, desbloquearZelda, estaZeldaDesbloqueado,
} from "./themes.js";
import { estaAutorizado, obtenerEmailUsuario } from "./auth.js";

const API_URL = "";          // relativo: misma origin que el servidor FastAPI
/* global __API_KEY__ */
const API_KEY = __API_KEY__; // inyectada por webpack DefinePlugin desde .env — nunca en el repo

// Estado de la edición pendiente de ubicar
let _datosModificados = null;
let _rangoAddress     = null;
let _rangoFilas       = 0;
let _rangoCols        = 0;

// Easter egg
let _eggInterval = null;

// Historial local
const _CLAVE_HISTORIAL   = "asistente-excel-historial";
const _MAX_HISTORIAL     = 20;
let _historialAbierto    = true;

Office.onReady(() => {
  if (!estaAutorizado()) {
    document.getElementById("panel").style.display = "none";
    const panelNoAuth = document.getElementById("panel-no-autorizado");
    panelNoAuth.style.display = "flex";
    const email = obtenerEmailUsuario();
    if (email) {
      document.getElementById("no-auth-email").textContent = email;
    }
    return;
  }

  cargarTema();
  construirSelectorTemas();
  _inicializarBarraArchivo();
  _renderizarHistorial();
  cargarConfigAddin();

  document.getElementById("pregunta").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && e.ctrlKey) {
      e.preventDefault();
      preguntar();
    }
  });
});

// ── Temas ─────────────────────────────────────────────────────────────────────

function construirSelectorTemas() {
  const contenedor = document.getElementById("selector-temas");
  contenedor.innerHTML = "";
  temasVisibles().forEach((tema) => {
    const btn = document.createElement("button");
    btn.className      = "boton-tema";
    btn.dataset.temaId = tema.id;
    btn.textContent    = tema.nombre;
    btn.addEventListener("click", () => aplicarTema(tema));
    contenedor.appendChild(btn);
  });
}

function toggleConfig() {
  const panel = document.getElementById("panel-config");
  panel.style.display = panel.style.display === "none" ? "block" : "none";
}

// ── Preguntar ─────────────────────────────────────────────────────────────────

const _PALABRAS_ZELDA = ["zelda", "link"];

async function preguntar() {
  const instruccion = document.getElementById("pregunta").value.trim();
  if (!instruccion) {
    mostrarEstado("Escribe una pregunta primero.");
    return;
  }

  // Easter egg: detectar palabras mágicas
  if (!estaZeldaDesbloqueado() && _PALABRAS_ZELDA.includes(instruccion.toLowerCase())) {
    mostrarEasterEgg();
    return;
  }

  const boton = document.getElementById("btn-preguntar");
  boton.disabled = true;
  mostrarEstado("Leyendo seleccion...");
  ocultarRespuesta();
  ocultarDialogo();

  try {
    const { valores, direccion } = await leerRangoSeleccionado();

    if (valores.length < 2) {
      mostrarEstado("Selecciona al menos una fila de cabeceras y una de datos.");
      return;
    }

    _rangoAddress = direccion;
    _rangoFilas   = valores.length;
    _rangoCols    = valores[0].length;

    mostrarEstado("Consultando al asistente... (rango: " + direccion + ")");

    const respuesta = await llamarApi("/edit", { datos: valores, instruccion });

    if (respuesta.tipo === "edicion") {
      _datosModificados = respuesta.datos_modificados;
      mostrarDialogo(respuesta.descripcion);
      mostrarEstado("Rango leido: " + direccion);
      _agregarAlHistorial(instruccion, "✏️ " + respuesta.descripcion);
    } else {
      mostrarRespuesta(respuesta.respuesta);
      mostrarEstado("Rango consultado: " + direccion);
      _agregarAlHistorial(instruccion, respuesta.respuesta);
    }

  } catch (error) {
    mostrarEstado("Error: " + error.message);
  } finally {
    boton.disabled = false;
  }
}

async function leerRangoSeleccionado() {
  return Excel.run(async (context) => {
    const rango = context.workbook.getSelectedRange();
    rango.load(["values", "address"]);
    await context.sync();
    return { valores: rango.values, direccion: rango.address };
  });
}

async function llamarApiGet(endpoint) {
  const respuesta = await fetch(API_URL + endpoint, {
    method: "GET",
    headers: { "X-API-Key": API_KEY },
  });
  if (!respuesta.ok) {
    const error = await respuesta.json().catch(() => ({}));
    throw new Error(error.detail || "HTTP " + respuesta.status);
  }
  return respuesta.json();
}

async function llamarApi(endpoint, payload) {
  const respuesta = await fetch(API_URL + endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-API-Key": API_KEY,
    },
    body: JSON.stringify(payload),
  });

  if (!respuesta.ok) {
    const error = await respuesta.json().catch(() => ({}));
    throw new Error(error.detail || "HTTP " + respuesta.status);
  }

  return respuesta.json();
}

// ── Escritura en Excel ────────────────────────────────────────────────────────

async function escribirEnExcel(destino) {
  ocultarDialogo();
  mostrarEstado("Escribiendo en Excel...");

  const datos = _datosModificados;
  const filas = datos.length;
  const cols  = datos[0].length;

  try {
    await Excel.run(async (context) => {
      let targetRange;

      if (destino === "nueva_hoja") {
        const nuevaHoja = context.workbook.worksheets.add();
        nuevaHoja.activate();
        targetRange = nuevaHoja.getRangeByIndexes(0, 0, filas, cols);
      } else {
        const sheet    = context.workbook.worksheets.getActiveWorksheet();
        const refRange = sheet.getRange(_rangoAddress);
        refRange.load(["rowIndex", "columnIndex"]);
        await context.sync();

        const anchorRow = refRange.rowIndex;
        const anchorCol = refRange.columnIndex;

        if (destino === "sustituir") {
          targetRange = sheet.getRangeByIndexes(anchorRow, anchorCol, filas, cols);
        } else if (destino === "derecha") {
          targetRange = sheet.getRangeByIndexes(anchorRow, anchorCol + _rangoCols, filas, cols);
        } else if (destino === "debajo") {
          targetRange = sheet.getRangeByIndexes(anchorRow + _rangoFilas, anchorCol, filas, cols);
        }
      }

      targetRange.values = datos;
      await context.sync();
    });

    mostrarEstado("Datos escritos correctamente.");
    _datosModificados = null;

  } catch (error) {
    mostrarEstado("Error al escribir: " + error.message);
  }
}

// ── Easter egg Zelda ──────────────────────────────────────────────────────────

function mostrarEasterEgg() {
  document.getElementById("easter-egg").style.display = "flex";
  _lanzarFuegosArtificiales();
  _tocarJingleZelda();
}

// Jingle "secreto encontrado" de Zelda — sintetizado con Web Audio API
function _tocarJingleZelda() {
  try {
    const AudioContext = window.AudioContext || window.webkitAudioContext;
    if (!AudioContext) return;

    const ctx = new AudioContext();

    // Las 4 notas ascendentes del jingle clásico (onda cuadrada = sonido NES)
    // B4 → F#5 → A#5 → B5 (la última sostenida)
    const notas = [
      { freq: 493.88, t: 0.00, dur: 0.10 },   // B4
      { freq: 739.99, t: 0.10, dur: 0.10 },   // F#5
      { freq: 932.33, t: 0.20, dur: 0.10 },   // A#5
      { freq: 987.77, t: 0.30, dur: 0.65 },   // B5 — sostenida
    ];

    notas.forEach(({ freq, t, dur }) => {
      const osc  = ctx.createOscillator();
      const gain = ctx.createGain();

      osc.connect(gain);
      gain.connect(ctx.destination);

      osc.type = "square";
      osc.frequency.setValueAtTime(freq, ctx.currentTime + t);

      // Ataque rápido, caída suave al final de cada nota
      const vol = 0.18;
      gain.gain.setValueAtTime(0, ctx.currentTime + t);
      gain.gain.linearRampToValueAtTime(vol, ctx.currentTime + t + 0.008);
      gain.gain.setValueAtTime(vol, ctx.currentTime + t + dur - 0.04);
      gain.gain.linearRampToValueAtTime(0, ctx.currentTime + t + dur);

      osc.start(ctx.currentTime + t);
      osc.stop(ctx.currentTime + t + dur + 0.01);
    });

  } catch (_e) {
    // Sin soporte de audio — el Easter egg sigue funcionando en silencio
  }
}

function cerrarEasterEgg() {
  document.getElementById("easter-egg").style.display = "none";
  _detenerFuegos();
}

function activarZeldaDesdeEasterEgg() {
  desbloquearZelda();          // solo se guarda si el usuario confirma
  cerrarEasterEgg();
  const zelda = TEMAS.find((t) => t.id === "zelda");
  if (zelda) {
    aplicarTema(zelda);
    construirSelectorTemas();  // mostrar el tema en el selector
  }
}

// Fuegos artificiales ─────────────────────────────────────────────────────────

const _COLORES_EGG = [
  "#C8A84B", "#FFD700", "#1B4D1B", "#FFFEF0",
  "#FF6B35", "#00CED1", "#FF1493", "#ADFF2F",
];

function _crearRafaga(contenedor) {
  const w = contenedor.offsetWidth  || 320;
  const h = contenedor.offsetHeight || 400;

  const cx = 20 + Math.random() * (w - 40);
  const cy = 20 + Math.random() * (h * 0.65);
  const n  = 14 + Math.floor(Math.random() * 8);

  for (let i = 0; i < n; i++) {
    const p      = document.createElement("div");
    const angulo = (i / n) * Math.PI * 2;
    const dist   = 50 + Math.random() * 90;
    const vx     = Math.cos(angulo) * dist;
    const vy     = Math.sin(angulo) * dist;
    const dur    = 700 + Math.random() * 600;
    const color  = _COLORES_EGG[Math.floor(Math.random() * _COLORES_EGG.length)];
    const tam    = 5 + Math.random() * 5;

    p.className = "ee-particula";
    p.style.cssText = [
      `left:${cx}px`, `top:${cy}px`,
      `width:${tam}px`, `height:${tam}px`,
      `background:${color}`,
      `--vx:${vx}px`, `--vy:${vy}px`,
      `animation-duration:${dur}ms`,
    ].join(";");

    contenedor.appendChild(p);
    setTimeout(() => p.remove(), dur + 50);
  }
}

function _lanzarFuegosArtificiales() {
  const contenedor = document.getElementById("ee-fuegos");
  if (!contenedor) return;

  // Ráfaga inicial inmediata
  _crearRafaga(contenedor);
  _crearRafaga(contenedor);

  _eggInterval = setInterval(() => {
    const n = 1 + Math.floor(Math.random() * 2);
    for (let i = 0; i < n; i++) {
      setTimeout(() => _crearRafaga(contenedor), i * 180);
    }
  }, 600);
}

function _detenerFuegos() {
  if (_eggInterval) {
    clearInterval(_eggInterval);
    _eggInterval = null;
  }
  const contenedor = document.getElementById("ee-fuegos");
  if (contenedor) contenedor.innerHTML = "";
}

// ── UI helpers ────────────────────────────────────────────────────────────────

function mostrarEstado(texto) {
  document.getElementById("estado").textContent = texto;
}

function mostrarRespuesta(texto) {
  /* global marked */
  document.getElementById("respuesta").innerHTML = marked.parse(texto);
  document.getElementById("bloque-respuesta").style.display = "block";
}

function ocultarRespuesta() {
  document.getElementById("bloque-respuesta").style.display = "none";
}

function mostrarDialogo(descripcion) {
  document.getElementById("desc-edicion").textContent = "Resultado: " + descripcion;
  document.getElementById("dialogo-destino").style.display = "block";
}

function ocultarDialogo() {
  document.getElementById("dialogo-destino").style.display = "none";
}

async function copiarRespuesta() {
  const texto = document.getElementById("respuesta").innerText;
  await navigator.clipboard.writeText(texto);
  const label = document.querySelector(".boton-copiar .ms-Button-label");
  label.textContent = "Copiado";
  setTimeout(() => { label.textContent = "Copiar respuesta"; }, 2000);
}

// ── Barra de archivo activo ───────────────────────────────────────────────────

async function _inicializarBarraArchivo() {
  await _actualizarArchivoActivo();

  // Actualizar rango en tiempo real al cambiar la selección
  try {
    await Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(_onSelectionChanged);
      await context.sync();
    });
  } catch (_e) {
    // Sin soporte de eventos — la barra muestra el estado inicial solamente
  }
}

async function _actualizarArchivoActivo() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet    = workbook.worksheets.getActiveWorksheet();
      const rango    = workbook.getSelectedRange();
      workbook.load("name");
      sheet.load("name");
      rango.load("address");
      await context.sync();

      const libro = workbook.name || "Libro";
      const hoja  = sheet.name   || "Hoja";
      const dir   = rango.address
        ? rango.address.replace(/^[^!]+!/, "")   // quitar prefijo de hoja
        : "—";

      document.getElementById("archivo-nombre").textContent = libro;
      document.getElementById("archivo-hoja").textContent   = hoja;
      document.getElementById("archivo-rango").textContent  = dir;
    });
  } catch (_e) {
    // Excel no disponible en este contexto (ej. preview)
  }
}

async function _onSelectionChanged() {
  try {
    await Excel.run(async (context) => {
      const rango = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      rango.load("address");
      sheet.load("name");
      await context.sync();

      const dir  = rango.address
        ? rango.address.replace(/^[^!]+!/, "")
        : "—";
      document.getElementById("archivo-rango").textContent = dir;
      document.getElementById("archivo-hoja").textContent  = sheet.name || "Hoja";
    });
  } catch (_e) { /* silencioso */ }
}

// ── Historial local ───────────────────────────────────────────────────────────

function _cargarHistorial() {
  try {
    return JSON.parse(localStorage.getItem(_CLAVE_HISTORIAL) || "[]");
  } catch (_e) {
    return [];
  }
}

function _guardarHistorial(entradas) {
  localStorage.setItem(_CLAVE_HISTORIAL, JSON.stringify(entradas.slice(-_MAX_HISTORIAL)));
}

function _agregarAlHistorial(pregunta, respuesta) {
  const entradas = _cargarHistorial();
  entradas.push({ pregunta, respuesta, ts: Date.now() });
  _guardarHistorial(entradas);
  _renderizarHistorial();
}

function _renderizarHistorial() {
  const entradas  = _cargarHistorial();
  const lista     = document.getElementById("historial-lista");
  const count     = document.getElementById("historial-count");
  const chevron   = document.getElementById("historial-chevron");

  count.textContent = entradas.length ? `(${entradas.length})` : "";
  chevron.textContent = _historialAbierto ? "▲" : "▼";
  lista.style.display = _historialAbierto ? "flex" : "none";

  lista.innerHTML = "";

  if (entradas.length === 0) {
    const vacio = document.createElement("div");
    vacio.className = "historial-vacio";
    vacio.textContent = "Aún no hay preguntas";
    lista.appendChild(vacio);
    return;
  }

  // Mostrar las últimas 10 entradas en orden cronológico
  const visibles = entradas.slice(-10);
  visibles.forEach(({ pregunta, respuesta }) => {
    const entrada = document.createElement("div");
    entrada.className = "historial-entrada";

    // Respuesta: quitar Markdown básico para lectura limpia
    const textoBot = respuesta
      .replace(/\*\*(.*?)\*\*/g, "$1")
      .replace(/\*(.*?)\*/g, "$1")
      .replace(/`(.*?)`/g, "$1")
      .replace(/#{1,3} /g, "")
      .substring(0, 200) + (respuesta.length > 200 ? "…" : "");

    entrada.innerHTML = `
      <span class="historial-usuario">Tú</span>
      <span class="historial-texto-usuario">${_escaparHtml(pregunta)}</span>
      <span class="historial-bot">Asistente</span>
      <span class="historial-texto-bot">${_escaparHtml(textoBot)}</span>
    `;
    lista.appendChild(entrada);
  });

  // Scroll al final automáticamente
  lista.scrollTop = lista.scrollHeight;
}

function _escaparHtml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function toggleHistorial() {
  _historialAbierto = !_historialAbierto;
  _renderizarHistorial();
}

function limpiarHistorial(event) {
  event.stopPropagation();   // no colapsar el panel al limpiar
  localStorage.removeItem(_CLAVE_HISTORIAL);
  _renderizarHistorial();
}

// ── Configuración del servidor ────────────────────────────────────────────────

/**
 * Consulta /addin-config al arrancar para:
 *  - Ocultar permanentemente la sección de Telegram si ENABLE_TELEGRAM=false
 *  - Aplicar el nombre de empresa configurado en el servidor (sin recompilar)
 */
async function cargarConfigAddin() {
  try {
    const config = await llamarApiGet("/addin-config");

    // ── Nombre de empresa ────────────────────────────────────────────────────
    if (config.nombre_empresa) {
      // Sobreescribir el nombre del tema Empresa con el valor del servidor
      const temaEmpresa = TEMAS.find((t) => t.id === "empresa");
      if (temaEmpresa) {
        temaEmpresa.nombre    = config.nombre_empresa;
        temaEmpresa.subtitulo = config.nombre_empresa;
        construirSelectorTemas();
        // Si el tema activo es empresa, re-aplicarlo para actualizar el subtítulo
        const activo = localStorage.getItem("asistente-excel-tema");
        if (activo === "empresa") {
          aplicarTema(temaEmpresa);
        }
      }
    }

    // ── Módulo Telegram ──────────────────────────────────────────────────────
    if (!config.telegram_habilitado) {
      // Modo empresa sin bot: ocultar permanentemente toda la sección Telegram
      document.getElementById("btn-enviar-bot").style.display = "none";
      document.getElementById("info-telegram").style.display  = "none";
      return;  // no llamar a comprobarVinculo
    }

    await comprobarVinculo();
  } catch {
    // Error de red o servidor antiguo sin este endpoint: flujo normal
    await comprobarVinculo();
  }
}

// ── Vínculo Telegram ──────────────────────────────────────────────────────────

const _CLAVE_EMAIL_VINCULO = "asistente-excel-email-vinculo";

/**
 * Devuelve el email efectivo para comprobar el vínculo Telegram.
 * Orden de prioridad:
 *   1. Office.context.userProfile.email  (disponible en Outlook, no en Excel)
 *   2. localStorage (guardado manualmente por el usuario)
 *   3. "" — sin email
 */
function obtenerEmailEfectivo() {
  const officeEmail = obtenerEmailUsuario();
  if (officeEmail) return officeEmail;
  return localStorage.getItem(_CLAVE_EMAIL_VINCULO) || "";
}

/** Guarda el email introducido manualmente y recomprueba el vínculo. */
async function guardarEmailVinculo() {
  const input = document.getElementById("input-email-vinculo");
  const email = (input?.value || "").trim().toLowerCase();
  if (!email) return;
  localStorage.setItem(_CLAVE_EMAIL_VINCULO, email);
  await comprobarVinculo();
}

async function comprobarVinculo() {
  const email  = obtenerEmailEfectivo();
  const boton  = document.getElementById("btn-enviar-bot");
  const info   = document.getElementById("info-telegram");
  const bloque = document.getElementById("bloque-email-vinculo");

  if (!email) {
    // Sin email: mostrar info + campo para introducirlo manualmente
    if (info)   info.style.display   = "";
    if (bloque) bloque.style.display = "";
    return;
  }
  // Tenemos email — ocultar el bloque de entrada manual si ya no hace falta
  if (bloque) bloque.style.display = "none";

  try {
    const resultado = await llamarApiGet(
      "/tiene-vinculo?email=" + encodeURIComponent(email)
    );
    if (resultado.vinculado) {
      boton.style.display = "";
      info.style.display  = "none";
    } else {
      boton.style.display = "none";
      info.style.display  = "";
    }
  } catch {
    // Error de red: mostrar el botón (mejor falso positivo que ocultar la función)
    boton.style.display = "";
    info.style.display  = "none";
  }
}

// ── Enviar al bot ─────────────────────────────────────────────────────────────

async function enviarAlBot() {
  const email = obtenerEmailEfectivo();
  if (!email) {
    mostrarEstado("Introduce tu email primero en el campo de vinculación.");
    return;
  }

  const boton = document.getElementById("btn-enviar-bot");
  boton.disabled = true;
  mostrarEstado("Leyendo selección...");

  try {
    const { valores, direccion } = await leerRangoSeleccionado();

    if (valores.length < 2) {
      mostrarEstado("Selecciona al menos una fila de cabeceras y una de datos.");
      return;
    }

    // Intentar obtener el nombre del libro
    let nombreArchivo = "datos_excel.xlsx";
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();
        if (workbook.name) {
          const base = workbook.name.replace(/\.xlsx?$/i, "");
          nombreArchivo = base + ".xlsx";
        }
      });
    } catch (_e) { /* usar nombre por defecto */ }

    mostrarEstado(`Enviando ${valores.length - 1} filas al bot...`);

    const resultado = await llamarApi("/enviar-al-bot", {
      datos:          valores,
      nombre_archivo: nombreArchivo,
      email:          email,
    });

    mostrarEstado("✅ " + (resultado.mensaje || "Archivo enviado a Telegram."));

  } catch (error) {
    const msg = error.message || "Error desconocido";
    if (msg.includes("vinculada") || msg.includes("vincular")) {
      mostrarEstado(
        "⚠️ Cuenta no vinculada. Escribe /vincular " + email + " en el bot de Telegram primero."
      );
    } else {
      mostrarEstado("Error al enviar: " + msg);
    }
  } finally {
    boton.disabled = false;
  }
}

// Exponer al HTML
window.preguntar                  = preguntar;
window.copiarRespuesta            = copiarRespuesta;
window.escribirEnExcel            = escribirEnExcel;
window.toggleConfig               = toggleConfig;
window.toggleHistorial            = toggleHistorial;
window.limpiarHistorial           = limpiarHistorial;
window.mostrarEasterEgg           = mostrarEasterEgg;
window.cerrarEasterEgg            = cerrarEasterEgg;
window.activarZeldaDesdeEasterEgg = activarZeldaDesdeEasterEgg;
window.enviarAlBot                 = enviarAlBot;
