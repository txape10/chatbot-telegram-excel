/* global Office, Excel, fetch, document, localStorage, navigator */

import {
  cargarTema, aplicarTema, temasVisibles,
  TEMAS, desbloquearZelda, estaZeldaDesbloqueado,
} from "./themes.js";
import { estaAutorizado, obtenerEmailUsuario } from "./auth.js";

const API_URL = "";        // relativo: webpack hace de proxy hacia localhost:8000
const API_KEY = "test123";

// Estado de la edición pendiente de ubicar
let _datosModificados = null;
let _rangoAddress     = null;
let _rangoFilas       = 0;
let _rangoCols        = 0;

// Easter egg
let _eggInterval = null;

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
    } else {
      mostrarRespuesta(respuesta.respuesta);
      mostrarEstado("Rango consultado: " + direccion);
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

// Exponer al HTML
window.preguntar                 = preguntar;
window.copiarRespuesta           = copiarRespuesta;
window.escribirEnExcel           = escribirEnExcel;
window.toggleConfig              = toggleConfig;
window.mostrarEasterEgg          = mostrarEasterEgg;
window.cerrarEasterEgg           = cerrarEasterEgg;
window.activarZeldaDesdeEasterEgg = activarZeldaDesdeEasterEgg;
