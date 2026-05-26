/* global Office, Excel, fetch, document, localStorage, navigator */

import {
  cargarTema, aplicarTema, aplicarTemaId, temasVisibles,
  TEMAS, desbloquearZelda, estaZeldaDesbloqueado,
} from "./themes.js";
import { estaAutorizado, obtenerEmailUsuario } from "./auth.js";

const API_URL = "";          // relativo: misma origin que el servidor FastAPI
/* global __API_KEY__ */
const API_KEY = __API_KEY__; // inyectada por webpack DefinePlugin desde .env — nunca en el repo

// Estado de la edición pendiente de ubicar
let _datosModificados = null;
let _rangoAddress     = null;
let _rangoHojaNombre  = null;
let _rangoFilas       = 0;
let _rangoCols        = 0;

// Feedback RAG: última pregunta+respuesta de texto para el botón 👍
let _feedbackPregunta  = "";
let _feedbackRespuesta = "";

// Easter egg
let _eggInterval = null;

// Historial local (UI)
const _CLAVE_HISTORIAL   = "asistente-excel-historial";
const _MAX_HISTORIAL     = 20;
let _historialAbierto    = true;

// Historial de conversación enviado al LLM (contexto multi-turno)
let _historialLLM        = [];
const _MAX_TURNOS_LLM    = 6; // 3 turnos (user+model) para no saturar el contexto

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

  // Ping para despertar Render antes de la primera pregunta (evita 502 en cold start)
  fetch("/health").catch(() => {});

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
  const select = document.getElementById("select-tema");
  if (!select) return;
  const temaActualId = localStorage.getItem("asistente-excel-tema") || "default";
  select.innerHTML = "";
  temasVisibles().forEach((tema) => {
    const opt = document.createElement("option");
    opt.value       = tema.id;
    opt.textContent = tema.nombre;
    if (tema.id === temaActualId) opt.selected = true;
    select.appendChild(opt);
  });
}

function aplicarTemaDesdeSelect(id) {
  aplicarTemaId(id);
}

function toggleConfig() {
  const panel = document.getElementById("panel-config");
  panel.style.display = panel.style.display === "none" ? "block" : "none";
}

function toggleTelegram() {
  const body    = document.getElementById("config-telegram-body");
  const chevron = document.getElementById("acordeon-chevron");
  const abierto = body.style.display !== "none";
  body.style.display    = abierto ? "none" : "block";
  chevron.textContent   = abierto ? "▶" : "▼";
}

// ── Formato condicional ───────────────────────────────────────────────────────

const _FORMATO_KEYWORDS = [
  "formato condicional", "colorea", "colorear", "resalta", "resaltar",
  "escala de color", "barra de datos", "iconos de", "conjunto de iconos",
  "semáforo", "semaforo", "flechas", "banderas", "estrellas", "clasificacion",
  "top 10", "top 5", "top 3", "10 superiores", "5 superiores", "inferiores",
  "superiores", "heatmap", "gradiente", "pinta", "colorar", "marca en rojo",
  "marca en verde", "marca en amarillo", "destaca", "destacar",
  "conditional format",
];

const _COLORES_CF = {
  rojo: "#FF0000", verde: "#00B050", amarillo: "#FFFF00",
  naranja: "#FF8C00", azul: "#4472C4", morado: "#7030A0",
  rosa: "#FF69B4", celeste: "#87CEEB", gris: "#808080",
  blanco: "#FFFFFF", negro: "#000000", dorado: "#FFD700",
};

const _OP_CF = {
  ">":     "greaterThan",
  "<":     "lessThan",
  ">=":    "greaterThanOrEqual",
  "<=":    "lessThanOrEqual",
  "==":    "equalTo",
  "!=":    "notEqualTo",
  "entre": "between",
  "fuera": "notBetween",
};

const _OP_TEXTO_CF = {
  contiene:    "contains",
  no_contiene: "notContains",
  empieza_por: "beginsWith",
  termina_en:  "endsWith",
};

const _ICONOS_CF = {
  flechas:       "threeArrows",
  semaforo:      "threeTrafficLights1",
  banderas:      "threeFlags",
  formas:        "threeSymbols",
  estrellas:     "fiveRating",
  clasificacion: "fourRating",
};

function _colorCf(nombre) {
  return _COLORES_CF[nombre] || nombre;
}

function _esFormatoCondicional(texto) {
  const t = texto.toLowerCase();
  return _FORMATO_KEYWORDS.some((kw) => t.includes(kw));
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
  ocultarRespuesta();
  ocultarDialogo();
  _ocultarFeedback();
  mostrarEstado("Leyendo selección...");

  try {
    const { valores, direccion, hojaNombre } = await leerRangoSeleccionado();
    const tienesDatos = valores && valores.length >= 2;

    if (tienesDatos) {
      // ── Flujo con datos: edición / consulta sobre la tabla seleccionada ──
      _rangoAddress    = direccion;
      _rangoHojaNombre = hojaNombre;
      _rangoFilas      = valores.length;
      _rangoCols       = (valores[0] || []).length;
      mostrarEstado("Consultando al asistente... (rango: " + direccion + ")");

      if (_esFormatoCondicional(instruccion)) {
        // ── Flujo de formato condicional ──────────────────────────────────────
        mostrarEstado("Aplicando formato condicional...");
        const respuesta = await llamarApi("/format", {
          datos: valores, instruccion,
          device_id: _obtenerOCrearDeviceId(),
          user_email: obtenerEmailUsuario(),
          excel_version: _obtenerExcelVersion(),
        });

        if (respuesta.tipo === "formato" && respuesta.regla) {
          await _aplicarFormatoCondicional(respuesta.regla);
          mostrarRespuesta("🎨 " + respuesta.descripcion);
          mostrarEstado("Formato aplicado · " + direccion);
          _agregarAlHistorial(instruccion, "🎨 " + respuesta.descripcion);
          _actualizarHistorialLLM(instruccion, respuesta.descripcion);
        } else {
          const msg = (respuesta && respuesta.respuesta) || "No se pudo interpretar la regla de formato.";
          mostrarRespuesta(msg);
          mostrarEstado("Listo · " + direccion);
          _agregarAlHistorial(instruccion, msg);
          _actualizarHistorialLLM(instruccion, msg);
        }
      } else {
        // ── Flujo de edición ──────────────────────────────────────────────────
        const respuesta = await llamarApi("/edit", {
          datos: valores, instruccion, historial: _historialLLM,
          device_id: _obtenerOCrearDeviceId(),
          user_email: obtenerEmailUsuario(),
          excel_version: _obtenerExcelVersion(),
        });

        if (respuesta.tipo === "edicion") {
          _datosModificados = respuesta.datos_modificados;
          mostrarRespuesta("✏️ " + respuesta.descripcion + "\n\n*Elige dónde escribir el resultado:*");
          mostrarDialogo(respuesta.descripcion);
          mostrarEstado("Edición lista · " + direccion);
          _agregarAlHistorial(instruccion, "✏️ " + respuesta.descripcion);
          _actualizarHistorialLLM(instruccion, respuesta.descripcion);
        } else {
          mostrarRespuesta(respuesta.respuesta);
          mostrarEstado("Listo · " + direccion);
          _agregarAlHistorial(instruccion, respuesta.respuesta);
          _actualizarHistorialLLM(instruccion, respuesta.respuesta);
          _mostrarFeedback(instruccion, respuesta.respuesta);
        }
      }

    } else {
      // ── Flujo sin datos: pregunta general o creación desde cero ──
      // Guardamos la dirección actual como ancla por si el LLM devuelve datos para escribir
      _rangoAddress    = direccion;
      _rangoHojaNombre = hojaNombre;
      _rangoFilas      = 0;
      _rangoCols       = 0;
      mostrarEstado("Consultando al asistente...");
      const respuesta = await llamarApi("/ask", {
        pregunta: instruccion, historial: _historialLLM,
        device_id: _obtenerOCrearDeviceId(),
        user_email: obtenerEmailUsuario(),
        excel_version: _obtenerExcelVersion(),
      });

      if (respuesta.tipo === "datos" && respuesta.datos_modificados) {
        _datosModificados = respuesta.datos_modificados;
        mostrarRespuesta("✏️ " + respuesta.descripcion + "\n\n*Elige dónde escribir el resultado:*");
        mostrarDialogo(respuesta.descripcion);
        mostrarEstado("Tabla lista para escribir.");
        _agregarAlHistorial(instruccion, "✏️ " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else {
        mostrarRespuesta(respuesta.respuesta);
        mostrarEstado("Listo");
        _agregarAlHistorial(instruccion, respuesta.respuesta);
        _actualizarHistorialLLM(instruccion, respuesta.respuesta);
        _mostrarFeedback(instruccion, respuesta.respuesta);
      }
    }

    document.getElementById("pregunta").value = "";

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
    rango.worksheet.load("name");
    await context.sync();
    return {
      valores:    rango.values,
      direccion:  rango.address,
      hojaNombre: rango.worksheet.name,
    };
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

/**
 * Devuelve un array 2D de formatos de número inferidos a partir de los datos.
 * Fila 0 (cabecera): "General". Columnas numéricas: "#,##0.00" o "#,##0".
 * Devuelve null si no hay ninguna columna numérica (nada que aplicar).
 */
function _inferirFormatos(datos) {
  if (!datos || datos.length < 2) return null;
  const filas = datos.length;
  const cols  = datos[0].length;

  // Inicializar con "General"
  const fmts = Array.from({ length: filas }, () => Array(cols).fill("General"));
  let hayAlguno = false;

  for (let c = 0; c < cols; c++) {
    const vals = datos.slice(1)
      .map(row => row[c])
      .filter(v => v !== null && v !== "" && v !== undefined);
    if (!vals.length || !vals.every(v => typeof v === "number")) continue;

    const fmt = vals.some(v => v % 1 !== 0) ? "#,##0.00" : "#,##0";
    for (let r = 1; r < filas; r++) fmts[r][c] = fmt;
    hayAlguno = true;
  }

  return hayAlguno ? fmts : null;
}

async function escribirEnExcel(destino) {
  ocultarDialogo();
  mostrarEstado("Escribiendo en Excel...");

  const datos = _datosModificados;
  if (!datos || datos.length === 0) {
    mostrarEstado("No hay datos modificados para escribir.");
    return;
  }
  const filas = datos.length;
  const cols  = (datos[0] || []).length;

  // Dirección local sin prefijo de hoja ("Sheet1!A1:D11" → "A1:D11")
  const localAddress = _rangoAddress ? _rangoAddress.split("!").pop() : null;

  try {
    await Excel.run(async (context) => {
      let targetRange;
      let sourceRange = null;

      if (destino === "nueva_hoja") {
        // Capturar el rango fuente antes de crear la nueva hoja
        if (_rangoHojaNombre && localAddress && _rangoFilas > 0) {
          const srcSheet = context.workbook.worksheets.getItem(_rangoHojaNombre);
          sourceRange = srcSheet.getRange(localAddress);
        }
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

        if (_rangoFilas > 0 && _rangoCols > 0) {
          sourceRange = sheet.getRangeByIndexes(anchorRow, anchorCol, _rangoFilas, _rangoCols);
        }

        if (destino === "sustituir") {
          targetRange = sheet.getRangeByIndexes(anchorRow, anchorCol, filas, cols);
        } else if (destino === "derecha") {
          targetRange = sheet.getRangeByIndexes(anchorRow, anchorCol + _rangoCols, filas, cols);
        } else if (destino === "debajo") {
          targetRange = sheet.getRangeByIndexes(anchorRow + _rangoFilas, anchorCol, filas, cols);
        }
      }

      // Opción A: copiar formato del rango origen (colores, bordes, número).
      // Si target es más ancho que source, Office.js repite el patrón (tiling).
      if (sourceRange) {
        targetRange.copyFrom(sourceRange, Excel.RangeCopyType.formats, false, false);
      }

      targetRange.values = datos;

      // Opción B: si no hay rango fuente (tabla creada desde cero), inferir
      // formato numérico por columna y aplicarlo automáticamente.
      if (!sourceRange) {
        const fmts = _inferirFormatos(datos);
        if (fmts) targetRange.numberFormat = fmts;
      }

      await context.sync();
    });

    mostrarEstado("✅ Datos escritos correctamente.");
    _datosModificados = null;

  } catch (error) {
    mostrarEstado("Error al escribir: " + error.message);
  }
}

async function _aplicarFormatoCondicional(regla) {
  if (!_rangoAddress) throw new Error("No hay rango seleccionado.");

  await Excel.run(async (context) => {
    const sheet    = context.workbook.worksheets.getActiveWorksheet();
    const selRange = sheet.getRange(_rangoAddress);
    selRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();

    const cabeceras = selRange.values[0] || [];
    let rangoFormato;

    if (regla.col) {
      const colIdx = cabeceras.findIndex((c) => String(c) === String(regla.col));
      if (colIdx < 0) throw new Error(`Columna '${regla.col}' no encontrada en el rango seleccionado.`);
      rangoFormato = sheet.getRangeByIndexes(
        selRange.rowIndex + 1,
        selRange.columnIndex + colIdx,
        selRange.rowCount - 1,
        1,
      );
    } else {
      rangoFormato = sheet.getRangeByIndexes(
        selRange.rowIndex + 1,
        selRange.columnIndex,
        selRange.rowCount - 1,
        selRange.columnCount,
      );
    }

    rangoFormato.conditionalFormats.clearAll();
    const cfs = rangoFormato.conditionalFormats;

    switch (regla.tipo) {
      case "valor": {
        const cf  = cfs.add(Excel.ConditionalFormatType.cellValue);
        const op  = _OP_CF[regla.op] || "greaterThan";
        const rule = { formula1: String(regla.valor), operator: op };
        if (regla.op === "entre" || regla.op === "fuera") {
          rule.formula2 = String(regla.valor2 ?? regla.valor);
        }
        cf.cellValue.rule = rule;
        cf.cellValue.format.fill.color = _colorCf(regla.color);
        break;
      }
      case "top_bottom": {
        const cf   = cfs.add(Excel.ConditionalFormatType.topBottom);
        const tipo = regla.porcentaje
          ? (regla.direccion === "top" ? Excel.ConditionalTopBottomCriterionType.topPercent    : Excel.ConditionalTopBottomCriterionType.bottomPercent)
          : (regla.direccion === "top" ? Excel.ConditionalTopBottomCriterionType.topItems      : Excel.ConditionalTopBottomCriterionType.bottomItems);
        cf.topBottom.rule = { rank: regla.n || 10, type: tipo };
        cf.topBottom.format.fill.color = _colorCf(regla.color);
        break;
      }
      case "escala": {
        const cf     = cfs.add(Excel.ConditionalFormatType.colorScale);
        const colors = regla.colores || ["rojo", "verde"];
        const [c0, c1, c2] = colors;
        if (colors.length === 2) {
          cf.colorScale.criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue,  color: _colorCf(c0) },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: _colorCf(c1) },
          };
        } else {
          cf.colorScale.criteria = {
            minimum:  { type: Excel.ConditionalFormatColorCriterionType.lowestValue,   color: _colorCf(c0) },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: _colorCf(c1) },
            maximum:  { type: Excel.ConditionalFormatColorCriterionType.highestValue,  color: _colorCf(c2) },
          };
        }
        break;
      }
      case "barra": {
        const cf = cfs.add(Excel.ConditionalFormatType.dataBar);
        cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
        const fillColor = _colorCf(regla.color || "azul");
        cf.dataBar.positiveFormat.fillColor   = fillColor;
        cf.dataBar.positiveFormat.borderColor = fillColor;
        break;
      }
      case "icono": {
        const cf     = cfs.add(Excel.ConditionalFormatType.iconSet);
        const estilo = _ICONOS_CF[regla.estilo] || "threeArrows";
        cf.iconSet.style = Excel.IconSet[estilo];
        break;
      }
      case "texto": {
        const cf = cfs.add(Excel.ConditionalFormatType.containsText);
        const op = _OP_TEXTO_CF[regla.op] || "contains";
        cf.textComparison.rule = {
          operator: Excel.ConditionalTextOperator[op],
          text: String(regla.valor),
        };
        cf.textComparison.format.fill.color = _colorCf(regla.color);
        break;
      }
      case "formula": {
        const cf = cfs.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = regla.formula;
        cf.custom.format.fill.color = _colorCf(regla.color);
        break;
      }
      default:
        throw new Error(`Tipo de formato desconocido: ${regla.tipo}`);
    }

    await context.sync();
  });
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

function _mostrarFeedback(pregunta, respuesta) {
  _feedbackPregunta  = pregunta;
  _feedbackRespuesta = respuesta;
  const btn = document.getElementById("btn-feedback");
  if (!btn) return;
  btn.style.display = "inline-block";
  btn.disabled = false;
  btn.querySelector(".ms-Button-label").textContent = "👍 Útil";
}

function _ocultarFeedback() {
  const btn = document.getElementById("btn-feedback");
  if (btn) btn.style.display = "none";
}

async function enviarFeedback() {
  if (!_feedbackPregunta || !_feedbackRespuesta) return;
  const btn = document.getElementById("btn-feedback");
  if (btn) {
    btn.disabled = true;
    btn.querySelector(".ms-Button-label").textContent = "✅ Guardado";
  }
  try {
    await llamarApi("/feedback", {
      device_id: _obtenerOCrearDeviceId(),
      pregunta:  _feedbackPregunta,
      respuesta: _feedbackRespuesta,
    });
  } catch (_e) {
    // silencioso — el feedback es opcional
  }
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

function _actualizarHistorialLLM(pregunta, respuesta) {
  _historialLLM.push(
    { role: "user",  parts: [pregunta]  },
    { role: "model", parts: [respuesta] },
  );
  // Mantener solo los últimos N turnos para no saturar el contexto
  if (_historialLLM.length > _MAX_TURNOS_LLM) {
    _historialLLM = _historialLLM.slice(_historialLLM.length - _MAX_TURNOS_LLM);
  }
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
  _historialLLM = [];        // también limpiar el contexto enviado al LLM
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
      document.getElementById("btn-enviar-bot").style.display = "none";
      const secTelegram = document.getElementById("config-telegram");
      if (secTelegram) secTelegram.style.display = "none";
      return;
    }

    await comprobarVinculo();
  } catch {
    // Error de red o servidor antiguo sin este endpoint: flujo normal
    await comprobarVinculo();
  }
}

// ── Device ID persistente ─────────────────────────────────────────────────────

/**
 * Devuelve el UUID único de este dispositivo/instalación del Add-in.
 * Se genera una vez y se persiste en localStorage indefinidamente.
 */
function _obtenerOCrearDeviceId() {
  let deviceId = localStorage.getItem("addin-device-id");
  if (!deviceId) {
    deviceId = typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
          const r = Math.random() * 16 | 0;
          return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
        });
    localStorage.setItem("addin-device-id", deviceId);
  }
  return deviceId;
}

function _obtenerExcelVersion() {
  try {
    const diag = Office.context.diagnostics;
    const plat  = diag.platform  || "";
    const build = diag.version   || "";
    return (plat + (build ? " " + build : "")).trim() || null;
  } catch {
    return null;
  }
}

// ── Vínculo Telegram ──────────────────────────────────────────────────────────

async function comprobarVinculo() {
  const deviceId   = _obtenerOCrearDeviceId();
  const boton      = document.getElementById("btn-enviar-bot");
  const divSi      = document.getElementById("config-telegram-vinculado");
  const divNo      = document.getElementById("config-telegram-no-vinculado");

  try {
    const resultado = await llamarApiGet(
      "/tiene-vinculo?device_id=" + encodeURIComponent(deviceId)
    );
    if (resultado.vinculado) {
      boton.style.display = "";
      if (divSi) divSi.style.display = "";
      if (divNo) divNo.style.display = "none";
    } else {
      boton.style.display = "none";
      if (divSi) divSi.style.display = "none";
      if (divNo) divNo.style.display = "";
    }
  } catch {
    boton.style.display = "none";
    if (divSi) divSi.style.display = "none";
    if (divNo) divNo.style.display = "";
  }
}

async function verificarCodigo() {
  const input        = document.getElementById("input-codigo");
  const estadoVinc   = document.getElementById("estado-vinculo");
  const btnVincular  = document.getElementById("btn-vincular");
  const codigo       = input.value.trim();

  if (!/^\d{6}$/.test(codigo)) {
    estadoVinc.textContent = "Introduce un código de 6 dígitos.";
    return;
  }

  btnVincular.disabled  = true;
  estadoVinc.textContent = "Verificando...";

  try {
    const deviceId = _obtenerOCrearDeviceId();
    await llamarApi("/verificar-codigo", { device_id: deviceId, codigo });
    // ¡Vinculado! Limpiar input y re-comprobar para mostrar el botón 📤
    input.value            = "";
    estadoVinc.textContent = "";
    await comprobarVinculo();
  } catch (error) {
    const msg = error.message || "";
    if (msg.includes("expirado") || msg.includes("Expirado")) {
      estadoVinc.textContent = "⚠️ Código expirado. Genera uno nuevo con /codigo en Telegram.";
    } else if (msg.includes("utilizado") || msg.includes("Utilizado")) {
      estadoVinc.textContent = "⚠️ Código ya utilizado. Genera uno nuevo con /codigo.";
    } else if (msg.includes("encontrado")) {
      estadoVinc.textContent = "❌ Código incorrecto.";
    } else {
      estadoVinc.textContent = "❌ Error al verificar. Inténtalo de nuevo.";
    }
  } finally {
    btnVincular.disabled = false;
  }
}

// ── Enviar al bot ─────────────────────────────────────────────────────────────

async function enviarAlBot() {
  const deviceId = _obtenerOCrearDeviceId();
  const boton    = document.getElementById("btn-enviar-bot");
  boton.disabled = true;
  mostrarEstado("Leyendo selección...");

  try {
    const { valores } = await leerRangoSeleccionado();

    if (!valores || valores.length < 2) {
      mostrarEstado("Selecciona al menos una fila de cabeceras y una de datos.");
      return;
    }

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
      device_id:      deviceId,
    });

    mostrarEstado("✅ " + (resultado.mensaje || "Archivo enviado a Telegram."));

  } catch (error) {
    const msg = error.message || "Error desconocido";
    if (msg.includes("vínculo") || msg.includes("código") || msg.includes("vincular")) {
      mostrarEstado("⚠️ Sin vínculo activo. Introduce el código del bot en el Add-in.");
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
window.toggleTelegram             = toggleTelegram;
window.aplicarTemaDesdeSelect     = aplicarTemaDesdeSelect;
window.toggleHistorial            = toggleHistorial;
window.limpiarHistorial           = limpiarHistorial;
window.mostrarEasterEgg           = mostrarEasterEgg;
window.cerrarEasterEgg            = cerrarEasterEgg;
window.activarZeldaDesdeEasterEgg = activarZeldaDesdeEasterEgg;
window.enviarAlBot                 = enviarAlBot;
window.verificarCodigo             = verificarCodigo;
