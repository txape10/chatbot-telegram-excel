/* global Office, Excel, fetch, document, localStorage, navigator */

import {
  cargarTema, aplicarTema, aplicarTemaId, temasVisibles,
  TEMAS, desbloquearZelda, estaZeldaDesbloqueado,
} from "./themes.js";
import { estaAutorizado, obtenerEmailUsuario, obtenerNombreUsuario } from "./auth.js";

const API_URL = "";          // relativo: misma origin que el servidor FastAPI
/* global __API_KEY__ */
const API_KEY = __API_KEY__; // inyectada por webpack DefinePlugin desde .env — nunca en el repo

// Estado de la edición pendiente de ubicar
let _datosModificados = null;
let _nombreHoja       = null;
let _rangoAddress     = null;
let _rangoHojaNombre  = null;
let _rangoFilas       = 0;
let _rangoCols        = 0;
let _operacionActual  = null;

// Feedback RAG: pregunta+respuesta se almacenan como data-* en zona-feedback
// para que cada respuesta conserve sus propios datos aunque llegue otra petición.

// Aclaración pendiente: instrucción original guardada hasta que el usuario responda
let _instruccionPreAclaracion = null;

// Easter egg
let _eggInterval = null;

// Historial local (UI)
const _CLAVE_HISTORIAL   = "asistente-excel-historial";
const _MAX_HISTORIAL     = 20;
let _historialAbierto    = true;

// Historial de conversación enviado al LLM (contexto multi-turno)
const _CLAVE_HISTORIAL_LLM = "asistente-excel-historial-llm";
let _historialLLM          = [];
const _MAX_TURNOS_LLM      = 6; // 3 turnos (user+model) para no saturar el contexto

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
  _cargarHistorialLLM();
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

// ── Preguntar ─────────────────────────────────────────────────────────────────

const _PALABRAS_ZELDA = ["zelda", "link"];

async function preguntar() {
  let instruccion = document.getElementById("pregunta").value.trim();
  if (!instruccion) {
    mostrarEstado("Escribe una pregunta primero.");
    return;
  }

  // Si hay una aclaración pendiente, fusionar instrucción original + respuesta del usuario
  if (_instruccionPreAclaracion) {
    instruccion = _instruccionPreAclaracion + " — " + instruccion;
    _instruccionPreAclaracion = null;
  }

  // Easter egg: detectar palabras mágicas (siempre, aunque ya esté desbloqueado)
  if (_PALABRAS_ZELDA.includes(instruccion.toLowerCase())) {
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
    const { valores: _valoresSel, direccion: _direccionSel, hojaNombre: _hojaNombreSel } = await leerRangoSeleccionado();
    let tienesDatos = _valoresSel && _valoresSel.length >= 2;
    let valores     = _valoresSel;
    let direccion   = _direccionSel;
    let hojaNombre  = _hojaNombreSel;
    let _usedRangeCapturado = null;  // para pasar a /ask si no hay suficientes datos

    if (!tienesDatos) {
      // Sin selección — intentar con el rango usado de la hoja activa
      try {
        mostrarEstado("Leyendo hoja activa...");
        const usedRange = await leerUsedRange();
        if (usedRange.valores && usedRange.valores.length >= 2) {
          tienesDatos = true;
          valores     = usedRange.valores;
          direccion   = usedRange.direccion;
          hojaNombre  = usedRange.hojaNombre;
        }
        _usedRangeCapturado = usedRange;  // capturar para fallback en /ask
      } catch (_e) { /* hoja vacía — continuar con /ask */ }
    }

    if (tienesDatos) {
      // ── Flujo con datos: edición / consulta sobre la tabla seleccionada ──
      _rangoAddress    = direccion;
      _rangoHojaNombre = hojaNombre;
      _rangoFilas      = valores.length;
      _rangoCols       = (valores[0] || []).length;
      mostrarEstado("Consultando al asistente... (rango: " + direccion + ")");

      // ── Flujo con datos: el backend decide edición, formato condicional o consulta ──
      mostrarEstado("Consultando al asistente... (rango: " + direccion + ")");
      const respuesta = await llamarApi("/edit", {
        datos: valores, instruccion, historial: _historialLLM,
        device_id:    _obtenerOCrearDeviceId(),
        user_email:   obtenerEmailUsuario(),
        display_name: obtenerNombreUsuario(),
        excel_version: _obtenerExcelVersion(),
      });

      if (respuesta.tipo === "pipeline" && respuesta.pasos) {
        // Pipeline multi-op: ejecutar cada paso en orden
        mostrarEstado("Ejecutando pipeline...");
        let ultimaEdicion = null;
        const resultadosQuery = [];
        let graficosInsertados = 0;
        let hojaAnalisisNombre = null;
        for (const paso of respuesta.pasos) {
          const res = await _aplicarPaso(paso, graficosInsertados);
          if (paso.tipo === "edicion") ultimaEdicion = paso;
          if (paso.tipo === "query_resultado") resultadosQuery.push(paso);
          if (paso.tipo === "analisis_hoja") resultadosQuery.push({ pregunta: "Análisis del archivo", resultado: paso.texto });
          if (paso.tipo === "grafico" && paso.datos_chart) graficosInsertados++;
          if (res?.hojaAnalisis) hojaAnalisisNombre = res.hojaAnalisis;
        }
        // Activar la hoja de análisis al final (el gráfico ya se insertó en la hoja original)
        if (hojaAnalisisNombre) {
          await Excel.run(async (ctx) => {
            ctx.workbook.worksheets.getItem(hojaAnalisisNombre).activate();
            await ctx.sync();
          });
        }
        if (ultimaEdicion) {
          _datosModificados = ultimaEdicion.datos_modificados;
          _operacionActual  = ultimaEdicion.operacion || null;
          const _da = { final: "debajo", principio: "principio" }[ultimaEdicion.destino] ?? "sustituir";
          await escribirEnExcel(_da);
          // Sobreescribir columnas aritméticas con fórmulas Excel vivas
          const _fmlCols = respuesta.pasos.filter(p => p.tipo === "formula_columna");
          for (const fml of _fmlCols) { await _aplicarFormulaColumna(fml); }
          // Aplicar formatos numéricos pendientes (columnas con N decimales)
          const _fmtNums = respuesta.pasos.filter(p => p.tipo === "formato_numero");
          for (const fmt of _fmtNums) { await _aplicarFormatoNumero(fmt); }
        }
        // Respuesta organizada: cambios + consultas
        const partesEdicion = respuesta.pasos
          .filter(p => p.tipo !== "query_resultado")
          .map(p => p.descripcion).filter(Boolean);
        let textoRespuesta = "";
        if (partesEdicion.length > 0) {
          textoRespuesta += "✅ **Cambios aplicados:**\n" + partesEdicion.map(d => `• ${d}`).join("\n");
        }
        if (resultadosQuery.length > 0) {
          if (textoRespuesta) textoRespuesta += "\n\n";
          textoRespuesta += "📊 **Consultas:**\n" + resultadosQuery.map(q => `**${q.pregunta}**\n${q.resultado}`).join("\n\n");
        }
        if (!textoRespuesta) textoRespuesta = "✅ " + respuesta.descripcion;
        mostrarRespuesta(textoRespuesta);
        mostrarEstado("Pipeline aplicada · " + direccion);
        _agregarAlHistorial(instruccion, textoRespuesta);
        _actualizarHistorialLLM(instruccion, textoRespuesta);
      } else if (respuesta.tipo === "analisis_hoja" && respuesta.hoja_datos) {
        const nombre = "Análisis " + (_rangoHojaNombre || "datos");
        mostrarEstado("Creando hoja de análisis...");
        await _insertarHojaAnalisis(respuesta.hoja_datos, nombre);
        await Excel.run(async (ctx) => {
          ctx.workbook.worksheets.getItem(nombre).activate();
          await ctx.sync();
        });
        mostrarRespuesta("📊 " + respuesta.descripcion + "\n\nHoja **\"" + nombre + "\"** creada con el análisis completo.\n\n" + (respuesta.texto || ""));
        mostrarEstado("Análisis listo · " + nombre);
        _agregarAlHistorial(instruccion, "📊 " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else if (respuesta.tipo === "tabla_dinamica" && respuesta.params) {
        mostrarEstado("Creando tabla dinámica...");
        await _insertarTablaDinamica(respuesta.params);
        mostrarRespuesta("📊 " + respuesta.descripcion + "\n\nTabla dinámica creada en hoja nueva.");
        mostrarEstado("Tabla dinámica lista.");
        _agregarAlHistorial(instruccion, "📊 " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else if (respuesta.tipo === "grafico" && respuesta.datos_chart) {
        mostrarEstado("Insertando gráfico...");
        await _insertarGrafico(respuesta.tipo_grafico, respuesta.datos_chart, respuesta.titulo);
        mostrarRespuesta("📊 " + respuesta.descripcion);
        mostrarEstado("Gráfico insertado · " + direccion);
        _agregarAlHistorial(instruccion, "📊 " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else if (respuesta.tipo === "formato" && (respuesta.reglas || respuesta.regla)) {
        // Formato condicional real de Office.js (reglas = array; regla = legacy)
        const reglas = respuesta.reglas || [respuesta.regla];
        await _aplicarFormatosCondicionales(reglas);
        mostrarRespuesta("🎨 " + respuesta.descripcion);
        mostrarEstado("Formato aplicado · " + direccion);
        _agregarAlHistorial(instruccion, "🎨 " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else if (respuesta.tipo === "edicion") {
        _datosModificados = respuesta.datos_modificados;
        _operacionActual  = respuesta.operacion || null;
        const _destinoAuto = { final: "debajo", principio: "principio" }[respuesta.destino] ?? null;
        if (_destinoAuto) {
          mostrarRespuesta("✏️ " + respuesta.descripcion);
          mostrarEstado("Escribiendo...");
          await escribirEnExcel(_destinoAuto);
        } else {
          mostrarRespuesta("✏️ " + respuesta.descripcion + "\n\n*Elige dónde escribir el resultado:*");
          mostrarDialogo(respuesta.descripcion);
          mostrarEstado("Edición lista · " + direccion);
        }
        _agregarAlHistorial(instruccion, "✏️ " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else if (respuesta.tipo === "aclaracion") {
        // Guardar instrucción original para fusionarla con la respuesta del usuario
        _instruccionPreAclaracion = instruccion;
        const preguntaTexto = respuesta.pregunta || "¿Puedes concretar más?";
        const opciones = respuesta.opciones || [];
        // Mostrar pregunta + botones pulsables (un clic envía directamente)
        let html = `<p>🤔 ${preguntaTexto}</p>`;
        if (opciones.length > 0) {
          html += '<div class="opciones-aclaracion">';
          opciones.forEach((o) => {
            const esc = o.replace(/"/g, "&quot;");
            html += `<button class="ms-Button btn-aclaracion" onclick="elegirAclaracion(&quot;${esc}&quot;)">
              <span class="ms-Button-label">${o}</span>
            </button>`;
          });
          html += "</div>";
        }
        document.getElementById("respuesta").innerHTML = html;
        document.getElementById("bloque-respuesta").style.display = "block";
        mostrarEstado("Se necesita aclaración");
      } else if (respuesta.tipo === "query_resultado") {
        const msg = respuesta.resultado || respuesta.descripcion || "Sin resultado";
        if (respuesta.datos_tabla && respuesta.datos_tabla.length > 1) {
          mostrarEstado("Escribiendo resultado en hoja...");
          await _escribirTablaEnHoja(respuesta.datos_tabla, "Consulta");
          const resumen = respuesta.texto_resumen ? respuesta.texto_resumen + "\n\n" : "";
          mostrarRespuesta("📊 " + resumen + "Resultado escrito en hoja **Consulta**.");
          mostrarEstado("Listo · hoja Consulta");
        } else {
          mostrarRespuesta("📊 " + msg);
          mostrarEstado("Listo · " + direccion);
        }
        _agregarAlHistorial(instruccion, msg);
        _actualizarHistorialLLM(instruccion, msg);
        _mostrarFeedback(instruccion, msg);
      } else {
        const msg = respuesta.respuesta || respuesta.mensaje || "Sin respuesta";
        mostrarRespuesta(msg);
        mostrarEstado("Listo · " + direccion);
        _agregarAlHistorial(instruccion, msg);
        _actualizarHistorialLLM(instruccion, msg);
        _mostrarFeedback(instruccion, msg);
      }

    } else {
      // ── Flujo sin datos: pregunta general o creación desde cero ──
      // Si tenemos datos de la hoja (pero no suficientes para /edit), los pasamos
      // para que el backend pueda detectar intención de edición y redirigir.
      _rangoAddress    = _usedRangeCapturado?.direccion  || direccion;
      _rangoHojaNombre = _usedRangeCapturado?.hojaNombre || hojaNombre;
      _rangoFilas      = 0;
      _rangoCols       = 0;
      mostrarEstado("Consultando al asistente...");
      const respuesta = await llamarApi("/ask", {
        pregunta: instruccion,
        datos:    _usedRangeCapturado?.valores || null,
        historial: _historialLLM,
        device_id:    _obtenerOCrearDeviceId(),
        user_email:   obtenerEmailUsuario(),
        display_name: obtenerNombreUsuario(),
        excel_version: _obtenerExcelVersion(),
      });

      if (respuesta.tipo === "datos" && respuesta.datos_modificados) {
        _datosModificados = respuesta.datos_modificados;
        _operacionActual  = respuesta.operacion || null;
        _nombreHoja = respuesta.nombre_hoja || null;
        if (respuesta.nueva_hoja) {
          await escribirEnExcel("nueva_hoja");
          mostrarRespuesta("✅ " + respuesta.descripcion);
          mostrarEstado("✅ Hoja '" + respuesta.nombre_hoja + "' creada.");
          _agregarAlHistorial(instruccion, "✅ " + respuesta.descripcion);
          _actualizarHistorialLLM(instruccion, respuesta.descripcion);
        } else {
          mostrarRespuesta("✏️ " + respuesta.descripcion + "\n\n*Elige dónde escribir el resultado:*");
          mostrarDialogo(respuesta.descripcion);
          mostrarEstado("Tabla lista para escribir.");
          _agregarAlHistorial(instruccion, "✏️ " + respuesta.descripcion);
          _actualizarHistorialLLM(instruccion, respuesta.descripcion);
        }
      } else if (respuesta.tipo === "edicion" && respuesta.datos_modificados) {
        // El backend detectó intención de edición desde /ask — procesar igual que /edit
        _datosModificados = respuesta.datos_modificados;
        _operacionActual  = respuesta.operacion || null;
        const _dest = { final: "debajo", principio: "principio" }[respuesta.destino] ?? null;
        if (_dest) {
          mostrarRespuesta("✏️ " + respuesta.descripcion);
          mostrarEstado("Escribiendo...");
          await escribirEnExcel(_dest);
        } else {
          mostrarRespuesta("✏️ " + respuesta.descripcion + "\n\n*Elige dónde escribir el resultado:*");
          mostrarDialogo(respuesta.descripcion);
          mostrarEstado("Edición lista");
        }
        _agregarAlHistorial(instruccion, "✏️ " + respuesta.descripcion);
        _actualizarHistorialLLM(instruccion, respuesta.descripcion);
      } else if (respuesta.tipo === "pipeline" && respuesta.pasos) {
        // Pipeline de edición detectado desde /ask
        mostrarEstado("Ejecutando...");
        let _ultEdit = null;
        for (const paso of respuesta.pasos) {
          await _aplicarPaso(paso);
          if (paso.tipo === "edicion") _ultEdit = paso;
        }
        if (_ultEdit) {
          _datosModificados = _ultEdit.datos_modificados;
          _operacionActual  = _ultEdit.operacion || null;
          await escribirEnExcel("sustituir");
        }
        const _desc = respuesta.descripcion || (_ultEdit?.descripcion ?? "Listo");
        mostrarRespuesta("✅ " + _desc);
        mostrarEstado("Listo");
        _agregarAlHistorial(instruccion, "✅ " + _desc);
        _actualizarHistorialLLM(instruccion, _desc);
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

/**
 * Combina values y text: sustituye seriales de fecha Excel (enteros 25000–60000)
 * por su representación de texto formateada, dejando el resto sin tocar.
 */
function _combinarFechas(values, text) {
  return values.map((fila, i) =>
    fila.map((val, j) => {
      if (typeof val === "number" && Number.isInteger(val) && val > 25000 && val < 60000) {
        return (text[i] && text[i][j] != null) ? text[i][j] : val;
      }
      return val;
    })
  );
}

async function leerRangoSeleccionado() {
  return Excel.run(async (context) => {
    const rango = context.workbook.getSelectedRange();
    rango.load(["values", "text", "address"]);
    rango.worksheet.load("name");
    await context.sync();
    return {
      valores:    _combinarFechas(rango.values, rango.text),
      direccion:  rango.address,
      hojaNombre: rango.worksheet.name,
    };
  });
}

async function leerUsedRange() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rango = sheet.getUsedRangeOrNullObject();
    rango.load(["isNullObject", "values", "text", "address"]);
    sheet.load("name");
    await context.sync();
    if (rango.isNullObject) {
      return { valores: null, direccion: null, hojaNombre: sheet.name };
    }
    return {
      valores:    _combinarFechas(rango.values, rango.text),
      direccion:  rango.address,
      hojaNombre: sheet.name,
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
const _DATE_HEADER_KW = ["fecha", "date", "día", "dia", "vencim", "alta", "baja", "inicio", "fin"];
// Reconoce "dd/mm/yyyy" y "yyyy-mm-dd"
const _DATE_RE = /^\d{1,2}\/\d{1,2}\/\d{4}$|^\d{4}-\d{2}-\d{2}$/;

/**
 * Convierte un string de fecha ("dd/mm/yyyy" o "yyyy-mm-dd") al número
 * serial de Excel (días desde 30/12/1899). Devuelve el string original
 * si no puede interpretarlo.
 */
function _dateStringToSerial(str) {
  const p = str.includes("/") ? str.split("/").map(Number) : null;
  const g = str.includes("-") ? str.split("-").map(Number) : null;
  let d, m, y;
  if (p && p.length === 3) { [d, m, y] = p; }
  else if (g && g.length === 3) { [y, m, d] = g; }
  else { return str; }
  if (isNaN(d + m + y) || y < 1900 || m < 1 || m > 12 || d < 1 || d > 31) return str;
  const ms = new Date(y, m - 1, d).getTime() - new Date(1899, 11, 30).getTime();
  return Math.round(ms / 86400000);
}

/**
 * Prepara la matriz de datos para escribir en Excel:
 * - Convierte strings de fecha a seriales (para alineación y filtros correctos)
 * - Devuelve también la matriz de formatos a aplicar con numberFormat
 *
 * Detecta columnas de fecha por nombre de cabecera O por patrón de valor.
 * Detecta columnas numéricas para aplicar separador de miles.
 */
function _prepararDatos(datos) {
  if (!datos || datos.length < 2) return { valores: datos, formatos: null };

  const filas   = datos.length;
  const cols    = datos[0].length;
  const headers = datos[0].map(h => String(h || "").toLowerCase());
  const valores = datos.map(row => [...row]);   // copia profunda
  const fmts    = Array.from({ length: filas }, () => Array(cols).fill("General"));
  let hayAlguno = false;

  for (let c = 0; c < cols; c++) {
    const colVals = datos.slice(1)
      .map(row => row[c])
      .filter(v => v !== null && v !== "" && v !== undefined);

    // ── Columna de fecha (por cabecera o por patrón de valores) ──────────────
    const esFechaHeader = _DATE_HEADER_KW.some(kw => headers[c].includes(kw));
    const esFechaValor  = colVals.length > 0 &&
      colVals.every(v => typeof v === "string" && _DATE_RE.test(v));

    if (esFechaHeader || esFechaValor) {
      for (let r = 1; r < filas; r++) {
        const v = datos[r][c];
        if (typeof v === "string" && _DATE_RE.test(v)) {
          valores[r][c] = _dateStringToSerial(v);
        }
        fmts[r][c] = "dd/mm/yyyy";
      }
      hayAlguno = true;
      continue;
    }

    // ── Columna numérica ─────────────────────────────────────────────────────
    if (!colVals.length || !colVals.every(v => typeof v === "number")) continue;
    const fmt = colVals.some(v => v % 1 !== 0) ? "#,##0.00" : "#,##0";
    for (let r = 1; r < filas; r++) fmts[r][c] = fmt;
    hayAlguno = true;
  }

  return { valores, formatos: hayAlguno ? fmts : null };
}

/** Convierte índice de columna 0-based a letra(s) Excel (0→A, 25→Z, 26→AA, …). */
function _idxToColLetter(colIdx) {
  let result = "";
  let n = colIdx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

/**
 * Añade una columna con fórmulas Excel vivas (=E2-F2) al final de los datos escritos.
 * Usa getUsedRange() para localizar la posición exacta: no depende de _rangoAddress.
 */
async function _aplicarFormulaColumna(paso) {
  await Excel.run(async (ctx) => {
    const sheet     = ctx.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load(["rowIndex", "columnIndex", "columnCount", "rowCount"]);
    await ctx.sync();

    const anchorRow  = usedRange.rowIndex;
    const anchorCol  = usedRange.columnIndex;
    const numDataRows = usedRange.rowCount - 1;   // excluir fila de cabecera
    const newColIdx  = anchorCol + usedRange.columnCount;  // siguiente columna tras los datos

    if (numDataRows <= 0) return;

    const col1Letra = _idxToColLetter(anchorCol + paso.col1_idx);
    const col2Letra = paso.col2_idx !== undefined
      ? _idxToColLetter(anchorCol + paso.col2_idx)
      : null;
    const op        = paso.operador;

    // Cabecera de la nueva columna
    sheet.getRangeByIndexes(anchorRow, newColIdx, 1, 1).values = [[paso.col]];

    // Fórmulas para cada fila de datos
    // col op col → =E2*F2 | col op constante → =E2*1.21
    const formulas = [];
    for (let i = 0; i < numDataRows; i++) {
      const rowNum    = anchorRow + i + 2;  // fila Excel 1-based: +1 cabecera, +1 base 1
      const operando2 = col2Letra ? `${col2Letra}${rowNum}` : String(paso.valor_fijo);
      formulas.push([`=${col1Letra}${rowNum}${op}${operando2}`]);
    }
    sheet.getRangeByIndexes(anchorRow + 1, newColIdx, numDataRows, 1).formulas = formulas;
    await ctx.sync();
  });
}

/**
 * Aplica formato numérico (#,##0.00) a una columna específica tras escribirEnExcel.
 * Usa _rangoAddress como ancla para localizar la columna en la hoja.
 */
async function _aplicarFormatoNumero(paso) {
  if (paso.col_idx == null || !_rangoAddress || !_rangoHojaNombre) return;
  const decimales = paso.decimales || 2;
  const fmtStr = "#,##0." + "0".repeat(decimales);
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(_rangoHojaNombre);
    const localAddress = _rangoAddress.split("!").pop();
    const refRange = sheet.getRange(localAddress);
    refRange.load("rowIndex, columnIndex");
    await ctx.sync();
    const numFilas = (_datosModificados?.length || 2) - 1;  // excluir cabecera
    if (numFilas <= 0) return;
    const colRange = sheet.getRangeByIndexes(
      refRange.rowIndex + 1,
      refRange.columnIndex + paso.col_idx,
      numFilas,
      1
    );
    colRange.numberFormat = Array(numFilas).fill([fmtStr]);
    await ctx.sync();
  });
}

async function escribirEnExcel(destino) {
  ocultarDialogo();
  mostrarEstado("Escribiendo en Excel...");

  const datos = _datosModificados;
  if (!datos || datos.length === 0) {
    mostrarEstado("No hay datos modificados para escribir.");
    return;
  }

  const localAddress = _rangoAddress ? _rangoAddress.split("!").pop() : null;

  try {
    await Excel.run(async (context) => {

      // ── Nueva hoja: flujo independiente ──────────────────────────────────────
      if (destino === "nueva_hoja") {
        let sourceRange = null;
        if (_rangoHojaNombre && localAddress && _rangoFilas > 0) {
          const srcSheet = context.workbook.worksheets.getItem(_rangoHojaNombre);
          sourceRange = srcSheet.getRange(localAddress);
        }
        const sheet       = context.workbook.worksheets.add(_nombreHoja || undefined);
        sheet.activate();
        const filas       = datos.length;
        const cols        = (datos[0] || []).length;
        const targetRange = sheet.getRangeByIndexes(0, 0, filas, cols);
        if (sourceRange) {
          targetRange.copyFrom(sourceRange, Excel.RangeCopyType.formats, false, false);
        }
        const { valores: datosPrep, formatos: fmtsPrep } = _prepararDatos(datos);
        targetRange.values = datosPrep;
        if (!sourceRange && fmtsPrep) targetRange.numberFormat = fmtsPrep;
        await context.sync();
        if (filas >= 2) {
          try {
            const tabla = sheet.tables.add(sheet.getRangeByIndexes(0, 0, filas, cols), true);
            tabla.style = "TableStyleMedium2";
            await context.sync();
            tabla.getRange().select();
          } catch (_) {
            sheet.getRangeByIndexes(0, 0, filas, cols).select();
          }
          await context.sync();
        }
        return;
      }

      // ── Destinos en hoja activa ───────────────────────────────────────────────
      const sheet    = context.workbook.worksheets.getActiveWorksheet();
      const refRange = sheet.getRange(localAddress || "A1");
      refRange.load(["rowIndex", "columnIndex"]);
      await context.sync();

      const anchorRow = refRange.rowIndex;
      const anchorCol = refRange.columnIndex;

      // Detectar tabla adyacente (debajo/derecha): necesario antes de escribir
      // para saber si hay que saltar la cabecera del resultado.
      let tablaInfo = null;
      if ((destino === "debajo" || destino === "derecha" || destino === "principio") && _rangoFilas > 0 && _rangoCols > 0) {
        const srcRange = sheet.getRangeByIndexes(anchorRow, anchorCol, _rangoFilas, _rangoCols);
        const tablas   = srcRange.getTables(false);
        tablas.load("items/id");
        await context.sync();
        if (tablas.items.length > 0) {
          const tablaRango = tablas.items[0].getRange();
          tablaRango.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
          await context.sync();
          tablaInfo = {
            tabla:       tablas.items[0],
            rowIndex:    tablaRango.rowIndex,
            rowCount:    tablaRango.rowCount,
            columnIndex: tablaRango.columnIndex,
            columnCount: tablaRango.columnCount,
          };
        }
      }

      // Los datos del backend siempre incluyen cabecera en [0].
      // Al escribir DEBAJO de una tabla existente hay dos casos:
      //   a) El resultado tiene MÁS filas que la selección original → operación que añade filas
      //      (duplicar_filas, añadir_fila_total…): escribir solo las filas extra.
      //   b) Igual o menos filas → resultado transformado (filtro, sort…): escribir sin cabecera
      //      para evitar duplicarla, pero como bloque separado (no extiende la tabla).
      const _filasNuevas = (_rangoFilas > 0 && datos.length > _rangoFilas)
        ? datos.length - _rangoFilas   // solo las filas añadidas
        : 0;
      const _conTablaFilas = (destino === "debajo" || destino === "principio") && tablaInfo;
      const datosAEscribir = _conTablaFilas
        ? (_filasNuevas > 0 ? datos.slice(_rangoFilas) : datos.slice(1))
        : datos;

      const filas = datosAEscribir.length;
      const cols  = (datosAEscribir[0] || []).length;
      if (filas === 0 || cols === 0) return;

      // Rango fuente para copyFrom (formato visual)
      const sourceRange = (_rangoFilas > 0 && _rangoCols > 0)
        ? sheet.getRangeByIndexes(anchorRow, anchorCol, _rangoFilas, _rangoCols)
        : null;

      // Rango destino según modo
      let targetRange;
      if (destino === "sustituir") {
        targetRange = sheet.getRangeByIndexes(anchorRow, anchorCol, filas, cols);
      } else if (destino === "derecha") {
        targetRange = sheet.getRangeByIndexes(anchorRow, anchorCol + _rangoCols, filas, cols);
      } else if (destino === "principio") {
        if (tablaInfo) {
          // rows.add() dentro de una tabla evita el error de range.insert() y expande el ListObject
          const { valores: vprinc } = _prepararDatos(datosAEscribir);
          tablaInfo.tabla.rows.add(0, vprinc);
          await context.sync();
          tablaInfo.tabla.getRange().select();
          await context.sync();
          return;
        }
        // Sin tabla: insertar filas desplazando hacia abajo y escribir normalmente
        targetRange = sheet.getRangeByIndexes(anchorRow + 1, anchorCol, filas, cols);
        targetRange.insert(Excel.InsertShiftDirection.down);
        await context.sync();
      } else {
        targetRange = sheet.getRangeByIndexes(anchorRow + _rangoFilas, anchorCol, filas, cols);
      }

      if (sourceRange) {
        targetRange.copyFrom(sourceRange, Excel.RangeCopyType.formats, false, false);
      }

      // _prepararDatos necesita la cabecera para detectar tipos de columna;
      // cuando saltamos filas pasamos datos completos y luego cortamos el resultado.
      const datosParaFormato = _conTablaFilas ? datos : datosAEscribir;
      const { valores: datosPrep, formatos: fmtsPrep } = _prepararDatos(datosParaFormato);
      // Índice de corte: _rangoFilas cuando hay filas nuevas (caso a), 1 en caso contrario
      const _sliceIdx = _conTablaFilas ? (_filasNuevas > 0 ? _rangoFilas : 1) : 0;
      const valoresAEscribir = _sliceIdx > 0 ? datosPrep.slice(_sliceIdx) : datosPrep;
      targetRange.values = valoresAEscribir;

      if (!sourceRange && fmtsPrep) {
        const fmtsAEscribir = _sliceIdx > 0 ? fmtsPrep.slice(_sliceIdx) : fmtsPrep;
        targetRange.numberFormat = fmtsAEscribir;
      }

      await context.sync();

      // Gestionar tabla Excel: extender la existente o crear nueva para sustituir.
      // Solo extender en debajo si hay filas verdaderamente nuevas (operación añade filas);
      // si el resultado es transformado (filtro, sort…) se escribe como bloque independiente.
      if (tablaInfo) {
        const filasExtra = (destino === "debajo" && _filasNuevas > 0) ? filas : 0;
        const colsExtra  = destino === "derecha" ? cols : 0;
        // "principio": el insert() dentro de la tabla ya expande el ListObject — no resize
        if (filasExtra > 0 || colsExtra > 0) {
          const nuevoRango = sheet.getRangeByIndexes(
            tablaInfo.rowIndex,
            tablaInfo.columnIndex,
            tablaInfo.rowCount  + filasExtra,
            tablaInfo.columnCount + colsExtra,
          );
          tablaInfo.tabla.resize(nuevoRango);
          await context.sync();
        }
        tablaInfo.tabla.getRange().select();
        await context.sync();
      } else if (destino === "sustituir" && filas >= 2) {
        try {
          const tabla = sheet.tables.add(
            sheet.getRangeByIndexes(anchorRow, anchorCol, filas, cols), true,
          );
          tabla.style = "TableStyleMedium2";
          await context.sync();
          tabla.getRange().select();
          await context.sync();
        } catch (_) {
          targetRange.select();
          await context.sync();
        }
      } else {
        targetRange.select();
        await context.sync();
      }
    });

    mostrarEstado("✅ Datos escritos correctamente.");
    _datosModificados = null;
    _nombreHoja = null;

  } catch (error) {
    mostrarEstado("Error al escribir: " + error.message);
  }
}

/** Aplica una regla DSL individual sobre un rango ya cargado (sin sync propio). */
function _aplicarUnaRegla(cfs, regla) {
  switch (regla.tipo) {
    case "valor": {
      const cf   = cfs.add(Excel.ConditionalFormatType.cellValue);
      const op   = _OP_CF[regla.op] || "greaterThan";
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
      const cf        = cfs.add(Excel.ConditionalFormatType.dataBar);
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
}

/** Aplica un array de reglas DSL. Agrupa por columna para hacer un solo clearAll por columna. */
/** Despacha un paso individual de un pipeline. Reutiliza los handlers existentes. */
async function _aplicarPaso(paso, chartIndex = 0) {
  switch (paso.tipo) {
    case "edicion":
      // Los datos se escriben al final (el caller maneja _datosModificados)
      break;
    case "formula":
      if (paso.col_nueva && paso.formulas) {
        await _aplicarFormula(paso);
      }
      break;
    case "formato":
      if (paso.reglas || paso.regla) {
        await _aplicarFormatosCondicionales(paso.reglas || [paso.regla]);
      }
      break;
    case "grafico":
      if (paso.datos_chart) {
        await _insertarGrafico(paso.tipo_grafico, paso.datos_chart, paso.titulo, chartIndex);
      }
      break;
    case "tabla_dinamica":
      if (paso.params) {
        await _insertarTablaDinamica(paso.params);
      }
      break;
    case "analisis_hoja": {
      const nombre = "Análisis " + (_rangoHojaNombre || "datos");
      await _insertarHojaAnalisis(paso.hoja_datos, nombre);
      return { hojaAnalisis: nombre };
    }
    case "query_resultado":
      if (paso.datos_tabla && paso.datos_tabla.length > 1) {
        await _escribirTablaEnHoja(paso.datos_tabla, "Consulta");
      }
      break;
    default:
      break;
  }
  return null;
}

async function _escribirTablaEnHoja(matriz, nombre) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    let hoja = wb.worksheets.getItemOrNullObject(nombre);
    await context.sync();
    if (hoja.isNullObject) {
      hoja = wb.worksheets.add(nombre);
    } else {
      hoja.getUsedRange().clear();
    }
    const nFilas = matriz.length;
    const nCols  = matriz[0].length;
    const rango  = hoja.getRange("A1").getResizedRange(nFilas - 1, nCols - 1);
    rango.values = matriz;

    // Cabecera negrita con fondo gris claro
    const cabecera = hoja.getRange("A1").getResizedRange(0, nCols - 1);
    cabecera.format.font.bold = true;
    cabecera.format.fill.color = "#D9D9D9";

    // Formato por columna según tipo de dato (detectado en la primera fila de datos)
    if (nFilas > 1) {
      const headers = matriz[0];
      for (let c = 0; c < nCols; c++) {
        const val = matriz[1][c];
        const colRng = hoja.getRange("A1").getOffsetRange(1, c).getResizedRange(nFilas - 2, 0);
        const headerLower = String(headers[c]).toLowerCase();
        if (headerLower.includes("fecha") || headerLower.includes("date")) {
          // Seriales de fecha de Excel → formato DD/MM/YYYY
          colRng.numberFormat = [["DD/MM/YYYY"]];
        } else if (typeof val === "number" && !Number.isInteger(val)) {
          colRng.numberFormat = [["#,##0.00"]];
        } else if (typeof val === "number") {
          colRng.numberFormat = [["#,##0"]];
        }
      }
    }

    hoja.getUsedRange().format.autofitColumns();
    hoja.activate();
    await context.sync();
  });
}

async function _insertarHojaAnalisis(datos, nombre) {
  await Excel.run(async (context) => {
    const wb = context.workbook;
    let hoja = wb.worksheets.getItemOrNullObject(nombre);
    hoja.load("isNullObject");
    await context.sync();

    if (hoja.isNullObject) {
      hoja = wb.worksheets.add(nombre);
    } else {
      hoja.getRange().clear();
    }

    if (!datos || datos.length === 0) { await context.sync(); return; }

    const nCols = Math.max(...datos.map(r => r.length));
    const matriz = datos.map(fila => {
      const norm = [...fila];
      while (norm.length < nCols) norm.push("");
      return norm;
    });

    hoja.getRangeByIndexes(0, 0, matriz.length, nCols).values = matriz;

    // Negritas + fondo azul claro para filas de sección (todo mayúsculas en col A)
    datos.forEach((fila, i) => {
      const celda = String(fila[0] || "").trim();
      if (celda.length > 2 && celda === celda.toUpperCase()) {
        const rFila = hoja.getRangeByIndexes(i, 0, 1, nCols);
        rFila.format.font.bold = true;
        rFila.format.fill.color = "#D9E1F2";
      }
    });

    hoja.getUsedRangeOrNullObject().format.autofitColumns();
    await context.sync();
  });
}

async function _aplicarFormula(paso) {
  const { col_nueva, formulas } = paso;
  if (!col_nueva || !formulas || formulas.length === 0) return;
  await Excel.run(async (context) => {
    const sheet     = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load(["columnCount", "rowCount"]);
    await context.sync();

    const nCols   = usedRange.columnCount;
    const nRows   = usedRange.rowCount;
    if (nRows < 2) return; // solo cabecera, sin datos

    // Cabecera de la nueva columna
    sheet.getRangeByIndexes(0, nCols, 1, 1).values = [[col_nueva]];

    // Fórmulas para cada fila de datos
    const dataRows  = Math.min(formulas.length, nRows - 1);
    const dataRange = sheet.getRangeByIndexes(1, nCols, dataRows, 1);
    dataRange.formulas = formulas.slice(0, dataRows);

    await context.sync();
  });
}

async function _aplicarFormatosCondicionales(reglas) {
  if (!_rangoAddress) throw new Error("No hay rango seleccionado.");

  await Excel.run(async (context) => {
    const sheet     = context.workbook.worksheets.getActiveWorksheet();
    const localAddr = _rangoAddress.includes("!") ? _rangoAddress.split("!")[1] : _rangoAddress;
    const selRange  = sheet.getRange(localAddr);
    selRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();

    const cabeceras = selRange.values[0] || [];

    // Agrupar reglas por columna objetivo para hacer clearAll una vez por columna
    const porColumna = new Map(); // colKey → { rangoFormato, reglas[] }

    for (const regla of reglas) {
      let colKey, rangoFormato;

      if (regla.col) {
        const colIdx = cabeceras.findIndex((c) => String(c) === String(regla.col));
        if (colIdx < 0) throw new Error(`Columna '${regla.col}' no encontrada en el rango seleccionado.`);
        colKey = String(colIdx);
        if (!porColumna.has(colKey)) {
          rangoFormato = sheet.getRangeByIndexes(
            selRange.rowIndex + 1,
            selRange.columnIndex + colIdx,
            selRange.rowCount - 1,
            1,
          );
          porColumna.set(colKey, { rangoFormato, reglas: [] });
        }
      } else {
        colKey = "__all__";
        if (!porColumna.has(colKey)) {
          rangoFormato = sheet.getRangeByIndexes(
            selRange.rowIndex + 1,
            selRange.columnIndex,
            selRange.rowCount - 1,
            selRange.columnCount,
          );
          porColumna.set(colKey, { rangoFormato, reglas: [] });
        }
      }
      porColumna.get(colKey).reglas.push(regla);
    }

    // Aplicar reglas por columna (un clearAll + N reglas añadidas)
    for (const { rangoFormato, reglas: reglasCol } of porColumna.values()) {
      rangoFormato.conditionalFormats.clearAll();
      const cfs = rangoFormato.conditionalFormats;
      for (const regla of reglasCol) {
        _aplicarUnaRegla(cfs, regla);
      }
    }

    await context.sync();
  });
}

// ── Gráficos nativos Office.js ────────────────────────────────────────────────

const _CHART_TYPES = {
  barras:     "columnClustered",
  lineas:     "line",
  sectores:   "pie",
  dispersion: "xyscatter",
};

async function _insertarGrafico(tipoGrafico, datosChart, titulo, chartIndex = 0) {
  const filas = datosChart.length;
  const cols  = (datosChart[0] || []).length;
  if (filas < 2) throw new Error("No hay suficientes datos para crear el gráfico.");

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Calcular la primera fila libre debajo de todo el contenido de la hoja
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load(["rowIndex", "rowCount"]);
    await context.sync();
    const lastDataRow = usedRange.isNullObject ? 0 : (usedRange.rowIndex + usedRange.rowCount);

    // Escribir datos en una hoja temporal oculta
    const tempNombre = "_chart_temp_" + Date.now();
    const tempSheet  = context.workbook.worksheets.add(tempNombre);
    tempSheet.visibility = Excel.SheetVisibility.veryHidden;
    const dataRange = tempSheet.getRangeByIndexes(0, 0, filas, cols);
    dataRange.values = datosChart;
    await context.sync();

    // Crear el gráfico en la hoja activa con los datos de la hoja temporal
    const chartType = _CHART_TYPES[tipoGrafico] || "columnClustered";
    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.columns);

    chart.title.text    = titulo;
    chart.title.visible = true;
    chart.legend.visible = cols > 2;

    // Posicionar debajo del último contenido de la hoja + desplazamiento por índice
    const anchorRow = lastDataRow + 2 + chartIndex * 17;
    const anchorCol = 0;
    chart.setPosition(
      sheet.getRangeByIndexes(anchorRow, anchorCol, 1, 1),
      sheet.getRangeByIndexes(anchorRow + 15, anchorCol + 8, 1, 1),
    );

    await context.sync();
  });
}

// ── Tabla dinámica nativa Office.js ──────────────────────────────────────────

const _AGGFUNC_MAP = {
  suma:     "Sum",
  promedio: "Average",
  contar:   "Count",
  max:      "Max",
  min:      "Min",
};

async function _insertarTablaDinamica(params) {
  const filas    = params.filas    || [];
  const columnas = params.columnas || [];
  const valores  = params.valores;
  const funcion  = _AGGFUNC_MAP[params.funcion] || "Sum";

  if (!valores) throw new Error("Falta la columna de valores para la tabla dinámica.");
  if (!filas.length) throw new Error("Falta al menos una columna de filas para la tabla dinámica.");

  await Excel.run(async (context) => {
    const sheet      = context.workbook.worksheets.getActiveWorksheet();
    const localAddr  = _rangoAddress.includes("!") ? _rangoAddress.split("!")[1] : _rangoAddress;
    const sourceRange = sheet.getRange(localAddr);

    // Crear tabla dinámica en una hoja nueva
    const tdNombre = "TD_" + Date.now().toString().slice(-6);
    const tdSheet  = context.workbook.worksheets.add(tdNombre);
    tdSheet.activate();

    const pivotTable = tdSheet.pivotTables.add(
      tdNombre,
      sourceRange,
      tdSheet.getCell(1, 0),  // destino: B1 (deja una fila de margen)
    );
    await context.sync();

    // Añadir campos de fila
    for (const col of filas) {
      try {
        pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(col));
      } catch (_) { /* columna no encontrada — ignorar */ }
    }

    // Añadir campos de columna (opcional)
    for (const col of columnas) {
      try {
        pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(col));
      } catch (_) { /* ignorar */ }
    }

    // Añadir campo de valores con la función de agregación
    try {
      const dataHierarchy = pivotTable.dataHierarchies.add(
        pivotTable.hierarchies.getItem(valores),
      );
      dataHierarchy.summarizeBy = Excel.AggregationFunction[funcion];
    } catch (_) { /* ignorar */ }

    pivotTable.layout.layoutType = Excel.PivotLayoutType.tabular;

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

/** El usuario pulsó una opción de aclaración → rellenar el campo y enviar. */
function elegirAclaracion(opcion) {
  document.getElementById("pregunta").value = opcion;
  preguntar();
}

function _mostrarFeedback(pregunta, respuesta) {
  const zona = document.getElementById("zona-feedback");
  if (!zona) return;
  zona.dataset.pregunta  = pregunta;
  zona.dataset.respuesta = respuesta;
  zona.style.display = "flex";
  // Resetear a estado inicial: ambos botones visibles, icono oculto
  const btnPos = document.getElementById("btn-feedback-pos");
  const btnNeg = document.getElementById("btn-feedback-neg");
  const icon   = document.getElementById("feedback-icon");
  if (btnPos) btnPos.style.display = "";
  if (btnNeg) btnNeg.style.display = "";
  if (icon)   icon.style.display   = "none";
}

function _ocultarFeedback() {
  const zona = document.getElementById("zona-feedback");
  if (zona) zona.style.display = "none";
}

async function enviarFeedback(tipo) {
  const zona = document.getElementById("zona-feedback");
  if (!zona) return;
  const pregunta  = zona.dataset.pregunta  || "";
  const respuesta = zona.dataset.respuesta || "";
  if (!pregunta || !respuesta) return;

  const btnPos = document.getElementById("btn-feedback-pos");
  const btnNeg = document.getElementById("btn-feedback-neg");
  const icon   = document.getElementById("feedback-icon");
  // El icono muestra el voto dado; el botón contrario sigue visible para cambiar
  if (tipo === "positivo") {
    if (btnPos) btnPos.style.display = "none";
    if (btnNeg) btnNeg.style.display = "";
    if (icon)   { icon.textContent = "✅"; icon.style.display = "inline"; }
  } else {
    if (btnNeg) btnNeg.style.display = "none";
    if (btnPos) btnPos.style.display = "";
    if (icon)   { icon.textContent = "❌"; icon.style.display = "inline"; }
  }
  try {
    await llamarApi("/feedback", {
      device_id: _obtenerOCrearDeviceId(),
      pregunta,
      respuesta,
      tipo: tipo || "positivo",
    });
  } catch (_e) {
    // silencioso — el feedback es opcional
  }
}

function mostrarRespuesta(texto) {
  /* global marked, DOMPurify */
  document.getElementById("respuesta").innerHTML = DOMPurify.sanitize(marked.parse(texto));
  document.getElementById("bloque-respuesta").style.display = "block";
}

function ocultarRespuesta() {
  document.getElementById("bloque-respuesta").style.display = "none";
}

function mostrarDialogo(descripcion) {
  document.getElementById("desc-edicion").textContent = "Resultado: " + descripcion;

  // Operaciones que solo permiten insertar filas (al principio / al final de la tabla)
  const soloFilas = ["duplicar_filas", "añadir_fila_total"];
  const restringido = soloFilas.includes(_operacionActual);

  document.getElementById("btn-sustituir").style.display  = restringido ? "none" : "";
  document.getElementById("btn-derecha").style.display    = restringido ? "none" : "";
  document.getElementById("btn-nueva-hoja").style.display = restringido ? "none" : "";
  document.getElementById("btn-principio").style.display  = restringido ? "" : "none";
  // "Al final" siempre visible; para ops normales etiquetarlo "Debajo"
  const lblDebajo = document.querySelector("#btn-debajo .ms-Button-label");
  if (lblDebajo) lblDebajo.textContent = restringido ? "Al final" : "Debajo";

  document.getElementById("dialogo-destino").style.display = "block";
  document.getElementById("btn-debajo").focus();
}

function ocultarDialogo() {
  document.getElementById("dialogo-destino").style.display = "none";
  _operacionActual = null;
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
  localStorage.setItem(_CLAVE_HISTORIAL_LLM, JSON.stringify(_historialLLM));
}

function _cargarHistorialLLM() {
  try {
    const guardado = localStorage.getItem(_CLAVE_HISTORIAL_LLM);
    if (guardado) _historialLLM = JSON.parse(guardado);
  } catch (_e) {
    _historialLLM = [];
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
  localStorage.removeItem(_CLAVE_HISTORIAL_LLM);
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
