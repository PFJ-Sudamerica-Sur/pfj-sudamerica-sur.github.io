/**
 * Capacitación en línea PFJ 2027 · Panel de administración + receptor de respuestas
 * --------------------------------------------------------------------------------
 * Esta planilla de Google es el "panel admin" de la página:
 *   - Lo que escribas en las pestañas Videos, Fotos staff, Info sesiones, Estacas y Preguntas
 *     aparece en la página (al recargarla).
 *   - Lo que hacen los consejeros llega a las pestañas Registros, Avance, Repasos, Respuestas y Resumen.
 * Solo pueden editar las personas con las que compartas esta planilla.
 *
 * Cómo instalarlo (una sola vez):
 *  1. Crea una planilla nueva en Google Drive, por ejemplo "Panel capacitación PFJ 2027".
 *  2. Menú Extensiones > Apps Script. Borra lo que haya, pega este archivo completo y guarda (ícono de disquete).
 *  3. Vuelve a la planilla y recárgala. Aparece el menú "Capacitación PFJ" > "Preparar panel (primera vez)".
 *     Autoriza los permisos cuando los pida. Se crean todas las pestañas con el contenido actual.
 *  4. En Apps Script: Implementar > Nueva implementación > tipo "Aplicación web".
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *  5. Copia la URL que termina en /exec y entrégala para pegarla en app.js (CONFIG.ENDPOINT).
 *  Si después cambias este código: Implementar > Gestionar implementaciones > editar (lápiz) >
 *  Versión: "Nueva versión" > Implementar. Así la URL no cambia.
 */

const SITIO = "https://pfj-sudamerica-sur.github.io/capacitacion-online/";
const TOTAL_SECCIONES = { C1: 10, C2: 0, C3: 0 }; // completar cuando existan la 2ª y 3ª

const PANEL = {
  "Videos":        ["Clave (no cambiar)", "Módulo", "Dónde aparece", "Título", "Enlace del video (YouTube o Drive)", "Solo para la sesión (vacío = todas)"],
  "Fotos staff":   ["Sesión", "Cargo", "Nombre", "Enlace de la foto (Drive)"],
  "Info sesiones": ["País", "Sesión", "Fecha de la sesión recuperativa", "Lugar o enlace de Meet/Zoom", "Enlace grupo de WhatsApp", "Enlace Classroom"],
  "Estacas":       ["País", "Sesión", "Estaca"],
  "Preguntas":     ["Módulo", "Pregunta", "Opción A", "Opción B", "Opción C", "Opción D", "Correcta (A, B, C o D)", "Explicación"]
};
const RESPUESTAS = {
  Registros:  ["Fecha", "Capacitación", "Nombre", "Correo", "País", "Sesión", "Estaca", "Estaca escrita a mano", "Modalidad"],
  Avance:     ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Estaca", "Sección", "Completado"],
  Repasos:    ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Estaca", "Módulo", "Correctas", "Total"],
  Respuestas: ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Estaca", "Módulo", "Pregunta", "Respuesta elegida", "¿Correcta?"],
  Resumen:    ["Sesión", "Estaca", "Nombre", "Correo", "Modalidad", "Capacitación", "Secciones completadas", "Total secciones",
               "% avance", "Repasos (correctas/total)", "Última actividad", "Estado"]
};
const LETRAS = ["A", "B", "C", "D"];

/* ================= Menú ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Capacitación PFJ")
    .addItem("Preparar panel (primera vez)", "prepararPanel")
    .addItem("Actualizar resumen", "actualizarResumen")
    .addToUi();
}

/** Crea las pestañas del panel con el contenido que hoy tiene la página (datos.json). */
function prepararPanel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existentes = Object.keys(PANEL).filter((n) => ss.getSheetByName(n));
  if (existentes.length && ui.alert("Ya existen las pestañas: " + existentes.join(", ") +
      ".\nSi continúas, esas pestañas no se tocan y solo se crean las que faltan.", ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const d = JSON.parse(UrlFetchApp.fetch(SITIO + "datos.json", { muteHttpExceptions: true }).getContentText());

  instrucciones_(ss);

  const filas = {
    "Videos": d.videos.map((v) => [v.clave, v.modulo, v.lugar, v.titulo, v.enlace, v.sesion]),
    "Fotos staff": [],
    "Info sesiones": [],
    "Estacas": [],
    "Preguntas": d.preguntas.map((q) => [q.modulo, q.pregunta, q.opciones[0] || "", q.opciones[1] || "", q.opciones[2] || "", q.opciones[3] || "",
                                         LETRAS[q.correcta], q.explicacion || ""])
  };
  Object.entries(d.estacas).forEach(([pais, ses]) => Object.entries(ses).forEach(([s, lista]) => {
    filas["Info sesiones"].push([pais, s, "", "", "", ""]);
    lista.forEach((e) => filas["Estacas"].push([pais, s, e]));
  }));

  Object.entries(PANEL).forEach(([nombre, enc]) => {
    if (ss.getSheetByName(nombre)) return;
    const h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight("bold").setBackground("#fff2cc");
    h.setFrozenRows(1);
    if (filas[nombre].length) h.getRange(2, 1, filas[nombre].length, enc.length).setValues(filas[nombre]);
    h.autoResizeColumns(1, enc.length);
  });
  // Lista desplegable para la respuesta correcta
  const p = ss.getSheetByName("Preguntas");
  p.getRange("G2:G500").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LETRAS, true).build());
  ss.getSheetByName("Videos").getRange("A:A").setBackground("#eeeeee");

  Object.keys(RESPUESTAS).forEach(hojaRespuestas_);
  const hoja1 = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (hoja1 && hoja1.getLastRow() === 0) ss.deleteSheet(hoja1);
  ss.setActiveSheet(ss.getSheetByName("Instrucciones"));
  ui.alert("Panel listo. Revisa la pestaña Instrucciones.");
}

function instrucciones_(ss) {
  if (ss.getSheetByName("Instrucciones")) return;
  const h = ss.insertSheet("Instrucciones", 0);
  const t = [
    ["PANEL DE LA CAPACITACIÓN EN LÍNEA PFJ 2027"],
    [""],
    ["Lo que cambies aquí aparece en la página al recargarla. Solo pueden editar las personas con quienes compartas esta planilla."],
    [""],
    ["VIDEOS: sube el video a YouTube como \"No listado\" (o a Google Drive, compartido con \"Cualquier persona con el enlace\")."],
    ["Copia el enlace y pégalo en la pestaña Videos, en la fila que corresponde. No cambies la columna Clave."],
    ["Si un video es distinto para cada sesión (por ejemplo, el mensaje final), copia la fila, pega el enlace y escribe la sesión en la última columna."],
    [""],
    ["FOTOS DEL STAFF: sube la foto a Google Drive, compártela con \"Cualquier persona con el enlace\" y pega el enlace en Fotos staff, con la sesión, el cargo y el nombre."],
    [""],
    ["INFO SESIONES: fecha y lugar de la sesión recuperativa, y enlaces de WhatsApp y Classroom de cada sesión. Cada consejero ve solo los de su sesión."],
    [""],
    ["ESTACAS: una fila por estaca. El nombre de la sesión debe ser igual en Estacas, Info sesiones, Fotos staff y Videos."],
    [""],
    ["PREGUNTAS: solo de alternativas. Escribe 2 a 4 opciones y en \"Correcta\" la letra de la buena. Para quitar una pregunta, borra la fila."],
    [""],
    ["RESPUESTAS DE LOS CONSEJEROS: llegan solas a Registros, Avance, Repasos y Respuestas. Resumen muestra una fila por consejero."],
    ["No compartas esta planilla con los consejeros: contiene nombres y correos."]
  ];
  h.getRange(1, 1, t.length, 1).setValues(t);
  h.getRange("A1").setFontWeight("bold").setFontSize(14);
  h.setColumnWidth(1, 900);
  h.getRange(1, 1, t.length, 1).setWrap(true);
}

/* ================= La página pide el contenido ================= */
function doGet(e) {
  const salida = { ok: true, datos: leerContenido_() };
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

function leerContenido_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filas = (n) => {
    const h = ss.getSheetByName(n);
    if (!h || h.getLastRow() < 2) return [];
    return h.getRange(2, 1, h.getLastRow() - 1, PANEL[n].length).getDisplayValues()
      .map((r) => r.map((c) => String(c).trim())).filter((r) => r.some(Boolean));
  };
  const datos = {};

  const est = filas("Estacas");
  if (est.length) {
    datos.estacas = {};
    est.forEach(([pais, ses, estaca]) => {
      if (!pais || !ses || !estaca) return;
      datos.estacas[pais] = datos.estacas[pais] || {};
      datos.estacas[pais][ses] = datos.estacas[pais][ses] || [];
      datos.estacas[pais][ses].push(estaca);
    });
  }
  const vid = filas("Videos");
  if (vid.length) datos.videos = vid.map(([clave, modulo, lugar, titulo, enlace, sesion]) => ({ clave, modulo, lugar, titulo, enlace, sesion }));

  datos.fotos = filas("Fotos staff").map(([sesion, cargo, nombre, foto]) => ({ sesion, cargo, nombre, foto }));
  datos.infoSesiones = filas("Info sesiones").map(([pais, sesion, recuperativa, lugar, whatsapp, classroom]) => ({ pais, sesion, recuperativa, lugar, whatsapp, classroom }));

  const pre = filas("Preguntas");
  if (pre.length) {
    datos.preguntas = pre.map(([modulo, pregunta, a, b, c, d, correcta, explicacion]) => {
      const todas = [a, b, c, d];
      const idx = LETRAS.indexOf(String(correcta).toUpperCase());
      const opciones = todas.filter(Boolean);
      return { modulo, pregunta, opciones, correcta: idx >= 0 && todas[idx] ? opciones.indexOf(todas[idx]) : 0, explicacion };
    }).filter((q) => q.pregunta && q.opciones.length >= 2);
  }
  return datos;
}

/* ================= La página envía avances y respuestas ================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fecha = new Date(d.fecha || Date.now());
    const base = [fecha, d.capacitacion, d.nombre, (d.correo || "").toLowerCase(), d.sesion, d.estaca].map(limpiar_);
    const agregar = (hoja, extra) => hojaRespuestas_(hoja).appendRow(base.concat(extra.map(limpiar_)));
    if (d.tipo === "registro") hojaRespuestas_("Registros").appendRow([fecha, d.capacitacion, d.nombre, (d.correo || "").toLowerCase(),
                                 d.pais, d.sesion, d.estaca, d.estacaEscrita ? "Sí" : "No", d.modalidad].map(limpiar_));
    if (d.tipo === "seccion")   agregar("Avance", [d.seccion, d.completado ? "Sí" : "No"]);
    if (d.tipo === "repaso")    agregar("Repasos", [d.modulo, d.puntaje, d.total]);
    if (d.tipo === "respuesta") agregar("Respuestas", [d.modulo, d.pregunta, d.respuesta, d.correcta ? "Sí" : "No"]);
    if (d.tipo === "registro" || d.tipo === "seccion" || d.tipo === "repaso") actualizarResumen();
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

/** Evita que un texto se interprete como fórmula y corta textos muy largos. */
function limpiar_(v) {
  if (v instanceof Date || typeof v === "number" || typeof v === "boolean") return v;
  let s = String(v == null ? "" : v).slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/** Reconstruye la hoja Resumen: una fila por consejero y capacitación. */
function actualizarResumen() {
  const personas = {};
  const persona = (correo, cap, fecha) => {
    const k = correo + "|" + cap;
    personas[k] = personas[k] || { cap, correo, nombre: "", sesion: "", estaca: "", modalidad: "", secciones: {}, repasos: {}, ultima: fecha };
    const p = personas[k];
    if (fecha > p.ultima) p.ultima = fecha;
    return p;
  };
  filas_("Registros").forEach((r) => { const p = persona(r[3], r[1], r[0]); p.nombre = r[2]; p.sesion = r[5]; p.estaca = r[6]; p.modalidad = r[8]; });
  filas_("Avance").forEach((r) => { const p = persona(r[3], r[1], r[0]); p.nombre = p.nombre || r[2]; p.secciones[r[6]] = r[7] === "Sí"; });
  filas_("Repasos").forEach((r) => { const p = persona(r[3], r[1], r[0]); p.nombre = p.nombre || r[2]; p.repasos[r[6]] = [r[7], r[8]]; });

  const salida = Object.values(personas).map((p) => {
    const hechas = Object.values(p.secciones).filter(Boolean).length;
    const total = TOTAL_SECCIONES[p.cap] || hechas;
    let ok = 0, tot = 0;
    Object.values(p.repasos).forEach(([a, b]) => { ok += Number(a); tot += Number(b); });
    const pct = total ? hechas / total : 0;
    const estado = pct >= 1 ? "Completó la parte en línea" : (hechas ? "En curso" : "Solo registrado");
    return [p.sesion, p.estaca, p.nombre, p.correo, p.modalidad, p.cap, hechas, total, pct, tot ? ok + "/" + tot : "", p.ultima, estado];
  }).sort((a, b) => (a[0] + a[1] + a[2]).localeCompare(b[0] + b[1] + b[2]));

  const h = hojaRespuestas_("Resumen");
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, RESPUESTAS.Resumen.length).clearContent();
  if (salida.length) {
    h.getRange(2, 1, salida.length, RESPUESTAS.Resumen.length).setValues(salida);
    h.getRange(2, 9, salida.length, 1).setNumberFormat("0%");
  }
}

function hojaRespuestas_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.appendRow(RESPUESTAS[nombre]);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, RESPUESTAS[nombre].length).setFontWeight("bold").setBackground("#d9ead3");
  }
  return h;
}

function filas_(nombre) {
  const h = hojaRespuestas_(nombre);
  return h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, RESPUESTAS[nombre].length).getValues() : [];
}
