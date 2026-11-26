/**
 * Capacitación en línea PFJ 2027 · Panel de administración + receptor de respuestas
 * --------------------------------------------------------------------------------
 * Esta planilla de Google es el "panel admin" de la página:
 *   - Lo que escribas en las pestañas Videos, Fotos staff, Info sesiones, Estacas y Preguntas
 *     aparece en la página (al recargarla).
 *   - Cuando un consejero presiona "Finalizar y enviar", su capacitación completa llega a
 *     Informe (una fila por consejero) y Respuestas (el detalle). Solo este panel guarda esos datos.
 *   - En la pestaña Coordinadores se anotan los correos de cada sesión. El menú
 *     "Crear o actualizar planillas por sesión" crea una planilla por sesión y la comparte solo con ellos.
 *     Cada planilla de sesión tiene dos hojas que se rehacen solas en cada envío: Resumen (cantidades
 *     por capacitación y por estaca) y Consejeros (quién terminó cada capacitación). Sin correos,
 *     respuestas ni puntajes.
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

// Se lee desde el repositorio (no desde la página publicada) para que funcione aunque GitHub Pages esté lento.
const DATOS_URL = "https://raw.githubusercontent.com/PFJ-Sudamerica-Sur/pfj-sudamerica-sur.github.io/main/capacitacion-online/datos.json";
const CARPETA_SESIONES = "Capacitación PFJ 2027 · Planillas por sesión";

const PANEL = {
  "Videos":        ["Clave (no cambiar)", "Módulo", "Dónde aparece", "Título", "Enlace del video (YouTube o Drive)", "Solo para la sesión (vacío = todas)"],
  "Fotos staff":   ["Sesión", "Cargo", "Nombre", "Enlace de la foto (Drive)"],
  "Info sesiones": ["País", "Sesión", "Fecha de la sesión recuperativa", "Lugar o enlace de Meet/Zoom", "Enlace grupo de WhatsApp", "Enlace Classroom (solo si tu sesión lo usa)"],
  "Estacas":       ["País", "Sesión", "Estaca"],
  "Preguntas":     ["Módulo", "Pregunta", "Opción A", "Opción B", "Opción C", "Opción D", "Correcta (A, B, C o D)", "Explicación"],
  "Coordinadores": ["País", "Sesión", "Correos de los coordinadores (separados por coma)", "Planilla de la sesión (se llena sola)"]
};
const RESPUESTAS = {
  Informe:    ["Fecha de envío", "Capacitación", "Nombre", "Correo", "País", "Sesión", "Estaca", "Estaca escrita a mano", "Modalidad",
               "Repaso módulo 1", "Repaso módulo 2", "Repaso módulo 3", "Repaso módulo 4", "Total correctas", "Total preguntas", "% correctas", "Bailes practicados"],
  Respuestas: ["Fecha de envío", "Capacitación", "Nombre", "Correo", "Sesión", "Estaca", "Módulo", "Pregunta", "Respuesta elegida", "¿Correcta?"]
};
const LETRAS = ["A", "B", "C", "D"];

/* ================= Menú ================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Capacitación PFJ")
    .addItem("Preparar panel (primera vez)", "prepararPanel")
    .addSeparator()
    .addItem("Crear planillas de las filas seleccionadas (sin compartir)", "crearPlanillasSeleccion")
    .addItem("Compartir las planillas de las filas seleccionadas", "compartirPlanillasSeleccion")
    .addItem("Crear y compartir planillas de TODAS las sesiones", "crearPlanillasSesion")
    .addSeparator()
    .addItem("Rehacer la hoja Resumen de este panel", "resumenMaestra")
    .addToUi();
}

/** Crea las pestañas del panel con el contenido que hoy tiene la página (datos.json). */
function prepararPanel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existentes = Object.keys(PANEL).filter((n) => ss.getSheetByName(n));
  if (existentes.length && ui.alert("Ya existen las pestañas: " + existentes.join(", ") +
      ".\nSi continúas, esas pestañas no se tocan y solo se crean las que faltan.", ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const d = JSON.parse(UrlFetchApp.fetch(DATOS_URL, { muteHttpExceptions: true }).getContentText());

  instrucciones_(ss);

  const filas = {
    "Videos": d.videos.map((v) => [v.clave, v.modulo, v.lugar, v.titulo, v.enlace, v.sesion]),
    "Fotos staff": [],
    "Info sesiones": [],
    "Estacas": [],
    "Coordinadores": [],
    "Preguntas": d.preguntas.map((q) => [q.modulo, q.pregunta, q.opciones[0] || "", q.opciones[1] || "", q.opciones[2] || "", q.opciones[3] || "",
                                         LETRAS[q.correcta], q.explicacion || ""])
  };
  Object.entries(d.estacas).forEach(([pais, ses]) => Object.entries(ses).forEach(([s, lista]) => {
    filas["Info sesiones"].push([pais, s, "", "", "", ""]);
    filas["Coordinadores"].push([pais, s, "", ""]);
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
  ss.getSheetByName("Preguntas").getRange("G2:G500")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LETRAS, true).build());
  ss.getSheetByName("Videos").getRange("A:A").setBackground("#eeeeee");
  ss.getSheetByName("Coordinadores").getRange("D:D").setBackground("#eeeeee");

  Object.keys(RESPUESTAS).forEach((n) => hoja_(n));
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
    ["Por defecto todos ven los videos del Área. Si una sesión graba su propio video (pensamiento espiritual o mensaje final), copia la fila, pega su enlace y escribe el nombre de la sesión en la última columna: solo esa sesión lo verá."],
    ["La línea de autoridad la graba el Setenta de Área y es igual para todos: no agregues versiones por sesión."],
    [""],
    ["FOTOS DEL STAFF: sube la foto a Google Drive, compártela con \"Cualquier persona con el enlace\" y pega el enlace en Fotos staff, con la sesión, el cargo y el nombre."],
    [""],
    ["INFO SESIONES: fecha y lugar de la sesión recuperativa, y enlace del grupo de WhatsApp de cada sesión (Classroom solo si la sesión lo usa). Cada consejero ve solo los de su sesión."],
    [""],
    ["ESTACAS: una fila por estaca. El nombre de la sesión debe ser igual en Estacas, Info sesiones, Fotos staff, Videos y Coordinadores."],
    [""],
    ["PREGUNTAS: solo de alternativas. Escribe 2 a 4 opciones y en \"Correcta\" la letra de la buena. Para quitar una pregunta, borra la fila."],
    [""],
    ["RESPUESTAS DE LOS CONSEJEROS: solo llegan cuando el consejero termina y presiona \"Finalizar y enviar\". Informe tiene una fila por consejero con sus puntajes; Respuestas tiene cada pregunta respondida."],
    ["No compartas esta planilla con los consejeros ni con las sesiones: contiene los datos de todas."],
    [""],
    ["PLANILLAS POR SESIÓN: en la pestaña Coordinadores escribe los correos de los coordinadores de cada sesión, separados por coma."],
    ["Luego usa el menú Capacitación PFJ > Crear o actualizar planillas por sesión. Se crea una planilla para cada sesión con correos (en la carpeta \"" + CARPETA_SESIONES + "\" de tu Drive),"],
    ["se comparte solo con esos correos (pueden verla, no editarla) y desde entonces recibe sola los informes de su sesión."],
    ["Si cambian los coordinadores, corrige los correos y vuelve a usar el mismo menú: se agregan los nuevos y se quita el acceso a quienes ya no están."]
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

function filasPanel_(nombre) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, PANEL[nombre].length).getDisplayValues()
    .map((r) => r.map((c) => String(c).trim())).filter((r) => r.some(Boolean));
}

function leerContenido_() {
  const datos = {};

  const est = filasPanel_("Estacas");
  if (est.length) {
    datos.estacas = {};
    est.forEach(([pais, ses, estaca]) => {
      if (!pais || !ses || !estaca) return;
      datos.estacas[pais] = datos.estacas[pais] || {};
      datos.estacas[pais][ses] = datos.estacas[pais][ses] || [];
      datos.estacas[pais][ses].push(estaca);
    });
  }
  const vid = filasPanel_("Videos");
  if (vid.length) datos.videos = vid.map(([clave, modulo, lugar, titulo, enlace, sesion]) => ({ clave, modulo, lugar, titulo, enlace, sesion }));

  datos.fotos = filasPanel_("Fotos staff").map(([sesion, cargo, nombre, foto]) => ({ sesion, cargo, nombre, foto }));
  datos.infoSesiones = filasPanel_("Info sesiones").map(([pais, sesion, recuperativa, lugar, whatsapp, classroom]) => ({ pais, sesion, recuperativa, lugar, whatsapp, classroom }));

  const pre = filasPanel_("Preguntas");
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

/* ================= La página envía la capacitación terminada ================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    // Solo se guardan capacitaciones terminadas: la página envía todo junto al presionar "Finalizar y enviar".
    if (d.tipo !== "final") return ContentService.createTextOutput("ignorado");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { informe, respuestas } = filasDeEnvio_(d);

    // 1) Planilla maestra (todas las sesiones): aquí quedan todos los datos
    agregarFilas_(ss, informe, respuestas);

    // 2) Planilla de la sesión y resumen del panel. Si algo falla, el envío igual quedó guardado.
    try { actualizarPlanillaSesion_(d.pais, d.sesion); } catch (err) { console.error("Planilla de " + d.sesion + ": " + err); }
    try { armarInformes_(ss, null); } catch (err) { console.error("Resumen del panel: " + err); }
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function filasDeEnvio_(d) {
  const fecha = new Date(d.fecha || Date.now());
  const correo = String(d.correo || "").toLowerCase();
  const mod = (n) => (d.modulos || []).find((m) => String(m.n) === String(n)) || {};
  const pm = (n) => { const m = mod(n); return m.total !== "" && m.total != null ? m.puntaje + " de " + m.total : ""; }; // "4/4" lo convertiría en fecha
  const pct = Number(d.totalPreguntas) ? Number(d.correctas) / Number(d.totalPreguntas) : "";
  const informe = [fecha, d.capacitacion, d.nombre, correo, d.pais, d.sesion, d.estaca, d.estacaEscrita ? "Sí" : "No",
    d.modalidad, pm(1), pm(2), pm(3), pm(4), d.correctas, d.totalPreguntas, pct, d.bailes].map(limpiar_);
  const respuestas = (d.respuestas || []).map((q) => [fecha, d.capacitacion, d.nombre, correo, d.sesion, d.estaca,
    q.modulo, q.pregunta, q.respuesta, q.correcta ? "Sí" : "No"].map(limpiar_));
  return { informe, respuestas };
}

function agregarFilas_(ss, informe, respuestas) {
  const h = hoja_("Informe", ss);
  h.appendRow(informe);
  h.getRange(h.getLastRow(), 16).setNumberFormat("0%");
  if (respuestas.length) {
    const r = hoja_("Respuestas", ss);
    r.getRange(r.getLastRow() + 1, 1, respuestas.length, respuestas[0].length).setValues(respuestas);
  }
}

/** Busca en la pestaña Coordinadores la fila de esa sesión: [fila, país, sesión, correos, url]. */
function filaCoordinadores_(pais, sesion) {
  const hc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinadores");
  if (!hc || hc.getLastRow() < 2) return null;
  const datos = hc.getRange(2, 1, hc.getLastRow() - 1, 4).getValues();
  const i = datos.findIndex((r) => String(r[1]).trim() === String(sesion || "").trim() && (!pais || !String(r[0]).trim() || String(r[0]).trim() === pais));
  return i < 0 ? null : { fila: i + 2, pais: String(datos[i][0]).trim(), sesion: String(datos[i][1]).trim(), correos: correosDe_(datos[i][2]), url: String(datos[i][3]).trim() };
}

/** Rehace el Resumen y la lista de Consejeros de la planilla de una sesión, si ya fue creada. */
function actualizarPlanillaSesion_(pais, sesion) {
  const f = filaCoordinadores_(pais, sesion);
  if (!f || !/^https:\/\//.test(f.url)) return;
  armarInformes_(SpreadsheetApp.openByUrl(f.url), { pais: f.pais, sesion: f.sesion });
}

/* ================= Planillas por sesión ================= */
/** Menú: todas las sesiones con correos, creando, actualizando y compartiendo. */
function crearPlanillasSesion() {
  const hc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coordinadores");
  if (!hc) { SpreadsheetApp.getUi().alert("Falta la pestaña Coordinadores. Usa primero \"Preparar panel (primera vez)\"."); return; }
  const filas = [];
  for (let f = 2; f <= hc.getLastRow(); f++) filas.push(f);
  procesarFilas_(filas, true);
}

/** Menú: solo las filas seleccionadas en Coordinadores, sin compartir (para revisar antes). */
function crearPlanillasSeleccion() {
  const filas = filasSeleccionadas_();
  if (filas) procesarFilas_(filas, false);
}

/** Menú: solo las filas seleccionadas en Coordinadores, compartiéndolas con sus coordinadores. */
function compartirPlanillasSeleccion() {
  const filas = filasSeleccionadas_();
  if (!filas) return;
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("Se compartirán " + filas.length + " planilla(s) con los correos de esas filas. Google les enviará un correo avisándoles. ¿Continuar?",
      ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  procesarFilas_(filas, true);
}

/** Números de fila seleccionados en la pestaña Coordinadores (puede ser un bloque o varias selecciones con Ctrl). */
function filasSeleccionadas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const rangos = hoja.getActiveRangeList() ? hoja.getActiveRangeList().getRanges() : [];
  const filas = [];
  rangos.forEach((r) => { for (let f = r.getRow(); f < r.getRow() + r.getNumRows(); f++) if (f >= 2 && !filas.includes(f)) filas.push(f); });
  if (hoja.getName() !== "Coordinadores" || !filas.length) {
    SpreadsheetApp.getUi().alert("Primero ve a la pestaña Coordinadores y selecciona las filas de las sesiones (por ejemplo, arrastrando sobre los números de fila de la izquierda).");
    return null;
  }
  return filas.sort((a, b) => a - b);
}

/** Crea (si falta) y actualiza la planilla de cada fila; si compartir = true, la comparte solo con sus coordinadores. */
function procesarFilas_(filas, compartir) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hc = ss.getSheetByName("Coordinadores");
  const yo = Session.getEffectiveUser().getEmail().toLowerCase();
  const resumen = [];

  filas.forEach((fila) => {
    const r = hc.getRange(fila, 1, 1, 4).getValues()[0];
    const pais = String(r[0]).trim(), sesion = String(r[1]).trim();
    const correos = correosDe_(r[2]);
    const url = String(r[3]).trim();
    if (!sesion) return;
    if (compartir && !correos.length && !url) { resumen.push(sesion + ": sin correos, no se creó"); return; }

    const planilla = url ? SpreadsheetApp.openByUrl(url) : nuevaPlanillaSesion_(hc, fila, pais, sesion);
    armarInformes_(planilla, { pais, sesion });

    let estado = "sin compartir";
    if (compartir) {
      // Compartir solo con los coordinadores de la lista (solo lectura) y quitar a quienes ya no están
      const archivo = DriveApp.getFileById(planilla.getId());
      const actuales = archivo.getViewers().map((u) => u.getEmail().toLowerCase());
      correos.filter((c) => !actuales.includes(c)).forEach((c) => archivo.addViewer(c));
      actuales.filter((c) => !correos.includes(c) && c !== yo).forEach((c) => archivo.removeViewer(c));
      estado = correos.length ? "compartida con " + correos.length + " coordinador(es)" : "sin correos";
    }
    resumen.push(sesion + ": " + estado + (url ? "" : " (nueva)"));
  });

  armarInformes_(ss, null);
  ss.setActiveSheet(hc);
  ui.alert(resumen.length ? "Listo:\n\n" + resumen.join("\n") + "\n\nEl enlace de cada planilla quedó en la columna D." : "No había sesiones en esas filas.");
}

/** Crea la planilla de una sesión en la carpeta de sesiones y anota su URL en Coordinadores. */
function nuevaPlanillaSesion_(hc, fila, pais, sesion) {
  const planilla = SpreadsheetApp.create("Capacitación PFJ 2027 · " + sesion + (pais ? " (" + pais + ")" : ""));
  DriveApp.getFileById(planilla.getId()).moveTo(carpetaSesiones_());
  hc.getRange(fila, 4).setValue(planilla.getUrl());
  return planilla;
}

/* ================= Informes: Resumen + Consejeros ================= */
// Colores de la página: 1ª amarillo, 2ª azul, 3ª verde (como las portadas de cada capacitación)
const CAPS = [
  { c: "C1", nombre: "Primera capacitación", corto: "1ª", fuerte: "#f9c80e", textoFuerte: "#333333", suave: "#fdefc2", textoSuave: "#5c4300" },
  { c: "C2", nombre: "Segunda capacitación", corto: "2ª", fuerte: "#0b6f9a", textoFuerte: "#ffffff", suave: "#d6ecf5", textoSuave: "#0b3f66" },
  { c: "C3", nombre: "Tercera capacitación", corto: "3ª", fuerte: "#0f8f45", textoFuerte: "#ffffff", suave: "#d9efdc", textoSuave: "#0b5a2c" }
];
const COLOR = { titulo: "#00627f", azul: "#0b4f77", encabezado: "#e8f1f5", linea: "#d5dadc", gris: "#6b7477", amarillo: "#f9c80e" };
const modalidadCorta_ = (m) => (/repaso/i.test(m) ? "Como repaso" : /lugar/i.test(m) ? "En lugar de la presencial" : String(m || ""));

/** Envíos de la planilla maestra, uno por consejero y capacitación (si alguien envió dos veces, vale el último). */
function envios_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Informe");
  if (!h || h.getLastRow() < 2) return [];
  const porClave = {};
  h.getRange(2, 1, h.getLastRow() - 1, RESPUESTAS.Informe.length).getValues().forEach((r) => {
    const [fecha, cap, nombre, correo, pais, sesion, estaca, , modalidad] = r;
    if (!cap || !nombre) return;
    const e = { fecha: fecha instanceof Date ? fecha : new Date(fecha), cap: String(cap), nombre: String(nombre).trim(),
      correo: String(correo).trim().toLowerCase(), pais: String(pais).trim(), sesion: String(sesion).trim(),
      estaca: String(estaca).trim(), modalidad: modalidadCorta_(modalidad) };
    const clave = (e.correo || e.nombre.toLowerCase()) + "|" + e.cap;
    if (!porClave[clave] || porClave[clave].fecha <= e.fecha) porClave[clave] = e;
  });
  return Object.values(porClave);
}

/**
 * Arma las hojas "Resumen" (solo cantidades) y "Consejeros" (una tabla por capacitación).
 * filtro = { pais, sesion } para la planilla de una sesión; null para el panel con todas las sesiones.
 * No incluye correos, respuestas ni puntajes.
 */
function armarInformes_(ss, filtro) {
  const maestra = !filtro;
  let env = envios_();
  if (!maestra) env = env.filter((e) => e.sesion === filtro.sesion && (!filtro.pais || !e.pais || e.pais === filtro.pais));
  const titulo = maestra ? "Capacitación PFJ 2027 · Todas las sesiones" : "Capacitación PFJ 2027 · " + filtro.sesion + (filtro.pais ? " (" + filtro.pais + ")" : "");
  const actualizado = "Se actualiza sola cada vez que un consejero presiona \"Finalizar y enviar\". Última actualización: " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const nombreSesion = (e) => e.sesion + (e.pais ? " (" + e.pais + ")" : "");

  /* ---------- Resumen ---------- */
  const hr = hojaInforme_(ss, "Resumen", 0);
  const AN = 4;
  let f = encabezadoHoja_(hr, titulo, actualizado, AN);

  f = seccion_(hr, f, "Consejeros que terminaron la parte en línea", AN);
  columnas_(hr, f++, ["Capacitación", "Total", "En lugar de la presencial", "Como repaso"]);
  CAPS.forEach((k) => {
    const de = env.filter((e) => e.cap === k.c);
    hr.getRange(f, 1, 1, AN).setValues([[k.nombre, de.length, de.filter((e) => e.modalidad === "En lugar de la presencial").length, de.filter((e) => e.modalidad === "Como repaso").length]]);
    hr.getRange(f, 1).setBackground(k.suave).setFontColor(k.textoSuave).setFontWeight("bold");
    filaDatos_(hr, f++, AN);
  });
  hr.getRange(f++, 1).setValue("\"En lugar de la presencial\" = deben asistir a la sesión recuperativa.").setFontColor(COLOR.gris).setFontStyle("italic").setFontSize(9);
  f++;

  // Por estaca (planilla de sesión) o por sesión (panel). Se listan también las que aún no tienen a nadie.
  const claves = [];
  if (maestra) filasPanel_("Coordinadores").forEach(([p, s]) => { if (s) claves.push(s + (p ? " (" + p + ")" : "")); });
  else filasPanel_("Estacas").forEach(([p, s, e]) => { if (s === filtro.sesion && (!filtro.pais || !p || p === filtro.pais) && e) claves.push(e); });
  const claveDe = (e) => (maestra ? nombreSesion(e) : e.estaca);
  env.forEach((e) => { if (!claves.includes(claveDe(e))) claves.push(claveDe(e)); });
  const filasClave = claves.map((k) => [k, ...CAPS.map((c) => env.filter((e) => e.cap === c.c && claveDe(e) === k).length)]);
  filasClave.sort((a, b) => (b[1] + b[2] + b[3]) - (a[1] + a[2] + a[3]) || String(a[0]).localeCompare(String(b[0]), "es"));

  f = seccion_(hr, f, maestra ? "Por sesión" : "Por estaca", AN);
  columnas_(hr, f++, [maestra ? "Sesión" : "Estaca", "1ª capacitación", "2ª capacitación", "3ª capacitación"]);
  CAPS.forEach((k, i) => hr.getRange(f - 1, i + 2).setBackground(k.suave).setFontColor(k.textoSuave));
  if (filasClave.length) {
    hr.getRange(f, 1, filasClave.length, AN).setValues(filasClave.map((r) => r.map(limpiar_)));
    for (let i = 0; i < filasClave.length; i++) {
      filaDatos_(hr, f + i, AN);
      if (!(filasClave[i][1] + filasClave[i][2] + filasClave[i][3])) hr.getRange(f + i, 1, 1, AN).setFontColor(COLOR.gris);
    }
    f += filasClave.length;
  }
  hr.getRange(f, 1, 1, AN).setValues([["Total", ...CAPS.map((c) => env.filter((e) => e.cap === c.c).length)]])
    .setFontWeight("bold").setBorder(true, null, null, null, null, null, COLOR.azul, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  hr.getRange(5, 2, f - 4, AN - 1).setHorizontalAlignment("center");
  hr.setColumnWidth(1, 360); hr.setColumnWidths(2, AN - 1, 170);
  terminarHoja_(hr);

  /* ---------- Consejeros: una tabla por capacitación ---------- */
  const hc = hojaInforme_(ss, "Consejeros", 1);
  const enc = [...(maestra ? ["Sesión"] : []), "Nombre", "Estaca", "Modalidad", "Último envío"];
  const AC = enc.length;
  f = encabezadoHoja_(hc, maestra ? "Consejeros · Todas las sesiones" : "Consejeros · " + filtro.sesion, actualizado, AC);

  CAPS.forEach((k) => {
    const lista = env.filter((e) => e.cap === k.c)
      .sort((a, b) => (maestra ? nombreSesion(a).localeCompare(nombreSesion(b), "es") : 0) || a.estaca.localeCompare(b.estaca, "es") ||
        a.nombre.localeCompare(b.nombre, "es"));
    // Franja con el color de la capacitación
    hc.getRange(f, 1, 1, AC).setBackground(k.fuerte).setFontColor(k.textoFuerte).setFontWeight("bold").setFontSize(12);
    hc.getRange(f, 1).setValue(k.nombre.toUpperCase());
    hc.getRange(f, AC).setValue(lista.length === 1 ? "1 consejero" : lista.length + " consejeros").setHorizontalAlignment("right");
    hc.setRowHeight(f++, 30);
    if (!lista.length) {
      hc.getRange(f++, 1).setValue("Todavía nadie ha terminado esta capacitación.").setFontColor(COLOR.gris).setFontStyle("italic");
    } else {
      columnas_(hc, f++, enc);
      hc.getRange(f - 1, 1, 1, AC).setBackground(k.suave).setFontColor(k.textoSuave);
      const filas = lista.map((e) => [...(maestra ? [nombreSesion(e)] : []), e.nombre, e.estaca, e.modalidad, e.fecha].map(limpiar_));
      hc.getRange(f, 1, filas.length, AC).setValues(filas);
      hc.getRange(f, AC, filas.length, 1).setNumberFormat("dd/mm/yyyy hh:mm").setHorizontalAlignment("center");
      for (let i = 0; i < filas.length; i++) filaDatos_(hc, f + i, AC);
      f += filas.length;
    }
    f += 2;
  });
  hc.setColumnWidths(1, AC, 200);
  hc.setColumnWidth(AC, 140);
  terminarHoja_(hc);

  // Las planillas de sesión solo llevan estas dos hojas (los datos completos quedan en el panel)
  if (!maestra) {
    ["Informe", "Respuestas", "Hoja 1", "Sheet1"].forEach((n) => { const h = ss.getSheetByName(n); if (h) ss.deleteSheet(h); });
  }
}

/* ---------- Ayudas de diseño ---------- */
/** Franja amarilla arriba (como la página), título y texto de actualización. Devuelve la siguiente fila libre. */
function encabezadoHoja_(h, titulo, subtitulo, ancho) {
  h.getRange(1, 1, 1, ancho).setBackground(COLOR.amarillo);
  h.setRowHeight(1, 6);
  h.getRange(2, 1).setValue(titulo).setFontFamily("Oswald").setFontSize(20).setFontColor(COLOR.titulo);
  h.setRowHeight(2, 44);
  h.getRange(3, 1).setValue(subtitulo).setFontColor(COLOR.gris).setFontSize(9);
  return 5;
}
/** Fila de sección en azul PFJ. */
function seccion_(h, f, texto, ancho) {
  h.getRange(f, 1, 1, ancho).setBackground(COLOR.azul).setFontColor("#ffffff").setFontWeight("bold");
  h.getRange(f, 1).setValue(texto.toUpperCase());
  h.setRowHeight(f, 28);
  return f + 1;
}
/** Encabezados de columnas. */
function columnas_(h, f, nombres) {
  h.getRange(f, 1, 1, nombres.length).setValues([nombres]).setBackground(COLOR.encabezado).setFontColor(COLOR.azul).setFontWeight("bold");
}
/** Línea fina bajo cada fila de datos. */
function filaDatos_(h, f, ancho) {
  h.getRange(f, 1, 1, ancho).setBorder(null, null, true, null, null, null, COLOR.linea, SpreadsheetApp.BorderStyle.SOLID);
  h.setRowHeight(f, 24);
}
/** Tipografía y detalles comunes. */
function terminarHoja_(h) {
  const r = h.getRange(1, 1, Math.max(h.getLastRow(), 1), Math.max(h.getLastColumn(), 1));
  r.setVerticalAlignment("middle");
  h.getRange(3, 1, Math.max(h.getLastRow() - 2, 1), Math.max(h.getLastColumn(), 1)).setFontFamily("Open Sans");
  h.setHiddenGridlines(true);
  h.setFrozenRows(3);
}

/** Devuelve la hoja vacía (la crea si no existe) en la posición indicada. */
function hojaInforme_(ss, nombre, posicion) {
  let h = ss.getSheetByName(nombre);
  if (h) { h.clear(); h.clearFormats(); h.setFrozenRows(0); } else h = ss.insertSheet(nombre, posicion);
  return h;
}

function resumenMaestra() {
  armarInformes_(SpreadsheetApp.getActiveSpreadsheet(), null);
}

function correosDe_(texto) {
  return String(texto).split(/[,;\s]+/).map((c) => c.trim().toLowerCase()).filter((c) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
}

function carpetaSesiones_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("CARPETA_SESIONES_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* la carpeta se borró: crear otra */ } }
  const carpeta = DriveApp.createFolder(CARPETA_SESIONES);
  props.setProperty("CARPETA_SESIONES_ID", carpeta.getId());
  return carpeta;
}

/* ================= Utilidades ================= */
/** Evita que un texto se interprete como fórmula y corta textos muy largos. */
function limpiar_(v) {
  if (v instanceof Date || typeof v === "number" || typeof v === "boolean") return v;
  let s = String(v == null ? "" : v).slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/** Devuelve la pestaña de respuestas (Informe o Respuestas) y la crea con encabezados si no existe. */
function hoja_(nombre, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.appendRow(RESPUESTAS[nombre]);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, RESPUESTAS[nombre].length).setFontWeight("bold").setBackground("#d9ead3");
  }
  return h;
}
