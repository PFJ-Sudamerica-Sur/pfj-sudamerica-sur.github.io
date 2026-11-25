/**
 * Capacitación en línea PFJ 2027 · Panel de administración + receptor de respuestas
 * --------------------------------------------------------------------------------
 * Esta planilla de Google es el "panel admin" de la página:
 *   - Lo que escribas en las pestañas Videos, Fotos staff, Info sesiones, Estacas y Preguntas
 *     aparece en la página (al recargarla).
 *   - Cuando un consejero presiona "Finalizar y enviar", su capacitación completa llega a
 *     Informe (una fila por consejero) y Respuestas (el detalle), y también a la planilla de su sesión.
 *   - En la pestaña Coordinadores se anotan los correos de cada sesión. El menú
 *     "Crear o actualizar planillas por sesión" crea una planilla por sesión y la comparte solo con ellos.
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
    .addItem("Crear o actualizar planillas por sesión", "crearPlanillasSesion")
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
    const { informe, respuestas } = filasDeEnvio_(d);

    // 1) Planilla maestra (todas las sesiones)
    agregarFilas_(SpreadsheetApp.getActiveSpreadsheet(), informe, respuestas);

    // 2) Planilla de la sesión, si ya existe. Si falla, la maestra igual quedó guardada.
    try {
      const url = planillaDeSesion_(d.pais, d.sesion);
      if (url) agregarFilas_(SpreadsheetApp.openByUrl(url), informe, respuestas);
    } catch (err) {
      console.error("No se pudo copiar a la planilla de la sesión " + d.sesion + ": " + err);
    }
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

/** Busca en la pestaña Coordinadores la planilla ya creada para esa sesión. */
function planillaDeSesion_(pais, sesion) {
  const fila = filasPanel_("Coordinadores").find((r) => r[1] === String(sesion || "").trim() && (!pais || !r[0] || r[0] === pais));
  return fila && /^https:\/\//.test(fila[3]) ? fila[3] : "";
}

/* ================= Planillas por sesión ================= */
function crearPlanillasSesion() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hc = ss.getSheetByName("Coordinadores");
  if (!hc) { ui.alert("Falta la pestaña Coordinadores. Usa primero \"Preparar panel (primera vez)\"."); return; }

  const carpeta = carpetaSesiones_();
  const yo = Session.getEffectiveUser().getEmail().toLowerCase();
  const maestraInforme = hoja_("Informe").getDataRange().getValues().slice(1);
  const maestraResp = hoja_("Respuestas").getDataRange().getValues().slice(1);
  const resumen = [];

  const datos = hc.getLastRow() > 1 ? hc.getRange(2, 1, hc.getLastRow() - 1, 4).getValues() : [];
  datos.forEach((r, i) => {
    const pais = String(r[0]).trim(), sesion = String(r[1]).trim();
    const correos = String(r[2]).split(/[,;\s]+/).map((c) => c.trim().toLowerCase()).filter((c) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
    let url = String(r[3]).trim();
    if (!sesion || (!correos.length && !url)) return;

    let planilla;
    if (url) {
      planilla = SpreadsheetApp.openByUrl(url);
    } else {
      // Crear la planilla de la sesión y copiar lo que ya hubiera llegado de esa sesión
      planilla = SpreadsheetApp.create("Capacitación PFJ 2027 · " + sesion + (pais ? " (" + pais + ")" : ""));
      DriveApp.getFileById(planilla.getId()).moveTo(carpeta);
      hoja_("Informe", planilla); hoja_("Respuestas", planilla);
      const inicial = planilla.getSheetByName("Hoja 1") || planilla.getSheetByName("Sheet1");
      if (inicial) planilla.deleteSheet(inicial);
      const deSesion = (filas, col) => filas.filter((f) => String(f[col]).trim() === sesion);
      const inf = deSesion(maestraInforme, 5), resp = deSesion(maestraResp, 4);
      if (inf.length) planilla.getSheetByName("Informe").getRange(2, 1, inf.length, inf[0].length).setValues(inf);
      if (resp.length) planilla.getSheetByName("Respuestas").getRange(2, 1, resp.length, resp[0].length).setValues(resp);
      url = planilla.getUrl();
      hc.getRange(i + 2, 4).setValue(url);
    }

    // Compartir solo con los coordinadores de la lista (solo lectura) y quitar a quienes ya no están
    const archivo = DriveApp.getFileById(planilla.getId());
    const actuales = archivo.getViewers().map((u) => u.getEmail().toLowerCase());
    correos.filter((c) => !actuales.includes(c)).forEach((c) => archivo.addViewer(c));
    actuales.filter((c) => !correos.includes(c) && c !== yo).forEach((c) => archivo.removeViewer(c));
    resumen.push(sesion + ": " + (correos.length ? correos.length + " coordinador(es)" : "sin correos"));
  });

  ui.alert(resumen.length
    ? "Listo. Planillas creadas o actualizadas:\n\n" + resumen.join("\n")
    : "No hay sesiones con correos en la pestaña Coordinadores.");
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
