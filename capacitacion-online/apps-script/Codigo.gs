/**
 * Capacitación en línea PFJ 2027 · receptor de avances
 * ----------------------------------------------------
 * Recibe lo que envía la página (registro, secciones completadas, repasos y tareas)
 * y lo guarda en esta planilla de Google. Además arma una hoja "Resumen" por consejero.
 *
 * Cómo instalarlo (una sola vez):
 *  1. Crea una planilla nueva en Google Drive (con la cuenta de PFJ), p. ej. "Capacitación en línea PFJ 2027 · Informe".
 *  2. Menú Extensiones > Apps Script. Borra lo que haya y pega este archivo completo. Guarda.
 *  3. Botón Implementar > Nueva implementación > tipo "Aplicación web".
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *  4. Autoriza los permisos y copia la URL que termina en /exec.
 *  5. Pega esa URL en app.js, en CONFIG.ENDPOINT.
 *  Si después cambias este código, usa Implementar > Gestionar implementaciones > editar > Nueva versión
 *  (así la URL no cambia).
 */

const TOTAL_SECCIONES = { C1: 10, C2: 0, C3: 0 }; // completar cuando existan la 2ª y 3ª

const HOJAS = {
  Registros: ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Estaca", "Modalidad"],
  Avance:    ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Sección", "Completado"],
  Repasos:   ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Sección", "Puntaje", "Total"],
  Tareas:    ["Fecha", "Capacitación", "Nombre", "Correo", "Sesión", "Tarea", "Pregunta", "Respuesta"],
  Resumen:   ["Sesión", "Nombre", "Correo", "Estaca", "Modalidad", "Capacitación", "Secciones completadas",
              "Total secciones", "% avance", "Repasos (correctas/total)", "Tareas respondidas", "Última actividad", "Estado"]
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fecha = new Date(d.fecha || Date.now());
    const base = [fecha, d.capacitacion, d.nombre, (d.correo || "").toLowerCase(), d.sesion];
    if (d.tipo === "registro") hoja_("Registros").appendRow([fecha, d.capacitacion, d.nombre, (d.correo || "").toLowerCase(), d.sesion, d.estaca, d.modalidad]);
    if (d.tipo === "seccion")  hoja_("Avance").appendRow(base.concat([d.seccion, d.completado ? "Sí" : "No"]));
    if (d.tipo === "repaso")   hoja_("Repasos").appendRow(base.concat([d.seccion, d.puntaje, d.total]));
    if (d.tipo === "tarea")    hoja_("Tareas").appendRow(base.concat([d.tarea, d.pregunta, d.respuesta]));
    actualizarResumen();
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

/** Reconstruye la hoja Resumen: una fila por consejero y capacitación, con el último estado de cada cosa. */
function actualizarResumen() {
  const personas = {}; // clave: correo|capacitación
  const clave = (correo, cap) => correo + "|" + cap;
  const persona = (r, cap) => {
    const k = clave(r[3], cap);
    personas[k] = personas[k] || { cap, nombre: r[2], correo: r[3], sesion: r[4], estaca: "", modalidad: "",
                                   secciones: {}, repasos: {}, tareas: {}, ultima: r[0] };
    if (r[0] > personas[k].ultima) personas[k].ultima = r[0];
    return personas[k];
  };
  filas_("Registros").forEach((r) => { const p = persona(r, r[1]); p.nombre = r[2]; p.sesion = r[4]; p.estaca = r[5]; p.modalidad = r[6]; });
  filas_("Avance").forEach((r) => { persona(r, r[1]).secciones[r[5]] = r[6] === "Sí"; });
  filas_("Repasos").forEach((r) => { persona(r, r[1]).repasos[r[5]] = [r[6], r[7]]; });
  filas_("Tareas").forEach((r) => { persona(r, r[1]).tareas[r[5]] = true; });

  const salida = Object.values(personas).map((p) => {
    const hechas = Object.values(p.secciones).filter(Boolean).length;
    const total = TOTAL_SECCIONES[p.cap] || hechas;
    let ok = 0, tot = 0;
    Object.values(p.repasos).forEach(([a, b]) => { ok += Number(a); tot += Number(b); });
    const pct = total ? hechas / total : 0;
    const estado = pct >= 1 ? "Completó la parte en línea" : (hechas ? "En curso" : "Solo registrado");
    return [p.sesion, p.nombre, p.correo, p.estaca, p.modalidad, p.cap, hechas, total, pct,
            tot ? ok + "/" + tot : "", Object.keys(p.tareas).length, p.ultima, estado];
  }).sort((a, b) => (a[0] + a[1]).localeCompare(b[0] + b[1]));

  const h = hoja_("Resumen");
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, HOJAS.Resumen.length).clearContent();
  if (salida.length) {
    h.getRange(2, 1, salida.length, HOJAS.Resumen.length).setValues(salida);
    h.getRange(2, 9, salida.length, 1).setNumberFormat("0%");
  }
}

/** Menú para actualizar el resumen a mano desde la planilla. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Capacitación PFJ").addItem("Actualizar resumen", "actualizarResumen").addToUi();
}

function hoja_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.appendRow(HOJAS[nombre]);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, HOJAS[nombre].length).setFontWeight("bold");
  }
  return h;
}

function filas_(nombre) {
  const h = hoja_(nombre);
  return h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, HOJAS[nombre].length).getValues() : [];
}
