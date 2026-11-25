/* ============================================================
   Capacitación en línea PFJ 2027 · lógica de la página
   - Registro del consejero (nombre, correo, sesión)
   - Marcar secciones como completadas
   - Repasos de opción múltiple por sección
   - Tareas de respuesta abierta
   - Envío de todo a una planilla de Google (Apps Script)
   ============================================================ */

const CONFIG = {
  // Pega aquí la URL del Apps Script publicado como aplicación web (termina en /exec).
  // Mientras esté vacía, todo se guarda solo en el dispositivo del consejero.
  ENDPOINT: "",
  CAPACITACION: "C1",
  SESIONES: ["La Serena", "Viña del Mar", "Santiago Norte", "Santiago Sur", "Santiago Este", "Santiago Oeste", "Santiago Centro", "Otra"]
};

const CLAVE = "pfj2027-" + CONFIG.CAPACITACION;
const estado = JSON.parse(localStorage.getItem(CLAVE) || "{}");
estado.secciones ||= {}; estado.repasos ||= {}; estado.tareas ||= {};
const guardar = () => localStorage.setItem(CLAVE, JSON.stringify(estado));

/* ---------- Envío a la planilla ---------- */
function enviar(tipo, datos) {
  if (!CONFIG.ENDPOINT || !estado.consejero) return;
  const cuerpo = { tipo, capacitacion: CONFIG.CAPACITACION, fecha: new Date().toISOString(), ...estado.consejero, ...datos };
  // text/plain evita la verificación CORS; Apps Script lo recibe igual en e.postData.contents
  fetch(CONFIG.ENDPOINT, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(cuerpo) })
    .catch(() => {});
}

/* ---------- Registro ---------- */
const reg = document.getElementById("registro");
const sel = document.getElementById("reg-sesion");
CONFIG.SESIONES.forEach((s) => sel.insertAdjacentHTML("beforeend", `<option>${s}</option>`));
function pintarRegistro() {
  const c = estado.consejero;
  reg.classList.toggle("registro--ok", !!c);
  if (c) document.getElementById("reg-nombre-saludo").textContent = `${c.nombre} · ${c.sesion}`;
}
document.getElementById("reg-form").addEventListener("submit", (ev) => {
  ev.preventDefault();
  const f = ev.target;
  estado.consejero = {
    nombre: f.nombre.value.trim(), correo: f.correo.value.trim().toLowerCase(),
    sesion: f.sesion.value, estaca: f.estaca.value.trim(),
    modalidad: f.modalidad.value
  };
  guardar(); pintarRegistro();
  enviar("registro", {});
  // reenviar lo que ya se hubiera hecho antes de registrarse
  Object.entries(estado.secciones).forEach(([id, v]) => v && enviar("seccion", { seccion: id, completado: true }));
  Object.entries(estado.repasos).forEach(([id, r]) => enviar("repaso", { seccion: id, puntaje: r.puntaje, total: r.total }));
  Object.entries(estado.tareas).forEach(([id, t]) => enviar("tarea", { tarea: id, pregunta: t.pregunta, respuesta: t.respuesta }));
});
document.getElementById("reg-cambiar").addEventListener("click", () => {
  const c = estado.consejero, f = document.getElementById("reg-form");
  f.nombre.value = c.nombre; f.correo.value = c.correo; f.sesion.value = c.sesion; f.estaca.value = c.estaca; f.modalidad.value = c.modalidad;
  delete estado.consejero; guardar(); pintarRegistro();
});
pintarRegistro();

function pedirRegistro() {
  if (estado.consejero) return true;
  alert("Antes de continuar, completa tus datos al inicio de la página para que tu avance quede registrado.");
  reg.scrollIntoView({ behavior: "smooth", block: "center" });
  return false;
}

/* ---------- Secciones completadas ---------- */
const secciones = [...document.querySelectorAll("[data-seccion]")];
const items = [...document.querySelectorAll("#indice-lista li")];
function actualizarProgreso() {
  let n = 0;
  secciones.forEach((s, i) => {
    const hecho = !!estado.secciones[s.id];
    s.querySelector(".completar input").checked = hecho;
    if (items[i]) items[i].classList.toggle("hecho", hecho);
    if (hecho) n++;
  });
  document.getElementById("progreso-texto").textContent = `${n} de ${secciones.length}`;
  document.getElementById("progreso-barra").style.width = (100 * n / secciones.length) + "%";
}
secciones.forEach((s) => {
  const cb = s.querySelector(".completar input");
  cb.addEventListener("change", () => {
    if (cb.checked && !pedirRegistro()) { cb.checked = false; return; }
    estado.secciones[s.id] = cb.checked; guardar(); actualizarProgreso();
    enviar("seccion", { seccion: s.id, completado: cb.checked });
  });
});
actualizarProgreso();

/* ---------- Repasos ---------- */
document.querySelectorAll(".repaso").forEach((cont) => {
  const id = cont.dataset.repaso;
  const preguntas = REPASOS[id] || [];
  const previo = estado.repasos[id];
  cont.innerHTML = `<p class="repaso__titulo"><span>Comprueba lo que aprendiste</span><span class="repaso__puntaje"></span></p>`;
  const respondidas = {};
  const puntaje = cont.querySelector(".repaso__puntaje");
  const pintarPuntaje = () => {
    const n = Object.keys(respondidas).length, ok = Object.values(respondidas).filter(Boolean).length;
    puntaje.textContent = n ? `${ok} de ${preguntas.length} correctas` : (previo ? `Último intento: ${previo.puntaje} de ${previo.total}` : "");
    if (n === preguntas.length) {
      estado.repasos[id] = { puntaje: ok, total: preguntas.length }; guardar();
      enviar("repaso", { seccion: id, puntaje: ok, total: preguntas.length });
    }
  };
  preguntas.forEach((q, i) => {
    const div = document.createElement("div");
    div.className = "pregunta";
    div.innerHTML = `<p>${i + 1}. ${q.p}</p>` + q.o.map((t, j) => `<button class="opcion" data-j="${j}">${t}</button>`).join("") + `<div class="explica">${q.e}</div>`;
    div.querySelectorAll(".opcion").forEach((b) => b.addEventListener("click", () => {
      if (i in respondidas) return; // solo cuenta el primer intento
      respondidas[i] = +b.dataset.j === q.c;
      b.classList.add(respondidas[i] ? "ok" : "mal");
      div.querySelector(`.opcion[data-j='${q.c}']`).classList.add("ok");
      div.querySelector(".explica").classList.add("ver");
      pintarPuntaje();
    }));
    cont.appendChild(div);
  });
  pintarPuntaje();
});

/* ---------- Tareas ---------- */
document.querySelectorAll(".tarea").forEach((t) => {
  const id = t.dataset.tarea;
  const pregunta = t.querySelector("p").textContent;
  t.insertAdjacentHTML("beforeend", `<textarea placeholder="Escribe aquí tu respuesta"></textarea><div class="tarea__pie"><button class="boton-borde">Guardar respuesta</button><span class="tarea__estado"></span></div>`);
  const ta = t.querySelector("textarea"), est = t.querySelector(".tarea__estado");
  if (estado.tareas[id]) { ta.value = estado.tareas[id].respuesta; est.textContent = "Respuesta guardada ✓"; }
  t.querySelector("button").addEventListener("click", () => {
    if (!ta.value.trim()) { est.textContent = "Escribe tu respuesta primero."; return; }
    if (!pedirRegistro()) return;
    estado.tareas[id] = { pregunta, respuesta: ta.value.trim() }; guardar();
    enviar("tarea", { tarea: id, pregunta, respuesta: ta.value.trim() });
    est.textContent = CONFIG.ENDPOINT ? "Respuesta enviada a tus coordinadores ✓" : "Respuesta guardada en este dispositivo ✓";
  });
});
