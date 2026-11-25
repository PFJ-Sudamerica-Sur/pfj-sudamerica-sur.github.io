/* ============================================================
   Capacitación en línea PFJ 2027 · lógica de la página
   - Carga el contenido (datos.json y, si está conectado, el panel de Google Sheets)
   - Videos y fotos editables desde el panel
   - Registro del consejero con sesión y estaca en lista
   - Módulos, secciones completadas y repasos de alternativas
   - Envío de avances y respuestas a la planilla (Apps Script)
   ============================================================ */

const CONFIG = {
  // URL del Apps Script publicado como aplicación web (termina en /exec).
  // Mientras esté vacía, se usa datos.json y todo se guarda solo en el dispositivo.
  ENDPOINT: "",
  CAPACITACION: "C1"
};

const CLAVE = "pfj2027-" + CONFIG.CAPACITACION;
const estado = JSON.parse(localStorage.getItem(CLAVE) || "{}");
estado.secciones ||= {}; estado.repasos ||= {}; estado.bailes ||= [];
const guardar = () => localStorage.setItem(CLAVE, JSON.stringify(estado));
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
let DATOS = { estacas: {}, videos: [], fotos: [], infoSesiones: [], preguntas: [], modulos: [] };

/* ---------- Contenido ---------- */
async function cargarDatos() {
  try { DATOS = { ...DATOS, ...(await fetch("datos.json", { cache: "no-cache" }).then((r) => r.json())) }; } catch (e) { /* sin datos por defecto */ }
  if (!CONFIG.ENDPOINT) return;
  try {
    const r = await fetch(CONFIG.ENDPOINT + "?accion=contenido").then((x) => x.json());
    if (r && r.ok && r.datos) {
      Object.entries(r.datos).forEach(([k, v]) => {
        const vacio = v == null || (Array.isArray(v) ? !v.length : typeof v === "object" && !Object.keys(v).length);
        if (!vacio) DATOS[k] = v;
      });
    }
  } catch (e) { /* si el panel no responde, se usa datos.json */ }
}

/* ---------- Envío a la planilla ---------- */
function enviar(tipo, datos) {
  if (!CONFIG.ENDPOINT || !estado.consejero) return;
  const cuerpo = { tipo, capacitacion: CONFIG.CAPACITACION, fecha: new Date().toISOString(), ...estado.consejero, ...datos };
  // text/plain evita la verificación CORS; Apps Script lo recibe igual en e.postData.contents
  fetch(CONFIG.ENDPOINT, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(cuerpo) }).catch(() => {});
}

/* ---------- Videos y fotos ---------- */
function urlVideo(u) {
  u = String(u || "").trim();
  let m = u.match(/(?:youtu\.be\/|[?&]v=|\/embed\/|\/shorts\/|\/live\/)([\w-]{11})/);
  if (m) return `https://www.youtube-nocookie.com/embed/${m[1]}?rel=0`;
  m = u.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([\w-]{20,})/);
  if (m) return `https://drive.google.com/file/d/${m[1]}/preview`;
  m = u.match(/vimeo\.com\/(?:video\/)?(\d+)/);
  if (m) return `https://player.vimeo.com/video/${m[1]}`;
  return "";
}
function urlFoto(u) {
  u = String(u || "").trim();
  const m = u.match(/drive\.google\.com\/.*(?:\/d\/|[?&]id=)([\w-]{20,})/);
  if (m) return `https://lh3.googleusercontent.com/d/${m[1]}`;
  return /^https:\/\//.test(u) ? u : "";
}
function videoPara(clave) {
  const ses = estado.consejero?.sesion || "";
  const vs = DATOS.videos.filter((v) => v.clave === clave && v.enlace);
  return vs.find((v) => v.sesion && v.sesion === ses) || vs.find((v) => !v.sesion);
}
function pintarVideos() {
  $$(".video[data-video]").forEach((cont) => {
    const v = videoPara(cont.dataset.video);
    const src = v && urlVideo(v.enlace);
    const poster = $(".poster", cont), marco = $(".video__marco", cont);
    if (src) {
      if (!marco || marco.dataset.src !== src) {
        marco?.remove();
        const div = document.createElement("div");
        div.className = "video__marco"; div.dataset.src = src;
        div.innerHTML = `<iframe src="${esc(src)}" title="${esc(v.titulo || "Video")}" loading="lazy" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>`;
        cont.insertBefore(div, cont.firstChild);
      }
      if (poster) poster.hidden = true;
      cont.classList.remove("video--pendiente");
    } else {
      marco?.remove();
      if (poster) poster.hidden = false;
      cont.classList.add("video--pendiente");
    }
  });
}
function pintarStaff() {
  const cont = $("#staff-sesion");
  if (!cont) return;
  const ses = estado.consejero?.sesion;
  const fotos = DATOS.fotos.filter((f) => ses && f.sesion === ses && (f.nombre || f.foto));
  if (!fotos.length) return; // se queda la lista de cargos
  cont.innerHTML = `<div class="staff__grilla">` + fotos.map((f) => {
    const src = urlFoto(f.foto);
    return `<figure class="staff__ficha">${src ? `<img src="${esc(src)}" alt="${esc(f.nombre)}" loading="lazy" referrerpolicy="no-referrer">` : `<div class="staff__sinfoto"></div>`}<figcaption><b>${esc(f.nombre)}</b><span>${esc(f.cargo)}</span></figcaption></figure>`;
  }).join("") + `</div>`;
}
function pintarInfoSesion() {
  const ses = estado.consejero?.sesion;
  const info = DATOS.infoSesiones.find((i) => ses && i.sesion === ses) || {};
  const enlace = (u, t) => (/^https:\/\//.test(u || "") ? `<a href="${esc(u)}" target="_blank" rel="noopener">${t}</a>` : "");
  $$("[data-info-sesion]").forEach((cont) => {
    if (cont.dataset.infoSesion === "contacto") {
      const links = [enlace(info.whatsapp, "Grupo de WhatsApp de tu sesión"), enlace(info.classroom, "Classroom de tu sesión")].filter(Boolean);
      cont.innerHTML = ses && links.length ? `<p><b>Únete:</b> ${links.join(" · ")}</p>` : "";
    } else {
      const lugar = /^https:\/\//.test(info.lugar || "") ? enlace(info.lugar, "Entrar a la reunión en línea") : esc(info.lugar || "");
      cont.innerHTML = info.recuperativa || info.lugar
        ? `<div class="aviso"><b>Tu sesión recuperativa${ses ? " (" + esc(ses) + ")" : ""}:</b><br>${info.recuperativa ? "Fecha: " + esc(info.recuperativa) + "<br>" : ""}${lugar ? "Lugar: " + lugar : ""}</div>`
        : `<div class="aviso">Los coordinadores de tu sesión te avisarán la fecha y el lugar de la sesión recuperativa.</div>`;
    }
  });
}
function pintarSegunSesion() { pintarVideos(); pintarStaff(); pintarInfoSesion(); }

/* ---------- Registro ---------- */
const OTRA_SESION = "__otra", OTRA_DE_SESION = "__todas", OTRA_ESCRITA = "__escribir";
let estacaModoTodas = false;

function opcionesSesion() {
  return `<option value="">Elige tu sesión</option>` +
    Object.entries(DATOS.estacas).map(([pais, ses]) =>
      `<optgroup label="${esc(pais)}">` + Object.keys(ses).map((s) => `<option value="${esc(pais + "|" + s)}">${esc(s)}</option>`).join("") + `</optgroup>`
    ).join("") +
    `<option value="${OTRA_SESION}">Mi sesión no está en la lista</option>`;
}
function opcionesTodasEstacas() {
  return Object.entries(DATOS.estacas).map(([pais, ses]) => {
    const lista = [...new Set(Object.values(ses).flat())].sort((a, b) => a.localeCompare(b, "es"));
    return `<optgroup label="${esc(pais)}">` + lista.map((e) => `<option>${esc(e)}</option>`).join("") + `</optgroup>`;
  }).join("");
}
function llenarEstacas() {
  const selS = $("#reg-sesion"), selE = $("#reg-estaca");
  const v = selS.value;
  if (!v) { selE.innerHTML = `<option value="">Primero elige tu sesión</option>`; selE.disabled = true; mostrarEstacaEscrita(false); return; }
  selE.disabled = false;
  if (v === OTRA_SESION || estacaModoTodas) {
    selE.innerHTML = `<option value="">Elige tu estaca</option>` + opcionesTodasEstacas() + `<option value="${OTRA_ESCRITA}">Mi estaca no aparece (escribirla)</option>`;
  } else {
    const [pais, ses] = v.split("|");
    const lista = (DATOS.estacas[pais] || {})[ses] || [];
    selE.innerHTML = `<option value="">Elige tu estaca</option>` + lista.map((e) => `<option>${esc(e)}</option>`).join("") + `<option value="${OTRA_DE_SESION}">Mi estaca es de otra sesión</option>`;
  }
  mostrarEstacaEscrita(false);
}
function mostrarEstacaEscrita(si) {
  $("#reg-estaca-otra-campo").hidden = !si;
  $("#reg-estaca-otra").required = si;
  if (!si) $("#reg-estaca-otra").value = "";
}

function iniciarRegistro() {
  const reg = $("#registro");
  if (!reg) return;
  const selS = $("#reg-sesion"), selE = $("#reg-estaca");
  selS.innerHTML = opcionesSesion();
  selS.addEventListener("change", () => { estacaModoTodas = false; llenarEstacas(); });
  selE.addEventListener("change", () => {
    if (selE.value === OTRA_DE_SESION) { estacaModoTodas = true; llenarEstacas(); selE.focus(); }
    else mostrarEstacaEscrita(selE.value === OTRA_ESCRITA);
  });

  const pintar = () => {
    const c = estado.consejero;
    reg.classList.toggle("registro--ok", !!c);
    if (c) $("#reg-nombre-saludo").textContent = `${c.nombre} · ${c.sesion} · ${c.estaca}`;
  };
  $("#reg-form").addEventListener("submit", (ev) => {
    ev.preventDefault();
    const f = ev.target;
    const otraSesion = selS.value === OTRA_SESION;
    const [pais, sesion] = otraSesion ? ["", "Otra"] : selS.value.split("|");
    const estaca = selE.value === OTRA_ESCRITA ? f.estacaOtra.value.trim() : selE.value;
    estado.consejero = {
      nombre: f.nombre.value.trim(), correo: f.correo.value.trim().toLowerCase(),
      pais, sesion, estaca, estacaEscrita: selE.value === OTRA_ESCRITA,
      modalidad: f.modalidad.value
    };
    guardar(); pintar(); pintarSegunSesion();
    enviar("registro", {});
    // reenviar lo que ya se hubiera hecho antes de registrarse
    Object.entries(estado.secciones).forEach(([id, v]) => v && enviar("seccion", { seccion: id, completado: true }));
    Object.entries(estado.repasos).forEach(([m, r]) => enviar("repaso", { modulo: m, puntaje: r.puntaje, total: r.total }));
  });
  $("#reg-cambiar").addEventListener("click", () => {
    const c = estado.consejero, f = $("#reg-form");
    f.nombre.value = c.nombre; f.correo.value = c.correo; f.modalidad.value = c.modalidad;
    selS.value = c.sesion === "Otra" ? OTRA_SESION : `${c.pais}|${c.sesion}`;
    estacaModoTodas = false; llenarEstacas();
    if (c.estacaEscrita) { estacaModoTodas = true; llenarEstacas(); selE.value = OTRA_ESCRITA; mostrarEstacaEscrita(true); f.estacaOtra.value = c.estaca; }
    else {
      selE.value = c.estaca;
      if (selE.value !== c.estaca) { estacaModoTodas = true; llenarEstacas(); selE.value = c.estaca; }
    }
    delete estado.consejero; guardar(); pintar();
  });
  pintar();
}

function pedirRegistro() {
  if (estado.consejero) return true;
  alert("Antes de continuar, completa tus datos al inicio de la página para que tu avance quede registrado.");
  $("#registro")?.scrollIntoView({ behavior: "smooth", block: "center" });
  return false;
}

/* ---------- Módulos ---------- */
const modulos = () => $$(".modulo");
function moduloDe(el) { return el?.closest(".modulo"); }
function activarModulo(n, destino) {
  modulos().forEach((m) => m.classList.toggle("activo", m.dataset.modulo === String(n)));
  $$(".indice-modulos .modbtn").forEach((b) => b.classList.toggle("activo", b.dataset.modulo === String(n)));
  const el = destino || $(`#modulo-${n}`);
  // salto directo (no suave): al cambiar de módulo cambia el alto de la página y un desplazamiento suave queda corrido
  if (el) setTimeout(() => el.scrollIntoView({ behavior: "instant", block: "start" }), 0);
}
function irAHash(desplazar) {
  const id = decodeURIComponent(location.hash.slice(1));
  const el = id && document.getElementById(id);
  const mod = el && (el.classList.contains("modulo") ? el : moduloDe(el));
  if (mod) activarModulo(mod.dataset.modulo, desplazar ? el : null);
  else if (!$(".modulo.activo")) {
    // primer módulo sin terminar
    const pendiente = modulos().find((m) => $$("[data-seccion]", m).some((s) => !estado.secciones[s.id]));
    modulos().forEach((m) => m.classList.toggle("activo", m === (pendiente || modulos()[0])));
    const act = $(".modulo.activo");
    $$(".indice-modulos .modbtn").forEach((b) => b.classList.toggle("activo", b.dataset.modulo === act?.dataset.modulo));
  }
}
function iniciarModulos() {
  const indice = $("#indice-modulos");
  if (!indice) return;
  indice.innerHTML = modulos().map((m) => {
    const n = m.dataset.modulo, titulo = $(".modulo__cabecera h2", m).textContent;
    return `<a class="modbtn" href="#modulo-${n}" data-modulo="${n}"><span class="modbtn__n">Módulo ${n}</span><span class="modbtn__t">${esc(titulo)}</span><span class="modbtn__p"></span></a>`;
  }).join("");
  modulos().forEach((m, i) => {
    const nav = $(".modulo__nav", m), lista = modulos();
    const ant = lista[i - 1], sig = lista[i + 1];
    nav.innerHTML = (ant ? `<a class="boton-borde" href="#modulo-${ant.dataset.modulo}">← Módulo ${ant.dataset.modulo}</a>` : `<span></span>`) +
      (sig ? `<a class="boton-rojo" href="#modulo-${sig.dataset.modulo}">Siguiente: Módulo ${sig.dataset.modulo} →</a>` : `<span></span>`);
  });
  window.addEventListener("hashchange", () => irAHash(true));
  irAHash(!!location.hash);
}

/* ---------- Secciones completadas ---------- */
function actualizarProgreso() {
  const secciones = $$("[data-seccion]");
  if (!secciones.length) return;
  let n = 0;
  secciones.forEach((s) => {
    const hecho = !!estado.secciones[s.id];
    $(".completar input", s).checked = hecho;
    if (hecho) n++;
  });
  modulos().forEach((m) => {
    const ss = $$("[data-seccion]", m), h = ss.filter((s) => estado.secciones[s.id]).length;
    const b = $(`.modbtn[data-modulo="${m.dataset.modulo}"]`);
    if (b) { $(".modbtn__p", b).textContent = h === ss.length ? "✓ Completado" : `${h} de ${ss.length}`; b.classList.toggle("hecho", h === ss.length); }
  });
  $("#progreso-texto").textContent = `${n} de ${secciones.length} secciones`;
  $("#progreso-barra").style.width = (100 * n / secciones.length) + "%";
}
function iniciarSecciones() {
  $$("[data-seccion]").forEach((s) => {
    const cb = $(".completar input", s);
    cb.addEventListener("change", () => {
      if (cb.checked && !pedirRegistro()) { cb.checked = false; return; }
      estado.secciones[s.id] = cb.checked; guardar(); actualizarProgreso(); pintarFinal();
      enviar("seccion", { seccion: s.id, completado: cb.checked });
    });
  });
  actualizarProgreso();
}

/* ---------- Repasos por módulo (solo alternativas) ---------- */
function iniciarRepasos() {
  $$("[data-modulo-repaso]").forEach((cont) => {
    const m = cont.dataset.moduloRepaso;
    const preguntas = DATOS.preguntas.filter((q) => String(q.modulo) === m && q.pregunta && q.opciones?.length >= 2);
    if (!preguntas.length) { cont.innerHTML = ""; return; }
    const previo = estado.repasos[m];
    cont.innerHTML = `<p class="repaso__titulo"><span>Repaso del módulo ${esc(m)}</span><span class="repaso__puntaje"></span></p>`;
    const respondidas = {};
    const puntaje = $(".repaso__puntaje", cont);
    const pintarPuntaje = () => {
      const n = Object.keys(respondidas).length, ok = Object.values(respondidas).filter(Boolean).length;
      puntaje.textContent = n ? `${ok} de ${preguntas.length} correctas` : (previo ? `Tu resultado: ${previo.puntaje} de ${previo.total}` : "");
      if (n === preguntas.length && !previo) {
        estado.repasos[m] = { puntaje: ok, total: preguntas.length }; guardar();
        enviar("repaso", { modulo: m, puntaje: ok, total: preguntas.length });
        pintarFinal();
      }
    };
    preguntas.forEach((q, i) => {
      const div = document.createElement("div");
      div.className = "pregunta";
      div.innerHTML = `<p>${i + 1}. ${esc(q.pregunta)}</p>` + q.opciones.map((t, j) => `<button type="button" class="opcion" data-j="${j}">${esc(t)}</button>`).join("") + `<div class="explica">${esc(q.explicacion || "")}</div>`;
      $$(".opcion", div).forEach((b) => b.addEventListener("click", () => {
        if (i in respondidas) return; // solo cuenta el primer intento
        if (!pedirRegistro()) return;
        const j = +b.dataset.j, correcta = j === Number(q.correcta);
        respondidas[i] = correcta;
        b.classList.add(correcta ? "ok" : "mal");
        $(`.opcion[data-j='${q.correcta}']`, div)?.classList.add("ok");
        $(".explica", div).classList.add("ver");
        if (!previo) enviar("respuesta", { modulo: m, pregunta: q.pregunta, respuesta: q.opciones[j], correcta });
        pintarPuntaje();
      }));
      cont.appendChild(div);
    });
    pintarPuntaje();
  });
}

/* ---------- Bailes en línea (un reproductor, varios botones) ---------- */
function iniciarBailes() {
  const sel = $("#baile-selector");
  if (!sel) return;
  const bailes = DATOS.videos.filter((v) => /^baile-\d+$/.test(v.clave) && !v.sesion)
    .sort((a, b) => a.clave.localeCompare(b.clave, "es", { numeric: true }));
  if (!bailes.length) return;
  const video = $(".video", sel), botones = $(".baile__botones", sel);
  botones.innerHTML = bailes.map((b, i) =>
    `<button type="button" role="tab" class="baile__btn" data-clave="${esc(b.clave)}"><span>Baile ${i + 1}</span>${esc(b.titulo)}</button>`).join("");
  const elegir = (clave) => {
    video.dataset.video = clave;
    $$(".baile__btn", botones).forEach((b) => b.setAttribute("aria-selected", b.dataset.clave === clave));
    pintarVideos();
  };
  $$(".baile__btn", botones).forEach((b) => b.addEventListener("click", () => elegir(b.dataset.clave)));
  elegir(bailes[0].clave);

  const lista = $("#baile-practicados");
  lista.innerHTML = bailes.map((b) =>
    `<label class="completar"><input type="checkbox" value="${esc(b.titulo)}"${estado.bailes.includes(b.titulo) ? " checked" : ""}> ${esc(b.titulo)}</label>`).join("");
  $$("input", lista).forEach((cb) => cb.addEventListener("change", () => {
    if (cb.checked && !pedirRegistro()) { cb.checked = false; return; }
    estado.bailes = $$("input:checked", lista).map((x) => x.value); guardar();
    enviar("seccion", { seccion: "Baile practicado: " + cb.value, completado: cb.checked });
  }));
}

/* ---------- Terminar la capacitación ---------- */
function resumenAvance() {
  return modulos().map((m) => {
    const n = m.dataset.modulo;
    const ss = $$("[data-seccion]", m);
    const hechas = ss.filter((s) => estado.secciones[s.id]).length;
    const tienePreguntas = DATOS.preguntas.some((q) => String(q.modulo) === n);
    const rep = estado.repasos[n];
    return { n, titulo: $(".modulo__cabecera h2", m).textContent, hechas, total: ss.length,
             repaso: rep, listo: hechas === ss.length && (!tienePreguntas || !!rep) };
  });
}
function pintarFinal() {
  const cont = $("#final-resumen");
  if (!cont) return;
  const r = resumenAvance(), todo = r.every((x) => x.listo);
  cont.innerHTML = `<ul class="final__lista">` + r.map((x) =>
    `<li class="${x.listo ? "ok" : ""}"><a href="#modulo-${x.n}">Módulo ${x.n}: ${esc(x.titulo)}</a><span>${x.hechas} de ${x.total} secciones · ` +
    (x.repaso ? `repaso ${x.repaso.puntaje} de ${x.repaso.total}` : "repaso pendiente") + `</span></li>`).join("") + `</ul>` +
    (todo ? "" : `<p class="final__falta">Te falta completar lo que está en gris. Haz clic en el módulo para terminarlo.</p>`);
  const btn = $("#final-enviar");
  btn.disabled = !todo || !!estado.finalizado;
  btn.textContent = estado.finalizado ? "Capacitación enviada ✓" : "Finalizar y enviar";
  const conf = $("#final-confirmacion");
  if (estado.finalizado) {
    conf.innerHTML = `<div class="aviso aviso--ok"><b>¡Listo, ${esc(estado.consejero?.nombre?.split(" ")[0] || "")}! Terminaste la parte en línea.</b><br>` +
      (CONFIG.ENDPOINT ? "Tus coordinadores ya recibieron tu capacitación." : "Tu avance quedó guardado en este dispositivo (el envío a los coordinadores aún no está conectado).") +
      `<br>No olvides asistir a la <a href="#sesion-recuperativa">sesión recuperativa</a>.</div>`;
  } else conf.innerHTML = "";
}
function iniciarFinal() {
  const btn = $("#final-enviar");
  if (!btn) return;
  btn.addEventListener("click", () => {
    if (!pedirRegistro()) return;
    const r = resumenAvance();
    if (!r.every((x) => x.listo)) return;
    estado.finalizado = new Date().toISOString(); guardar();
    const correctas = r.reduce((a, x) => a + (x.repaso?.puntaje || 0), 0), total = r.reduce((a, x) => a + (x.repaso?.total || 0), 0);
    enviar("final", { correctas, totalPreguntas: total, bailes: estado.bailes.join(", ") });
    pintarFinal();
  });
  pintarFinal();
}

/* ---------- Inicio ---------- */
(async () => {
  await cargarDatos();
  iniciarRegistro();
  iniciarModulos();
  iniciarSecciones();
  iniciarRepasos();
  iniciarBailes();
  iniciarFinal();
  pintarSegunSesion();
})();
