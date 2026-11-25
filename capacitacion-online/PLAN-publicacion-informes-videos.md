# Capacitación en línea PFJ 2027: cómo publicarla, cómo recibir informes y qué videos grabar

## 1. Cómo publicar la página sin pagar dominio

La página son archivos simples (HTML, CSS y JS). No necesita servidor ni base de datos propia, así que se puede alojar gratis.

| Opción | Dirección que tendría | Costo | A favor | En contra |
|---|---|---|---|---|
| **GitHub Pages** (recomendada) | `https://pfj-capacitacion.github.io/` (o similar) | Gratis | Dirección limpia, rápida y estable. Se actualiza subiendo los archivos. Se le puede poner un dominio propio después. | Hay que crear una cuenta de GitHub para el equipo (puede ser con el correo de PFJ). |
| **Netlify** o **Cloudflare Pages** | `https://pfj-capacitacion.netlify.app` | Gratis | Se publica arrastrando la carpeta al navegador. | Otra cuenta más que administrar. |
| **Google Apps Script** (todo dentro de Google) | `https://script.google.com/macros/s/…/exec` | Gratis | Todo queda en la cuenta de Google de PFJ, junto a la planilla de informes. | Dirección larga y fea. Google muestra un aviso arriba de la página. Más difícil de editar. |
| **Google Sites** | `https://sites.google.com/view/…` | Gratis | Muy fácil de editar por cualquiera del equipo. | No permite marcar como completado, ni repasos, ni enviar respuestas. Solo serviría para mostrar los videos. |

**Recomendación:** publicar en **GitHub Pages** y guardar los informes en una **planilla de Google** (punto 2). Así quedan separadas la página (pública, sin datos personales) y la información de los consejeros (privada, en el Drive de PFJ).

**Si después se quiere dominio propio:** cuesta entre 10 y 15 USD al año (por ejemplo `capacitacionpfj.org`). Se conecta a GitHub Pages en unos minutos sin cambiar nada de la página. No es necesario para empezar.

**Mientras no haya permiso para usar la marca de la Iglesia:** la página ya no tiene encabezado ni nombre de la Iglesia, solo los títulos de la capacitación. Conviene compartir el enlace únicamente con los consejeros inscritos (por el Classroom o el grupo de WhatsApp de cada sesión), no publicarlo en redes.

## 2. Cómo recibir el informe de quién hizo qué

### Lo que ya hace la página
1. Al empezar, el consejero escribe su **nombre, correo, sesión, estaca** y si hace la capacitación **en lugar de la presencial** o **como repaso**.
2. Cada vez que **marca una sección como completada**, **termina un repaso** o **guarda una tarea**, la página lo envía a la planilla.
3. Si el consejero cambia de computador o borra el navegador, sus datos anteriores igual quedan en la planilla.

### La planilla de Google (archivo `apps-script/Codigo.gs`)
Se crea sola la primera vez que llega un dato. Tiene 5 hojas:

- **Resumen**: una fila por consejero, con su sesión, el avance (por ejemplo 7 de 10, 70 %), los repasos (por ejemplo 24/29 correctas), las tareas respondidas, la última actividad y su estado ("Solo registrado", "En curso" o "Completó la parte en línea").
- **Registros**: quién se inscribió y cuándo.
- **Avance**: cada sección marcada o desmarcada, con fecha y hora.
- **Repasos**: el puntaje de cada repaso. Cuenta solo el primer intento de cada pregunta.
- **Tareas**: la respuesta completa de cada tarea, para que los coordinadores la lean.

### Cómo lo ven los coordinadores
- **Opción simple:** compartir la planilla completa (solo lectura) con los coordinadores de área.
- **Por sesión:** en una planilla aparte para cada sesión, usar `=QUERY(IMPORTRANGE("url de la planilla"; "Resumen!A:M"); "select * where Col1 = 'Santiago Este'")`. Así cada sesión ve solo a sus consejeros y sus tareas.
- Los coordinadores de sesión pueden usar la hoja **Tareas** para preparar la **sesión recuperativa**: ver qué temas costaron más y qué respuestas conviene comentar.

### Instalación (unos 10 minutos)
Los pasos están al inicio de `Codigo.gs`. En resumen: crear la planilla, pegar el código en *Extensiones > Apps Script*, implementar como *Aplicación web* (ejecutar como "Yo", acceso "Cualquier persona"), copiar la URL `/exec` y pegarla en `app.js`, en `CONFIG.ENDPOINT`.

### Cuidado de los datos
- La planilla guarda nombres, correos y respuestas de los consejeros. Debe quedar en la cuenta de PFJ y compartirse solo con quienes la necesitan.
- La página no guarda nada de los jóvenes participantes.
- La URL del Apps Script solo puede **agregar** filas a la planilla. Nadie puede leer la planilla a través de ella.

## 3. Qué videos hay que tener

Hay dos tipos de videos: unos los **graba una sola vez el Área** (iguales para todas las sesiones) y otros los **graba cada sesión** (para que los consejeros conozcan a su staff). Sugerencia: que el Área grabe la mayor parte, para que el contenido sea uniforme y las sesiones no carguen con todo.

| # | Video | Duración sugerida | Lo graba | Estado |
|---|---|---|---|---|
| 1 | Bienvenida del élder y la hermana Renlund | existente | Video oficial | Ya existe |
| 2 | Pensamiento espiritual "En sus asuntos" | 5–8 min | Codirector de cada sesión | Por grabar |
| 3 | Presentación del staff de sesión (organigrama con fotos) | 3–5 min | Matrimonio Director de cada sesión | Por grabar |
| 4 | Línea de autoridad, del Setenta al consejero | 5 min | Área (coordinador de área) | Por grabar |
| 5 | Adelanto de la semana | existente | Video oficial | Ya existe |
| 6 | Check-in y Reúnete con tu consejero (con simulación de la lista de verificación) | 8–10 min | Área (coordinadores auxiliares de ejemplo) | Por grabar |
| 7 | Reúnete con tu compañía y Orientación | 6–8 min | Área | Por grabar |
| 8 | Noche de hogar: lección, un juego con su análisis y metas con "Un modelo para el crecimiento" | 10 min | Área | Por grabar |
| 9 | Taller: Inclusión durante el check-in | 10 min | **Equipo de Inclusión del Área** (texto sin cambios) | Por grabar |
| 10 | Baile en línea paso a paso | existente | Videos del paso a paso del Drive | Ya existe |
| 11 | Estandartes, versos cantados y comportamiento en los bailes | 6–8 min | Área | Por grabar |
| 12 | Taller: Sugerencias para enseñar a los jóvenes | 8–10 min | Área o Directora de sesión | Por grabar |
| 13 | Taller: Principios de vestimenta y apariencia | 10 min | Área o Codirectora de sesión | Por grabar |
| 14 | Taller: Tiempo libre de consejeros | 5–8 min | Área o Codirector de sesión | Por grabar |
| 15 | Mensaje final "Regocíjate en Cristo" | 10–15 min | Director de cada sesión | Por grabar |

**Sesión por sesión solo habría que grabar 3 videos (2, 3 y 15).** El resto se graba una vez para toda el Área.

**Dónde subirlos:** en YouTube como **"No listado"**, en un canal de la cuenta de PFJ. Se insertan en la página con el enlace y no aparecen en búsquedas. Otra opción es Google Drive, pero carga más lento y a veces pide permisos.

## 4. Antes de usarla con consejeros

- [ ] **Revisar las respuestas de los repasos con el Manual para el Personal 2027.** Varias preguntas vienen de los formularios de 2024-2025, y algunas cosas cambiaron (por ejemplo, ahora los consejeros eligen el nombre de la compañía). Las que más conviene revisar: quién supervisa cada actividad y la regla de comer con otra compañía.
- [ ] Completar los enlaces pendientes: manuales, Guía de Aplicación, discurso "En sus asuntos", música del baile, Kit de Salud Mental y capacitación de vestimenta.
- [ ] Grabar los videos y reemplazar las portadas de muestra.
- [ ] Instalar la planilla y pegar la URL en `app.js`.
- [ ] Publicar en GitHub Pages y probar con 2 o 3 consejeros antes de abrirla a todos.
- [ ] Definir con cada sesión la fecha de su **sesión recuperativa**.
