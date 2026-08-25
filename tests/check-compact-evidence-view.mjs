import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");

if (html !== operative) throw new Error("index.html y operativo.html no son identicos");

for (const id of ["personalBreakPanel", "overtimePanel"]) {
  if (!html.includes(`<details id="${id}" class="panel collapsible-panel">`)) {
    throw new Error(`${id} no es desplegable`);
  }
}

const tasksStart = html.indexOf('<section id="tasksView"');
const evidenceStart = html.indexOf('<section id="evidenceView"');
const teamStart = html.indexOf('<section id="teamView"');
const trackingStart = html.indexOf('<aside id="taskTrackingPanel"');
const uploadStart = html.indexOf('<section id="evidenceUploadPanel"');

if (!(tasksStart >= 0 && evidenceStart > tasksStart && teamStart > evidenceStart)) {
  throw new Error("No se encontraron las vistas principales en el orden esperado");
}
if (!(trackingStart > uploadStart && trackingStart < teamStart)) {
  throw new Error("Mis tareas / Seguimiento no esta dentro de Sustentos despues del formulario");
}
if (html.slice(tasksStart, evidenceStart).includes("Mis tareas / Seguimiento")) {
  throw new Error("Mis tareas / Seguimiento continua dentro de Tareas");
}
if (!html.includes('activateView("evidenceView", { updateUrl: true, historyMode: "replace" });')) {
  throw new Error("El envio de sustentos no permanece en la vista Sustentos");
}

console.log("CompactPanels=OK");
console.log("TaskTrackingView=Sustentos");
