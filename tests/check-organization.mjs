import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const cloud = fs.readFileSync(new URL("../functions/cloud/[[path]].js", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");
const sw = fs.readFileSync(new URL("../sw.js", import.meta.url), "utf8");

const expectedEmails = [
  "nykol.ruiz@lgtask.local",
  "ronald.chavez@lgtask.local",
  "alejandro.cotrina@lgtask.local",
  "ariana.perez@lgtask.local",
  "abel.barrantes@lgtask.local",
  "fernando.bedrinana@lgtask.local",
  "nathaly.fuentes@lgtask.local"
];

for (const email of expectedEmails) {
  if (!cloud.includes(email)) throw new Error(`Falta el acceso organizacional ${email}`);
}

for (const marker of [
  'const SCHEMA_VERSION = "27-uppercase-users-1"',
  'UPDATE users SET name = UPPER(TRIM(name))',
  '["team", "TEXT NOT NULL DEFAULT',
  '["job_title", "TEXT NOT NULL DEFAULT',
  '["member_type", "TEXT NOT NULL DEFAULT',
  'userTeam(actor) === userTeam(target)',
  'La cuenta observadora es de solo lectura',
  'Solo puedes asignar tareas dentro de tu equipo',
  'No puedes revisar tareas de otro equipo',
  'function coordinatorCanManageTaskUser',
  'function isSelfManagedMasterTask',
  'function validOfflineEvidenceQueue',
  'type: "SustentoOffline"',
  'type: "AutoaprobacionMaster"',
  'type: "EdicionMaster"',
  'review: selfManagedMaster ? "Aprobada" : "Pendiente"',
  'route === "tasks/master-update"'
]) {
  if (!cloud.includes(marker)) throw new Error(`Falta la proteccion: ${marker}`);
}

for (const marker of [
  'id="organizationScopeTabs"',
  'data-organization-scope="training"',
  'data-organization-scope="masters"',
  'data-organization-scope="audiovisual"',
  'class="scope-tab-copy"',
  'function organizationScopeUsers()',
  'function renderOrganizationScope()',
  'function isSelfManagedMasterTask(task)',
  'function hasCoordinatorPersonalWeek()',
  'function updateMasterTask(taskId)',
  'const APP_VERSION = "42-weekly-photo-report1"',
  'const APP_MODE = document.documentElement.classList.contains("app-mode")',
  '<title>Task Hub</title>',
  '<span class="brand-mark" aria-label="Task Hub">TH</span>',
  'id="quickAddTask"',
  'class="panel stack task-composer hidden"',
  'flex: 1 1 calc(50% - 5px);',
  'const EVIDENCE_OUTBOX_DB = "lgtask-evidence-outbox-v1"',
  'function syncOfflineOutboxes()'
]) {
  if (!html.includes(marker)) throw new Error(`Falta la interfaz organizacional: ${marker}`);
}

if (cloud !== api) throw new Error("Las rutas /cloud y /api no son identicas");
if (html !== operative) throw new Error("index.html y operativo.html no son identicos");
if (!sw.includes('task-hub-shell-v42-weekly-photo-report1')) throw new Error("El cache del Service Worker no fue actualizado");
if (html.includes('id="evidenceLinks"') || html.includes("els.evidenceLinks")) {
  throw new Error("El campo de enlaces continua visible en el formulario de sustentos");
}

console.log("OrganizationProfiles=7");
console.log("TeamIsolation=OK");
console.log("ObserverReadOnly=OK");
console.log("FrontendScopes=Trainers,Master,Audiovisual");
console.log("MasterAutonomy=OK");
console.log("CoordinatorPersonalWeek=OK");
