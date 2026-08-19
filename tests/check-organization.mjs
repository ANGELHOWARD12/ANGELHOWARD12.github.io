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
  'const SCHEMA_VERSION = "24-org-teams-1"',
  '["team", "TEXT NOT NULL DEFAULT',
  '["job_title", "TEXT NOT NULL DEFAULT',
  '["member_type", "TEXT NOT NULL DEFAULT',
  'userTeam(actor) === userTeam(target)',
  'La cuenta observadora es de solo lectura',
  'Solo puedes asignar tareas dentro de tu equipo',
  'No puedes revisar tareas de otro equipo'
]) {
  if (!cloud.includes(marker)) throw new Error(`Falta la proteccion: ${marker}`);
}

for (const marker of [
  'id="organizationScope"',
  '<option value="training">Training</option>',
  '<option value="masters">Masters</option>',
  '<option value="audiovisual">Audiovisuales</option>',
  'function organizationScopeUsers()',
  'function renderOrganizationScope()',
  'const APP_VERSION = "24-org-teams1"'
]) {
  if (!html.includes(marker)) throw new Error(`Falta la interfaz organizacional: ${marker}`);
}

if (cloud !== api) throw new Error("Las rutas /cloud y /api no son identicas");
if (html !== operative) throw new Error("index.html y operativo.html no son identicos");
if (!sw.includes('lgtask-shell-v24-org-teams1')) throw new Error("El cache del Service Worker no fue actualizado");

console.log("OrganizationProfiles=7");
console.log("TeamIsolation=OK");
console.log("ObserverReadOnly=OK");
console.log("FrontendScopes=Training,Masters,Audiovisuales");
