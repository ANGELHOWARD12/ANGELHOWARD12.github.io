import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

for (const marker of [
  'route === "tasks/action"',
  "async function mutateTaskAction",
  '["start", "rename", "reschedule", "reassign", "reminder"]',
  "loadData(db, { ownerIds: actionOwnerIds })",
  "clean(entry.mutationId) === mutationId",
  'action === "start"',
  'action === "rename"',
  'action === "reschedule"',
  'action === "reassign"',
  'action === "reminder"'
]) {
  if (!api.includes(marker)) throw new Error(`Falta accion idempotente en API: ${marker}`);
}

for (const marker of [
  'const APP_VERSION = "47-idempotent-task-actions"',
  "async function requestTaskAction",
  'apiRequest("/tasks/action"',
  'requestTaskAction("start"',
  '"reassign",',
  '"rename",',
  '"reschedule",',
  '"reminder",',
  "applyOffline(mutationId)",
  "Cambio guardado. Task Hub lo sincronizara al volver la nube."
]) {
  if (!html.includes(marker)) throw new Error(`Falta accion pequena u offline: ${marker}`);
}

console.log("TaskActions=Start+Rename+Reschedule+Reassign+Reminder");
console.log("Idempotency=MutationId");
console.log("OfflineFallback=Preserved");
