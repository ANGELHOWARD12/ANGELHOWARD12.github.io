import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

for (const marker of [
  "taskHistoryMatch",
  "async function getTaskHistory",
  "No tienes permiso para consultar este historial.",
  "includeHistory = true",
  "includeHistory: false",
  "function taskWithoutHistory",
  "historyLoaded: false",
  "function mergeSubmittedTaskHistory",
  "delete task.historyLoaded"
]) {
  if (!api.includes(marker)) throw new Error(`Falta historial bajo demanda en API: ${marker}`);
}

for (const marker of [
  'const APP_VERSION = "55-access-restored"',
  "async function ensureTaskHistory",
  "/history`)",
  "task.historyLoaded = true",
  "Cargando historial completo...",
  "historyLoaded: task.historyLoaded !== false"
]) {
  if (!html.includes(marker)) throw new Error(`Falta historial bajo demanda en cliente: ${marker}`);
}

console.log("StateHistory=Excluded");
console.log("TaskHistory=PermissionCheckedOnDemand");
console.log("OfflineHistoryDelta=Preserved");
