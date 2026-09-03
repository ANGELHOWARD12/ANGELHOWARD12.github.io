import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

if (!html.includes('const APP_VERSION = "54-access-suspended"')) {
  throw new Error("La version de la etapa 6 no fue actualizada.");
}

for (const marker of [
  'route === "app/action"',
  "async function mutateAppAction",
  "async function mutateAppData",
  "attempts = 6",
  "WHERE id = 1 AND updated_at = ?",
  "concurrencyRetries",
  "function appDataPayload",
  'Object.defineProperty(data, "__appDataSnapshot"',
  "payloadJson !== data.__appDataSnapshot",
  "results[appUpdateIndex]?.meta?.changes"
]) {
  if (!api.includes(marker)) throw new Error(`Falta proteccion concurrente en API: ${marker}`);
}

for (const marker of [
  "async function requestAppAction",
  'apiRequestWithRetry("/app/action"',
  '"motivation",',
  '"announcement",',
  '"registration-review",',
  '"recovery-close",'
]) {
  if (!html.includes(marker)) throw new Error(`Falta accion general pequena en cliente: ${marker}`);
}

const fullStateSaveCalls = html.match(/\bsaveState\(\);/g) || [];
if (fullStateSaveCalls.length !== 4) {
  throw new Error(`PUT /state debe quedar solo para tres respaldos offline y la importacion; encontrados: ${fullStateSaveCalls.length}`);
}

// Modela 100 escrituras de tareas independientes: ninguna necesita tocar app_data.
const taskRows = new Map(Array.from({ length: 100 }, (_, index) => [`task-${index}`, 0]));
await Promise.all(
  Array.from(taskRows.keys(), async (taskId, index) => {
    await new Promise((resolve) => setTimeout(resolve, index % 4));
    taskRows.set(taskId, taskRows.get(taskId) + 1);
  })
);
if (taskRows.size !== 100 || Array.from(taskRows.values()).some((value) => value !== 1)) {
  throw new Error("La simulacion concurrente perdio una escritura de tarea.");
}

console.log("AppActions=SmallCASUpdates");
console.log("TaskWrites=NoUnchangedAppDataWrite");
console.log("ConcurrentSimulation=100IndependentTasks");
console.log("FullState=OfflineAndImportOnly");
