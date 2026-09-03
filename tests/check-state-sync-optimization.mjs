import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

for (const marker of [
  'const SCHEMA_VERSION = "30-user-retirement-date-1"',
  "async function readStateVersion(db)",
  'searchParams.get("since")',
  "notModified: true",
  "loadData(db, { ownerIds: requestedOwnerIds, includeHistory: false })",
  "INNER JOIN task_records task ON task.id = history.task_id",
  "idx_task_records_updated",
  'async function dueReminderNotifications(db, now, ownerId = "")',
  "json_each(task.data, '$.reminders')",
  "dispatchDueReminders(db, null, user.id)",
  "loadData(db, { ownerIds: [user.id] })"
]) {
  if (!api.includes(marker)) throw new Error(`Falta optimizacion D1: ${marker}`);
}

if (/\(route === "state" && request\.method === "GET"\) \|\| route === "storage\/status"/.test(api)) {
  throw new Error("GET /state no debe disparar mantenimiento pesado en cada consulta");
}

for (const marker of [
  'const APP_VERSION = "54-access-suspended"',
  "const STATE_REFRESH_INTERVAL_MS = 60 * 1000",
  "const STATE_REFRESH_JITTER_MS = 15 * 1000",
  "let serverStateVersion = \"\"",
  "/state?since=",
  "if (data.notModified)",
  "function schedulePeriodicStateRefresh()"
]) {
  if (!html.includes(marker)) throw new Error(`Falta sincronizacion incremental: ${marker}`);
}

if (html.includes("}, 30000);")) {
  throw new Error("Permanece la consulta fija de estado cada 30 segundos");
}

console.log("StateVersion=ConditionalRefresh");
console.log("D1Scope=RoleOwnerFilter");
console.log("ReminderRead=NoTaskHistory");
console.log("Refresh=60s+Jitter");
