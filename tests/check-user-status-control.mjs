import assert from "node:assert/strict";
import fs from "node:fs";

const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");
const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");

assert.match(api, /route === "admin\/user-status"/);
assert.match(api, /target\.role !== "Trainer" \|\| !coordinatorCanManageUser\(actor, target\)/);
assert.match(api, /UPDATE users SET status = \?, retired_at = \? WHERE id = \?/);
assert.match(api, /UPDATE app_data SET updated_at = CASE WHEN updated_at >= \?/);
assert.match(api, /DELETE FROM sessions WHERE user_id = \?/);
assert.match(api, /tasksPreserved: true/);
assert.match(api, /nextStatus = retire \? "Retirado"/);
assert.match(api, /retired_at = \?/);
assert.doesNotMatch(api.slice(api.indexOf("async function setUserStatus"), api.indexOf("async function resetPassword")), /DELETE FROM task_records|DELETE FROM evidence/);
assert.match(html, /Desactivar acceso/);
assert.match(html, /Reactivar acceso/);
assert.match(html, /Sus tareas y sustentos se conservaran/);
assert.match(html, /Retirar del Task/);
assert.match(html, /user\.status !== "Retirado"/);
assert.match(html, /dateValue < retiredDate && state\.tasks\.some/);

console.log("TeamAccess=CoordinatorScopedDisableReactivate");
console.log("HistoricalTasksAndEvidence=Preserved");
console.log("RetiredMembers=HiddenFromActiveLists");
