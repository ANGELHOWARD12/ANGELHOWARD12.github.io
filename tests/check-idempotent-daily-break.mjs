import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

for (const marker of [
  'route === "schedule/break"',
  "async function saveDailyBreakSchedule",
  "loadData(db, { ownerIds: [user.id], includeHistory: false })",
  "dateValue < today && !data.lateEvidenceUploadsEnabled",
  "hasTaskConflict(data.tasks, user.id, dateValue, breakStart, breakEnd)",
  "const alreadySaved =",
  "if (!alreadySaved)",
  "idempotent: alreadySaved"
]) {
  if (!api.includes(marker)) throw new Error(`Falta guardado diario idempotente en API: ${marker}`);
}

for (const marker of [
  'const APP_VERSION = "52-team-member-retirement"',
  "async function saveDailyBreakSetting",
  'apiRequestWithRetry("/schedule/break"',
  "applyDailyBreakLocally(dateValue, settings)",
  "persistOfflineState(JSON.parse(JSON.stringify(state)))",
  "Break protegido para sincronizar.",
  'els.breakSettingsForm.addEventListener("submit", async (event) =>',
  "await saveDailyBreakSetting(dateValue, normalized)"
]) {
  if (!html.includes(marker)) throw new Error(`Falta guardado diario pequeno u offline: ${marker}`);
}

console.log("DailyBreak=SmallValidatedOperation");
console.log("Idempotency=DesiredState");
console.log("OfflineFallback=Preserved");
