import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");

if (html !== operative) throw new Error("index.html y operativo.html no son identicos");

for (const marker of [
  'id="weeklyEvidenceDownloadPanel"',
  'id="previousEvidenceWeek"',
  'id="currentEvidenceWeek"',
  'id="nextEvidenceWeek"',
  'id="downloadWeeklyEvidence"',
  "Descargar sustentos de la semana completa",
  'els.weeklyEvidenceDownloadPanel.classList.toggle("hidden", !canViewAll())',
  "function renderWeeklyEvidenceDownload()",
  "function downloadWeeklyEvidencePackage()",
  "rangeByTaskDate: true",
  'els.evidenceDateTo.value = addDaysIso(start, 5)',
  'els.evidenceTrainerFilter.value = "all"'
]) {
  if (!html.includes(marker)) throw new Error(`Falta la descarga semanal: ${marker}`);
}

const viewStart = html.indexOf('<section id="evidenceView"');
const panelStart = html.indexOf('id="weeklyEvidenceDownloadPanel"');
const uploadStart = html.indexOf('id="evidenceUploadPanel"');
if (!(viewStart >= 0 && panelStart > viewStart && panelStart < uploadStart)) {
  throw new Error("La descarga semanal no aparece de forma visible al inicio de Sustentos");
}

console.log("WeeklyEvidenceDownload=OK");
console.log("Roles=Coordinador,Admin");
console.log("Range=Monday-Saturday");
