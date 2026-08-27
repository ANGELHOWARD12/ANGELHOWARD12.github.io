import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");

if (html !== operative) throw new Error("index.html y operativo.html deben permanecer identicos");

const required = [
  "const EVIDENCE_UPLOAD_TIMEOUT_MS = 180_000",
  "window.setTimeout(() => controller.abort(), EVIDENCE_UPLOAD_TIMEOUT_MS)",
  "async function apiBinaryRequestWithRetry(path, file, mimeType, attempts = 3)",
  'id="latestEvidenceReceipt"',
  "function renderLatestEvidenceReceipt()",
  "openEvidenceEntries.add(evidence.id)",
  "Acceso protegido activo",
  "function uploadFileDirectlyToR2",
  "function uploadMetricText",
  "Subiendo directo a R2",
  "restante ${etaLabel}",
  '"/evidence/upload/r2/presign"',
  '"/evidence/upload/r2/confirm"'
];

for (const marker of required) {
  if (!html.includes(marker)) throw new Error(`Falta mejora de resiliencia: ${marker}`);
}

const evidenceSubmitStart = html.indexOf('els.evidenceForm.addEventListener("submit"');
const evidenceSubmitEnd = html.indexOf('els.createTrainerForm.addEventListener("submit"', evidenceSubmitStart);
const evidenceSubmitSource = html.slice(evidenceSubmitStart, evidenceSubmitEnd);
if (evidenceSubmitSource.includes('cloudStatus === "error"')) {
  throw new Error("Un estado antiguo de nube no debe impedir un nuevo intento de carga");
}

console.log("UploadTimeout=180s");
console.log("UploadRetries=3");
console.log("EvidenceReceipt=OK");
console.log("StaleCloudStatusGate=None");
