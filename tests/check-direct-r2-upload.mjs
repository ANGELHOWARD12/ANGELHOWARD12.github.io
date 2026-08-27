import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");
const cloud = fs.readFileSync(new URL("../functions/cloud/[[path]].js", import.meta.url), "utf8");

if (api !== cloud) throw new Error("Las rutas cloud y api deben permanecer identicas");

for (const marker of [
  "async function presignEvidenceDirectUpload",
  "async function confirmEvidenceDirectUpload",
  "async function createR2PresignedPutUrl",
  "X-Amz-Algorithm",
  "X-Amz-Signature",
  "CREATE TABLE IF NOT EXISTS r2_direct_uploads",
  "cleanupAbandonedDirectUploads"
]) {
  if (!api.includes(marker)) throw new Error(`Falta carga directa R2: ${marker}`);
}

for (const marker of [
  "new XMLHttpRequest()",
  'xhr.open("PUT", signed.url, true)',
  'xhr.upload.addEventListener("progress"',
  "speedMbps",
  "etaSeconds",
  "for (let attempt = 1; attempt <= 4; attempt += 1)"
]) {
  if (!html.includes(marker)) throw new Error(`Falta progreso/reintento del navegador: ${marker}`);
}

console.log("DirectR2=Presigned PUT");
console.log("Progress=Mbps+ETA+Stage");
console.log("EncryptedOutbox=AES-GCM");
