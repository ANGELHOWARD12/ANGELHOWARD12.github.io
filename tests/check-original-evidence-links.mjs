import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");
const cloud = fs.readFileSync(new URL("../functions/cloud/[[path]].js", import.meta.url), "utf8");
const api = fs.readFileSync(new URL("../functions/api/[[path]].js", import.meta.url), "utf8");

if (html !== operative) throw new Error("index.html y operativo.html no son identicos");
if (cloud !== api) throw new Error("Los endpoints cloud y api no son identicos");

for (const marker of [
  'const APP_VERSION = "48-lazy-task-history"',
  "function originalEvidenceFileUrl(file, { absolute = false } = {})",
  "return new URL(source, PUBLIC_APP_ORIGIN).href",
  '{ header: "Direccion del original", width: 48 }',
  "text: row.fileUrl",
  "hyperlinks: reportRow.fileUrl",
  'target="_blank" rel="noopener">Ver archivo original</a>',
  'target="_blank" rel="noopener" title="Abrir imagen original en otra pestana"',
  'id="openFilePreviewOriginal"',
  "originalEvidenceFileUrl(file, { absolute: true })"
]) {
  if (!html.includes(marker)) throw new Error(`Falta el enlace al original: ${marker}`);
}

for (const marker of [
  'const evidenceFileMatch = route.match(/^evidence\\/([^/]+)\\/(?:file|photo)$/)',
  "const session = await authenticate(request, env.DB)",
  '"Content-Disposition": `${disposition}; filename="${safeName}"`',
  "return r2FileResponse(env, storage, row)"
]) {
  if (!cloud.includes(marker)) throw new Error(`Falta la proteccion o entrega original: ${marker}`);
}

console.log("OriginalEvidenceLinks=OK");
console.log("Excel=DirectAuthenticatedOriginal");
console.log("TaskHub=NewTabOriginal");
