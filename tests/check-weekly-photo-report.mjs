import fs from "node:fs";
import path from "node:path";

const root = process.cwd();
const index = fs.readFileSync(path.join(root, "index.html"), "utf8");
const operativo = fs.readFileSync(path.join(root, "operativo.html"), "utf8");
const sw = fs.readFileSync(path.join(root, "sw.js"), "utf8");
const cloud = fs.readFileSync(path.join(root, "functions", "cloud", "[[path]].js"), "utf8");
const api = fs.readFileSync(path.join(root, "functions", "api", "[[path]].js"), "utf8");

if (index !== operativo) throw new Error("index.html y operativo.html deben ser identicos");
if (cloud !== api) throw new Error("Las rutas cloud y api deben ser identicas");

const requiredMarkers = [
  'id="downloadWeeklyPhotoReport"',
  "Descargar reporte fotografico PowerPoint",
  '<script src="./vendor/pptxgen-4.0.1.min.js"></script>',
  "async function downloadWeeklyPhotoReport()",
  "function buildWeeklyPhotoPresentation(",
  "function groupPhotoReportRows(",
  "function addPhotoReportImage(",
  "photo-report-cover.png",
  "photo-report-evidence.png",
  "rangeByTaskDate: true",
  "await pptx.writeFile({ fileName, compression: true })",
  "options.hyperlink",
  "const end = addDaysIso(start, 5)"
];

for (const marker of requiredMarkers) {
  if (!index.includes(marker)) throw new Error(`Falta marcador del reporte fotografico: ${marker}`);
}

if (!index.includes("if (!canViewAll())")) throw new Error("El reporte debe estar limitado a Coordinacion y Admin");
if (!/\^image\\\/\(\?:jpeg\|jpg\|png\|gif\|webp\)/.test(index)) {
  throw new Error("El reporte debe filtrar formatos fotograficos compatibles");
}

for (const relative of [
  "vendor/pptxgen-4.0.1.min.js",
  "vendor/pptxgenjs-LICENSE.txt",
  "report-assets/photo-report-cover.png",
  "report-assets/photo-report-evidence.png"
]) {
  const file = path.join(root, relative);
  if (!fs.existsSync(file) || fs.statSync(file).size === 0) throw new Error(`Recurso faltante: ${relative}`);
}

for (const resource of [
  "/vendor/pptxgen-4.0.1.min.js",
  "/report-assets/photo-report-cover.png",
  "/report-assets/photo-report-evidence.png"
]) {
  if (!sw.includes(resource)) throw new Error(`El Service Worker no almacena ${resource}`);
}

console.log("Reporte fotografico semanal PowerPoint: OK");
