import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const serviceWorker = fs.readFileSync(new URL("../sw.js", import.meta.url), "utf8");

const requiredHtmlMarkers = [
  'const HOURLY_CLOUD_MAINTENANCE_MS = 60 * 60 * 1000',
  'async function runHourlyCloudMaintenance({ force = false } = {})',
  'await registration?.update().catch(() => {})',
  'await syncOfflineOutboxes()',
  'await refreshStateFromServer({ quiet: true })',
  'Math.ceil(timeoutMs / apiBases.length)',
  'if (protectedWorkInProgress())',
  'pendingAppReload = true',
  'window.setInterval(() => runHourlyCloudMaintenance({ force: true })'
];

for (const marker of requiredHtmlMarkers) {
  if (!html.includes(marker)) throw new Error(`Falta proteccion horaria: ${marker}`);
}

if (!serviceWorker.includes('for (const base of ["/cloud", "/api"])')) {
  throw new Error("El service worker no tiene ruta secundaria de nube");
}

const protectedCheck = html.indexOf('if (protectedWorkInProgress())', html.indexOf('controllerchange'));
const reload = html.indexOf('window.location.reload()', protectedCheck);
if (protectedCheck < 0 || reload < protectedCheck) {
  throw new Error("La actualizacion podria recargar durante una operacion protegida");
}

console.log("HourlyCloudWatchdog=OK");
console.log("SafeReloadGuard=OK");
console.log("ServiceWorkerFailover=cloud->api");
