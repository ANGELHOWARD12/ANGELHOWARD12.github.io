import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const operative = fs.readFileSync(new URL("../operativo.html", import.meta.url), "utf8");

if (html !== operative) throw new Error("index.html y operativo.html deben permanecer identicos");

const required = [
  'const OFFLINE_SESSION_DB = "taskhub-offline-session-v1"',
  'const OFFLINE_SESSION_STORE = "encrypted-sessions"',
  'iterations: 180_000',
  '{ name: "AES-GCM", length: 256 }',
  "async function unlockOfflineSession",
  "async function enterProtectedOfflineSession",
  "async function restoreCloudSessionIfNeeded",
  "Modo sin conexion.",
  "Sustento protegido en este dispositivo",
  "if (offlineSessionMode || !navigator.onLine)"
];

for (const marker of required) {
  if (!html.includes(marker)) throw new Error(`Falta continuidad offline: ${marker}`);
}

const transientGuard = html.indexOf("if (isTransientCloudError(error))");
const loginFallback = html.indexOf("if (await enterProtectedOfflineSession(email, password)) return;", transientGuard);
if (transientGuard < 0 || loginFallback < transientGuard) {
  throw new Error("El acceso local solo debe activarse ante errores transitorios de nube");
}

console.log("Continuidad offline cifrada y colas locales verificadas.");
