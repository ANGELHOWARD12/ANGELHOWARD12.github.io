import fs from "node:fs";

const html = fs.readFileSync(new URL("../index.html", import.meta.url), "utf8");
const start = html.indexOf("async function apiRequest(path, options = {})");
const end = html.indexOf("function cloudRetryDelay(attempt)", start);

if (start < 0 || end < 0) throw new Error("No se encontro apiRequest en index.html");

const source = html.slice(start, end);
const statusChanges = [];
const apiRequest = new Function(
  "API_BASE",
  "API_FALLBACK_BASE",
  "window",
  "navigator",
  "setCloudStatus",
  `${source}; return apiRequest;`
)(
  "/cloud",
  "/api",
  { setTimeout, clearTimeout },
  { onLine: true },
  (...args) => statusChanges.push(args)
);

const calls = [];
globalThis.fetch = async (url) => {
  calls.push(url);
  if (url.startsWith("/cloud")) {
    return new Response("temporary gateway failure", { status: 503 });
  }
  return Response.json({ ok: true, source: "fallback" });
};

const login = await apiRequest("/auth/login", {
  method: "POST",
  body: JSON.stringify({ email: "test@lgtask.local", password: "secret" }),
  allowFallback: true
});

if (!login.ok || login.source !== "fallback") throw new Error("El login no uso la ruta de respaldo");
if (calls.join(",") !== "/cloud/auth/login,/api/auth/login") {
  throw new Error(`Secuencia de login inesperada: ${calls.join(",")}`);
}

calls.length = 0;
const state = await apiRequest("/state");
if (!state.ok || calls.join(",") !== "/cloud/state,/api/state") {
  throw new Error("Las consultas GET no usan la ruta de respaldo");
}

if (!statusChanges.some(([status]) => status === "online")) {
  throw new Error("La conexion recuperada no vuelve a estado online");
}

console.log("CloudFallback=OK");
console.log("LoginFailover=cloud->api");
console.log("ReadFailover=cloud->api");
