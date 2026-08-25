const CACHE_NAME = "task-hub-shell-v39-compact-evidence-view1";
const APP_SHELL = [
  "/",
  "/index.html",
  "/operativo.html",
  "/manifest.webmanifest",
  "/vendor/exceljs-4.4.0.min.js",
  "/vendor/jszip-3.10.1.min.js",
  "/icons/task-hub-96.png",
  "/icons/task-hub-192.png",
  "/icons/task-hub-512.png"
];

function notificationOptions(notification = {}) {
  return {
    body: notification.body || "Tienes una actualizacion pendiente.",
    icon: "/icons/task-hub-192.png",
    badge: "/icons/task-hub-96.png",
    tag: notification.id || `lgtask-${Date.now()}`,
    renotify: true,
    requireInteraction: true,
    vibrate: [180, 90, 180, 90, 360],
    silent: false,
    timestamp: Date.now(),
    actions: [{ action: "open", title: "Abrir Task Hub" }],
    data: { url: notification.url || "/?view=tasksView" }
  };
}

async function fetchCloudWithFallback(path) {
  let lastError;
  for (const base of ["/cloud", "/api"]) {
    try {
      const response = await fetch(`${base}${path}`, {
        credentials: "include",
        cache: "no-store"
      });
      if (response.ok) return response;
      lastError = new Error(`Cloud route failed with HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Cloud unavailable");
}

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || url.origin !== self.location.origin || /^\/(?:cloud|api)\//.test(url.pathname)) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put("/index.html", clone));
          return response;
        })
        .catch(() => caches.match("/index.html"))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then(
      (cached) =>
        cached ||
        fetch(request).then((response) => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, clone));
          }
          return response;
        })
    )
  );
});

async function showPendingNotifications() {
  try {
    const response = await fetchCloudWithFallback("/notifications/pending");
    const data = await response.json();
    for (const notification of data.notifications || []) {
      await self.registration.showNotification(notification.title || "Task Hub", notificationOptions(notification));
    }
  } catch {
    await self.registration.showNotification(
      "Nueva actividad en Task Hub",
      notificationOptions({
        id: "lgtask-generic",
        body: "Abre Task Hub para revisar tus tareas y recordatorios.",
        url: "/?view=tasksView"
      })
    );
  }
}

self.addEventListener("push", (event) => {
  event.waitUntil(showPendingNotifications());
});

self.addEventListener("periodicsync", (event) => {
  if (event.tag !== "lgtask-reminders") return;
  event.waitUntil(
    fetchCloudWithFallback("/state")
      .then(() => new Promise((resolve) => setTimeout(resolve, 1200)))
      .then(showPendingNotifications)
      .catch(() => {})
  );
});

self.addEventListener("sync", (event) => {
  if (event.tag !== "lgtask-cloud-sync") return;
  event.waitUntil(
    fetchCloudWithFallback("/state")
      .then(() => self.clients.matchAll({ type: "window", includeUncontrolled: true }))
      .then((clients) => clients.forEach((client) => client.postMessage({ type: "CLOUD_ONLINE" })))
      .catch(() => {})
  );
});

self.addEventListener("message", (event) => {
  if (event.data?.type !== "SHOW_NOTIFICATION") return;
  event.waitUntil(
    self.registration.showNotification(
      event.data.title || "Task Hub",
      notificationOptions({
        id: event.data.tag,
        body: event.data.body,
        url: event.data.url
      })
    )
  );
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = new URL(event.notification.data?.url || "/?view=tasksView", self.location.origin).href;
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(async (clients) => {
      for (const client of clients) {
        if ("navigate" in client) await client.navigate(targetUrl);
        if ("focus" in client) return client.focus();
      }
      return self.clients.openWindow(targetUrl);
    })
  );
});
