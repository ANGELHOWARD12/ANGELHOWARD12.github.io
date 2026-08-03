const CACHE_NAME = "lgtask-shell-v19";
const APP_SHELL = [
  "/",
  "/index.html",
  "/operativo.html",
  "/manifest.webmanifest",
  "/vendor/exceljs-4.4.0.min.js",
  "/vendor/jszip-3.10.1.min.js",
  "/icons/lgtask-96.png",
  "/icons/lgtask-192.png",
  "/icons/lgtask-512.png"
];

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
    const response = await fetch("/cloud/notifications/pending", {
      credentials: "include",
      cache: "no-store"
    });
    if (!response.ok) throw new Error("No active session");
    const data = await response.json();
    for (const notification of data.notifications || []) {
      await self.registration.showNotification(notification.title || "LGTASK", {
        body: notification.body || "Tienes una actualizacion pendiente.",
        icon: "/icons/lgtask-192.png",
        badge: "/icons/lgtask-96.png",
        tag: notification.id,
        renotify: true,
        data: { url: notification.url || "/?view=tasksView" }
      });
    }
  } catch {
    await self.registration.showNotification("Nueva actividad en LGTASK", {
      body: "Abre LGTASK para revisar tus tareas y recordatorios.",
      icon: "/icons/lgtask-192.png",
      badge: "/icons/lgtask-96.png",
      tag: "lgtask-generic",
      data: { url: "/?view=tasksView" }
    });
  }
}

self.addEventListener("push", (event) => {
  event.waitUntil(showPendingNotifications());
});

self.addEventListener("periodicsync", (event) => {
  if (event.tag !== "lgtask-reminders") return;
  event.waitUntil(
    fetch("/cloud/state", { credentials: "include", cache: "no-store" })
      .then(() => new Promise((resolve) => setTimeout(resolve, 1200)))
      .then(showPendingNotifications)
      .catch(() => {})
  );
});

self.addEventListener("message", (event) => {
  if (event.data?.type !== "SHOW_NOTIFICATION") return;
  event.waitUntil(
    self.registration.showNotification(event.data.title || "LGTASK", {
      body: event.data.body || "",
      icon: "/icons/lgtask-192.png",
      badge: "/icons/lgtask-96.png",
      tag: event.data.tag || `lgtask-${Date.now()}`,
      data: { url: event.data.url || "/?view=tasksView" }
    })
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
