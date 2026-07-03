// Visit Report — Service Worker
// Caches the app shell so the form opens instantly even with weak/no signal.
// Form submissions still need internet (they email + write to your Sheet live).

const CACHE_NAME = "visit-report-v3";
const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;

  // Never cache POSTs to the Apps Script backend — always go to network
  if (req.method !== "GET") return;

  const isHTML = req.mode === "navigate" || req.destination === "document" || req.url.endsWith(".html") || req.url.endsWith("/");

  if (isHTML) {
    // Network-first for the app page itself: always get the latest version when
    // online (so new features show up immediately without needing a cache bump).
    // Falls back to the last cached copy only when offline.
    event.respondWith(
      fetch(req)
        .then((res) => {
          if (res.ok) {
            const resClone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone));
          }
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

  // Cache-first for static assets (icons, manifest) — fine to reuse these
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req).then((res) => {
        if (res.ok && req.url.startsWith(self.location.origin)) {
          const resClone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone));
        }
        return res;
      }).catch(() => cached);
    })
  );
});
