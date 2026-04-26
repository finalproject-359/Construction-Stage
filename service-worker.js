const CACHE_VERSION = "floodcontrol-v4";
const APP_SHELL_FILES = [
  "./",
  "./index.html",
  "./projects.html",
  "./activities.html",
  "./cost-management.html",
  "./reports.html",
  "./style.css",
  "./script.js",
  "./manifest.webmanifest"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(APP_SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_VERSION).map((key) => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  const requestUrl = new URL(request.url);

  if (request.method !== "GET") return;

  // Always bypass caching for external API/data sources so dashboard values stay fresh.
  if (requestUrl.origin !== self.location.origin) {
    return;
  }

  const isNavigationRequest = request.mode === "navigate" || request.destination === "document";

  if (isNavigationRequest) {
    event.respondWith(
      fetch(request, { cache: "no-store" })
        .then((response) => {
          if (response && response.ok && !response.redirected) {
            const responseClone = response.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(request, responseClone));
          }
          return response;
        })
        .catch(async () => {
          const cachedPage = await caches.match(request, { ignoreSearch: true });
          if (cachedPage) return cachedPage;
          return caches.match("./index.html");
        })
    );
    return;
  }

  event.respondWith(
    fetch(request)
      .then((networkResponse) => {
        if (
          networkResponse &&
          networkResponse.status === 200 &&
          networkResponse.type !== "opaque" &&
          !networkResponse.redirected
        ) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put(request, responseClone));
        }
        return networkResponse;
      })
      .catch(async () => {
        const cachedResponse = await caches.match(request, { ignoreSearch: true });
        return cachedResponse;
      })
  );
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});
