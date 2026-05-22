const CACHE_VERSION = "floodcontrol-v11";
const APP_SHELL_FILES = [
  "./",
  "./index.html",
  "./projects.html",
  "./activities.html",
  "./add-activity.html",
  "./cost-management.html",
  "./style.css",
  "./index.css",
  "./projects.css",
  "./activities.css",
  "./cost-management.css",
  "./professional-ui.css",
  "./ui-feedback.js",
  "./data-service.js",
  "./script.js",
  "./projects.js",
  "./activities.js",
  "./add-activity.js",
  "./cost-management.js",
  "./page-loader.js",
  "./manifest.webmanifest",
  "./assets/logo-uploads/app-logo-v20260522.png"
];

const STATIC_ASSET_DESTINATIONS = new Set(["style", "script", "manifest", "image", "font"]);

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

const cacheFreshResponse = async (request, response) => {
  if (
    response &&
    response.status === 200 &&
    response.type !== "opaque" &&
    !response.redirected
  ) {
    const responseClone = response.clone();
    const cache = await caches.open(CACHE_VERSION);
    await cache.put(request, responseClone);
  }
  return response;
};

const getCachedResponse = (request) => caches.match(request, { ignoreSearch: true });

self.addEventListener("fetch", (event) => {
  const { request } = event;
  const requestUrl = new URL(request.url);

  if (request.method !== "GET") return;

  // Always bypass caching for external API/data sources so dashboard values stay fresh.
  if (requestUrl.origin !== self.location.origin) {
    return;
  }

  const isNavigationRequest = request.mode === "navigate" || request.destination === "document";
  const isStaticAsset = STATIC_ASSET_DESTINATIONS.has(request.destination);

  if (isNavigationRequest) {
    event.respondWith(
      fetch(request, { cache: "no-cache" })
        .then((response) => cacheFreshResponse(request, response))
        .catch(async () => {
          const cachedPage = await getCachedResponse(request);
          if (cachedPage) return cachedPage;
          return caches.match("./index.html");
        })
    );
    return;
  }

  if (isStaticAsset) {
    event.respondWith(
      (async () => {
        const cachedResponse = await getCachedResponse(request);
        const networkFetch = fetch(request, { cache: "no-cache" })
          .then((response) => cacheFreshResponse(request, response))
          .catch(() => cachedResponse);

        return cachedResponse || networkFetch;
      })()
    );
    return;
  }

  event.respondWith(
    fetch(request, { cache: "no-cache" })
      .then((networkResponse) => cacheFreshResponse(request, networkResponse))
      .catch(() => getCachedResponse(request))
  );
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});
