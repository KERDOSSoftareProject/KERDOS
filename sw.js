// KERDOS offline shell.
//
// Keeps a copy of the app's own files (the page, its scripts, styles and
// OCR assets) in the browser's cache so KERDOS opens with no signal. Only
// same-origin GET requests are handled; every request to the backend
// passes straight through untouched. Built assets carry a content hash
// in their name, so serving them cache-first is always correct; the page
// itself is fetched fresh when possible so a new deployment is picked up
// on the next online load.
const CACHE = "kerdos-shell-v1";

self.addEventListener("install", () => self.skipWaiting());

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  if (new URL(request.url).origin !== self.location.origin) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then(response => { keep(request, response.clone()); return response; })
        .catch(() => caches.match(request).then(hit => hit || caches.match(new URL(self.registration.scope).pathname)))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then(hit => hit || fetch(request).then(response => {
      if (response.ok) keep(request, response.clone());
      return response;
    }))
  );
});

function keep(request, response) {
  caches.open(CACHE).then(cache => cache.put(request, response)).catch(() => {});
}
