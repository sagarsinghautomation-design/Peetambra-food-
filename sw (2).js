// P Food'z Cafe — service worker (v6)
// HTML: stale-while-revalidate (instant open, background update)
// Other same-origin: stale-while-revalidate
// Cross-origin (Apps Script API, Google Fonts, Chart.js): never touched
const CACHE_NAME = 'pfp-cafe-shell-v6';   // ← v5 से v6 — पुराना cache auto-clear
const APP_SHELL = ['./', './index.html', './manifest.json'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((c) => c.addAll(APP_SHELL)).catch(() => {}));
  self.skipWaiting();
});
self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))));
  self.clients.claim();
});

// HTML: cache-first, background refresh — app turant khulta hai
function staleWhileRevalidate(req) {
  return caches.open(CACHE_NAME).then((cache) =>
    cache.match(req).then((cached) => {
      const net = fetch(req).then((res) => {
        if (res && res.status === 200) cache.put(req, res.clone());
        return res;
      }).catch(() => cached || caches.match('./index.html'));
      return cached || net;
    })
  );
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  const isHtml = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('.html');
  if (isHtml) { event.respondWith(staleWhileRevalidate(req)); return; }
  event.respondWith(
    caches.match(req).then((cached) => {
      const net = fetch(req).then((res) => {
        if (res && res.status === 200) { const clone = res.clone(); caches.open(CACHE_NAME).then((c) => c.put(req, clone)); }
        return res;
      }).catch(() => cached);
      return cached || net;
    })
  );
});