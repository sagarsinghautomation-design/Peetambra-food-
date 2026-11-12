// P Food'z Cafe — service worker (v11)
// HTML: NETWORK-FIRST — har baar naya version, offline hone par hi cache se
// Baaki same-origin files: stale-while-revalidate
// Cross-origin (Render backend API, Google Fonts, Chart.js): kabhi cache nahi
const CACHE_NAME = 'pfp-cafe-shell-v11';   // version badalte hi purana cache auto-delete
const APP_SHELL = ['./', './index.html', './manifest.json'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((c) => c.addAll(APP_SHELL)).catch(() => {}));
  self.skipWaiting();
});
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function networkFirst(req) {
  return fetch(req, { cache: 'no-store' }).then((res) => {
    if (res && res.status === 200) { const clone = res.clone(); caches.open(CACHE_NAME).then((c) => c.put(req, clone)); }
    return res;
  }).catch(() => caches.match(req).then((c) => c || caches.match('./index.html')));
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  const isHtml = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('.html');
  if (isHtml || url.pathname.endsWith('sw.js')) { event.respondWith(networkFirst(req)); return; }
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
