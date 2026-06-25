const CACHE_NAME = 'opaustro-unified-v6';
const APP_SHELL = [
  './',
  './index.html',
  './cabinets.html',
  './tablero.html',
  './revision.html',
  './ventas.html',
  './ventas-hub.html',
  './ventas-reporte.html',
  './flujo-caja-opaustro.html',
  './ventas_data.js',
  './manifest.webmanifest',
  './EMPRESA.jpeg',
  './logo_opaustro.png',
  './logo.png',
  './magnum.avif',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
    ))
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  event.respondWith(
    fetch(req).then(response => {
      const copy = response.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(req, copy));
      return response;
    }).catch(() => caches.match(req).then(cached => cached || caches.match('./index.html')))
  );
});
