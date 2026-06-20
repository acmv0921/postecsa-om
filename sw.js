// POSTECSA OM — Service Worker v6.1
const CACHE_NAME = 'postecsa-om-v61-1781920798';

self.addEventListener('install', e => {
  console.log('[SW] Instalando cache:', CACHE_NAME);
  e.waitUntil(
    caches.open(CACHE_NAME).then(c => c.addAll(['/', '/index.html', '/manifest.json']))
    .catch(err => console.warn('[SW] Cache parcial:', err))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  console.log('[SW] Activando, limpiando caches viejos...');
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => {
        console.log('[SW] Borrando cache viejo:', k);
        return caches.delete(k);
      }))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Network first — siempre intenta red primero
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
