// ══════════════════════════════════════════════════════════
// SERVICE WORKER — POSTECSA OM v5
// Estrategia: Network First con fallback a cache
// Siempre intenta red primero → actualización inmediata
// ══════════════════════════════════════════════════════════

const CACHE_NAME = 'postecsa-om-v6-1781837208';
const CACHE_PREV = 'postecsa-om-v';

// Al instalar: cachear el index.html y manifest
self.addEventListener('install', event => {
  console.log('[SW] Instalando v5...');
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(['./index.html', './manifest.json']);
    })
  );
  // Activar inmediatamente sin esperar
  self.skipWaiting();
});

// Al activar: eliminar caches anteriores
self.addEventListener('activate', event => {
  console.log('[SW] Activando, limpiando caches viejos...');
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.filter(k => k.startsWith(CACHE_PREV) && k !== CACHE_NAME)
            .map(k => {
              console.log('[SW] Eliminando cache viejo:', k);
              return caches.delete(k);
            })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch: Network First — siempre intenta red, cache como respaldo
self.addEventListener('fetch', event => {
  // Solo manejar requests del mismo origen
  if (!event.request.url.startsWith(self.location.origin)) return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Si la red responde, actualizar cache y devolver
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // Sin red: usar cache
        return caches.match(event.request).then(cached => {
          if (cached) return cached;
          // Si no hay cache para esta URL, devolver index.html (SPA fallback)
          return caches.match('./index.html');
        });
      })
  );
});

// Escuchar mensajes para forzar actualización
self.addEventListener('message', event => {
  if (event.data === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
