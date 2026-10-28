// Service Worker checklist v1.0
const CACHE_NAME = 'checklist-cache-v1';
const PRECACHE = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(PRECACHE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== CACHE_NAME).map(x => caches.delete(x)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = e.request.url;
  if (e.request.method !== 'GET' || u.includes('script.google.com') || u.includes('docs.google.com') || u.includes('googleapis.com') || u.includes('wa.me')) return;
  e.respondWith(fetch(e.request).then(r => { if (r && r.status === 200) { const c = r.clone(); caches.open(CACHE_NAME).then(x => x.put(e.request, c)); } return r; }).catch(() => caches.match(e.request)));
});
