const CACHE_NAME='postecsa-om-v62-1782477628';
self.addEventListener('install',e=>{self.skipWaiting();});
self.addEventListener('activate',e=>{
  e.waitUntil(
    caches.keys()
      .then(ks=>Promise.all(ks.map(k=>caches.delete(k))))
      .then(()=>self.clients.claim())
  );
});
// No cachear nada — siempre red
self.addEventListener('fetch',e=>{
  e.respondWith(fetch(e.request).catch(()=>new Response('offline',{status:503})));
});
