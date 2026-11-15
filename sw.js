// POSTECSA OM - Service Worker "red primero" (permite instalar la app y abrir sin conexion)
const CACHE='postecsa-om-v8';
self.addEventListener('install',e=>{ self.skipWaiting(); });
self.addEventListener('activate',e=>{
  e.waitUntil(
    caches.keys()
      .then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
      .then(()=>self.clients.claim())
  );
});
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET') return;
  const u=new URL(e.request.url);
  if(u.origin!==location.origin) return;   // GAS, CDN, etc.: directo a la red
  e.respondWith(
    fetch(e.request).then(r=>{
      if(r&&r.ok){ const c=r.clone(); caches.open(CACHE).then(ch=>ch.put(e.request,c)); }
      return r;
    }).catch(()=>caches.match(e.request).then(m=>m||new Response('Sin conexion',{status:503})))
  );
});
