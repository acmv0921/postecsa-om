// Redirección: este service worker se desinstala solo y limpia cachés (checklist v2.2 vive en /postecsa-checklist/)
self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil((async()=>{
  const ks=await caches.keys(); await Promise.all(ks.map(k=>caches.delete(k)));
  await self.registration.unregister();
  const cs=await self.clients.matchAll({type:'window'});
  cs.forEach(c=>c.navigate('https://acmv0921.github.io/postecsa-checklist/'));
})()));
