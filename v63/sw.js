self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',()=>clients.claim());
// POSTECSA OM v6.3 — sin cache (siempre red)
