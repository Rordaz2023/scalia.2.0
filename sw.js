/* sw.js — "interruptor de apagado".
   Súbelo en lugar de tu sw.js actual. Cuando un dispositivo lo reciba, borra todas las cachés viejas,
   se desinstala solo y recarga la página para que cargue la versión real del servidor. */
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) {
    e.waitUntil((async function () {
        try { var ks = await caches.keys(); await Promise.all(ks.map(function (k) { return caches.delete(k); })); } catch (x) {}
        try { await self.registration.unregister(); } catch (x) {}
        try { var cs = await self.clients.matchAll({ type: 'window' }); cs.forEach(function (c) { c.navigate(c.url); }); } catch (x) {}
    })());
});
/* Mientras siga activo, nunca sirve nada guardado */
self.addEventListener('fetch', function (e) { e.respondWith(fetch(e.request, { cache: 'no-store' })); });
