/* Service worker de Datos Fijos.
   Sirve para que el navegador permita INSTALAR la página como app.
   No guarda nada en caché: la página siempre se carga fresca desde el servidor.
   Solo si no hay internet muestra un aviso en lugar del error del navegador. */
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });

var SIN_RED = '<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>Sin conexión · Datos Fijos</title><style>body{margin:0;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;' +
  'font-family:system-ui,sans-serif;background:linear-gradient(#e2262c,#80101a);color:#fff;text-align:center;padding:24px}img{width:min(70vw,320px)}' +
  'button{font:700 1rem system-ui;padding:12px 26px;border:0;border-radius:999px;background:#ffd24a;color:#3a2600}</style></head><body>' +
  '<img src="icon-512.png" alt="Datos Fijos"><h1>Sin conexión</h1><p>Revisa tu internet y vuelve a intentar.</p>' +
  '<button onclick="location.reload()">Reintentar</button></body></html>';

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).catch(function () {
      return new Response(SIN_RED, { headers: { 'Content-Type': 'text/html; charset=utf-8' } });
    }));
  }
});
