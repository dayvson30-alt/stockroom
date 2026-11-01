const V = 'stockroom-v5';
const SHELL = ['./', './index.html', './painel.html', './config.js', './manifest.json', './html5-qrcode.min.js', './icon-192.png', './icon-512.png', './ff-logo.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))));
  self.clients.claim();
});
// abre do cache na hora (sinal fraco não pode travar a abertura) e atualiza por trás: versão nova vale na próxima abertura
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  const update = fetch(e.request).then(r => {
    if (r.ok) { const cp = r.clone(); caches.open(V).then(c => c.put(e.request, cp)); }
    return r;
  });
  e.waitUntil(update.catch(() => {}));
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then(hit => hit || update));
});
