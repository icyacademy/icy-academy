// Guarda a app no telemóvel para abrir rápido e funcionar sem rede.
const VERSAO = 'icy-v18';
const FICHEIROS = ['./', './index.html', './config.js', './manifest.webmanifest', './logo-branco.png',
  './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => c.addAll(FICHEIROS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (url.hostname.includes('script.google')) return;
  // Rede primeiro (para receber atualizações), cache se não houver rede.
  e.respondWith(
    fetch(e.request).then(r => {
      if (r.ok && (url.origin === location.origin || url.hostname.includes('fonts.'))) {
        const copia = r.clone();
        caches.open(VERSAO).then(c => c.put(e.request, copia));
      }
      return r;
    }).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
  );
});
