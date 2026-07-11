const CACHE_NAME = 'ctu-rota-v1.9.2';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest?v=1.9.2',
  './icons/header-logo.png?v=1.9.2',
  './icons/favicon.png?v=1.9.2',
  './icons/apple-touch-icon.png?v=1.9.2',
  './icons/icon-192.png?v=1.9.2',
  './icons/icon-512.png?v=1.9.2'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if(event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if(url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        if(response && response.ok){
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        }
        return response;
      })
      .catch(() => caches.match(event.request).then(cached => cached || caches.match('./index.html')))
  );
});
