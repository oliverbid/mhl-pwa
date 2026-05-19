// MHL Load Planner — Service Worker
// Network-first strategy: always fetches fresh content, falls back to cache offline.

const CACHE = 'mhl-lp-v2';

self.addEventListener('install', () => self.skipWaiting());

self.addEventListener('activate', event => {
  // Remove old caches and take control immediately
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // Network-first: try the network, cache the result, fall back to cache if offline
  event.respondWith(
    fetch(event.request, { cache: 'no-cache' })
      .then(response => {
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE).then(c => c.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
