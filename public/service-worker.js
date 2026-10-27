const CACHE_NAME = 'acceso-oeste-v20';
const ASSETS_TO_CACHE = [
  '/scan.html',
  '/css/scan.css',
  '/js/scan.js',
  '/assets/icon-192x192.png',
  '/assets/icon-512x512.png'
];

// Instalar el service worker y forzar activación inmediata
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
});

// Limpiar cachés antiguos y tomar control inmediatamente de los clientes
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Estrategia Network-First: Primero busca en red; si falla (offline), responde con caché
self.addEventListener('fetch', (event) => {
  // Ignorar llamadas locales de hardware (huella) o firebase backend
  if (event.request.url.includes('localhost:5050') || event.request.method !== 'GET') {
    return;
  }

  // Para peticiones HTML y JS: Network-First garantizado
  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        // Si la red responde bien, actualizamos la copia en caché silenciosamente
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      })
      .catch(() => {
        // En caso de corte de internet (Offline), usar caché local
        return caches.match(event.request);
      })
  );
});
