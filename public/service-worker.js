const CACHE_NAME = 'acceso-oeste-v4';
const ASSETS_TO_CACHE = [
  '/scan.html',
  '/css/scan.css',
  '/js/scan.js',
  '/assets/icon-192x192.png',
  '/assets/icon-512x512.png'
];

// Instalar el service worker y cachear los archivos estáticos
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        return cache.addAll(ASSETS_TO_CACHE);
      })
  );
});

// Interceptar peticiones de red
self.addEventListener('fetch', (event) => {
  // Para las peticiones a la API local de huella, ignorar el caché
  if (event.request.url.includes('localhost:5050')) {
    return;
  }
  
  event.respondWith(
    caches.match(event.request)
      .then((cachedResponse) => {
        // Devuelve el archivo del caché si existe, sino lo busca en internet
        return cachedResponse || fetch(event.request);
      })
  );
});

// Limpiar cachés antiguos
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
    })
  );
});
