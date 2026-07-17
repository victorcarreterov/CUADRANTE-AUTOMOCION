const CACHE_NAME = 'mora-app-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icono.png'
];

// Instalar el Service Worker y almacenar en caché los archivos locales
self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    }).then(() => self.skipWaiting())
  );
});

// Activar el Service Worker y limpiar cachés antiguas
self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Estrategia de red con caída a caché para los recursos de la PWA
self.addEventListener('fetch', (e) => {
  // No intentar cachear las peticiones externas de Google Apps Script
  if (e.request.url.includes('script.google.com')) {
    return;
  }
  
  e.respondWith(
    fetch(e.request).catch(() => {
      return caches.match(e.request);
    })
  );
});