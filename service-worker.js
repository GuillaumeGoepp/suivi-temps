const CACHE_NAME = 'cdt-suivi-v2';
const ASSETS_TO_CACHE = [
  'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono:wght@500&display=swap',
  'https://fonts.gstatic.com'
];

self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  
  // Ne jamais intercepter les appels GAS
  if (url.includes('script.google.com')) return;
  
  // Ne jamais intercepter index.html (pour préserver le token dans l'URL)
  if (url.includes('index.html') || url.endsWith('/suivi-temps/') || url.endsWith('/suivi-temps/?')) return;
  
  // Cacher uniquement les polices Google
  if (url.includes('fonts.googleapis.com') || url.includes('fonts.gstatic.com')) {
    e.respondWith(
      caches.match(e.request).then(cached => cached || 
        fetch(e.request).then(response => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
          return response;
        })
      )
    );
  }
});
