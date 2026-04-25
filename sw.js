const CACHE_NAME = 'eullon-v1.9.4';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './manifest.json',
    './assets/css/tailwind.css',
    './assets/css/main.css',
    './assets/js/main.js',
    './assets/icons/icon-192.png',
    './assets/icons/icon-512.png'
];

// Instala e faz cache dos assets principais
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(ASSETS_TO_CACHE))
            .then(() => self.skipWaiting())
    );
});

// Remove caches antigos
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
            )
        ).then(() => self.clients.claim())
    );
});

// Páginas e assets versionados buscam a rede primeiro para receber mudanças do app sem reinstalar.
self.addEventListener('fetch', event => {
    if (event.request.method !== 'GET') return;

    const url = new URL(event.request.url);
    const shouldPreferNetwork = event.request.mode === 'navigate' || url.searchParams.has('v') || url.searchParams.has('atualizar');

    event.respondWith(
        (shouldPreferNetwork ? fetch(event.request).then(response => {
            if (response && response.status === 200 && response.type !== 'opaque') {
                const responseClone = response.clone();
                caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseClone));
            }
            return response;
        }).catch(() => caches.match(event.request)) : caches.match(event.request)).then(cached => {
            if (cached) return cached;
            return fetch(event.request).then(response => {
                if (!response || response.status !== 200 || response.type === 'opaque') {
                    return response;
                }
                const responseClone = response.clone();
                caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseClone));
                return response;
            });
        })
    );
});
// Escuta mensagens do cliente
self.addEventListener('message', event => {
    if (event.data === 'SKIP_WAITING') {
        self.skipWaiting();
    }
});
