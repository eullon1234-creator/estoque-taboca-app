const CACHE_NAME = 'eullon-app-v1';

// Assets to pre-cache on install for immediate offline fallback
const PRECACHE_ASSETS = [
    './',
    './index.html',
    './manifest.json',
    './assets/css/tailwind.css',
    './assets/css/main.css',
    './assets/js/main.js',
    './assets/icons/icon-192.png',
    './assets/icons/icon-512.png'
];

self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(PRECACHE_ASSETS))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
            )
        ).then(() => self.clients.claim())
    );
});

function isCriticalAsset(url) {
    return url.mode === 'navigate' ||
        /\.(html?|js|css|json)$/i.test(url.pathname);
}

function isStaticAsset(url) {
    return /\.(png|jpg|jpeg|gif|svg|ico|webp|woff2?|ttf|eot)$/i.test(url.pathname);
}

self.addEventListener('fetch', event => {
    if (event.request.method !== 'GET') return;

    const url = new URL(event.request.url);
    if (url.origin !== self.location.origin) {
        event.respondWith(fetch(event.request));
        return;
    }

    if (isCriticalAsset(event.request)) {
        event.respondWith(
            fetch(event.request)
                .then(response => {
                    if (response && response.status === 200 && response.type !== 'opaque') {
                        const clone = response.clone();
                        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone)).catch(() => {});
                    }
                    return response;
                })
                .catch(() => caches.match(event.request))
        );
    } else if (isStaticAsset(event.request)) {
        event.respondWith(
            caches.match(event.request).then(cached => {
                if (cached) return cached;
                return fetch(event.request).then(response => {
                    if (response && response.status === 200 && response.type !== 'opaque') {
                        const clone = response.clone();
                        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone)).catch(() => {});
                    }
                    return response;
                });
            })
        );
    } else {
        event.respondWith(
            caches.match(event.request).then(cached => {
                if (cached) return cached;
                return fetch(event.request).then(response => {
                    if (response && response.status === 200 && response.type !== 'opaque') {
                        const clone = response.clone();
                        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone)).catch(() => {});
                    }
                    return response;
                });
            })
        );
    }
});

self.addEventListener('message', event => {
    if (event.data === 'SKIP_WAITING') {
        self.skipWaiting();
    }
});
