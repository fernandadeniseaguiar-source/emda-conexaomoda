/* ========================================
   EMDA - Service Worker
   Auto-update enabled
======================================== */

const CACHE_NAME = 'emda-talentos-v10';

const STATIC_ASSETS = [
    '/',
    '/index.html',
    '/css/style.css',
    '/js/app.js',
    '/manifest.json',
    '/img/logo-white.png',
    '/img/logo-dress.png',
    '/img/logo-text.png',
    '/img/icon-192.png',
    '/img/icon-512.png'
];

// ========================================
// Install - Força atualização imediata
// ========================================

self.addEventListener('install', (event) => {
    console.log('[SW] Instalando...');
    
    // Pula a espera e ativa imediatamente
    self.skipWaiting();
    
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('[SW] Cacheando arquivos');
                return cache.addAll(STATIC_ASSETS);
            })
    );
});

// ========================================
// Activate - Limpa caches antigos
// ========================================

self.addEventListener('activate', (event) => {
    console.log('[SW] Ativando...');
    
    event.waitUntil(
        caches.keys()
            .then((cacheNames) => {
                return Promise.all(
                    cacheNames
                        .filter((name) => name !== CACHE_NAME)
                        .map((name) => {
                            console.log('[SW] Deletando cache antigo:', name);
                            return caches.delete(name);
                        })
                );
            })
            .then(() => {
                // Força todas as abas a usar o novo SW
                return self.clients.claim();
            })
    );
});

// ========================================
// Fetch - Network first, fallback to cache
// ========================================

self.addEventListener('fetch', (event) => {
    // Ignora requisições não-GET
    if (event.request.method !== 'GET') return;
    
    event.respondWith(
        // Tenta buscar da rede primeiro
        fetch(event.request)
            .then((response) => {
                // Se conseguiu, atualiza o cache
                if (response.ok) {
                    const responseClone = response.clone();
                    caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, responseClone);
                    });
                }
                return response;
            })
            .catch(() => {
                // Se falhou (offline), busca do cache
                return caches.match(event.request);
            })
    );
});
