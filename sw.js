const CACHE_NAME = 'equate-v1.2.0';

const PRECACHE_URLS = [
  './',
  './index.html',
  './css/styles.css',
  './vendor/xlsx.full.min.js',
  './vendor/plotly.min.js',
  './ESPP_nb.png',
  './visuate-rb.png',
  './icons/icon-192x192.png',
  './icons/icon-512x512.png',
  './js/config.js',
  './js/cacheManager.js',
  './js/translationData.js',
  './js/translations.js',
  './js/database.js',
  './js/currencyMappings.js',
  './js/currencyDetector.js',
  './js/currencyConverter.js',
  './js/historicalPriceGenerator.js',
  './js/fileAnalyzer.js',
  './js/fileParser.js',
  './js/investmentRules.js',
  './js/referencePointsBuilder.js',
  './js/calculator.js',
  './js/calculationsTab.js',
  './js/exporter.js',
  './js/exportCustomizer.js',
  './js/EventBus.js',
  './js/DOMManager.js',
  './js/CurrencyService.js',
  './js/app.js'
];

// Install: precache app shell
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
  );
});

// Activate: clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

// Fetch: route requests
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Skip non-GET requests
  if (event.request.method !== 'GET') return;

  // Skip cross-origin requests (analytics, external CDNs)
  if (url.origin !== location.origin) return;

  // Network-first for historical data files (updated daily)
  if (url.pathname.includes('hist_') && url.pathname.endsWith('.xlsx')) {
    event.respondWith(networkFirst(event.request));
    return;
  }

  // Cache-first for everything else (app shell, vendor libs, images)
  event.respondWith(cacheFirst(event.request));
});

async function cacheFirst(request) {
  const cached = await caches.match(request, { ignoreSearch: true });
  if (cached) return cached;

  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return new Response('Offline', { status: 503, statusText: 'Service Unavailable' });
  }
}

async function networkFirst(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request, { ignoreSearch: true });
    return cached || new Response('Offline', { status: 503, statusText: 'Service Unavailable' });
  }
}