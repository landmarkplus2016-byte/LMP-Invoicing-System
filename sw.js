// =============================================================================
// LMP Invoicing System — sw.js
// Service worker: caches all app shell files for offline support.
// CDN libraries (SheetJS, ExcelJS) are NOT cached — pulled fresh each session.
// =============================================================================

const CACHE = 'lmp-invoicing-v6';
const FILES = [
  './',
  './index.html',
  './poc-app.js',
  './tsr-app.js',
  './contractor-app.js',
  './finance-app.js',
  './styles.css',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './LMP Big Logo.jpg'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request))
  );
});
