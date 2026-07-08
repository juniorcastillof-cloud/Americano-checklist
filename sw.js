/**
 * sw.js — minimal offline app-shell cache.
 * Caches the app's own files on install so the checklist still opens with
 * no signal on a jobsite. Excel export / OneDrive sync still require a
 * connection when it's time to actually push data up.
 */
const CACHE_NAME = 'americano-checklist-v1';
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './css/styles.css',
  './js/checklist-data.js',
  './js/i18n.js',
  './js/storage.js',
  './js/graph.js',
  './js/export.js',
  './js/app.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  // Network-first for CDN libraries (SheetJS/MSAL) and Graph API calls;
  // cache-first for same-origin app shell files so the app opens offline.
  const url = new URL(event.request.url);
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(event.request).then((cached) => cached || fetch(event.request))
    );
  }
});
