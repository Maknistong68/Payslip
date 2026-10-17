/*
 * Payslip Generator service worker: caches every file of the app so it opens and works without
 * internet once it has been visited (or installed). Payroll data never passes through here; the app
 * makes no network requests. Bump VERSION whenever a file below changes (tools/test.js checks the list).
 */
'use strict';

var VERSION = '2026-10-17.1';
var CACHE = 'payslip-' + VERSION;
var ASSETS = [
  './',
  'index.html',
  'manifest.webmanifest',
  'css/styles.css',
  'css/fonts/Sora.ttf',
  'js/vendor/xlsx.full.min.js',
  'js/vendor/jspdf.umd.min.js',
  'js/vendor/jszip.min.js',
  'js/vendor/exceljs.min.js',
  'js/fonts/amiri.js',
  'js/schema.js',
  'js/xlsx-images.js',
  'js/pdf.js',
  'js/delivery.js',
  'js/template-xlsx.js',
  'js/art.js',
  'js/vault.js',
  'js/guide.js',
  'js/app.js',
  'js/pwa.js',
  'icons/icon.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png',
  'icons/apple-touch-icon.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(caches.open(CACHE).then(function (cache) {
    return cache.addAll(ASSETS.map(function (url) { return new Request(url, { cache: 'reload' }); }));
  }).then(function () {
    // the very first install takes over at once; later versions wait until the user reloads
    if (!self.registration.active) return self.skipWaiting();
  }));
});

self.addEventListener('activate', function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('payslip-') === 0 && k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('message', function (event) {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  var scope = new URL(self.registration.scope);
  var path = url.pathname.slice(scope.pathname.length);

  // the app shell: index.html for the start page, whatever the query string (?channels=all)
  if (req.mode === 'navigate' && (path === '' || path === 'index.html')) {
    event.respondWith(caches.match('index.html', { ignoreSearch: true }).then(function (hit) {
      return hit || fetch(req);
    }).catch(function () { return fetch(req); }));
    return;
  }

  event.respondWith(caches.match(req, { ignoreSearch: true }).then(function (hit) {
    if (hit) return hit;
    return fetch(req).then(function (res) {
      if (res && res.ok && res.type === 'basic' && ASSETS.indexOf(path) >= 0) {
        var copy = res.clone();
        caches.open(CACHE).then(function (cache) { cache.put(req, copy); });
      }
      return res;
    });
  }));
});
