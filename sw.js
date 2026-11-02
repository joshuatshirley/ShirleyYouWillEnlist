"use strict";

// Bump this on every deploy that changes any precached file -- clients keep
// serving whatever this cache holds until the name changes, so a stale
// CACHE_NAME means an iPad that's already offline-installed never sees the
// update even after a new push to main.
var CACHE_NAME = "florence-est-v34";

var PRECACHE_URLS = [
  "./",
  "./index.html",
  "./est-kiosk-standalone.html",
  "./question-data.js",
  "./manifest.json",
  "./icons/icon-180.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./assets/army_logo_horiz.svg",
  "./assets/m2s_qr.png",
  "./assets/qrcode.js"
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function (cache) { return cache.addAll(PRECACHE_URLS); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (key) {
        if (key !== CACHE_NAME) return caches.delete(key);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

// Cache-first for everything the kiosk needs to run (this is a single
// self-contained HTML file with all test data embedded, so precaching it
// is all offline requires). Anything not precached -- e.g. the optional
// api/results POST -- just passes through to the network and fails
// normally when offline; renderResults() already handles that failure.
self.addEventListener("fetch", function (event) {
  if (event.request.method !== "GET") return;

  event.respondWith(
    caches.match(event.request).then(function (cached) {
      if (cached) return cached;
      return fetch(event.request).then(function (response) {
        if (response && response.ok && response.type === "basic") {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            cache.put(event.request, copy);
          });
        }
        return response;
      }).catch(function () { return cached; });
    })
  );
});
