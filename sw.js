const CACHE='station-reporting-v2';
const FILES=['./','/stationwise/','/stationwise/index.html'];
self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)).catch(()=>{}));
  self.skipWaiting();
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch',e=>{
  // network-first so new versions always win; fall back to cache offline
  e.respondWith(fetch(e.request).catch(()=>caches.match(e.request)));
});
