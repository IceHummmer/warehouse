const CACHE='hh-warehouse-shell-v1';
self.addEventListener('install',event=>{event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(['/','/favicon.svg','/manifest.webmanifest'])));self.skipWaiting();});
self.addEventListener('activate',event=>{event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(key=>key!==CACHE).map(key=>caches.delete(key)))));self.clients.claim();});
self.addEventListener('fetch',event=>{const u=new URL(event.request.url);if(event.request.method!=='GET'||u.origin!==self.location.origin||u.pathname.startsWith('/api/'))return;event.respondWith(fetch(event.request).catch(()=>caches.match(event.request)));});
