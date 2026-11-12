var V="incompany-v1",SHELL=["./","index.html","manifest.webmanifest","icons/icon-192.png","icons/icon-512.png"];
self.addEventListener("install",function(e){e.waitUntil(caches.open(V).then(function(c){return c.addAll(SHELL);}).then(function(){return self.skipWaiting();}));});
self.addEventListener("activate",function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==V;}).map(function(k){return caches.delete(k);}));}).then(function(){return self.clients.claim();}));});
self.addEventListener("fetch",function(e){
  var u=new URL(e.request.url);
  if(e.request.method!=="GET"||u.origin!==location.origin)return; // API do Google nunca passa pelo cache
  e.respondWith(fetch(e.request).then(function(r){var c=r.clone();caches.open(V).then(function(x){x.put(e.request,c);});return r;}).catch(function(){return caches.match(e.request).then(function(r){return r||caches.match("index.html");});}));
});
