const C="mkhzn-shell-v5",F=["./","index.html","manifest.json","icon.svg"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(F)));self.skipWaiting()});
self.addEventListener("activate",e=>e.waitUntil(clients.claim()));
self.addEventListener("fetch",e=>{const u=new URL(e.request.url);
 if(/data\.json|version\.json|events|files\.json|main\.json|\/xlsx\//.test(u.pathname))return;   // البيانات وتحميل الإكسيل دايمًا من الجهاز الرئيسي (مش بيتخزّنوا)
 e.respondWith(fetch(e.request).then(r=>{const k=r.clone();caches.open(C).then(c=>c.put(e.request,k));return r}).catch(()=>caches.match(e.request)))});
