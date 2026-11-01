// Service worker tối giản: chỉ lưu "vỏ" ứng dụng để mở nhanh/cài được lên màn hình chính.
// Không bao giờ lưu dữ liệu /api (file của người dùng).
const CACHE = "pdf-giao-an-v1";
const SHELL = ["/", "/static/manifest.json", "/static/icon-192.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.pathname.startsWith("/api/")) return;
  e.respondWith(fetch(e.request).then(r => { if (u.origin === location.origin && r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); } return r; })
    .catch(() => caches.match(e.request)));
});
