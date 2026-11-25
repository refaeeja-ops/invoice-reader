/* قارئ الفواتير: يعمل بدون إنترنت. ملفات التطبيق: الشبكة أولًا (لتصل التحديثات فورًا) ثم النسخة المحفوظة.
   ملفات vendor الكبيرة (قارئ النص واللغات): المحفوظة أولًا. الخطوط: المحفوظة مع تحديثها في الخلفية. */
const APP = 'ir-app-v1', VENDOR = 'ir-vendor-v1', FONTS = 'ir-fonts-v1';
const SHELL = ['./', 'index.html', 'app.js', 'parse.js', 'export.js', 'digits.js', 'config.js', 'ref.js', 'store.js', 'manifest.webmanifest', 'icon.svg', 'icon-192.png'];
const VEND = ['vendor/pdf.min.js', 'vendor/pdf.worker.min.js', 'vendor/jsQR.js', 'vendor/tesseract.min.js', 'vendor/worker.min.js', 'vendor/digits-model.json'];
self.addEventListener('install', e => {
  e.waitUntil(Promise.all([caches.open(APP).then(c => c.addAll(SHELL)), caches.open(VENDOR).then(c => c.addAll(VEND))]).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => ![APP, VENDOR, FONTS].includes(k)).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin) {
    if (url.pathname.includes('/vendor/')) {
      e.respondWith(caches.open(VENDOR).then(c => c.match(req).then(hit => hit || fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; }))));
      return;
    }
    e.respondWith(caches.open(APP).then(c => fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; })
      .catch(() => c.match(req, { ignoreSearch: true }).then(hit => hit || (req.mode === 'navigate' ? c.match('index.html') : Response.error())))));
    return;
  }
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    e.respondWith(caches.open(FONTS).then(c => c.match(req).then(hit => {
      const net = fetch(req).then(r => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    })));
  }
});
