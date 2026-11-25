/* أُلغيت ميزة العمل دون اتصال. هذا الملف يزيل الـ Service Worker القديم ويمسح ذاكرته حتى يرى الجميع آخر نسخة من التطبيق. */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.map(k => caches.delete(k))))
    .then(() => self.registration.unregister())
    .then(() => self.clients.matchAll({ type: 'window' }))
    .then(cs => cs.forEach(c => c.navigate(c.url))));
});
