// sw.js — Guarda solo la "cáscara" de la app para que abra rápido y sin señal,
// y muestra los recordatorios de medicamentos. Nunca guarda respuestas del servidor: son datos de salud.
const VERSION = 'ms-pacientes-v7';
const CASCARA = ['./', './index.html', './config.js', './manifest.webmanifest',
  './icono-192.png', './icono-512.png', './apple-touch-icon.png', './favicon.png',
  './logo-ms.png', './logo-horizontal.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(CASCARA)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return; // API y fuentes: siempre red
  // Red primero (para recibir versiones nuevas), caché si no hay señal.
  e.respondWith(
    fetch(e.request)
      .then(r => { const copia = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copia)); return r; })
      .catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
  );
});

// ── Notificaciones ──
self.addEventListener('push', e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch { d = { titulo: 'Dr. Miguel Soto', cuerpo: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.titulo || 'Dr. Miguel Soto', {
    body: d.cuerpo || '',
    icon: './icono-192.png',
    badge: './favicon.png',
    tag: d.toma ? 'toma-' + d.toma : (d.tipo || 'aviso'),
    renotify: true,
    requireInteraction: d.tipo === 'toma',
    data: { url: d.url || './' },
  }));
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const destino = new URL((e.notification.data && e.notification.data.url) || './', self.location.href).href;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(ventanas => {
    for (const v of ventanas) {
      if (v.url.startsWith(self.registration.scope)) { v.navigate(destino).catch(() => {}); return v.focus(); }
    }
    return self.clients.openWindow(destino);
  }));
});
