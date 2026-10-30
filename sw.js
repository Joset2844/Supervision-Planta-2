// Supervisión de Planta: siempre busca la versión más nueva y funciona sin internet.
const CACHE = 'supervision-planta-v2';
const PRECARGA = [
  './',
  '/index.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches
      .open(CACHE)
      .then((c) =>
        Promise.all(
          PRECARGA.map((u) =>
            fetch(u, { mode: u.startsWith('http') ? 'no-cors' : 'same-origin', cache: 'no-store' })
              .then((r) => c.put(u, r))
              .catch(() => {})
          )
        )
      )
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches
      .keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (e) => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const propio = url.origin === self.location.origin;
  const libreria = /fonts\.googleapis\.com|fonts\.gstatic\.com/.test(url.host);
  if (!propio && !libreria) return; // Sheets, Supabase, etc.: siempre a la red

  if (propio) {
    // Archivos de la app: red primero (versión nueva), caché si no hay internet
    e.respondWith(
      fetch(req, { cache: 'no-store' })
        .then((r) => {
          const copia = r.clone();
          caches.open(CACHE).then((c) => c.put(req, copia));
          return r;
        })
        .catch(() => caches.match(req, { ignoreSearch: true }))
    );
  } else {
    // Librerías y fuentes: caché primero
    e.respondWith(
      caches.match(req).then(
        (hit) =>
          hit ||
          fetch(req).then((r) => {
            const copia = r.clone();
            caches.open(CACHE).then((c) => c.put(req, copia));
            return r;
          })
      )
    );
  }
});
