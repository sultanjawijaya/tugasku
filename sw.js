// sw.js — Service worker sederhana.
// Menyimpan file pembungkus (ikon, halaman awal) agar aplikasi cepat terbuka.
// Data tugas TIDAK disimpan di sini; tetap diambil langsung dari Google Sheets.

const CACHE = 'tugasku-v1';
const FILE = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILE)));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) =>
    Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

// Ambil dari internet dulu; jika offline, pakai salinan cache
self.addEventListener('fetch', (e) => {
  if (new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});
