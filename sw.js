/* ninetto-converter – service worker per l'uso offline.
   Percorsi relativi: funziona sia alla radice del sito sia in una sottocartella
   di GitHub Pages (es. /ninetto-converter/). Lo scope è la cartella di sw.js.
   Ad ogni modifica dei file dell'app aumenta VERSIONE: i telefoni scaricano la nuova versione. */
'use strict';

const VERSIONE = 'nc-v4';

// File dell'app (stessa origine)
const FILE_APP = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './manifest.webmanifest',
  './icon.svg',
  './icon-180.png',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png'
];

// Librerie da CDN: URL con versione fissa, quindi non cambiano mai → cache permanente
const FILE_CDN = [
  'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  // Strumenti Word ed Excel
  'https://cdn.jsdelivr.net/npm/docx-preview@0.4.1/dist/docx-preview.min.js',
  'https://cdn.jsdelivr.net/npm/html2pdf.js@0.14.0/dist/html2pdf.bundle.min.js',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/jspdf@4.2.1/dist/jspdf.umd.min.js',
  'https://cdn.jsdelivr.net/npm/jspdf-autotable@5.0.9/dist/jspdf.plugin.autotable.min.js',
  'https://unpkg.com/docx@8.5.0/build/index.umd.js'
];

self.addEventListener('install', (evento) => {
  evento.waitUntil((async () => {
    const cache = await caches.open(VERSIONE);
    await cache.addAll(FILE_APP);
    // Le librerie una per una: se il CDN non risponde l'installazione non fallisce,
    // verranno salvate al primo utilizzo.
    await Promise.all(FILE_CDN.map(async (url) => {
      try {
        const r = await fetch(url, { mode: 'cors' });
        if (r.ok) await cache.put(url, r);
      } catch (e) { /* riprova al prossimo uso */ }
    }));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (evento) => {
  evento.waitUntil((async () => {
    const nomi = await caches.keys();
    await Promise.all(nomi.filter((n) => n.startsWith('nc-') && n !== VERSIONE).map((n) => caches.delete(n)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (evento) => {
  const richiesta = evento.request;
  if (richiesta.method !== 'GET') return;
  const url = new URL(richiesta.url);

  // Librerie CDN: prima la cache, poi la rete (e si salva)
  if (FILE_CDN.includes(url.href)) {
    evento.respondWith((async () => {
      const cache = await caches.open(VERSIONE);
      const trovata = await cache.match(url.href);
      if (trovata) return trovata;
      const r = await fetch(url.href, { mode: 'cors' });
      if (r.ok) cache.put(url.href, r.clone());
      return r;
    })());
    return;
  }

  if (url.origin !== self.location.origin) return;

  // File dell'app: prima la rete (così vedi subito gli aggiornamenti), se offline la cache
  evento.respondWith((async () => {
    const cache = await caches.open(VERSIONE);
    try {
      const r = await fetch(richiesta);
      if (r.ok && url.pathname.startsWith(new URL('./', self.location).pathname)) {
        cache.put(richiesta.mode === 'navigate' ? './index.html' : richiesta, r.clone());
      }
      return r;
    } catch (e) {
      const trovata = await cache.match(richiesta, { ignoreSearch: true });
      if (trovata) return trovata;
      if (richiesta.mode === 'navigate') return cache.match('./index.html');
      throw e;
    }
  })());
});
