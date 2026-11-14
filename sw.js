/* ninetto-converter – service worker per l'uso offline.
   Percorsi relativi: funziona sia alla radice del sito sia in una sottocartella
   di GitHub Pages (es. /ninetto-converter/). Lo scope è la cartella di sw.js.
   Ad ogni modifica dei file dell'app aumenta VERSIONE: i telefoni scaricano la nuova versione. */
'use strict';

const VERSIONE = 'nc-v10';

// File dell'app (stessa origine): prima la rete, se offline la cache
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

// Librerie incluse nell'app (cartella lib/, versioni fissate): nessun CDN esterno.
// Non cambiano tra un avvio e l'altro: prima la cache (avvio veloce e offline), poi la rete.
const FILE_LIB = [
  './lib/pdf-lib.min.js',
  './lib/pdf.min.js',
  './lib/pdf.worker.min.js',
  './lib/jszip.min.js',
  // Strumenti Word ed Excel
  './lib/docx-preview.min.js',
  './lib/html2pdf.bundle.min.js',
  './lib/xlsx.full.min.js',
  './lib/jspdf.umd.min.js',
  './lib/jspdf.plugin.autotable.min.js',
  './lib/docx.umd.js'
];
const URL_LIB = new Set(FILE_LIB.map((f) => new URL(f, self.location).href));

self.addEventListener('install', (evento) => {
  evento.waitUntil((async () => {
    const cache = await caches.open(VERSIONE);
    // App e librerie principali: devono esserci tutte, altrimenti l'installazione si ripete
    // cache: 'reload' = sempre dal server, mai una copia vecchia dalla cache HTTP del browser
    // (GitHub Pages fa tenere i file al browser per 10 minuti)
    const nuovo = (f) => new Request(f, { cache: 'reload' });
    await cache.addAll(FILE_APP.concat(FILE_LIB.slice(0, 4)).map(nuovo));
    // Librerie Word/Excel: se una non arriva, verrà salvata al primo utilizzo
    await Promise.all(FILE_LIB.slice(4).map((f) => cache.add(nuovo(f)).catch(() => {})));
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
  if (url.origin !== self.location.origin) return;

  // Librerie: prima la cache, poi la rete (e si salva)
  if (URL_LIB.has(url.origin + url.pathname)) {
    evento.respondWith((async () => {
      const cache = await caches.open(VERSIONE);
      const trovata = await cache.match(url.origin + url.pathname);
      if (trovata) return trovata;
      const r = await fetch(new Request(url.origin + url.pathname, { cache: 'reload' }));
      if (r.ok) cache.put(url.origin + url.pathname, r.clone());
      return r;
    })());
    return;
  }

  // File dell'app: prima la rete (così vedi subito gli aggiornamenti), se offline la cache
  evento.respondWith((async () => {
    const cache = await caches.open(VERSIONE);
    try {
      // 'no-cache': il browser chiede sempre al server se il file è cambiato (risposta 304 se è uguale)
      const r = await fetch(new Request(richiesta.url, { cache: 'no-cache', credentials: 'same-origin' }));
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
