# ninetto-converter

PWA 100% client-side: Immagini in PDF, PDF in Immagini, Unisci PDF, Comprimi PDF.
Nessun file lascia il dispositivo. Funziona offline dopo la prima apertura.

## File
- `index.html` struttura, `style.css` grafica (chiaro/scuro), `app.js` logica, `sw.js` offline
- `manifest.webmanifest` + icone (`icon.svg`, `icon-180/192/512.png`, `icon-maskable-512.png`)

## Configurazione
- Link "Apri Editor PDF": in `app.js`, `CONFIG.EDITOR_PDF_URL`.
- Livelli di compressione: `CONFIG.COMPRESSIONE` in `app.js`.
- Dopo ogni modifica ai file aumenta `VERSIONE` in `sw.js` (es. `nc-v2`), così i telefoni scaricano la nuova versione.

## Librerie (CDN cdnjs, versioni fissate)
pdf-lib 1.17.1 · pdf.js 3.11.174 · JSZip 3.10.1
