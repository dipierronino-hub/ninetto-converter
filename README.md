# ninetto-converter

PWA 100% client-side: Immagini in PDF, PDF in Immagini, Unisci PDF, Dividi PDF (per pagina, intervalli o ogni N pagine, con nome per ogni file), Comprimi PDF, Converti immagini, Word/Excel in PDF, Estrai in Word.
Nessun file lascia il dispositivo. Funziona offline dopo la prima apertura.

## File
- `index.html` struttura, `style.css` grafica (chiaro/scuro), `app.js` logica, `sw.js` offline
- `manifest.webmanifest` + icone (`icon.svg`, `icon-180/192/512.png`, `icon-maskable-512.png`)

## Configurazione
- Link "Apri Editor PDF": in `app.js`, `CONFIG.EDITOR_PDF_URL`.
- Livelli di compressione: `CONFIG.COMPRESSIONE` in `app.js`.
- Dopo ogni modifica ai file aumenta `VERSIONE` in `sw.js` (es. `nc-v2`), così i telefoni scaricano la nuova versione.

## Librerie (incluse nella cartella `lib/`, versioni fissate)
pdf-lib 1.17.1 · pdf.js 3.11.174 (+ worker) · JSZip 3.10.1 · docx-preview 0.4.1 · html2pdf.js 0.14.0 · SheetJS xlsx 0.18.5 · jsPDF 4.2.1 · jsPDF-AutoTable 5.0.9 · docx 8.5.0.
Sono copie identiche dei pacchetti npm ufficiali (licenze in `lib/LICENZE.txt`): l'app non dipende da nessun CDN esterno,
quindi funziona anche se un CDN è bloccato o irraggiungibile. Il service worker le salva per l'uso offline.
