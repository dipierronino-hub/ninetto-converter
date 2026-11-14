/* =====================================================================
   ninetto-converter – logica dell'app
   Tutto avviene nel browser: i file non lasciano mai il dispositivo.
   Librerie (incluse nella cartella lib/, caricate in index.html):
     - pdf-lib  → window.PDFLib     (creare / unire PDF)
     - pdf.js   → window.pdfjsLib   (leggere e disegnare le pagine PDF)
     - JSZip    → window.JSZip      (creare archivi .zip)
   ===================================================================== */
'use strict';

/* ---------------------------------------------------------------------
   CONFIGURAZIONE – modifica qui
   --------------------------------------------------------------------- */
const VERSIONE_APP = 'nc-v10'; // uguale a VERSIONE in sw.js

const CONFIG = {
  // Indirizzo aperto dal pulsante "Apri Editor PDF" (es. 'https://tuosito.github.io/editor-pdf/').
  // Se resta vuoto il pulsante avvisa che l'indirizzo non è ancora impostato.
  EDITOR_PDF_URL: 'https://dipierronino-hub.github.io/editor-pdf/',
  // Librerie per Word ed Excel: si scaricano solo quando apri lo strumento che le usa
  // (copie locali in lib/ con versioni fissate: docx-preview 0.4.1, html2pdf.js 0.14.0, xlsx 0.18.5,
  //  jspdf 4.2.1, jspdf-autotable 5.0.9, docx 8.5.0; il service worker le salva per l'uso offline)
  LIB: {
    docxPreview: 'lib/docx-preview.min.js',
    html2pdf: 'lib/html2pdf.bundle.min.js',
    xlsx: 'lib/xlsx.full.min.js',
    jspdf: 'lib/jspdf.umd.min.js',
    autotable: 'lib/jspdf.plugin.autotable.min.js',
    docx: 'lib/docx.umd.js'
  },
  // Worker di pdf.js 3.11.174: stessa versione di lib/pdf.min.js
  PDFJS_WORKER: 'lib/pdf.worker.min.js',
  // Limite di pixel per un singolo canvas: oltre, i telefoni (soprattutto iPhone) si bloccano
  MAX_PIXEL: 16000000,
  MAX_LATO: 8192,
  A4: [595.28, 841.89],      // punti tipografici (1 pt = 1/72 di pollice)
  MARGINE_A4: 28.35,         // 1 cm
  // Livelli di compressione: qualità JPEG e risoluzione di ricampionamento
  COMPRESSIONE: {
    leggera: null,                       // senza perdita: il testo resta vero, si ottimizza solo la struttura
    bassa:   { qualita: 0.85, dpi: 200 },
    media:   { qualita: 0.70, dpi: 150 },
    massima: { qualita: 0.50, dpi: 100 }
  }
};

/* ---------------------------------------------------------------------
   Riferimenti al DOM
   --------------------------------------------------------------------- */
const $ = (id) => document.getElementById(id);
const el = {
  vistaHome: $('vistaHome'), vistaStrumento: $('vistaStrumento'),
  titolo: $('titoloStrumento'), descrizione: $('descrizioneStrumento'), icona: $('iconaStrumento'),
  areaFile: $('areaFile'), inputFile: $('inputFile'), testoArea: $('testoArea'), sottotestoArea: $('sottotestoArea'),
  contenitoreLista: $('contenitoreLista'), lista: $('listaFile'), conteggio: $('conteggioFile'),
  suggerimento: $('suggerimentoOrdine'), svuota: $('bottoneSvuota'),
  opzioni: $('pannelloOpzioni'), azione: $('bottoneAzione'),
  barra: $('barraAvanzamento'), riempimento: $('riempimento'), percentuale: $('percentuale'), testoAvanz: $('testoAvanzamento'),
  risultato: $('risultato'), toast: $('toast')
};

/* ---------------------------------------------------------------------
   Stato dell'app
   --------------------------------------------------------------------- */
const stato = {
  strumento: null,     // id dello strumento aperto
  voci: [],            // file caricati: { id, file, tipo, url, pagine, errore, password }
  lavoro: null,        // { annullato } dell'elaborazione in corso
  urlRisultato: null,  // blob URL del file pronto da scaricare
  prossimoId: 1
};

// Piano dello strumento «Dividi PDF»
const divisione = {
  voceId: null,          // PDF a cui si riferisce il piano
  chiave: null,          // modo + intervalli: se non cambia, il piano non si ridisegna
  parti: [],             // [{ da, a }] pagine 1..n incluse
  errore: null,          // intervalli scritti male
  nomi: new Map(),       // "da-a" → nome scelto dall'utente (resta anche cambiando modo)
  escluse: new Set(),    // "da-a" dei file senza spunta
  miniature: new Map(),  // numero pagina → { blob, url }
  timer: null
};
const chiaveParte = (p) => `${p.da}-${p.a}`;
let miniatureDivisioneInCorso = false;

/* ---------------------------------------------------------------------
   Definizione degli strumenti
   --------------------------------------------------------------------- */
const SCELTA = (nome, valore, testo, nota, scelto) =>
  `<label class="scelta"><input type="radio" name="${nome}" value="${valore}"${scelto ? ' checked' : ''}>` +
  `<span>${testo}${nota ? `<small>${nota}</small>` : ''}</span></label>`;

const STRUMENTI = {
  img2pdf: {
    titolo: 'Immagini in PDF',
    descrizione: 'JPG, PNG e WebP in un PDF unico multipagina o in PDF separati.',
    classe: 't-viola', tipo: 'immagine', minimo: 1,
    testoArea: 'Seleziona le immagini',
    accept: 'image/jpeg,image/png,image/webp,image/gif,image/bmp,image/avif,.jpg,.jpeg,.png,.webp,.gif,.bmp,.avif',
    opzioni: `
      <fieldset class="gruppo-opzione"><legend>Orientamento pagina</legend><div class="scelte">
        ${SCELTA('orientamento', 'auto', 'Automatico', 'segue la foto', true)}
        ${SCELTA('orientamento', 'verticale', 'Verticale')}
        ${SCELTA('orientamento', 'orizzontale', 'Orizzontale')}
      </div></fieldset>
      <fieldset class="gruppo-opzione"><legend>Dimensione pagina</legend><div class="scelte">
        ${SCELTA('formato', 'a4m', 'A4 con margini', 'bordo bianco di 1 cm', true)}
        ${SCELTA('formato', 'a4', 'A4 senza margini')}
        ${SCELTA('formato', 'adatta', "Adatta all'immagine", "la pagina prende la forma dell'immagine")}
      </div></fieldset>
      <fieldset class="gruppo-opzione"><legend>File da creare</legend><div class="scelte">
        ${SCELTA('unione', 'unico', 'Un PDF unico', 'tutte le immagini, una per pagina', true)}
        ${SCELTA('unione', 'separati', 'Un PDF per immagine', 'scaricati insieme in un .ZIP')}
      </div></fieldset>`,
    etichetta: () => (valoreOpzione('unione') === 'separati' && stato.voci.length > 1) ? 'Crea PDF separati (ZIP)' : 'Converti in PDF',
    esegui: eseguiImmaginiInPdf
  },
  pdf2img: {
    titolo: 'PDF in Immagini',
    descrizione: 'Ogni pagina diventa un\'immagine. Più pagine? Le ricevi in un unico .ZIP.',
    classe: 't-arancio', tipo: 'pdf', minimo: 1,
    testoArea: 'Seleziona i PDF',
    accept: 'application/pdf,.pdf',
    opzioni: `
      <fieldset class="gruppo-opzione"><legend>Formato immagine</legend><div class="scelte">
        ${SCELTA('formatoImg', 'jpeg', 'JPEG', 'più leggero, ideale per foto', true)}
        ${SCELTA('formatoImg', 'png', 'PNG', 'nitido, ideale per testo e grafici')}
      </div></fieldset>
      <fieldset class="gruppo-opzione"><legend>Risoluzione</legend><div class="scelte">
        ${SCELTA('dpi', '150', 'Standard', '150 DPI', true)}
        ${SCELTA('dpi', '300', 'Alta definizione', '300 DPI, file più pesanti')}
      </div></fieldset>`,
    etichetta: () => totalePagine() === 1 ? 'Scarica immagine' : 'Scarica ZIP',
    esegui: eseguiPdfInImmagini
  },
  unisci: {
    titolo: 'Unisci PDF',
    descrizione: 'Combina 2 o più PDF. Le pagine mantengono formato e orientamento originali.',
    classe: 't-azzurro', tipo: 'pdf', minimo: 2,
    testoArea: 'Seleziona i PDF da unire',
    accept: 'application/pdf,.pdf',
    opzioni: '',
    etichetta: () => stato.voci.length < 2 ? 'Aggiungi almeno 2 PDF' : `Unisci ${stato.voci.length} PDF`,
    esegui: eseguiUnisci
  },
  comprimi: {
    titolo: 'Comprimi PDF',
    descrizione: 'Riduci il peso ricampionando le pagine. Perfetto per email e PEC.',
    classe: 't-verde', tipo: 'pdf', minimo: 1,
    testoArea: 'Seleziona i PDF da comprimere',
    accept: 'application/pdf,.pdf',
    opzioni: `
      <fieldset class="gruppo-opzione"><legend>Livello di compressione</legend><div class="scelte">
        ${SCELTA('livello', 'leggera', 'Leggera · mantiene il testo', 'nessuna perdita di qualità: ottimizza la struttura del file')}
        ${SCELTA('livello', 'bassa', 'Bassa compressione', 'alta qualità visiva · JPEG 85%')}
        ${SCELTA('livello', 'media', 'Media · consigliata', 'per email · JPEG 70% a 150 DPI', true)}
        ${SCELTA('livello', 'massima', 'Massima compressione', 'file minimo · JPEG 50%')}
      </div>
      <p class="nota-opzione">Con i livelli Bassa, Media e Massima le pagine diventano immagini: il testo non sarà più selezionabile né ricercabile. Il livello Leggera conserva tutto ma riduce meno. Se il risultato non è più leggero, te lo segnalo e ti conviene tenere l'originale.</p>
      </fieldset>`,
    etichetta: () => stato.voci.length > 1 ? `Comprimi ${stato.voci.length} PDF (ZIP)` : 'Comprimi PDF',
    esegui: eseguiComprimi
  },
  docx2pdf: {
    titolo: 'Word in PDF',
    descrizione: 'Converte documenti .docx in PDF impaginati, pagina per pagina.',
    classe: 't-blu', tipo: 'docx', minimo: 1,
    testoArea: 'Seleziona i file Word',
    accept: '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    opzioni: `
      <fieldset class="gruppo-opzione"><legend>Metodo</legend><div class="scelte">
        ${SCELTA('metodoWord', 'stampa', 'PDF di qualità · consigliato', 'testo vero e nitido, margini e salti pagina di Word. Si apre la finestra di stampa: scegli "Salva come PDF"', true)}
        ${SCELTA('metodoWord', 'immagine', 'Download diretto', 'il PDF si scarica subito, ma ogni pagina è un\'immagine (testo non selezionabile)')}
      </div>
      <p class="nota-opzione">Funziona solo con .docx (i vecchi .doc vanno prima salvati come .docx). I font che non sono sul telefono vengono sostituiti con font simili.</p>
      </fieldset>`,
    etichetta: () => valoreOpzione('metodoWord') === 'immagine'
      ? (stato.voci.length > 1 ? `Converti ${stato.voci.length} file (ZIP)` : 'Scarica PDF')
      : 'Crea PDF (Salva come PDF)',
    esegui: eseguiWordInPdf
  },
  xlsx2pdf: {
    titolo: 'Excel in PDF',
    descrizione: 'Trasforma fogli .xlsx, .xls o .csv in tabelle PDF nitide e leggibili.',
    classe: 't-smeraldo', tipo: 'foglio', minimo: 1, singolo: true,
    testoArea: 'Seleziona il file Excel o CSV',
    accept: '.xlsx,.xls,.csv,.ods,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv',
    opzioni: `
      <div class="gruppo-opzione" id="gruppoFoglio" hidden>
        <label class="etichetta" for="sceltaFoglio">Foglio da convertire</label>
        <select id="sceltaFoglio" class="selettore"></select>
        <div id="anteprimaFoglio" class="anteprima-foglio"></div>
      </div>
      <fieldset class="gruppo-opzione"><legend>Orientamento pagina A4</legend><div class="scelte">
        ${SCELTA('orientamentoXls', 'orizzontale', 'Orizzontale', 'consigliato per fogli larghi', true)}
        ${SCELTA('orientamentoXls', 'verticale', 'Verticale')}
      </div></fieldset>
      <fieldset class="gruppo-opzione"><legend>Larghezza colonne</legend><div class="scelte">
        ${SCELTA('adatta', 'si', 'Adatta alla pagina', 'tutte le colonne in larghezza, testo ridotto se serve', true)}
        ${SCELTA('adatta', 'no', 'Dimensione naturale', 'le colonne in eccesso continuano nella pagina dopo')}
      </div></fieldset>`,
    etichetta: () => 'Converti in PDF',
    alCambioOpzioni: aggiornaAnteprimaFoglio,
    esegui: eseguiExcelInPdf
  },
  img2img: {
    titolo: 'Converti immagini',
    descrizione: 'Cambia formato (JPG, PNG, WebP), ridimensiona e alleggerisci una o tante immagini insieme.',
    classe: 't-rosa', tipo: 'immagine', minimo: 1,
    testoArea: 'Seleziona le immagini',
    accept: 'image/jpeg,image/png,image/webp,image/gif,image/bmp,image/avif,.jpg,.jpeg,.png,.webp,.gif,.bmp,.avif',
    opzioni: `
      <fieldset class="gruppo-opzione"><legend>Formato di destinazione</legend><div class="scelte">
        ${SCELTA('formatoDest', 'jpeg', 'JPG', 'il più compatibile, ideale per foto', true)}
        ${SCELTA('formatoDest', 'png', 'PNG', 'senza perdita, mantiene la trasparenza')}
        ${SCELTA('formatoDest', 'webp', 'WebP', 'moderno e leggerissimo')}
      </div></fieldset>
      <fieldset class="gruppo-opzione"><legend>Dimensione</legend><div class="scelte">
        ${SCELTA('lato', '0', 'Originale', 'nessun ridimensionamento', true)}
        ${SCELTA('lato', '2560', 'Grande', 'lato lungo 2560 px')}
        ${SCELTA('lato', '1600', 'Media', 'lato lungo 1600 px · per email')}
        ${SCELTA('lato', '1080', 'Piccola', 'lato lungo 1080 px · per WhatsApp e social')}
      </div></fieldset>
      <fieldset class="gruppo-opzione" id="gruppoQualita"><legend>Qualità (JPG e WebP)</legend><div class="scelte">
        ${SCELTA('qualitaImg', '0.92', 'Alta', 'differenza invisibile', true)}
        ${SCELTA('qualitaImg', '0.8', 'Bilanciata', 'file più leggeri')}
        ${SCELTA('qualitaImg', '0.65', 'Leggera', 'massimo risparmio')}
      </div></fieldset>`,
    etichetta: () => {
      const f = (valoreOpzione('formatoDest') || 'jpeg').replace('jpeg', 'jpg').toUpperCase();
      return stato.voci.length > 1 ? `Converti ${stato.voci.length} immagini in ${f} (ZIP)` : `Converti in ${f}`;
    },
    alCambioOpzioni: () => { const g = $('gruppoQualita'); if (g) g.hidden = valoreOpzione('formatoDest') === 'png'; },
    esegui: eseguiConvertiImmagini
  },
  dividi: {
    titolo: 'Dividi PDF',
    descrizione: 'Separa le pagine di un PDF in più file e dai a ciascuno il nome che vuoi.',
    classe: 't-ambra', tipo: 'pdf', minimo: 1, singolo: true,
    testoArea: 'Seleziona il PDF da dividere',
    accept: 'application/pdf,.pdf',
    opzioni: `
      <fieldset class="gruppo-opzione"><legend>Come dividere</legend><div class="scelte">
        ${SCELTA('modoDividi', 'pagina', 'Una pagina per file', 'ogni pagina diventa un PDF', true)}
        ${SCELTA('modoDividi', 'intervalli', 'Intervalli scelti', 'es. 1-3, 4, 5-8')}
        ${SCELTA('modoDividi', 'ogni', 'Ogni N pagine', 'blocchi di pagine uguali')}
      </div>
      <div class="campo-dividi" id="campoIntervalli" hidden>
        <label class="etichetta" for="testoIntervalli">Intervalli di pagine</label>
        <input id="testoIntervalli" class="selettore" type="text" inputmode="text" placeholder="es. 1-3, 4, 5-8" autocomplete="off" autocapitalize="off" spellcheck="false">
        <p class="nota-opzione">Un file per ogni intervallo, separati dalla virgola. «5-» vuol dire dalla pagina 5 all'ultima. Le pagine non indicate restano fuori.</p>
        <p class="errore-dividi" id="erroreIntervalli" hidden></p>
      </div>
      <div class="campo-dividi" id="campoOgni" hidden>
        <label class="etichetta" for="numeroOgni">Pagine per ogni file</label>
        <input id="numeroOgni" class="selettore" type="number" inputmode="numeric" min="1" step="1" value="2">
      </div>
      </fieldset>
      <fieldset class="gruppo-opzione" id="gruppoPiano" hidden><legend>File da creare e nomi</legend>
        <div class="serie-nomi">
          <input id="prefissoSerie" class="selettore" type="text" placeholder="Es. Fattura" autocomplete="off" spellcheck="false" aria-label="Nome comune per numerare tutti i file">
          <button id="applicaSerie" class="bottone-secondario" type="button">Numera tutti</button>
        </div>
        <p class="nota-opzione">Scrivi il nome di ogni file qui sotto, oppure un nome comune e tocca «Numera tutti» (Fattura_01, Fattura_02…). Togli la spunta ai file che non ti servono. Tocca la miniatura per vedere la pagina in grande.</p>
        <p id="riepilogoPiano" class="riepilogo-piano"></p>
        <ol id="pianoDivisione" class="piano-divisione"></ol>
      </fieldset>`,
    etichetta: () => {
      const v = stato.voci[0];
      if (!v || !v.pagine) return 'Lettura del PDF…';
      const n = partiIncluse().length;
      if (divisione.errore || !n) return 'Scegli le pagine da dividere';
      return n === 1 ? 'Crea il PDF' : `Dividi in ${n} PDF (ZIP)`;
    },
    alAggiorna: aggiornaPiano,
    alDigitare: (e) => { if (e.target.id === 'testoIntervalli' || e.target.id === 'numeroOgni') { clearTimeout(divisione.timer); divisione.timer = setTimeout(() => { aggiornaPiano(); aggiornaInterfaccia(); }, 250); } },
    pronto: () => !divisione.errore && partiIncluse().length > 0,
    esegui: eseguiDividi
  },
  pdf2docx: {
    titolo: 'Estrai in Word',
    descrizione: 'Estrae il testo del PDF in un documento .docx modificabile.',
    classe: 't-indaco', tipo: 'pdf', minimo: 1,
    testoArea: 'Seleziona i PDF',
    accept: 'application/pdf,.pdf',
    opzioni: `
      <div class="gruppo-opzione">
        <p class="nota-opzione" style="margin:0">Vengono estratti i paragrafi nell'ordine di lettura, con le dimensioni del testo e un salto pagina per ogni pagina del PDF. Immagini, colonne e tabelle complesse non vengono ricostruite. I PDF scansionati (solo immagini) non contengono testo da estrarre.</p>
      </div>`,
    etichetta: () => stato.voci.length > 1 ? `Estrai ${stato.voci.length} PDF (ZIP)` : 'Estrai in Word',
    esegui: eseguiPdfInWord
  }
};

/* =====================================================================
   AVVIO
   ===================================================================== */
const librerieMancanti = [['pdf-lib', window.PDFLib], ['pdf.js', window.pdfjsLib], ['JSZip', window.JSZip]].filter((x) => !x[1]).map((x) => x[0]);
const librerieOk = !librerieMancanti.length;
if (librerieOk) { try { sessionStorage.removeItem('nc-riparazione'); } catch (e) { /* ignora */ } }
else riparaLibrerie();

/** Librerie mancanti: quasi sempre è una copia vecchia rimasta in cache (browser o service worker).
    La prima volta svuota le cache dell'app e ricarica da GitHub Pages; se non basta mostra il motivo preciso. */
async function riparaLibrerie() {
  let giaProvato = false;
  try { giaProvato = sessionStorage.getItem('nc-riparazione') === '1'; sessionStorage.setItem('nc-riparazione', '1'); } catch (e) { giaProvato = true; }
  if (!giaProvato && navigator.onLine !== false) {
    try {
      if ('serviceWorker' in navigator) for (const r of await navigator.serviceWorker.getRegistrations()) await r.unregister();
      if (window.caches) for (const n of await caches.keys()) if (n.startsWith('nc-')) await caches.delete(n);
    } catch (e) { /* si ricarica comunque */ }
    location.reload();
    return;
  }
  const box = $('avvisoLibrerie');
  const dettagli = document.createElement('small');
  dettagli.style.cssText = 'display:block;margin-top:6px;opacity:.85;word-break:break-word';
  const errori = (window.__erroriCaricamento || []).slice(0, 4).join(' · ');
  dettagli.textContent = `Dettagli: mancano ${librerieMancanti.join(', ')}${errori ? ' – ' + errori : ''} – versione ${VERSIONE_APP} – ${navigator.userAgent}`;
  box.append(dettagli);
  box.hidden = false;
}

// pdf.js usa un "worker" separato. Lo scarichiamo con fetch (così passa dal service worker
// e funziona offline) e lo avviamo da un blob URL dello stesso sito.
const workerPronto = (async () => {
  if (!window.pdfjsLib) return;
  try {
    const risposta = await fetch(CONFIG.PDFJS_WORKER);
    if (!risposta.ok) throw new Error('HTTP ' + risposta.status);
    const blob = await risposta.blob();
    pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([blob], { type: 'text/javascript' }));
  } catch (e) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = CONFIG.PDFJS_WORKER;
  }
})();
// Un solo worker pdf.js riusato per tutti i documenti (più veloce e meno memoria)
let workerPdf = null;

impostaTema();
impostaEditor();
impostaInstallazione();
impostaAreaFile();
el.svuota.addEventListener('click', () => { svuotaVoci(); aggiornaInterfaccia(); });
el.azione.addEventListener('click', avviaElaborazione);
$('bottoneIndietro').addEventListener('click', () => { location.hash = ''; });
$('bottoneAnnulla').addEventListener('click', () => { if (stato.lavoro) { stato.lavoro.annullato = true; el.testoAvanz.textContent = 'Annullamento…'; } });
// errori imprevisti: un messaggio chiaro invece del silenzio
window.addEventListener('unhandledrejection', (e) => { console.error(e.reason); if (e.reason && e.reason.name !== 'Annullato') notifica('Si è verificato un errore imprevisto. Se continua, ricarica la pagina.', true); });
el.opzioni.addEventListener('change', () => {
  const s = STRUMENTI[stato.strumento];
  if (s && s.alCambioOpzioni) s.alCambioOpzioni();
  aggiornaInterfaccia();
});
el.opzioni.addEventListener('input', (e) => {
  const s = STRUMENTI[stato.strumento];
  if (s && s.alDigitare) s.alDigitare(e);
});
window.addEventListener('hashchange', instrada);
instrada();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then((r) => r.update()).catch(() => {}));
}

/* =====================================================================
   NAVIGAZIONE (#img2pdf, #pdf2img, #unisci, #comprimi, #dividi, #docx2pdf, #xlsx2pdf, #pdf2docx)
   ===================================================================== */
function instrada() {
  const id = location.hash.replace('#', '');
  if (stato.lavoro && id !== stato.strumento) {
    // Uscendo durante un'elaborazione la fermiamo
    stato.lavoro.annullato = true;
  }
  if (STRUMENTI[id]) apriStrumento(id);
  else mostraHome();
}

function mostraHome() {
  svuotaVoci();
  stato.strumento = null;
  el.vistaStrumento.hidden = true;
  el.vistaHome.hidden = false;
  document.title = 'ninetto-converter';
  window.scrollTo(0, 0);
}

function apriStrumento(id) {
  if (stato.strumento === id) return;
  svuotaVoci();
  stato.strumento = id;
  const s = STRUMENTI[id];
  el.vistaStrumento.className = 'vista ' + s.classe;
  el.titolo.textContent = s.titolo;
  el.descrizione.textContent = s.descrizione;
  el.icona.innerHTML = document.querySelector(`.tessera[href="#${id}"] .tessera-icona`).innerHTML;
  el.testoArea.textContent = s.testoArea;
  el.inputFile.accept = s.accept;
  el.opzioni.innerHTML = s.opzioni;
  el.vistaHome.hidden = true;
  el.vistaStrumento.hidden = false;
  document.title = s.titolo + ' · ninetto-converter';
  aggiornaInterfaccia();
  window.scrollTo(0, 0);
}

/* =====================================================================
   CARICAMENTO FILE
   ===================================================================== */
function impostaAreaFile() {
  el.inputFile.addEventListener('change', () => {
    aggiungiFile(el.inputFile.files);
    el.inputFile.value = ''; // permette di riselezionare lo stesso file
  });
  el.areaFile.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); el.inputFile.click(); }
  });
  ['dragenter', 'dragover'].forEach((t) => el.areaFile.addEventListener(t, (e) => {
    if (!e.dataTransfer || !Array.from(e.dataTransfer.types).includes('Files')) return;
    e.preventDefault(); el.areaFile.classList.add('sopra');
  }));
  ['dragleave', 'drop'].forEach((t) => el.areaFile.addEventListener(t, () => el.areaFile.classList.remove('sopra')));
  el.areaFile.addEventListener('drop', (e) => {
    if (!e.dataTransfer || !e.dataTransfer.files.length) return;
    e.preventDefault();
    aggiungiFile(e.dataTransfer.files);
  });
  // Evita che un file lasciato fuori dall'area venga aperto dal browser al posto dell'app
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
}

function tipoFile(file) {
  const nome = file.name.toLowerCase();
  if (file.type === 'application/pdf' || nome.endsWith('.pdf')) return 'pdf';
  if (/^image\/(jpeg|png|webp|gif|bmp|avif)$/.test(file.type) || /\.(jpe?g|png|webp|gif|bmp|avif)$/.test(nome)) return 'immagine';
  if (nome.endsWith('.docx')) return 'docx';
  if (/\.(xlsx|xls|csv|ods)$/.test(nome)) return 'foglio';
  return null;
}

function aggiungiFile(elencoFile) {
  if (!stato.strumento) instrada(); // l'hash è cambiato ma l'evento non è ancora arrivato
  if (!stato.strumento || stato.lavoro) return;
  const s = STRUMENTI[stato.strumento];
  const scartati = [];
  let nuovi = Array.from(elencoFile);
  if (s.singolo) {
    // Strumento a file singolo: il nuovo file sostituisce quello già caricato
    const valido = nuovi.find((f) => tipoFile(f) === s.tipo);
    nuovi.filter((f) => f !== valido).forEach((f) => { if (tipoFile(f) !== s.tipo) scartati.push(f.name); });
    nuovi = valido ? [valido] : [];
    if (valido) { stato.voci.forEach(liberaVoce); stato.voci = []; el.lista.replaceChildren(); }
  }
  for (const file of nuovi) {
    if (tipoFile(file) !== s.tipo) { scartati.push(file.name); continue; }
    const voce = { id: stato.prossimoId++, file, tipo: s.tipo, url: null, pagine: null, errore: null, password: null };
    voce.copia = copiaInMemoria(voce);
    stato.voci.push(voce);
    el.lista.appendChild(creaElementoVoce(voce));
  }
  if (scartati.length) {
    const atteso = { pdf: 'solo file PDF', immagine: 'solo immagini JPG, PNG, WebP, GIF, BMP o AVIF', docx: 'solo file Word .docx', foglio: 'solo file .xlsx, .xls, .csv o .ods' }[s.tipo];
    notifica(`Formato non supportato (${atteso}): ${scartati.join(', ')}`, true);
  }
  nascondiRisultato();
  aggiornaInterfaccia();
  generaAnteprime();
}

/**
 * Copia subito il contenuto del file in memoria.
 * Su Android, i file scelti da Galleria, Drive o Download sono "prestati" a Chrome solo per pochi istanti:
 * se si leggono più tardi compare l'errore "The requested file could not be read".
 */
async function copiaInMemoria(voce) {
  const f = voce.file;
  try {
    const byte = await f.arrayBuffer();
    voce.file = new File([byte], f.name, { type: f.type, lastModified: f.lastModified });
  } catch (e) {
    voce.errore = 'il telefono non permette di leggere questo file. Prova a scaricarlo prima nella memoria del telefono (es. cartella Download) e sceglilo da lì.';
  }
}

/* ---------- Elemento della lista con miniatura e comandi ---------- */
const ICONA = {
  su: '<svg viewBox="0 0 24 24"><path d="M6 15l6-6 6 6"/></svg>',
  giu: '<svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg>',
  rimuovi: '<svg viewBox="0 0 24 24"><path d="M5 7h14M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
  maniglia: '<svg viewBox="0 0 24 24"><circle cx="9" cy="6" r="1"/><circle cx="15" cy="6" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="9" cy="18" r="1"/><circle cx="15" cy="18" r="1"/></svg>',
  scarica: '<svg viewBox="0 0 24 24"><path d="M12 4v11m0 0-4-4m4 4 4-4M5 20h14"/></svg>',
  penna: '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9l-4-4L4 16v4zM14 6l4 4"/></svg>',
  spunta: '<svg viewBox="0 0 24 24"><path d="m5 12.5 4.5 4.5L19 7.5"/></svg>',
  condividi: '<svg viewBox="0 0 24 24"><circle cx="18" cy="5.5" r="2.6"/><circle cx="6" cy="12" r="2.6"/><circle cx="18" cy="18.5" r="2.6"/><path d="m8.4 13.3 7.2 4M15.6 6.8l-7.2 4"/></svg>',
  ricomincia: '<svg viewBox="0 0 24 24"><path d="M4 12a8 8 0 1 0 2.4-5.7"/><path d="M4 4v5h5"/></svg>'
};

function creaElementoVoce(voce) {
  const li = document.createElement('li');
  li.className = 'voce-file';
  li.dataset.id = voce.id;
  li.draggable = true;
  li.innerHTML = `
    <span class="maniglia" aria-hidden="true">${ICONA.maniglia}</span>
    <span class="numero"></span>
    <div class="miniatura"><div class="attesa"></div></div>
    <div class="info-file"><div class="nome"></div><div class="dettagli"></div></div>
    <div class="comandi-file">
      <button type="button" class="su" aria-label="Sposta su">${ICONA.su}</button>
      <button type="button" class="giu" aria-label="Sposta giù">${ICONA.giu}</button>
      <button type="button" class="rimuovi" aria-label="Rimuovi">${ICONA.rimuovi}</button>
    </div>`;
  li.querySelector('.nome').textContent = voce.file.name;
  li.querySelector('.nome').title = voce.file.name;
  li.querySelector('.su').addEventListener('click', () => sposta(voce.id, -1));
  li.querySelector('.giu').addEventListener('click', () => sposta(voce.id, +1));
  li.querySelector('.rimuovi').addEventListener('click', () => rimuovi(voce.id));
  abilitaTrascinamento(li);
  aggiornaDettagli(voce, li);
  return li;
}

function aggiornaDettagli(voce, li = trovaLi(voce.id)) {
  if (!li) return;
  const d = li.querySelector('.dettagli');
  d.classList.toggle('errore', !!voce.errore);
  if (voce.errore) { d.textContent = voce.errore; return; }
  let testo = formattaPeso(voce.file.size);
  if (voce.pagine) testo += ` · ${voce.pagine} ${voce.pagine === 1 ? 'pagina' : 'pagine'}`;
  if (voce.fogli) testo += ` · ${voce.fogli.length} ${voce.fogli.length === 1 ? 'foglio' : 'fogli'}`;
  if (voce.password) testo += ' · 🔒 sbloccato';
  d.textContent = testo;
}

const trovaLi = (id) => el.lista.querySelector(`[data-id="${id}"]`);

function sposta(id, delta) {
  if (stato.lavoro) return;
  const i = stato.voci.findIndex((v) => v.id === id);
  const j = i + delta;
  if (i < 0 || j < 0 || j >= stato.voci.length) return;
  [stato.voci[i], stato.voci[j]] = [stato.voci[j], stato.voci[i]];
  riordinaDom();
}

function rimuovi(id) {
  if (stato.lavoro) return;
  const i = stato.voci.findIndex((v) => v.id === id);
  if (i < 0) return;
  liberaVoce(stato.voci[i]);
  stato.voci.splice(i, 1);
  const li = trovaLi(id);
  if (li) li.remove();
  nascondiRisultato();
  aggiornaInterfaccia();
}

function riordinaDom() {
  stato.voci.forEach((v) => el.lista.appendChild(trovaLi(v.id)));
  nascondiRisultato();
  aggiornaInterfaccia();
}

/* ---------- Trascinamento per riordinare (computer) ---------- */
let idTrascinato = null;
function abilitaTrascinamento(li) {
  li.addEventListener('dragstart', (e) => {
    if (stato.lavoro) { e.preventDefault(); return; }
    idTrascinato = Number(li.dataset.id);
    li.classList.add('trascinata');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', li.dataset.id);
  });
  li.addEventListener('dragend', () => {
    idTrascinato = null;
    li.classList.remove('trascinata');
    el.lista.querySelectorAll('.bersaglio').forEach((x) => x.classList.remove('bersaglio'));
  });
  li.addEventListener('dragover', (e) => {
    if (idTrascinato === null) return;
    e.preventDefault();
    li.classList.add('bersaglio');
  });
  li.addEventListener('dragleave', () => li.classList.remove('bersaglio'));
  li.addEventListener('drop', (e) => {
    if (idTrascinato === null) return;
    e.preventDefault();
    e.stopPropagation();
    li.classList.remove('bersaglio');
    const da = stato.voci.findIndex((v) => v.id === idTrascinato);
    const a = stato.voci.findIndex((v) => v.id === Number(li.dataset.id));
    if (da < 0 || a < 0 || da === a) return;
    const [voce] = stato.voci.splice(da, 1);
    stato.voci.splice(a, 0, voce);
    riordinaDom();
  });
}

/* ---------- Miniature (una alla volta, per non saturare la memoria del telefono) ---------- */
let anteprimeInCorso = false;
async function generaAnteprime() {
  if (anteprimeInCorso) return;
  anteprimeInCorso = true;
  try {
    let voce;
    while ((voce = stato.voci.find((v) => !v.anteprimaFatta))) {
      voce.anteprimaFatta = true;
      await voce.copia;
      const li = trovaLi(voce.id);
      if (!li) continue;
      if (voce.errore) { li.querySelector('.miniatura').innerHTML = '<span aria-hidden="true">⚠️</span>'; aggiornaDettagli(voce); aggiornaInterfaccia(); continue; }
      const box = li.querySelector('.miniatura');
      try {
        if (voce.tipo === 'immagine') {
          voce.url = URL.createObjectURL(voce.file);
          const img = new Image();
          img.alt = '';
          img.decoding = 'async';
          await new Promise((ok, ko) => { img.onload = ok; img.onerror = () => ko(new Error('immagine')); img.src = voce.url; });
          box.replaceChildren(img);
        } else if (voce.tipo === 'docx') {
          box.innerHTML = '<span class="badge-tipo tipo-word">DOCX</span>';
        } else if (voce.tipo === 'foglio') {
          box.innerHTML = `<span class="badge-tipo tipo-excel">${escapeHtml(estensione(voce.file.name).toUpperCase())}</span>`;
          await preparaCartella(voce);
        } else {
          const pdf = await apriPdfJs(voce);
          try {
            voce.pagine = pdf.numPages;
            const pagina = await pdf.getPage(1);
            const vp = pagina.getViewport({ scale: 1 });
            const canvas = await disegnaPagina(pagina, 150 / Math.max(vp.width, vp.height));
            box.replaceChildren(canvas);
          } finally {
            pdf.destroy();
          }
        }
      } catch (e) {
        voce.errore = messaggioErrore(e, voce);
        box.innerHTML = '<span aria-hidden="true">⚠️</span>';
      }
      if (stato.voci.includes(voce)) { aggiornaDettagli(voce); aggiornaInterfaccia(); }
    }
  } finally {
    anteprimeInCorso = false;
  }
}

/* ---------- Pulizia memoria ---------- */
function liberaVoce(voce) {
  if (voce.url) URL.revokeObjectURL(voce.url);
  voce.url = null;
  const li = trovaLi(voce.id);
  if (li) li.querySelectorAll('canvas').forEach(liberaCanvas);
}

function svuotaVoci() {
  if (stato.lavoro) stato.lavoro.annullato = true;
  stato.voci.forEach(liberaVoce);
  stato.voci = [];
  el.lista.replaceChildren();
  nascondiRisultato();
  nascondiAvanzamento();
}

/* =====================================================================
   INTERFACCIA
   ===================================================================== */
function aggiornaInterfaccia() {
  if (!stato.strumento) return;
  const s = STRUMENTI[stato.strumento];
  const n = stato.voci.length;
  if (s.alAggiorna) s.alAggiorna();
  el.contenitoreLista.hidden = n === 0;
  el.conteggio.textContent = `${n} ${n === 1 ? 'file caricato' : 'file caricati'}`;
  el.suggerimento.hidden = n < 2;
  el.sottotestoArea.textContent = s.singolo
    ? (n ? 'scegline un altro per sostituirlo' : 'oppure trascinalo qui')
    : (n ? 'puoi aggiungerne altri, anche trascinandoli qui' : 'oppure trascinali qui · selezione multipla');
  el.lista.querySelectorAll('.voce-file').forEach((li, i) => {
    li.querySelector('.numero').textContent = i + 1;
    li.querySelector('.su').disabled = i === 0;
    li.querySelector('.giu').disabled = i === n - 1;
  });
  const conErrori = stato.voci.some((v) => v.errore);
  el.azione.textContent = stato.lavoro ? 'Elaborazione in corso…' : (n ? s.etichetta() : s.testoArea);
  el.azione.disabled = !librerieOk || !!stato.lavoro || n < s.minimo || conErrori || (!!s.pronto && !s.pronto());
  if (conErrori && n && !stato.lavoro) el.azione.textContent = 'Rimuovi i file con errore';
}

function valoreOpzione(nome) {
  const scelto = el.opzioni.querySelector(`input[name="${nome}"]:checked`);
  return scelto ? scelto.value : null;
}

const totalePagine = () => stato.voci.reduce((t, v) => t + (v.pagine || 0), 0);

function mostraAvanzamento(frazione, testo) {
  el.barra.hidden = false;
  const p = Math.max(0, Math.min(100, Math.round(frazione * 100)));
  el.riempimento.style.width = p + '%';
  el.percentuale.textContent = p + '%';
  if (testo) el.testoAvanz.textContent = testo;
}
function nascondiAvanzamento() { el.barra.hidden = true; el.riempimento.style.width = '0'; }

function nascondiRisultato() {
  if (stato.urlRisultato) URL.revokeObjectURL(stato.urlRisultato);
  (stato.urlAnteprime || []).forEach((u) => URL.revokeObjectURL(u)); stato.urlAnteprime = [];
  stato.urlRisultato = null;
  el.risultato.hidden = true;
  el.risultato.replaceChildren();
}

/**
 * Mostra il riquadro del risultato con il pulsante di download.
 * @param {{titolo:string, righe:string[], blob:Blob, nome:string, scaricaSubito:boolean}} r
 */
function mostraRisultato(r) {
  nascondiRisultato();
  stato.urlRisultato = URL.createObjectURL(r.blob);
  el.risultato.classList.remove('errore');
  const punto = r.nome.lastIndexOf('.'), base = punto > 0 ? r.nome.slice(0, punto) : r.nome, est = punto > 0 ? r.nome.slice(punto) : '';
  el.risultato.innerHTML = `<div class="risultato-testa"><span class="spunta">${ICONA.spunta}</span><div><h2></h2>${r.righe.map((x) => `<p>${x}</p>`).join('')}</div></div>
    <div class="anteprima-risultato" hidden></div>
    <label class="rinomina"><span>Nome del file</span><span class="campo-nome"><input type="text" spellcheck="false" autocomplete="off"><em></em></span></label>
    <div class="azioni-risultato">
      <a class="scarica" download>${ICONA.scarica}<span></span></a>
      <button type="button" class="bottone-secondario condividi" hidden>${ICONA.condividi}<span>Condividi</span></button>
      <button type="button" class="bottone-secondario ancora">${ICONA.ricomincia}<span>Nuova conversione</span></button>
    </div>`;
  el.risultato.querySelector('h2').textContent = r.titolo;
  const a = el.risultato.querySelector('.scarica'), campo = el.risultato.querySelector('.rinomina input');
  el.risultato.querySelector('.campo-nome em').textContent = est;
  campo.value = base;
  const nomeFinale = () => (campo.value.trim().replace(/[\\/:*?"<>|]+/g, '_') || base) + est;
  const aggiornaNome = () => { a.download = nomeFinale(); a.querySelector('span').textContent = `Scarica (${formattaPeso(r.blob.size)})`; a.title = 'Scarica ' + a.download; };
  campo.addEventListener('input', aggiornaNome);
  campo.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); a.click(); } });
  a.href = stato.urlRisultato;
  aggiornaNome();
  a.addEventListener('click', () => setTimeout(() => notifica(`Download avviato: ${a.download}`), 50));
  // Condividi (telefono: WhatsApp, email, Drive…)
  const cond = el.risultato.querySelector('.condividi');
  const tipo = r.blob.type || (est === '.zip' ? 'application/zip' : 'application/octet-stream');
  const prova = () => new File([r.blob], nomeFinale(), { type: tipo });
  // r.condividiFile: si condividono i singoli file (es. i PDF divisi) invece dello ZIP
  const daCondividere = () => r.condividiFile && r.file ? r.file.map((f) => new File([f.blob], f.nome.split('/').pop(), { type: f.blob.type || 'application/pdf' })) : [prova()];
  if (r.condividiFile && r.file) cond.querySelector('span').textContent = `Condividi i ${r.file.length} PDF`;
  try { if (navigator.canShare && navigator.canShare({ files: daCondividere() })) cond.hidden = false; } catch (e) { /* non disponibile */ }
  cond.addEventListener('click', async () => {
    try { await navigator.share({ files: daCondividere(), title: nomeFinale() }); }
    catch (e) { if (!e || e.name !== 'AbortError') notifica('Condivisione non riuscita: usa Scarica', true); }
  });
  el.risultato.querySelector('.ancora').addEventListener('click', () => { svuotaVoci(); aggiornaInterfaccia(); window.scrollTo({ top: 0, behavior: 'smooth' }); });
  if (/\.pdf$/i.test(r.nome)) aggiungiLinkEditor(r.blob, r.nome); // solo se il risultato è un PDF (non ZIP, immagini o .docx)
  el.risultato.hidden = false;
  el.risultato.scrollIntoView({ behavior: 'smooth', block: 'start' });
  mostraAnteprima(r).catch((e) => console.warn('anteprima', e));
}

/** Anteprima del risultato: pagine del PDF, l'immagine, oppure i file contenuti nello ZIP (ognuno scaricabile da solo). */
async function mostraAnteprima(r) {
  const box = el.risultato.querySelector('.anteprima-risultato');
  if (!box) return;
  const urls = (stato.urlAnteprime = stato.urlAnteprime || []);
  const scheda = (contenuto, etichetta, blob, nome) => {
    const s = document.createElement(blob ? 'a' : 'div');
    s.className = 'scheda-anteprima';
    if (blob) { const u = URL.createObjectURL(blob); urls.push(u); s.href = u; s.download = nome.split('/').pop(); s.title = 'Scarica ' + s.download; }
    const f = document.createElement('div'); f.className = 'foto'; f.append(contenuto);
    const t = document.createElement('span'); t.textContent = etichetta;
    s.append(f, t); box.append(s); box.hidden = false;
  };
  const immagine = (blob) => { const u = URL.createObjectURL(blob); urls.push(u); const i = new Image(); i.alt = ''; i.decoding = 'async'; i.src = u; return i; };
  if (r.file && r.file.length) {
    const max = r.tutti ? r.file.length : 12;
    for (const f of r.file.slice(0, max)) {
      if (!el.risultato.contains(box)) return;
      const contenuto = /^image\//.test(f.blob.type) ? immagine(f.blob) : f.anteprima ? immagine(f.anteprima) : Object.assign(document.createElement('b'), { className: 'badge-tipo tipo-pdf', textContent: (f.nome.split('.').pop() || '').toUpperCase() });
      scheda(contenuto, `${f.nome.split('/').pop()} · ${formattaPeso(f.blob.size)}`, f.blob, f.nome);
    }
    if (r.file.length > max) { const p = document.createElement('p'); p.className = 'altri'; p.textContent = `…e altri ${r.file.length - max} file nello ZIP`; box.append(p); }
    return;
  }
  if (/^image\//.test(r.blob.type)) { scheda(immagine(r.blob), 'Anteprima'); return; }
  if (r.blob.type === 'application/pdf' || /\.pdf$/i.test(r.nome)) {
    await workerPronto;
    if (!workerPdf || workerPdf.destroyed) workerPdf = new pdfjsLib.PDFWorker();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await r.blob.arrayBuffer()), worker: workerPdf, isEvalSupported: false }).promise;
    try {
      const n = Math.min(pdf.numPages, 4);
      for (let i = 1; i <= n; i++) {
        if (!el.risultato.contains(box)) return;
        const pagina = await pdf.getPage(i), vp = pagina.getViewport({ scale: 1 });
        const canvas = await disegnaPagina(pagina, 260 / Math.max(vp.width, vp.height));
        pagina.cleanup();
        scheda(canvas, `Pagina ${i}${i === n && pdf.numPages > n ? ` di ${pdf.numPages}` : ''}`);
      }
    } finally { pdf.destroy(); }
  }
}

/**
 * Aggiunge sotto il download il pulsante secondario "Apri in ninetto-pdf".
 * Le due app stanno sullo stesso dominio (dipierronino-hub.github.io): il PDF viene messo
 * nella cache "editor-pdf-condivisi", la stessa che ninetto-pdf usa per i file condivisi,
 * e l'editor aperto con ?shared=1 lo carica da solo. Niente esce dal dispositivo.
 * @param {Blob|null} blob PDF appena creato (null = link semplice all'editor)
 */
function aggiungiLinkEditor(blob, nome) {
  if (!CONFIG.EDITOR_PDF_URL) return;
  const box = document.createElement('div');
  box.className = 'verso-editor';
  box.innerHTML = `<span>Vuoi firmarlo o annotarlo?</span>
    <a class="bottone-verso-editor" target="_blank" rel="noopener">${ICONA.penna}<span>Apri ninetto-pdf</span></a>`;
  const link = box.querySelector('a');
  link.href = CONFIG.EDITOR_PDF_URL;
  el.risultato.appendChild(box);
  if (!blob || !('caches' in window)) return;
  const chiave = new URL('condiviso-0', CONFIG.EDITOR_PDF_URL).href; // stessa cartella dell'editor
  (async () => {
    try {
      const cache = await caches.open('editor-pdf-condivisi');
      for (const k of await cache.keys()) await cache.delete(k); // un solo file alla volta
      await cache.put(chiave, new Response(blob, { headers: { 'content-type': 'application/pdf', 'x-nome': encodeURIComponent(nome) } }));
      if (box.isConnected) {
        link.href = CONFIG.EDITOR_PDF_URL + '?shared=1'; // l'editor apre subito questo PDF
        link.querySelector('span').textContent = 'Apri in ninetto-pdf';
      }
    } catch (e) { /* cache non disponibile: resta il link semplice all'editor */ }
  })();
}

function mostraErrore(testo) {
  nascondiRisultato();
  el.risultato.classList.add('errore');
  el.risultato.innerHTML = '<h2>Non è stato possibile completare</h2><p></p>';
  el.risultato.querySelector('p').textContent = testo;
  el.risultato.hidden = false;
}

let timerToast = null;
function notifica(testo, errore = false) {
  el.toast.textContent = testo;
  el.toast.classList.toggle('errore', errore);
  el.toast.classList.add('visibile');
  clearTimeout(timerToast);
  timerToast = setTimeout(() => el.toast.classList.remove('visibile'), errore ? 5000 : 3000);
}

/* =====================================================================
   ESECUZIONE
   ===================================================================== */
async function avviaElaborazione() {
  if (stato.lavoro || !stato.strumento) return;
  const s = STRUMENTI[stato.strumento];
  const lavoro = { annullato: false };
  stato.lavoro = lavoro;
  nascondiRisultato();
  mostraAvanzamento(0, 'Preparazione…');
  aggiornaInterfaccia();
  try {
    await Promise.all(stato.voci.map((v) => v.copia));
    const illeggibile = stato.voci.find((v) => v.errore);
    if (illeggibile) { mostraErrore(`"${illeggibile.file.name}": ${illeggibile.errore}`); return; }
    await s.esegui(lavoro);
  } catch (e) {
    if (e && e.name === 'Annullato') notifica('Elaborazione annullata');
    else {
      console.error(e);
      mostraErrore(messaggioErrore(e));
    }
  } finally {
    if (stato.lavoro === lavoro) stato.lavoro = null;
    nascondiAvanzamento();
    aggiornaInterfaccia();
  }
}

function controllaAnnullato(lavoro) {
  if (lavoro.annullato) { const e = new Error('annullato'); e.name = 'Annullato'; throw e; }
}

/* Lascia respirare l'interfaccia (aggiorna la barra) tra un passo pesante e l'altro */
const pausa = () => new Promise((r) => setTimeout(r, 0));

/* ---------------------------------------------------------------------
   MODULO 1 – Immagini in PDF
   --------------------------------------------------------------------- */
async function eseguiImmaginiInPdf(lavoro) {
  const { PDFDocument } = PDFLib;
  const orientamento = valoreOpzione('orientamento');
  const formato = valoreOpzione('formato');
  const separati = valoreOpzione('unione') === 'separati' && stato.voci.length > 1;
  const voci = stato.voci.slice();
  const nomeBase = nomeSenzaEstensione(voci[0].file.name);

  let docUnico = separati ? null : await PDFDocument.create();
  const zip = separati ? new JSZip() : null;

  for (let i = 0; i < voci.length; i++) {
    controllaAnnullato(lavoro);
    mostraAvanzamento(i / voci.length, `Immagine ${i + 1} di ${voci.length}`);
    await pausa();
    const doc = docUnico || await PDFDocument.create();
    const img = await incorporaImmagine(doc, voci[i].file);
    aggiungiPaginaConImmagine(doc, img, formato, orientamento);
    if (separati) {
      const byte = await doc.save();
      zip.file(nomeUnico(zip, nomeSenzaEstensione(voci[i].file.name) + '.pdf'), byte);
    }
  }

  controllaAnnullato(lavoro);
  mostraAvanzamento(0.97, 'Salvataggio…');
  let blob, nome;
  if (separati) {
    blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
    nome = nomeBase + '-pdf.zip';
  } else {
    blob = new Blob([await docUnico.save()], { type: 'application/pdf' });
    nome = (voci.length === 1 ? nomeBase : 'immagini') + '.pdf';
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({
    titolo: 'PDF pronto! 🎉',
    righe: [separati ? `${voci.length} PDF separati in un archivio ZIP.` : `${voci.length} ${voci.length === 1 ? 'pagina' : 'pagine'} in un unico PDF.`],
    blob, nome, scaricaSubito: true
  });
}

/** Inserisce l'immagine nel PDF. JPG e PNG "normali" passano intatti (nessuna perdita di qualità);
 *  foto ruotate (EXIF), WebP e casi particolari passano dal canvas e diventano JPEG. */
async function incorporaImmagine(doc, file) {
  const byte = new Uint8Array(await file.arrayBuffer());
  const tipo = riconosciImmagine(byte);
  if (tipo === 'jpg' && orientamentoExif(byte) <= 1) {
    try { return await doc.embedJpg(byte); } catch (e) { /* passa al canvas */ }
  }
  if (tipo === 'png') {
    try { return await doc.embedPng(byte); } catch (e) { /* passa al canvas */ }
  }
  const jpeg = await ricodificaInJpeg(file, 0.92);
  return doc.embedJpg(new Uint8Array(await jpeg.arrayBuffer()));
}

function aggiungiPaginaConImmagine(doc, img, formato, orientamento) {
  const iw = img.width, ih = img.height;
  let pw, ph, margine = 0;
  if (formato === 'adatta') {
    // Pagina con le proporzioni dell'immagine, lato lungo pari al lato lungo di un A4
    const k = CONFIG.A4[1] / Math.max(iw, ih);
    pw = iw * k; ph = ih * k;
  } else {
    const orizzontale = orientamento === 'orizzontale' || (orientamento === 'auto' && iw > ih);
    [pw, ph] = orizzontale ? [CONFIG.A4[1], CONFIG.A4[0]] : CONFIG.A4;
    if (formato === 'a4m') margine = CONFIG.MARGINE_A4;
  }
  const page = doc.addPage([pw, ph]);
  const k = Math.min((pw - 2 * margine) / iw, (ph - 2 * margine) / ih);
  const w = iw * k, h = ih * k;
  page.drawImage(img, { x: (pw - w) / 2, y: (ph - h) / 2, width: w, height: h });
}

function riconosciImmagine(b) {
  if (b[0] === 0xFF && b[1] === 0xD8) return 'jpg';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
  return 'altro';
}

/** Legge il tag "Orientation" dell'EXIF di un JPEG (1 = dritto). */
function orientamentoExif(b) {
  try {
    const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let p = 2;
    while (p + 4 < v.byteLength) {
      const marker = v.getUint16(p);
      const lung = v.getUint16(p + 2);
      if (marker === 0xFFE1 && v.getUint32(p + 4) === 0x45786966) { // "Exif"
        const tiff = p + 10;
        const little = v.getUint16(tiff) === 0x4949;
        const dir = tiff + v.getUint32(tiff + 4, little);
        const n = v.getUint16(dir, little);
        for (let i = 0; i < n; i++) {
          const voce = dir + 2 + i * 12;
          if (v.getUint16(voce, little) === 0x0112) return v.getUint16(voce + 8, little);
        }
        return 1;
      }
      if ((marker & 0xFF00) !== 0xFF00 || marker === 0xFFDA) break;
      p += 2 + lung;
    }
  } catch (e) { /* EXIF illeggibile: consideriamo l'immagine dritta */ }
  return 1;
}

/** Decodifica un'immagine (rispettando la rotazione EXIF) e la ricodifica in JPEG su sfondo bianco. */
async function ricodificaInJpeg(file, qualita) {
  let sorgente, w, h;
  if (window.createImageBitmap) {
    sorgente = await createImageBitmap(file, { imageOrientation: 'from-image' });
    w = sorgente.width; h = sorgente.height;
  } else {
    const url = URL.createObjectURL(file);
    try {
      sorgente = new Image();
      sorgente.src = url;
      await sorgente.decode();
      w = sorgente.naturalWidth; h = sorgente.naturalHeight;
    } finally { URL.revokeObjectURL(url); }
  }
  const k = fattoreLimite(w, h);
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(w * k); canvas.height = Math.round(h * k);
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(sorgente, 0, 0, canvas.width, canvas.height);
  if (sorgente.close) sorgente.close();
  try { return await canvasInBlob(canvas, 'image/jpeg', qualita); }
  finally { liberaCanvas(canvas); }
}

/* ---------------------------------------------------------------------
   MODULO 2 – PDF in Immagini
   --------------------------------------------------------------------- */
async function eseguiPdfInImmagini(lavoro) {
  const png = valoreOpzione('formatoImg') === 'png';
  const dpi = Number(valoreOpzione('dpi'));
  const est = png ? 'png' : 'jpg';
  const mime = png ? 'image/png' : 'image/jpeg';
  const voci = stato.voci.slice();
  const totale = totalePagine() || voci.length;
  const zip = new JSZip();
  let fatte = 0, ultimaImmagine = null, ridotte = false;
  const elenco = [];

  for (const voce of voci) {
    const pdf = await apriPdfJs(voce);
    try {
      const cifre = Math.max(2, String(pdf.numPages).length);
      // Con più PDF ogni file ha la sua cartella nello ZIP
      const cartella = voci.length > 1 ? nomeUnico(zip, nomeSenzaEstensione(voce.file.name)) + '/' : '';
      for (let n = 1; n <= pdf.numPages; n++) {
        controllaAnnullato(lavoro);
        mostraAvanzamento(fatte / totale, `Pagina ${fatte + 1} di ${totale}`);
        await pausa();
        const pagina = await pdf.getPage(n);
        const canvas = await disegnaPagina(pagina, dpi / 72);
        if (canvas.ridotta) ridotte = true;
        const blob = await canvasInBlob(canvas, mime, 0.92);
        liberaCanvas(canvas);
        pagina.cleanup();
        ultimaImmagine = blob;
        const nomeImg = `${cartella}pagina-${String(n).padStart(cifre, '0')}.${est}`;
        zip.file(nomeImg, blob); elenco.push({ nome: nomeImg, blob });
        fatte++;
      }
    } finally {
      pdf.destroy();
    }
  }

  controllaAnnullato(lavoro);
  const nomeBase = nomeSenzaEstensione(voci[0].file.name);
  const nota = ridotte ? ['Alcune pagine molto grandi sono state ridotte per non superare i limiti di memoria del dispositivo.'] : [];
  if (fatte === 1) {
    mostraAvanzamento(1, 'Fatto');
    mostraRisultato({ titolo: 'Immagine pronta! 🎉', righe: ['Il PDF ha una sola pagina: ecco l\'immagine.', ...nota],
      blob: ultimaImmagine, nome: `${nomeBase}.${est}`, scaricaSubito: true });
    return;
  }
  mostraAvanzamento(0.98, 'Creazione dello ZIP…');
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({ titolo: 'Immagini pronte! 🎉', righe: [`${fatte} pagine convertite in ${est.toUpperCase()} a ${dpi} DPI, in un unico archivio ZIP.`, ...nota],
    blob, nome: `${voci.length > 1 ? 'pdf' : nomeBase}-immagini.zip`, file: elenco, scaricaSubito: true });
}

/* ---------------------------------------------------------------------
   MODULO 3 – Unisci PDF
   --------------------------------------------------------------------- */
async function eseguiUnisci(lavoro) {
  const { PDFDocument } = PDFLib;
  const voci = stato.voci.slice();
  const finale = await PDFDocument.create();
  for (let i = 0; i < voci.length; i++) {
    controllaAnnullato(lavoro);
    mostraAvanzamento(i / voci.length, `File ${i + 1} di ${voci.length}`);
    await pausa();
    let sorgente;
    try {
      sorgente = await PDFDocument.load(await voci[i].file.arrayBuffer());
    } catch (e) {
      e.nomeFile = voci[i].file.name;
      throw e;
    }
    // copyPages mantiene formato, rotazione e contenuto vettoriale di ogni pagina
    const pagine = await finale.copyPages(sorgente, sorgente.getPageIndices());
    pagine.forEach((p) => finale.addPage(p));
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(0.95, 'Salvataggio…');
  await pausa();
  const blob = new Blob([await finale.save({ useObjectStreams: true })], { type: 'application/pdf' });
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({
    titolo: 'PDF uniti! 🎉',
    righe: [`${voci.length} file, ${finale.getPageCount()} pagine in totale.`],
    blob, nome: 'pdf-uniti.pdf', scaricaSubito: true
  });
}

/* ---------------------------------------------------------------------
   MODULO 9 – Dividi PDF (per pagina, per intervalli o ogni N pagine) e rinomina ogni file
   --------------------------------------------------------------------- */

function azzeraDivisione() {
  divisione.miniature.forEach((m) => URL.revokeObjectURL(m.url));
  divisione.miniature.clear(); divisione.nomi.clear(); divisione.escluse.clear();
  divisione.voceId = null; divisione.chiave = null; divisione.parti = []; divisione.errore = null;
}

/** Calcola i file da creare. Restituisce { parti, errore }. */
function calcolaParti(modo, testo, n) {
  if (modo === 'ogni') {
    const k = Math.floor(Number(testo));
    if (!(k >= 1)) return { parti: [], errore: 'Scrivi quante pagine mettere in ogni file (1 o più).' };
    const parti = [];
    for (let da = 1; da <= n; da += k) parti.push({ da, a: Math.min(n, da + k - 1) });
    return { parti, errore: null };
  }
  if (modo === 'intervalli') {
    // trattini "lunghi" (correttore del telefono) e spazi intorno al trattino
    const pulito = String(testo || '').replace(/[–—−]/g, '-').replace(/\s*-\s*/g, '-').trim();
    if (!pulito) return { parti: [], errore: '' }; // ancora vuoto: nessun messaggio in rosso
    const parti = [];
    for (const pezzo of pulito.split(/[,;\s]+/).filter(Boolean)) {
      const m = pezzo.match(/^(\d*)-(\d*)$/) || pezzo.match(/^(\d+)$/);
      if (!m || (m[1] === '' && m[2] === '')) return { parti: [], errore: `«${pezzo}» non è un intervallo valido. Esempio: 1-3, 4, 5-8` };
      const da = m[1] === '' ? 1 : Number(m[1]);
      const a = m.length === 2 ? da : (m[2] === '' ? n : Number(m[2]));
      if (da < 1 || a < 1) return { parti: [], errore: 'Le pagine partono da 1.' };
      if (da > n || a > n) return { parti: [], errore: `Pagina ${Math.max(da, a)} inesistente: il PDF ha ${n} ${n === 1 ? 'pagina' : 'pagine'}.` };
      if (da > a) return { parti: [], errore: `«${pezzo}»: la prima pagina deve venire prima dell'ultima.` };
      parti.push({ da, a });
    }
    return { parti, errore: null };
  }
  return { parti: Array.from({ length: n }, (_, i) => ({ da: i + 1, a: i + 1 })), errore: null };
}

function nomePredefinito(base, p, n) {
  const c = String(n).length, z = (x) => String(x).padStart(c, '0');
  return p.da === p.a ? `${base}_pag${z(p.da)}` : `${base}_pag${z(p.da)}-${z(p.a)}`;
}

/** Nome del file (senza .pdf) ripulito dai caratteri non ammessi. */
function nomeParte(voce, p) {
  const scelto = (divisione.nomi.get(chiaveParte(p)) || '').trim().replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|]+/g, '_').trim();
  return scelto || nomePredefinito(nomeSenzaEstensione(voce.file.name), p, voce.pagine);
}

const partiIncluse = () => divisione.parti.filter((p) => !divisione.escluse.has(chiaveParte(p)));

/** Ricostruisce l'elenco dei file da creare quando cambiano PDF, modo o intervalli. */
function aggiornaPiano() {
  if (stato.strumento !== 'dividi') return;
  const gruppo = $('gruppoPiano');
  if (!gruppo) return;
  const voce = stato.voci[0];
  const modo = valoreOpzione('modoDividi') || 'pagina';
  $('campoIntervalli').hidden = modo !== 'intervalli';
  $('campoOgni').hidden = modo !== 'ogni';
  if (!voce) { azzeraDivisione(); gruppo.hidden = true; return; }
  if (!voce.pagine || voce.errore) { gruppo.hidden = true; divisione.parti = []; divisione.chiave = null; return; }
  const testo = modo === 'intervalli' ? $('testoIntervalli').value : modo === 'ogni' ? $('numeroOgni').value : '';
  const chiave = [voce.id, modo, testo].join('|');
  if (chiave === divisione.chiave) return;
  if (divisione.voceId !== voce.id) { azzeraDivisione(); divisione.voceId = voce.id; }
  divisione.chiave = chiave;
  const { parti, errore } = calcolaParti(modo, testo, voce.pagine);
  divisione.parti = parti;
  divisione.errore = errore;
  const err = $('erroreIntervalli');
  err.hidden = !errore; err.textContent = errore || '';
  gruppo.hidden = !parti.length;
  disegnaPiano(voce);
  generaMiniatureDivisione(voce);
}

function disegnaPiano(voce) {
  const lista = $('pianoDivisione');
  if (!lista) return;
  const n = voce.pagine;
  lista.replaceChildren(...divisione.parti.map((p, i) => {
    const k = chiaveParte(p), quante = p.a - p.da + 1;
    const li = document.createElement('li');
    li.className = 'parte' + (divisione.escluse.has(k) ? ' esclusa' : '');
    li.dataset.chiave = k;
    li.dataset.pagina = p.da;
    li.innerHTML = `
      <input type="checkbox" class="parte-includi"${divisione.escluse.has(k) ? '' : ' checked'} aria-label="Crea il file ${i + 1}">
      <button type="button" class="parte-mini" aria-label="Vedi la pagina ${p.da} in grande"><div class="attesa"></div></button>
      <div class="parte-info">
        <span class="parte-pagine"></span>
        <span class="campo-nome"><input type="text" spellcheck="false" autocomplete="off" autocapitalize="off" aria-label="Nome del file ${i + 1}"><em>.pdf</em></span>
      </div>`;
    li.querySelector('.parte-pagine').textContent = `${i + 1}. ${p.da === p.a ? `Pagina ${p.da}` : `Pagine ${p.da}–${p.a}`} · ${quante} ${quante === 1 ? 'pagina' : 'pagine'}`;
    const campo = li.querySelector('input[type=text]');
    campo.value = divisione.nomi.get(k) || '';
    campo.placeholder = nomePredefinito(nomeSenzaEstensione(voce.file.name), p, n);
    if (!campo.value) campo.value = campo.placeholder;
    campo.addEventListener('input', () => divisione.nomi.set(k, campo.value));
    campo.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); const s = li.nextElementSibling; if (s) s.querySelector('input[type=text]').focus(); else campo.blur(); } });
    li.querySelector('.parte-includi').addEventListener('change', (e) => {
      if (e.target.checked) divisione.escluse.delete(k); else divisione.escluse.add(k);
      li.classList.toggle('esclusa', !e.target.checked);
      aggiornaRiepilogo();
    });
    li.querySelector('.parte-mini').addEventListener('click', () => apriLente(voce, p));
    const m = divisione.miniature.get(p.da);
    if (m) mostraMiniaturaParte(li, m.url);
    return li;
  }));
  const bottone = $('applicaSerie');
  if (bottone && !bottone.dataset.pronto) {
    bottone.dataset.pronto = '1';
    bottone.addEventListener('click', applicaSerie);
    $('prefissoSerie').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); applicaSerie(); } });
  }
  aggiornaRiepilogo();
}

function aggiornaRiepilogo() {
  const r = $('riepilogoPiano');
  if (!r) return;
  const incluse = partiIncluse();
  const pagine = incluse.reduce((t, p) => t + p.a - p.da + 1, 0);
  r.textContent = `${incluse.length} di ${divisione.parti.length} file da creare · ${pagine} ${pagine === 1 ? 'pagina' : 'pagine'}`;
}

/** «Numera tutti»: Nome_01, Nome_02… sui file spuntati. Campo vuoto = torna ai nomi automatici. */
function applicaSerie() {
  const voce = stato.voci[0];
  if (!voce) return;
  const prefisso = $('prefissoSerie').value.trim().replace(/\.pdf$/i, '');
  const incluse = partiIncluse();
  if (!prefisso) divisione.nomi.clear();
  else {
    const c = Math.max(2, String(incluse.length).length);
    incluse.forEach((p, i) => divisione.nomi.set(chiaveParte(p), `${prefisso}_${String(i + 1).padStart(c, '0')}`));
  }
  disegnaPiano(voce);
  notifica(prefisso ? `Nomi aggiornati: ${prefisso}_01, ${prefisso}_02…` : 'Nomi automatici ripristinati');
}

function mostraMiniaturaParte(li, url) {
  const b = li.querySelector('.parte-mini');
  const img = new Image(); img.alt = ''; img.decoding = 'async'; img.src = url;
  b.replaceChildren(img);
}

/* Miniature della prima pagina di ogni file, una alla volta e piccole (memoria del telefono) */
async function generaMiniatureDivisione(voce) {
  if (miniatureDivisioneInCorso) return;
  miniatureDivisioneInCorso = true;
  let pdf = null;
  try {
    const mancante = () => divisione.voceId === voce.id && stato.voci[0] === voce
      ? divisione.parti.find((p) => !divisione.miniature.has(p.da)) : null;
    let p;
    while ((p = mancante())) {
      if (!pdf) pdf = await apriPdfJs(voce);
      const pagina = await pdf.getPage(p.da);
      const vp = pagina.getViewport({ scale: 1 });
      const canvas = await disegnaPagina(pagina, 160 / Math.max(vp.width, vp.height));
      pagina.cleanup();
      const blob = await canvasInBlob(canvas, 'image/jpeg', 0.8);
      liberaCanvas(canvas);
      if (divisione.voceId !== voce.id) break;
      const m = { blob, url: URL.createObjectURL(blob) };
      divisione.miniature.set(p.da, m);
      const lista = $('pianoDivisione');
      if (lista) lista.querySelectorAll(`.parte[data-pagina="${p.da}"]`).forEach((li) => mostraMiniaturaParte(li, m.url));
      await pausa();
    }
  } catch (e) {
    console.warn('miniature divisione', e);
  } finally {
    if (pdf) pdf.destroy();
    miniatureDivisioneInCorso = false;
  }
  // nel frattempo il piano può essere cambiato: ricontrolla
  if (stato.voci[0] && divisione.voceId === stato.voci[0].id && divisione.parti.some((p) => !divisione.miniature.has(p.da))) {
    setTimeout(() => generaMiniatureDivisione(stato.voci[0]), 0);
  }
}

/** Pagina in grande per capire che nome dare al file; frecce per scorrere le pagine del file. */
async function apriLente(voce, parte) {
  const ov = document.createElement('div');
  ov.className = 'lente';
  ov.innerHTML = `<div class="lente-box" role="dialog" aria-modal="true" aria-label="Anteprima pagina">
      <div class="lente-testa"><strong></strong><button type="button" class="lente-chiudi" aria-label="Chiudi">✕</button></div>
      <div class="lente-pagina"><div class="attesa"></div></div>
      <div class="lente-nav" ${parte.da === parte.a ? 'hidden' : ''}>
        <button type="button" class="bottone-secondario prec" aria-label="Pagina precedente"><svg viewBox="0 0 24 24"><path d="M15 5l-7 7 7 7"/></svg></button>
        <span></span>
        <button type="button" class="bottone-secondario succ" aria-label="Pagina successiva"><svg viewBox="0 0 24 24"><path d="M9 5l7 7-7 7"/></svg></button>
      </div></div>`;
  document.body.append(ov);
  document.body.classList.add('lente-aperta');
  let pdf = null, chiusa = false, numero = parte.da, turno = 0;
  const chiudi = () => {
    if (chiusa) return;
    chiusa = true;
    ov.querySelectorAll('canvas').forEach(liberaCanvas);
    ov.remove();
    document.body.classList.remove('lente-aperta');
    document.removeEventListener('keydown', tasti);
    if (pdf) pdf.destroy();
  };
  const tasti = (e) => {
    if (e.key === 'Escape') chiudi();
    else if (e.key === 'ArrowLeft') vai(-1);
    else if (e.key === 'ArrowRight') vai(+1);
  };
  document.addEventListener('keydown', tasti);
  ov.addEventListener('click', (e) => { if (e.target === ov) chiudi(); });
  ov.querySelector('.lente-chiudi').addEventListener('click', chiudi);
  ov.querySelector('.prec').addEventListener('click', () => vai(-1));
  ov.querySelector('.succ').addEventListener('click', () => vai(+1));
  const box = ov.querySelector('.lente-pagina');
  async function mostra() {
    const mio = ++turno;
    ov.querySelector('.lente-testa strong').textContent = `Pagina ${numero} di ${voce.pagine}`;
    ov.querySelector('.lente-nav span').textContent = `${numero - parte.da + 1} / ${parte.a - parte.da + 1}`;
    ov.querySelector('.prec').disabled = numero <= parte.da;
    ov.querySelector('.succ').disabled = numero >= parte.a;
    try {
      if (!pdf) pdf = await apriPdfJs(voce);
      if (chiusa) { pdf.destroy(); return; }
      const pagina = await pdf.getPage(numero);
      const vp = pagina.getViewport({ scale: 1 });
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const w = box.clientWidth || 320, h = box.clientHeight || 480;
      const canvas = await disegnaPagina(pagina, Math.min(w / vp.width, h / vp.height) * dpr);
      pagina.cleanup();
      if (chiusa || mio !== turno) { liberaCanvas(canvas); return; }
      box.querySelectorAll('canvas').forEach(liberaCanvas);
      box.replaceChildren(canvas);
    } catch (e) {
      if (!chiusa) box.textContent = 'Anteprima non disponibile';
    }
  }
  function vai(d) {
    const nuovo = Math.min(parte.a, Math.max(parte.da, numero + d));
    if (nuovo !== numero) { numero = nuovo; mostra(); }
  }
  ov.querySelector('.lente-chiudi').focus();
  mostra();
}

async function eseguiDividi(lavoro) {
  const { PDFDocument } = PDFLib;
  const voce = stato.voci[0];
  const parti = partiIncluse().map((p) => ({ ...p, nome: nomeParte(voce, p) }));
  if (!parti.length) throw new Error('nessuna pagina scelta');
  mostraAvanzamento(0, 'Lettura del PDF…');
  await pausa();
  let sorgente;
  try { sorgente = await PDFDocument.load(await voce.file.arrayBuffer(), { updateMetadata: false }); }
  catch (e) { e.nomeFile = voce.file.name; throw e; }
  const zip = new JSZip(), file = [];
  for (let i = 0; i < parti.length; i++) {
    controllaAnnullato(lavoro);
    mostraAvanzamento(i / parti.length, `File ${i + 1} di ${parti.length}`);
    await pausa();
    const p = parti[i];
    const nuovo = await PDFDocument.create();
    // copyPages mantiene formato, rotazione, testo e contenuto vettoriale delle pagine
    const pagine = await nuovo.copyPages(sorgente, Array.from({ length: p.a - p.da + 1 }, (_, j) => p.da - 1 + j));
    pagine.forEach((x) => nuovo.addPage(x));
    nuovo.setTitle(p.nome);
    const blob = new Blob([await nuovo.save({ useObjectStreams: true })], { type: 'application/pdf' });
    const nome = nomeUnico(zip, p.nome + '.pdf');
    zip.file(nome, blob);
    const m = divisione.miniature.get(p.da);
    file.push({ nome, blob, anteprima: m && m.blob, parte: p });
  }
  controllaAnnullato(lavoro);
  const n = voce.pagine || sorgente.getPageCount();
  if (file.length === 1) {
    const p = file[0].parte;
    mostraAvanzamento(1, 'Fatto');
    mostraRisultato({ titolo: 'PDF pronto! 🎉', righe: [`${p.da === p.a ? `Pagina ${p.da}` : `Pagine ${p.da}–${p.a}`} di ${n}, testo e qualità originali.`],
      blob: file[0].blob, nome: file[0].nome, scaricaSubito: true });
    return;
  }
  mostraAvanzamento(0.97, 'Creazione dello ZIP…');
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({
    titolo: 'PDF divisi! 🎉',
    righe: [`${file.length} file creati dalle ${n} pagine di "${escapeHtml(voce.file.name)}", con i nomi che hai scelto.`, 'Scarica tutto nello ZIP, oppure tocca un file qui sotto per scaricarlo da solo.'],
    blob, nome: nomeSenzaEstensione(voce.file.name) + '-diviso.zip', file, tutti: true, condividiFile: true, scaricaSubito: true
  });
}

/* ---------------------------------------------------------------------
   MODULO 4 – Comprimi PDF
   --------------------------------------------------------------------- */
async function eseguiComprimi(lavoro) {
  const { PDFDocument } = PDFLib;
  if (valoreOpzione('livello') === 'leggera') return comprimiSenzaPerdita(lavoro);
  const livello = CONFIG.COMPRESSIONE[valoreOpzione('livello')] || CONFIG.COMPRESSIONE.media;
  const voci = stato.voci.slice();
  const totale = totalePagine() || voci.length;
  const risultati = [];
  let fatte = 0;

  for (const voce of voci) {
    const pdf = await apriPdfJs(voce);
    const nuovo = await PDFDocument.create();
    try {
      for (let n = 1; n <= pdf.numPages; n++) {
        controllaAnnullato(lavoro);
        mostraAvanzamento(fatte / totale, `Pagina ${fatte + 1} di ${totale}`);
        await pausa();
        const pagina = await pdf.getPage(n);
        const vp = pagina.getViewport({ scale: 1 }); // misure in punti, rotazione già applicata
        const canvas = await disegnaPagina(pagina, livello.dpi / 72);
        const jpeg = await canvasInBlob(canvas, 'image/jpeg', livello.qualita);
        liberaCanvas(canvas);
        pagina.cleanup();
        const img = await nuovo.embedJpg(new Uint8Array(await jpeg.arrayBuffer()));
        nuovo.addPage([vp.width, vp.height]).drawImage(img, { x: 0, y: 0, width: vp.width, height: vp.height });
        fatte++;
      }
    } finally {
      pdf.destroy();
    }
    const byte = await nuovo.save({ useObjectStreams: true });
    risultati.push({ voce, byte, prima: voce.file.size, dopo: byte.byteLength });
  }

  return concludiCompressione(lavoro, risultati);
}

/** Livello «Leggera»: riscrive il PDF con flussi di oggetti compressi e toglie i dati inutili.
 *  Il testo resta vero (selezionabile), immagini e qualità restano identiche. */
async function comprimiSenzaPerdita(lavoro) {
  const { PDFDocument } = PDFLib;
  const voci = stato.voci.slice(), risultati = [];
  for (let i = 0; i < voci.length; i++) {
    controllaAnnullato(lavoro);
    mostraAvanzamento(i / voci.length, `File ${i + 1} di ${voci.length}: ottimizzo la struttura`);
    await pausa();
    let doc;
    try { doc = await PDFDocument.load(await voci[i].file.arrayBuffer(), { updateMetadata: false }); }
    catch (e) { e.nomeFile = voci[i].file.name; throw e; }
    // via le miniature incorporate e i metadati XMP (spesso pesanti e inutili)
    try { doc.catalog.delete(PDFLib.PDFName.of('Metadata')); } catch (e) { /* ignora */ }
    doc.getPages().forEach((p) => { try { p.node.delete(PDFLib.PDFName.of('Thumb')); } catch (e) { /* ignora */ } });
    const byte = await doc.save({ useObjectStreams: true, addDefaultPage: false });
    risultati.push({ voce: voci[i], byte, prima: voci[i].file.size, dopo: byte.byteLength });
  }
  return concludiCompressione(lavoro, risultati);
}

async function concludiCompressione(lavoro, risultati) {
  const voci = risultati.map((r) => r.voce);
  // se un file non diventa più leggero, nel risultato resta l'originale
  risultati.forEach((r) => { if (r.dopo >= r.prima) { r.originale = true; } });
  controllaAnnullato(lavoro);
  mostraAvanzamento(0.98, 'Salvataggio…');
  const prima = risultati.reduce((t, r) => t + r.prima, 0);
  const dopo = risultati.reduce((t, r) => t + (risultati.length > 1 ? Math.min(r.dopo, r.prima) : r.dopo), 0);
  const righe = [confronto(prima, dopo)];
  if (risultati.length > 1) {
    risultati.forEach((r) => righe.push(`<small>${escapeHtml(r.voce.file.name)}: ${formattaPeso(r.prima)} → ${r.originale ? 'già ottimizzato, tenuto l\'originale' : formattaPeso(r.dopo)}</small>`));
  }
  const migliorato = dopo < prima;
  if (!migliorato) righe.push(valoreOpzione('livello') === 'leggera' ? 'Il file era già ottimizzato: per ridurlo davvero prova il livello Media o Massima.' : 'Il file era già ottimizzato: la versione compressa non è più leggera. Ti conviene tenere l\'originale o provare la compressione massima.');

  let blob, nome;
  if (risultati.length === 1) {
    blob = new Blob([risultati[0].byte], { type: 'application/pdf' });
    nome = nomeSenzaEstensione(voci[0].file.name) + '-compresso.pdf';
  } else {
    const zip = new JSZip();
    for (const r of risultati) zip.file(nomeUnico(zip, nomeSenzaEstensione(r.voce.file.name) + (r.originale ? '.pdf' : '-compresso.pdf')), r.originale ? await r.voce.file.arrayBuffer() : r.byte);
    blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
    nome = 'pdf-compressi.zip';
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({ titolo: migliorato ? 'PDF compresso! 🎉' : 'Compressione completata', righe, blob, nome, scaricaSubito: migliorato });
}

/* ---------------------------------------------------------------------
   MODULO 8 – Converti immagini (JPG, PNG, WebP; ridimensiona; in blocco)
   --------------------------------------------------------------------- */
async function eseguiConvertiImmagini(lavoro) {
  const formato = valoreOpzione('formatoDest') || 'jpeg';
  const lato = Number(valoreOpzione('lato')) || 0;
  const qualita = formato === 'png' ? undefined : Number(valoreOpzione('qualitaImg')) || 0.92;
  const mime = 'image/' + formato, est = formato === 'jpeg' ? 'jpg' : formato;
  const voci = stato.voci.slice(), file = [];
  let prima = 0, dopo = 0, webpNonSupportato = false;
  for (let i = 0; i < voci.length; i++) {
    controllaAnnullato(lavoro);
    mostraAvanzamento(i / voci.length, `Immagine ${i + 1} di ${voci.length}`);
    await pausa();
    const voce = voci[i];
    let sorgente;
    try { sorgente = await createImageBitmap(voce.file, { imageOrientation: 'from-image' }); }
    catch (e) { const err = new Error('immagine'); err.nomeFile = voce.file.name; throw err; }
    let k = lato ? Math.min(1, lato / Math.max(sorgente.width, sorgente.height)) : 1;
    k = Math.min(k, fattoreLimite(sorgente.width, sorgente.height));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(sorgente.width * k)); canvas.height = Math.max(1, Math.round(sorgente.height * k));
    const ctx = canvas.getContext('2d');
    if (formato === 'jpeg') { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height); } // JPG non ha trasparenza: sfondo bianco
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(sorgente, 0, 0, canvas.width, canvas.height);
    if (sorgente.close) sorgente.close();
    let blob = await canvasInBlob(canvas, mime, qualita);
    liberaCanvas(canvas);
    if (blob.type !== mime) webpNonSupportato = true; // es. Safari vecchi: WebP non creabile
    prima += voce.file.size; dopo += blob.size;
    file.push({ nome: nomeSenzaEstensione(voce.file.name) + '.' + (blob.type === mime ? est : 'png'), blob });
  }
  controllaAnnullato(lavoro);
  const righe = [`${voci.length} ${voci.length === 1 ? 'immagine convertita' : 'immagini convertite'} in ${est.toUpperCase()}${lato ? `, lato lungo massimo ${lato} px` : ''}.`, confronto(prima, dopo).replace('File originale', 'Prima').replace('File compresso', 'Dopo')];
  if (webpNonSupportato) righe.push('Questo browser non sa creare file WebP: le immagini sono state salvate in PNG.');
  if (file.length === 1) {
    mostraAvanzamento(1, 'Fatto');
    mostraRisultato({ titolo: 'Immagine pronta! 🎉', righe, blob: file[0].blob, nome: file[0].nome, scaricaSubito: false });
    return;
  }
  mostraAvanzamento(0.97, 'Creazione dello ZIP…');
  const zip = new JSZip();
  file.forEach((f) => { f.nome = nomeUnico(zip, f.nome); zip.file(f.nome, f.blob); });
  const blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({ titolo: 'Immagini pronte! 🎉', righe, blob, nome: 'immagini-convertite.zip', file, scaricaSubito: false });
}

function confronto(prima, dopo) {
  const perc = Math.round((1 - dopo / prima) * 100);
  const esito = perc > 0
    ? `<span class="risparmio">(-${perc}%)</span>`
    : `<span class="peggio">(+${Math.abs(perc)}%)</span>`;
  return `<span class="confronto">File originale: ${formattaPeso(prima)} → File compresso: ${formattaPeso(dopo)} ${esito}</span>`;
}

/* =====================================================================
   UTILITÀ PDF.JS E CANVAS
   ===================================================================== */

/** Apre un PDF con pdf.js; se è protetto chiede la password (una volta, poi la ricorda). */
async function apriPdfJs(voce) {
  await workerPronto;
  if (!workerPdf || workerPdf.destroyed) workerPdf = new pdfjsLib.PDFWorker();
  const dati = new Uint8Array(await voce.file.arrayBuffer());
  const task = pdfjsLib.getDocument({ data: dati, worker: workerPdf, password: voce.password || undefined, isEvalSupported: false });
  task.onPassword = (aggiorna, motivo) => {
    const errata = motivo === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD;
    const pw = window.prompt(`${errata ? 'Password errata. ' : ''}"${voce.file.name}" è protetto. Inserisci la password per aprirlo:`);
    if (pw === null) {
      task.destroy();
    } else {
      voce.password = pw;
      aggiorna(pw);
    }
  };
  try {
    return await task.promise;
  } catch (e) {
    if (voce.password && e && e.name === 'PasswordException') voce.password = null;
    if (task.destroyed && (!e || e.name !== 'InvalidPDFException')) {
      const err = new Error('password annullata'); err.name = 'PasswordException'; throw err;
    }
    throw e;
  }
}

/** Disegna una pagina su un canvas bianco, riducendo la scala se supera i limiti del dispositivo. */
async function disegnaPagina(pagina, scala) {
  let vp = pagina.getViewport({ scale: scala });
  const k = fattoreLimite(vp.width, vp.height);
  if (k < 1) vp = pagina.getViewport({ scale: scala * k });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(vp.width));
  canvas.height = Math.max(1, Math.floor(vp.height));
  canvas.ridotta = k < 1;
  const ctx = canvas.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await pagina.render({ canvasContext: ctx, viewport: vp }).promise;
  return canvas;
}

function fattoreLimite(w, h) {
  return Math.min(1, Math.sqrt(CONFIG.MAX_PIXEL / (w * h)), CONFIG.MAX_LATO / w, CONFIG.MAX_LATO / h);
}

function canvasInBlob(canvas, tipo, qualita) {
  return new Promise((ok, ko) => canvas.toBlob((b) => b ? ok(b) : ko(new Error('memoria')), tipo, qualita));
}

/** Azzera il canvas: su Safari libera subito la memoria video invece di aspettare il garbage collector. */
function liberaCanvas(canvas) { canvas.width = 0; canvas.height = 0; }

/* =====================================================================
   UTILITÀ VARIE
   ===================================================================== */
function messaggioErrore(e, voce) {
  const nome = (voce && voce.file.name) || (e && e.nomeFile);
  const chi = nome ? `"${nome}": ` : '';
  const testo = String((e && e.message) || e || '');
  if (e && e.name === 'LibreriaNonCaricata') return `${chi}non riesco a scaricare i componenti per questo strumento. Serve la connessione la prima volta che lo usi, poi funziona anche offline.`;
  if (e && e.name === 'DocxNonValido') return `${chi}documento Word non leggibile. Serve un file .docx valido (i vecchi .doc vanno prima salvati come .docx).`;
  if (e && e.name === 'FoglioNonValido') return `${chi}foglio di calcolo non leggibile o protetto da password.`;
  if (e && e.name === 'FoglioVuoto') return `${chi}i fogli scelti sono vuoti.`;
  if (e && e.name === 'NessunTesto') return `${chi}il PDF non contiene testo (è una scansione): non c'è nulla da estrarre in Word.`;
  if (e && (e.name === 'NotReadableError' || e.name === 'NotFoundError')) return `${chi}il telefono non permette più di leggere il file. Toccalo con il cestino, poi sceglilo di nuovo (meglio dalla cartella Download).`;
  if (e && e.name === 'PasswordException') return `${chi}PDF protetto da password, non sbloccato.`;
  if (/encrypt/i.test(testo)) return `${chi}PDF protetto da password: non si può unire né dividere. Rimuovi la protezione (o comprimilo prima con questa app, che chiede la password) e riprova.`;
  if (e && (e.name === 'InvalidPDFException' || /invalid pdf|no pdf header|failed to parse/i.test(testo))) return `${chi}il file non è un PDF valido o è danneggiato.`;
  if (testo === 'immagine' || /decode|source image/i.test(testo)) return `${chi}immagine non leggibile o formato non supportato.`;
  if (testo === 'memoria' || /memory|allocation/i.test(testo)) return `${chi}memoria del dispositivo esaurita. Prova con meno file o con una risoluzione più bassa.`;
  return `${chi}errore imprevisto (${testo || 'sconosciuto'}).`;
}

function formattaPeso(byte) {
  if (byte < 1024) return byte + ' B';
  if (byte < 1024 * 1024) return (byte / 1024).toLocaleString('it-IT', { maximumFractionDigits: 0 }) + ' KB';
  return (byte / 1048576).toLocaleString('it-IT', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' MB';
}

function nomeSenzaEstensione(nome) {
  return (nome.replace(/\.[^.]+$/, '') || 'file').replace(/[\\/:*?"<>|]+/g, '_');
}

/** Evita due file (o cartelle) con lo stesso nome nello stesso ZIP */
function nomeUnico(zip, nome) {
  zip.nomiUsati = zip.nomiUsati || new Set();
  const m = nome.match(/^(.*?)(\.[^.]+)?$/);
  let prova = nome;
  for (let i = 2; zip.nomiUsati.has(prova.toLowerCase()); i++) prova = `${m[1]}-${i}${m[2] || ''}`;
  zip.nomiUsati.add(prova.toLowerCase());
  return prova;
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ---------- Tema chiaro / scuro ---------- */
function impostaTema() {
  let salvato = null;
  try { salvato = localStorage.getItem('ninetto.tema') || localStorage.getItem('nc-tema'); } catch (e) { /* storage non disponibile */ }
  if (salvato === 'light' || salvato === 'dark') document.documentElement.dataset.theme = salvato;
  $('bottoneTema').addEventListener('click', () => {
    const scuroAttuale = document.documentElement.dataset.theme
      ? document.documentElement.dataset.theme === 'dark'
      : matchMedia('(prefers-color-scheme: dark)').matches;
    const nuovo = scuroAttuale ? 'light' : 'dark';
    document.documentElement.dataset.theme = nuovo;
    try { localStorage.setItem('ninetto.tema', nuovo); } catch (e) { /* ignora */ }
  });
}

/* ---------- Link all'editor PDF ---------- */
function impostaEditor() {
  const link = $('linkEditor');
  if (CONFIG.EDITOR_PDF_URL) {
    link.href = CONFIG.EDITOR_PDF_URL;
  } else {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      notifica("L'indirizzo dell'Editor PDF non è ancora impostato (CONFIG.EDITOR_PDF_URL in app.js).");
    });
  }
}

/* ---------- Pulsante "Installa app" (Chrome / Edge / Android) ---------- */
function impostaInstallazione() {
  let richiesta = null;
  const bottone = $('bottoneInstalla');
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    richiesta = e;
    bottone.hidden = false;
  });
  bottone.addEventListener('click', async () => {
    if (!richiesta) return;
    richiesta.prompt();
    await richiesta.userChoice.catch(() => {});
    richiesta = null;
    bottone.hidden = true;
  });
  window.addEventListener('appinstalled', () => { bottone.hidden = true; notifica('App installata! 🎉'); });
}

/* =====================================================================
   STRUMENTI OFFICE (Word ed Excel)
   Le librerie si scaricano solo al primo uso di questi strumenti.
   ===================================================================== */

/* ---------- Caricamento librerie su richiesta ---------- */
const scriptCaricati = {};
function caricaScript(url) {
  if (!scriptCaricati[url]) {
    scriptCaricati[url] = new Promise((ok, ko) => {
      const s = document.createElement('script');
      s.src = url;
      s.onload = ok;
      s.onerror = () => {
        delete scriptCaricati[url];
        s.remove();
        const e = new Error('libreria');
        e.name = 'LibreriaNonCaricata';
        ko(e);
      };
      document.head.appendChild(s);
    });
  }
  return scriptCaricati[url];
}

// docx-preview e la libreria docx usano entrambe il nome globale "docx":
// ognuna viene salvata nella sua variabile e il nome globale viene ripristinato.
let libDocxPreview = null, libDocx = null;
async function caricaConNomeDocx(url) {
  const prima = window.docx;
  await caricaScript(url);
  const lib = window.docx;
  window.docx = prima;
  return lib;
}

async function librerieWord() {
  if (!libDocxPreview) libDocxPreview = await caricaConNomeDocx(CONFIG.LIB.docxPreview);
  await caricaScript(CONFIG.LIB.html2pdf);
  return libDocxPreview;
}

async function librerieExcel() {
  await caricaScript(CONFIG.LIB.xlsx);
  await caricaScript(CONFIG.LIB.jspdf);
  await caricaScript(CONFIG.LIB.autotable);
  return { XLSX: window.XLSX, jsPDF: window.jspdf.jsPDF };
}

async function libreriaDocx() {
  if (!libDocx) libDocx = await caricaConNomeDocx(CONFIG.LIB.docx);
  return libDocx;
}

function estensione(nome) {
  const m = nome.match(/\.([^.]+)$/);
  return m ? m[1].toLowerCase() : '';
}

/* ---------------------------------------------------------------------
   MODULO A – Word in PDF
   docx-preview disegna il documento in un riquadro fuori schermo, html2pdf
   (html2canvas) lo fotografa sezione per sezione, pdf-lib compone il PDF.
   --------------------------------------------------------------------- */
async function eseguiWordInPdf(lavoro) {
  if (valoreOpzione('metodoWord') !== 'immagine') return eseguiWordStampa(lavoro);
  const anteprima = await librerieWord();
  const scalaBase = 2;
  const voci = stato.voci.slice();
  const zip = voci.length > 1 ? new JSZip() : null;
  let ultimo = null, pagineTot = 0;

  for (let i = 0; i < voci.length; i++) {
    controllaAnnullato(lavoro);
    const voce = voci[i];
    const base = i / voci.length, quota = 1 / voci.length;
    mostraAvanzamento(base, `Lettura di ${voce.file.name}…`);
    const dati = await leggiArrayBuffer(voce.file);
    const palco = creaPalcoscenico();
    try {
      try {
        await anteprima.renderAsync(dati, palco, null, {
          inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: false,
          renderHeaders: true, renderFooters: true, renderFootnotes: true,
          useBase64URL: true, experimental: true
        });
      } catch (e) {
        const err = new Error('docx'); err.name = 'DocxNonValido'; err.nomeFile = voce.file.name; throw err;
      }
      await attendiFontEImmagini(palco);
      const sezioni = Array.from(palco.querySelectorAll('section.docx'));
      if (!sezioni.length) { const err = new Error('docx'); err.name = 'DocxNonValido'; err.nomeFile = voce.file.name; throw err; }

      const pdf = await PDFLib.PDFDocument.create();
      for (let n = 0; n < sezioni.length; n++) {
        controllaAnnullato(lavoro);
        mostraAvanzamento(base + quota * (n / sezioni.length), `${voce.file.name}: pagina ${n + 1} di ${sezioni.length}`);
        await pausa();
        const sezione = sezioni[n];
        const r = sezione.getBoundingClientRect();
        const scala = Math.min(scalaBase, Math.sqrt(CONFIG.MAX_PIXEL / (r.width * r.height)), CONFIG.MAX_LATO / r.height);
        const canvas = await html2pdf().set({
          html2canvas: { scale: scala, backgroundColor: '#ffffff', useCORS: true, logging: false }
        }).from(sezione).toCanvas().get('canvas');
        const cs = getComputedStyle(sezione);
        await aggiungiCanvasComePagine(pdf, canvas, r.width * 0.75, altezzaPagina(sezione),
          (parseFloat(cs.paddingTop) || 0) * 0.75, (parseFloat(cs.paddingBottom) || 0) * 0.75);
        liberaCanvas(canvas);
      }
      pagineTot += pdf.getPageCount();
      const byte = await pdf.save({ useObjectStreams: true });
      const nome = nomeSenzaEstensione(voce.file.name) + '.pdf';
      if (zip) zip.file(nomeUnico(zip, nome), byte);
      else ultimo = { blob: new Blob([byte], { type: 'application/pdf' }), nome };
    } finally {
      palco.remove();
    }
  }

  controllaAnnullato(lavoro);
  mostraAvanzamento(0.98, 'Salvataggio…');
  let blob, nome;
  if (zip) { blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' }); nome = 'word-in-pdf.zip'; }
  else ({ blob, nome } = ultimo);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({
    titolo: 'PDF pronto! 🎉',
    righe: [zip ? `${voci.length} documenti convertiti, ${pagineTot} pagine in totale.` : `${pagineTot} ${pagineTot === 1 ? 'pagina' : 'pagine'}.`],
    blob, nome, scaricaSubito: true
  });
}

/**
 * Metodo "PDF di qualità": il documento viene impaginato dal motore di stampa del browser,
 * con il formato pagina e i margini del file Word. Il risultato ha testo vero (selezionabile e nitido).
 */
async function eseguiWordStampa(lavoro) {
  const anteprima = await librerieWord();
  const voci = stato.voci.slice();
  const stampa = document.createElement('div');
  stampa.id = 'areaStampa';
  let regolePagina = '';
  for (let i = 0; i < voci.length; i++) {
    controllaAnnullato(lavoro);
    mostraAvanzamento(i / voci.length, `Impaginazione di ${voci[i].file.name}…`);
    await pausa();
    const contenitore = document.createElement('div');
    contenitore.className = 'documento-stampa';
    try {
      await anteprima.renderAsync(await leggiArrayBuffer(voci[i].file), contenitore, null, {
        inWrapper: false, breakPages: true, ignoreLastRenderedPageBreak: true,
        renderHeaders: false, renderFooters: false, renderFootnotes: true,
        useBase64URL: true, experimental: true
      });
    } catch (e) {
      const err = new Error('docx'); err.name = 'DocxNonValido'; err.nomeFile = voci[i].file.name; throw err;
    }
    const sezioni = contenitore.querySelectorAll('section.docx');
    if (!sezioni.length) { const err = new Error('docx'); err.name = 'DocxNonValido'; err.nomeFile = voci[i].file.name; throw err; }
    // Formato e margini della prima sezione diventano quelli della pagina di stampa
    if (!regolePagina) {
      const st = sezioni[0].style;
      const m = (st.padding || '').split(/\s+/).filter(Boolean);
      const [su, dx, giu, sx] = m.length === 4 ? m : m.length === 2 ? [m[0], m[1], m[0], m[1]] : ['2cm', '2cm', '2cm', '2cm'];
      // docx-preview indica la pagina intera (margini compresi) in larghezza e altezza minima
      const pt = (v) => (v.endsWith('pt') ? parseFloat(v) : v.endsWith('px') ? parseFloat(v) * 0.75 : NaN);
      const w = pt(st.width || ''), h = pt(st.minHeight || '');
      regolePagina = `@page { size: ${isFinite(w) && isFinite(h) ? `${w}pt ${h}pt` : 'A4'}; margin: ${su} ${dx} ${giu} ${sx}; }`;
    }
    stampa.appendChild(contenitore);
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Apertura della stampa…');
  const stile = document.createElement('style');
  stile.id = 'stilePagina';
  stile.textContent = regolePagina;
  document.head.appendChild(stile);
  document.body.appendChild(stampa);
  await attendiFontEImmagini(stampa);
  const titoloPrima = document.title;
  document.title = voci.length === 1 ? nomeSenzaEstensione(voci[0].file.name) : 'documenti-word';
  const pulisci = () => {
    document.title = titoloPrima;
    stampa.remove();
    stile.remove();
    window.removeEventListener('afterprint', pulisci);
  };
  window.addEventListener('afterprint', pulisci);
  await pausa();
  window.print();
  // Su alcuni telefoni "afterprint" non arriva: puliamo comunque dopo un po'
  setTimeout(() => { if (document.body.contains(stampa)) pulisci(); }, 60000);
  el.risultato.classList.remove('errore');
  el.risultato.innerHTML = `<h2>Ultimo passaggio 🖨️</h2>
    <p>Nella finestra di stampa scegli <strong>Salva come PDF</strong> come stampante, poi tocca il pulsante <strong>PDF</strong> (o "Salva").</p>
    <p>La finestra non si è aperta? Tocca di nuovo il pulsante qui sotto.</p>`;
  aggiungiLinkEditor(null); // il PDF lo salva la stampa: link semplice, poi lo si apre con «Apri»
  el.risultato.hidden = false;
}

/** Riquadro fuori schermo in cui docx-preview disegna il documento */
function creaPalcoscenico() {
  const div = document.createElement('div');
  div.className = 'palcoscenico';
  div.setAttribute('aria-hidden', 'true');
  document.body.appendChild(div);
  return div;
}

/** Altezza di pagina (in punti) dichiarata dal documento Word per quella sezione */
function altezzaPagina(sezione) {
  const px = parseFloat(getComputedStyle(sezione).minHeight) || sezione.getBoundingClientRect().height;
  return px * 0.75;
}

async function attendiFontEImmagini(contenitore) {
  if (document.fonts && document.fonts.ready) await document.fonts.ready;
  const immagini = Array.from(contenitore.querySelectorAll('img'));
  await Promise.all(immagini.map((img) => img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; })));
}

/** Inserisce il canvas nel PDF; se il contenuto supera l'altezza di una pagina lo divide su più pagine,
 *  tagliando dove c'è uno spazio bianco (tra due righe di testo) e saltando le parti vuote. */
async function aggiungiCanvasComePagine(pdf, canvas, larghezzaPt, altezzaPt, margineSu = 0, margineGiu = 0) {
  const k = canvas.width / larghezzaPt;          // pixel per punto
  const pxPerPagina = Math.floor(altezzaPt * k);
  const pxSu = Math.round(margineSu * k), pxGiu = Math.round(margineGiu * k);
  // Le pagine dopo la prima riprendono i margini di Word (in alto e in basso)
  const pxContenuto = Math.max(pxPerPagina * 0.5, pxPerPagina - pxSu - pxGiu);
  const ctx = canvas.getContext('2d');
  const rigaBianca = (y) => {
    const d = ctx.getImageData(0, y, canvas.width, 1).data;
    for (let k = 0; k < d.length; k += 4) if (d[k] < 245 || d[k + 1] < 245 || d[k + 2] < 245) return false;
    return true;
  };
  const zonaVuota = (da, a) => { for (let y = da; y < a; y += 3) if (!rigaBianca(y)) return false; return true; };

  const fette = [];
  let y = 0;
  const spazio = () => (fette.length ? pxContenuto : pxPerPagina - pxGiu);
  while (canvas.height - y > (fette.length ? pxContenuto + pxGiu : pxPerPagina) * 1.02) {
    let taglio = y + spazio();
    for (let t = taglio; t > taglio - spazio() * 0.15; t--) { if (rigaBianca(t)) { taglio = t; break; } }
    fette.push([y, taglio]);
    y = taglio;
    while (y < canvas.height && rigaBianca(y) && y - taglio < pxSu) y++; // salta lo spazio bianco a inizio pagina
  }
  if (!fette.length || !zonaVuota(y, canvas.height)) fette.push([y, canvas.height]);

  for (let n = 0; n < fette.length; n++) {
    const [da, a] = fette[n];
    const spostamento = n === 0 ? 0 : margineSu;
    const h = a - da;
    const fetta = document.createElement('canvas');
    fetta.width = canvas.width; fetta.height = h;
    const c = fetta.getContext('2d', { alpha: false });
    c.fillStyle = '#fff'; c.fillRect(0, 0, fetta.width, h);
    c.drawImage(canvas, 0, da, canvas.width, h, 0, 0, canvas.width, h);
    const jpeg = await canvasInBlob(fetta, 'image/jpeg', 0.92);
    liberaCanvas(fetta);
    const img = await pdf.embedJpg(new Uint8Array(await jpeg.arrayBuffer()));
    const pagina = pdf.addPage([larghezzaPt, altezzaPt]);
    const hPt = Math.min(altezzaPt - spostamento, h / k);
    pagina.drawImage(img, { x: 0, y: altezzaPt - spostamento - hPt, width: larghezzaPt, height: hPt });
  }
}

/** Lettura del file come ArrayBuffer con FileReader (compatibile anche con browser datati) */
function leggiArrayBuffer(file) {
  return new Promise((ok, ko) => {
    const r = new FileReader();
    r.onload = () => ok(r.result);
    r.onerror = () => ko(r.error || new Error('lettura'));
    r.readAsArrayBuffer(file);
  });
}

/* ---------------------------------------------------------------------
   MODULO B – Excel in PDF (SheetJS + jsPDF + AutoTable)
   --------------------------------------------------------------------- */

/** Legge la cartella di lavoro, popola il selettore dei fogli e mostra l'anteprima */
async function preparaCartella(voce) {
  const { XLSX } = await librerieExcel();
  const dati = await leggiArrayBuffer(voce.file);
  let cartella;
  try {
    if (estensione(voce.file.name) === 'csv') {
      const testo = decodificaTesto(new Uint8Array(dati));
      const primaRiga = testo.split(/\r?\n/, 1)[0] || '';
      const sep = [';', ',', '\t'].reduce((a, b) => (primaRiga.split(b).length > primaRiga.split(a).length ? b : a));
      cartella = XLSX.read(testo, { type: 'string', FS: sep, raw: true });
    } else {
      cartella = XLSX.read(new Uint8Array(dati), { type: 'array', cellDates: true });
    }
  } catch (e) {
    const err = new Error('foglio'); err.name = 'FoglioNonValido'; throw err;
  }
  voce.cartella = cartella;
  voce.fogli = cartella.SheetNames.slice();
  if (!stato.voci.includes(voce)) return;
  const select = $('sceltaFoglio');
  if (!select) return;
  select.innerHTML = (voce.fogli.length > 1 ? '<option value="*">Tutti i fogli</option>' : '') +
    voce.fogli.map((n, i) => `<option value="${i}">${escapeHtml(n)}</option>`).join('');
  if (voce.fogli.length > 1) select.value = '0';
  $('gruppoFoglio').hidden = false;
  aggiornaAnteprimaFoglio();
}

/** CSV: prova UTF-8, altrimenti la codifica di Excel per Windows */
function decodificaTesto(byte) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(byte).replace(/^﻿/, ''); }
  catch (e) { return new TextDecoder('windows-1252').decode(byte); }
}

/** Righe del foglio come testo già formattato (date, valute, percentuali come in Excel) */
function righeFoglio(foglio) {
  const righe = XLSX.utils.sheet_to_json(foglio, { header: 1, raw: false, defval: '', blankrows: false });
  let colonne = 0;
  righe.forEach((r) => r.forEach((c, i) => { if (String(c).trim() !== '') colonne = Math.max(colonne, i + 1); }));
  return righe
    .map((r) => Array.from({ length: colonne }, (_, i) => (r[i] === undefined || r[i] === null ? '' : String(r[i]))))
    .filter((r) => r.some((c) => c.trim() !== ''));
}

function fogliScelti(voce) {
  const v = $('sceltaFoglio') ? $('sceltaFoglio').value : '0';
  return v === '*' ? voce.fogli.map((_, i) => i) : [Number(v) || 0];
}

function aggiornaAnteprimaFoglio() {
  const voce = stato.voci[0];
  const box = $('anteprimaFoglio');
  if (!box || !voce || !voce.cartella) return;
  const indice = fogliScelti(voce)[0];
  const righe = righeFoglio(voce.cartella.Sheets[voce.fogli[indice]]);
  if (!righe.length) { box.innerHTML = '<p class="nota-opzione">Questo foglio è vuoto.</p>'; return; }
  const MAXR = 12, MAXC = 8;
  const cella = (t, tag) => `<${tag}>${escapeHtml(t.length > 30 ? t.slice(0, 29) + '…' : t)}</${tag}>`;
  const intestazione = righe[0].slice(0, MAXC).map((c) => cella(c, 'th')).join('');
  const corpo = righe.slice(1, MAXR).map((r) => `<tr>${r.slice(0, MAXC).map((c) => cella(c, 'td')).join('')}</tr>`).join('');
  const altro = [];
  if (righe.length > MAXR) altro.push(`${righe.length} righe`);
  if (righe[0].length > MAXC) altro.push(`${righe[0].length} colonne`);
  box.innerHTML = `<div class="tabella-scorrevole"><table><thead><tr>${intestazione}</tr></thead><tbody>${corpo}</tbody></table></div>` +
    `<p class="nota-opzione">Anteprima di "${escapeHtml(voce.fogli[indice])}"${altro.length ? ` · in totale ${altro.join(' e ')}` : ''}</p>`;
}

async function eseguiExcelInPdf(lavoro) {
  const { jsPDF } = await librerieExcel();
  const voce = stato.voci[0];
  if (!voce.cartella) await preparaCartella(voce);
  const orizzontale = valoreOpzione('orientamentoXls') !== 'verticale';
  const adatta = valoreOpzione('adatta') !== 'no';
  const indici = fogliScelti(voce);
  const doc = new jsPDF({ orientation: orizzontale ? 'landscape' : 'portrait', unit: 'pt', format: 'a4' });
  const disegnaTabella = typeof window.autoTable === 'function' ? (opz) => window.autoTable(doc, opz) : (opz) => doc.autoTable(opz);
  const margine = { top: 46, right: 28, bottom: 36, left: 28 };
  let disegnati = 0, righeTot = 0;

  for (let k = 0; k < indici.length; k++) {
    controllaAnnullato(lavoro);
    const nomeFoglio = voce.fogli[indici[k]];
    mostraAvanzamento(k / indici.length, `Foglio "${nomeFoglio}"`);
    await pausa();
    const righe = righeFoglio(voce.cartella.Sheets[nomeFoglio]);
    if (!righe.length) continue;
    if (disegnati > 0) doc.addPage();
    const colonne = righe[0].length;
    // Con "adatta alla pagina" il testo si riduce quando le colonne sono tante
    const corpo = adatta ? Math.max(5, Math.min(9, 9 - (colonne - (orizzontale ? 10 : 6)) * 0.4)) : 8;
    disegnaTabella({
      head: [righe[0]],
      body: righe.slice(1),
      startY: margine.top,
      margin: margine,
      theme: 'grid',
      tableWidth: adatta ? 'auto' : 'wrap',
      horizontalPageBreak: !adatta,
      horizontalPageBreakRepeat: adatta ? null : 0,
      styles: { font: 'helvetica', fontSize: corpo, cellPadding: corpo < 7 ? 2 : 3, overflow: 'linebreak', valign: 'middle', lineColor: [205, 213, 225], lineWidth: 0.4, textColor: [30, 27, 75] },
      headStyles: { fillColor: [21, 128, 61], textColor: 255, fontStyle: 'bold', lineColor: [21, 128, 61] },
      alternateRowStyles: { fillColor: [240, 253, 244] },
      didDrawPage: () => {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        doc.setTextColor(30, 27, 75);
        doc.text(`${nomeSenzaEstensione(voce.file.name)} · ${nomeFoglio}`, margine.left, 28);
      }
    });
    disegnati++;
    righeTot += righe.length - 1;
  }
  if (!disegnati) { const e = new Error('vuoto'); e.name = 'FoglioVuoto'; throw e; }

  // Numeri di pagina in basso a destra
  const tot = doc.getNumberOfPages();
  for (let i = 1; i <= tot; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(120, 120, 140);
    const w = doc.internal.pageSize.getWidth(), h = doc.internal.pageSize.getHeight();
    doc.text(`Pagina ${i} di ${tot}`, w - margine.right, h - 16, { align: 'right' });
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  const blob = doc.output('blob');
  mostraRisultato({
    titolo: 'PDF pronto! 🎉',
    righe: [`${disegnati} ${disegnati === 1 ? 'foglio' : 'fogli'}, ${righeTot} righe, ${tot} ${tot === 1 ? 'pagina' : 'pagine'} A4 ${orizzontale ? 'orizzontali' : 'verticali'}.`],
    blob, nome: nomeSenzaEstensione(voce.file.name) + '.pdf', scaricaSubito: true
  });
}

/* ---------------------------------------------------------------------
   MODULO C – PDF in Word (pdf.js + docx)
   --------------------------------------------------------------------- */
async function eseguiPdfInWord(lavoro) {
  const D = await libreriaDocx();
  const voci = stato.voci.slice();
  const totale = totalePagine() || voci.length;
  const zip = voci.length > 1 ? new JSZip() : null;
  let fatte = 0, ultimo = null, senzaTesto = 0;

  for (const voce of voci) {
    const pdf = await apriPdfJs(voce);
    const paragrafi = [];
    let pagineVuote = 0;
    try {
      for (let n = 1; n <= pdf.numPages; n++) {
        controllaAnnullato(lavoro);
        mostraAvanzamento(fatte / totale, `Pagina ${fatte + 1} di ${totale}`);
        await pausa();
        const pagina = await pdf.getPage(n);
        const blocchi = await estraiParagrafi(pagina);
        pagina.cleanup();
        if (!blocchi.length) {
          pagineVuote++;
          blocchi.push({ testo: `[Pagina ${n}: nessun testo, probabilmente è un'immagine o una scansione]`, dimensione: 10, grassetto: false, rientro: 0, corsivo: true });
        }
        blocchi.forEach((b, i) => paragrafi.push(new D.Paragraph({
          pageBreakBefore: i === 0 && n > 1,
          spacing: { after: Math.round(b.dimensione * 6) },
          indent: b.rientro > 12 ? { left: Math.round(b.rientro * 20) } : undefined,
          children: [new D.TextRun({ text: b.testo, size: Math.round(b.dimensione * 2), bold: b.grassetto, italics: !!b.corsivo })]
        })));
        fatte++;
      }
    } finally {
      pdf.destroy();
    }
    if (pagineVuote === pdf.numPages) senzaTesto++;
    const documento = new D.Document({
      creator: 'ninetto-converter',
      title: nomeSenzaEstensione(voce.file.name),
      styles: { default: { document: { run: { font: 'Calibri' } } } },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        children: paragrafi
      }]
    });
    const blob = await D.Packer.toBlob(documento);
    const nome = nomeSenzaEstensione(voce.file.name) + '.docx';
    if (zip) zip.file(nomeUnico(zip, nome), blob);
    else ultimo = { blob, nome };
  }

  if (senzaTesto === voci.length) {
    const e = new Error('scansione'); e.name = 'NessunTesto'; throw e;
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(0.98, 'Salvataggio…');
  let blob, nome;
  if (zip) { blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' }); nome = 'pdf-in-word.zip'; }
  else ({ blob, nome } = ultimo);
  mostraAvanzamento(1, 'Fatto');
  const righe = [`${fatte} ${fatte === 1 ? 'pagina estratta' : 'pagine estratte'} in un documento Word modificabile.`];
  if (senzaTesto) righe.push(`${senzaTesto} file senza testo (scansioni): per quelli serve un OCR.`);
  mostraRisultato({ titolo: 'Documento Word pronto! 🎉', righe, blob, nome, scaricaSubito: true });
}

/**
 * Estrae il testo di una pagina e lo ricompone in paragrafi:
 * 1) frammenti → righe (stessa altezza), ordinate dall'alto in basso e da sinistra a destra;
 * 2) righe vicine con lo stesso corpo → stesso paragrafo.
 */
async function estraiParagrafi(pagina) {
  const vp = pagina.getViewport({ scale: 1 });
  const contenuto = await pagina.getTextContent();
  const pezzi = [];
  for (const it of contenuto.items) {
    if (!it.str || !it.str.trim()) continue;
    const [x, y] = pdfjsLib.Util.applyTransform([it.transform[4], it.transform[5]], vp.transform);
    const h = Math.hypot(it.transform[2], it.transform[3]) || it.height || 10;
    const stile = contenuto.styles[it.fontName] || {};
    pezzi.push({ x, y, h, w: it.width || 0, testo: it.str, grassetto: /bold|black|heavy/i.test(it.fontName + ' ' + (stile.fontFamily || '')) });
  }
  if (!pezzi.length) return [];
  pezzi.sort((a, b) => a.y - b.y || a.x - b.x);

  // 1) righe
  const righe = [];
  for (const p of pezzi) {
    const riga = righe.find((r) => Math.abs(r.y - p.y) < Math.max(2, Math.min(r.h, p.h) * 0.5));
    if (riga) { riga.pezzi.push(p); riga.h = Math.max(riga.h, p.h); }
    else righe.push({ y: p.y, h: p.h, pezzi: [p] });
  }
  righe.sort((a, b) => a.y - b.y);
  for (const r of righe) {
    r.pezzi.sort((a, b) => a.x - b.x);
    let testo = '', fine = null;
    for (const p of r.pezzi) {
      if (fine !== null && p.x - fine > p.h * 0.2 && !testo.endsWith(' ') && !p.testo.startsWith(' ')) testo += ' ';
      testo += p.testo;
      fine = p.x + p.w;
    }
    r.testo = testo.replace(/\s+/g, ' ').trim();
    r.x = r.pezzi[0].x;
    r.grassetto = r.pezzi.every((p) => p.grassetto);
  }
  const valide = righe.filter((r) => r.testo);
  if (!valide.length) return [];

  // Corpo del testo "normale" = altezza più frequente
  const conteggio = {};
  valide.forEach((r) => { const k = Math.round(r.h); conteggio[k] = (conteggio[k] || 0) + r.testo.length; });
  const corpoBase = Number(Object.keys(conteggio).sort((a, b) => conteggio[b] - conteggio[a])[0]) || 10;
  const sinistra = Math.min(...valide.map((r) => r.x));

  // 2) paragrafi
  const paragrafi = [];
  let corrente = null, rigaPrec = null;
  for (const r of valide) {
    const stessoCorpo = corrente && Math.abs(r.h - corrente.h) <= Math.max(1, corrente.h * 0.15);
    const vicina = rigaPrec && (r.y - rigaPrec.y) <= rigaPrec.h * 1.7;
    if (corrente && stessoCorpo && vicina) {
      corrente.testo = corrente.testo.endsWith('-') ? corrente.testo.slice(0, -1) + r.testo : corrente.testo + ' ' + r.testo;
      corrente.grassetto = corrente.grassetto && r.grassetto;
    } else {
      corrente = { testo: r.testo, h: r.h, x: r.x, grassetto: r.grassetto };
      paragrafi.push(corrente);
    }
    rigaPrec = r;
  }
  return paragrafi.map((p) => ({
    testo: p.testo,
    dimensione: Math.max(6, Math.min(40, Math.round(p.h * 2) / 2)),
    grassetto: p.grassetto || p.h >= corpoBase * 1.25,
    rientro: p.x - sinistra
  }));
}
