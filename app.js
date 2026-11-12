/* =====================================================================
   ninetto-converter – logica dell'app
   Tutto avviene nel browser: i file non lasciano mai il dispositivo.
   Librerie (caricate da CDN in index.html):
     - pdf-lib  → window.PDFLib     (creare / unire PDF)
     - pdf.js   → window.pdfjsLib   (leggere e disegnare le pagine PDF)
     - JSZip    → window.JSZip      (creare archivi .zip)
   ===================================================================== */
'use strict';

/* ---------------------------------------------------------------------
   CONFIGURAZIONE – modifica qui
   --------------------------------------------------------------------- */
const CONFIG = {
  // Indirizzo aperto dal pulsante "Apri Editor PDF" (es. 'https://tuosito.github.io/editor-pdf/').
  // Se resta vuoto il pulsante avvisa che l'indirizzo non è ancora impostato.
  EDITOR_PDF_URL: '',
  // Worker di pdf.js: stessa versione della libreria in index.html
  PDFJS_WORKER: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  // Limite di pixel per un singolo canvas: oltre, i telefoni (soprattutto iPhone) si bloccano
  MAX_PIXEL: 16000000,
  MAX_LATO: 8192,
  A4: [595.28, 841.89],      // punti tipografici (1 pt = 1/72 di pollice)
  MARGINE_A4: 28.35,         // 1 cm
  // Livelli di compressione: qualità JPEG e risoluzione di ricampionamento
  COMPRESSIONE: {
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

/* ---------------------------------------------------------------------
   Definizione dei 4 strumenti
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
    accept: 'image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp',
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
        ${SCELTA('livello', 'bassa', 'Bassa compressione', 'alta qualità visiva · JPEG 85%')}
        ${SCELTA('livello', 'media', 'Media · consigliata', 'per email · JPEG 70% a 150 DPI', true)}
        ${SCELTA('livello', 'massima', 'Massima compressione', 'file minimo · JPEG 50%')}
      </div>
      <p class="nota-opzione">Le pagine vengono trasformate in immagini: il testo non sarà più selezionabile né ricercabile, e i link interni si perdono.</p>
      </fieldset>`,
    etichetta: () => stato.voci.length > 1 ? `Comprimi ${stato.voci.length} PDF (ZIP)` : 'Comprimi PDF',
    esegui: eseguiComprimi
  }
};

/* =====================================================================
   AVVIO
   ===================================================================== */
const librerieOk = !!(window.PDFLib && window.pdfjsLib && window.JSZip);
if (!librerieOk) $('avvisoLibrerie').hidden = false;

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
el.opzioni.addEventListener('change', aggiornaInterfaccia);
window.addEventListener('hashchange', instrada);
instrada();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
}

/* =====================================================================
   NAVIGAZIONE (#img2pdf, #pdf2img, #unisci, #comprimi)
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
  if (/^image\/(jpeg|png|webp)$/.test(file.type) || /\.(jpe?g|png|webp)$/.test(nome)) return 'immagine';
  return null;
}

function aggiungiFile(elencoFile) {
  if (!stato.strumento || stato.lavoro) return;
  const s = STRUMENTI[stato.strumento];
  const scartati = [];
  for (const file of elencoFile) {
    if (tipoFile(file) !== s.tipo) { scartati.push(file.name); continue; }
    const voce = { id: stato.prossimoId++, file, tipo: s.tipo, url: null, pagine: null, errore: null, password: null };
    stato.voci.push(voce);
    el.lista.appendChild(creaElementoVoce(voce));
  }
  if (scartati.length) {
    const atteso = s.tipo === 'pdf' ? 'solo file PDF' : 'solo immagini JPG, PNG o WebP';
    notifica(`Formato non supportato (${atteso}): ${scartati.join(', ')}`, true);
  }
  nascondiRisultato();
  aggiornaInterfaccia();
  generaAnteprime();
}

/* ---------- Elemento della lista con miniatura e comandi ---------- */
const ICONA = {
  su: '<svg viewBox="0 0 24 24"><path d="M6 15l6-6 6 6"/></svg>',
  giu: '<svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg>',
  rimuovi: '<svg viewBox="0 0 24 24"><path d="M5 7h14M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
  maniglia: '<svg viewBox="0 0 24 24"><circle cx="9" cy="6" r="1"/><circle cx="15" cy="6" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="9" cy="18" r="1"/><circle cx="15" cy="18" r="1"/></svg>',
  scarica: '<svg viewBox="0 0 24 24"><path d="M12 4v11m0 0-4-4m4 4 4-4M5 20h14"/></svg>'
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
      const li = trovaLi(voce.id);
      if (!li) continue;
      const box = li.querySelector('.miniatura');
      try {
        if (voce.tipo === 'immagine') {
          voce.url = URL.createObjectURL(voce.file);
          const img = new Image();
          img.alt = '';
          img.decoding = 'async';
          await new Promise((ok, ko) => { img.onload = ok; img.onerror = () => ko(new Error('immagine')); img.src = voce.url; });
          box.replaceChildren(img);
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
  el.contenitoreLista.hidden = n === 0;
  el.conteggio.textContent = `${n} ${n === 1 ? 'file caricato' : 'file caricati'}`;
  el.suggerimento.hidden = n < 2;
  el.sottotestoArea.textContent = n ? 'puoi aggiungerne altri, anche trascinandoli qui' : 'oppure trascinali qui · selezione multipla';
  el.lista.querySelectorAll('.voce-file').forEach((li, i) => {
    li.querySelector('.numero').textContent = i + 1;
    li.querySelector('.su').disabled = i === 0;
    li.querySelector('.giu').disabled = i === n - 1;
  });
  const conErrori = stato.voci.some((v) => v.errore);
  el.azione.textContent = stato.lavoro ? 'Elaborazione in corso…' : (n ? s.etichetta() : s.testoArea);
  el.azione.disabled = !librerieOk || !!stato.lavoro || n < s.minimo || conErrori;
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
  el.risultato.innerHTML = `<h2></h2>${r.righe.map((x) => `<p>${x}</p>`).join('')}
    <a class="scarica" download>${ICONA.scarica}<span></span></a>`;
  el.risultato.querySelector('h2').textContent = r.titolo;
  const a = el.risultato.querySelector('.scarica');
  a.href = stato.urlRisultato;
  a.download = r.nome;
  a.querySelector('span').textContent = `Scarica ${r.nome} (${formattaPeso(r.blob.size)})`;
  el.risultato.hidden = false;
  el.risultato.scrollIntoView({ behavior: 'smooth', block: 'center' });
  if (r.scaricaSubito) a.click();
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
        zip.file(`${cartella}pagina-${String(n).padStart(cifre, '0')}.${est}`, blob);
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
    blob, nome: `${voci.length > 1 ? 'pdf' : nomeBase}-immagini.zip`, scaricaSubito: true });
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
   MODULO 4 – Comprimi PDF
   --------------------------------------------------------------------- */
async function eseguiComprimi(lavoro) {
  const { PDFDocument } = PDFLib;
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

  controllaAnnullato(lavoro);
  mostraAvanzamento(0.98, 'Salvataggio…');
  const prima = risultati.reduce((t, r) => t + r.prima, 0);
  const dopo = risultati.reduce((t, r) => t + r.dopo, 0);
  const righe = [confronto(prima, dopo)];
  if (risultati.length > 1) {
    risultati.forEach((r) => righe.push(`<small>${escapeHtml(r.voce.file.name)}: ${formattaPeso(r.prima)} → ${formattaPeso(r.dopo)}</small>`));
  }
  const migliorato = dopo < prima;
  if (!migliorato) righe.push('Il file era già ottimizzato: la versione compressa non è più leggera. Ti conviene tenere l\'originale o provare la compressione massima.');

  let blob, nome;
  if (risultati.length === 1) {
    blob = new Blob([risultati[0].byte], { type: 'application/pdf' });
    nome = nomeSenzaEstensione(voci[0].file.name) + '-compresso.pdf';
  } else {
    const zip = new JSZip();
    risultati.forEach((r) => zip.file(nomeUnico(zip, nomeSenzaEstensione(r.voce.file.name) + '-compresso.pdf'), r.byte));
    blob = await zip.generateAsync({ type: 'blob', compression: 'STORE' });
    nome = 'pdf-compressi.zip';
  }
  controllaAnnullato(lavoro);
  mostraAvanzamento(1, 'Fatto');
  mostraRisultato({ titolo: migliorato ? 'PDF compresso! 🎉' : 'Compressione completata', righe, blob, nome, scaricaSubito: migliorato });
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
  if (e && e.name === 'PasswordException') return `${chi}PDF protetto da password, non sbloccato.`;
  if (/encrypt/i.test(testo)) return `${chi}PDF protetto da password: non si può unire. Rimuovi la protezione (o comprimilo prima con questa app, che chiede la password) e riprova.`;
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
  try { salvato = localStorage.getItem('nc-tema'); } catch (e) { /* storage non disponibile */ }
  if (salvato === 'light' || salvato === 'dark') document.documentElement.dataset.theme = salvato;
  $('bottoneTema').addEventListener('click', () => {
    const scuroAttuale = document.documentElement.dataset.theme
      ? document.documentElement.dataset.theme === 'dark'
      : matchMedia('(prefers-color-scheme: dark)').matches;
    const nuovo = scuroAttuale ? 'light' : 'dark';
    document.documentElement.dataset.theme = nuovo;
    try { localStorage.setItem('nc-tema', nuovo); } catch (e) { /* ignora */ }
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
