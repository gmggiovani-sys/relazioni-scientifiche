/* global Word, Office */
// Relazioni Tecniche: tabelle di foto con didascalia "Foto N – descrizione".
// Layout 1: una foto per riga, didascalia nella colonna a fianco.
// Layout 2 o 4: righe di foto alternate a righe di didascalie.
// Le operazioni che ridispongono la tabella la ricostruiscono dalle foto che contiene (stesso ordine e didascalie).

import { supporta } from "./word.js";

const CM = 28.3465;                          // punti per centimetro
const LARGHEZZA_FOTO = 7 * CM;
const MARGINE_CELLA = 0.5 * CM;
export const TITOLO_CAPITOLO = "DOCUMENTAZIONE FOTOGRAFICA";
const RE_DIDASCALIA = /^\s*Foto\s+(\d+)\s*(?:[–-]|$)/;
const RE_FOTO_N = /Foto\s+\d+/;
const PER_SYNC = 4;                          // foto inviate a Word per ogni richiesta

function richiedi() {
  if (!supporta("WordApi", "1.3"))
    throw new Error("Questa versione di Word non permette agli add-in di lavorare sulle tabelle: aggiorna Word.");
}

// "IMG_0012_vista_nord.jpg" → "IMG 0012 vista nord"
export const nomeFoto = file => file.replace(/\.[^.]+$/, "").replace(/_/g, " ");

// Celle di foto e didascalia dell'i-esima foto (da 0, indici di riga e colonna da 0)
function posizioni(i, colonne) {
  if (colonne === 1) return { fr: i, fc: 0, dr: i, dc: 1 };
  const fr = Math.floor(i / colonne) * 2;
  return { fr, fc: i % colonne, dr: fr + 1, dc: i % colonne };
}

const righeTabella = (n, colonne) => (colonne === 1 ? n : Math.ceil(n / colonne) * 2);
const colonneTabella = colonne => (colonne === 1 ? 2 : colonne);

// ===============================
// LETTURA DI UNA TABELLA FOTO
// ===============================

// Foto della tabella in ordine di lettura: [{ fr, fc, dr, dc, foto (InlinePicture), didascalia }]
// e foto per riga (1 se le didascalie stanno nella colonna a fianco)
async function leggiTabella(context, tabella) {
  tabella.load("rowCount,values");
  await context.sync();

  const celle = tabella.values.map((riga, r) => riga.map((_, c) => {
    const pics = tabella.getCell(r, c).body.inlinePictures;
    pics.load("items/width,items/height");
    return pics;
  }));
  await context.sync();

  const haFoto = (r, c) => !!(celle[r] && celle[r][c] && celle[r][c].items.length > 0);
  const testo = (r, c) => ((tabella.values[r] || [])[c] || "").trim();
  const nColonne = Math.max(...tabella.values.map(r => r.length));

  // foto e didascalie affiancate: due colonne e nessuna immagine nella seconda
  const affiancate = nColonne === 2 && !tabella.values.some((_, r) => haFoto(r, 1));
  const elementi = [];

  if (affiancate) {
    for (let r = 0; r < tabella.rowCount; r++)
      if (haFoto(r, 0)) elementi.push({ fr: r, fc: 0, dr: r, dc: 1, foto: celle[r][0].items[0], didascalia: testo(r, 1) });
    return { colonne: 1, elementi };
  }

  for (let r = 0; r < tabella.rowCount; r++) {
    if (!tabella.values[r].some((_, c) => haFoto(r, c))) continue;
    const rigaDid = r + 1 < tabella.rowCount && !tabella.values[r + 1].some((_, c) => haFoto(r + 1, c)) ? r + 1 : -1;
    for (let c = 0; c < tabella.values[r].length; c++)
      if (haFoto(r, c))
        elementi.push({ fr: r, fc: c, dr: rigaDid, dc: c, foto: celle[r][c].items[0], didascalia: rigaDid >= 0 ? testo(rigaDid, c) : "" });
  }
  return { colonne: nColonne, elementi };
}

async function tabellaAlCursore(context) {
  const t = context.document.getSelection().parentTableOrNullObject;
  t.load("rowCount");
  await context.sync();
  if (t.isNullObject) throw new Error("Posiziona il cursore nella tabella delle foto.");
  return t;
}

// Immagini delle foto (base64), lette a gruppi
async function immagini(context, elementi) {
  const res = [];
  for (let i = 0; i < elementi.length; i += PER_SYNC) {
    const parte = elementi.slice(i, i + PER_SYNC).map(e => e.foto.getBase64ImageSrc());
    await context.sync();
    res.push(...parte.map(x => x.value));
  }
  return res;
}

// ===============================
// CREAZIONE DI UNA TABELLA FOTO
// ===============================

// Tabella con le foto [{ base64, didascalia }] e il layout indicato, prima del paragrafo dato
async function creaTabella(context, paragrafo, foto, colonne, avanzamento) {
  const righe = righeTabella(foto.length, colonne), col = colonneTabella(colonne);
  const valori = Array.from({ length: righe }, () => Array(col).fill(""));
  foto.forEach((f, i) => { const p = posizioni(i, colonne); valori[p.dr][p.dc] = f.didascalia || "Foto 0"; });

  const t = paragrafo.insertTable(righe, col, "Before", valori);
  t.load("width");
  await context.sync();

  t.alignment = "Centered";
  for (const dove of ["Outside", "Inside"]) t.getBorder(dove).type = "Single";
  const larghezza = Math.min(LARGHEZZA_FOTO, (t.width || 481) / col - MARGINE_CELLA);

  for (let i = 0; i < foto.length; i++) {
    const p = posizioni(i, colonne);
    const cf = t.getCell(p.fr, p.fc), cd = t.getCell(p.dr, p.dc);
    const img = cf.body.insertInlinePictureFromBase64(foto[i].base64, "Start");
    img.lockAspectRatio = true;
    img.width = larghezza;
    cf.horizontalAlignment = "Centered";
    cd.horizontalAlignment = "Centered";
    cd.body.font.italic = true;
    cd.body.font.size = 10;
    if (i % PER_SYNC === PER_SYNC - 1) {
      if (avanzamento) avanzamento(i + 1, foto.length);
      await context.sync();
    }
  }
  await context.sync();
  return t;
}

// Ricostruisce la tabella con le foto indicate e il layout scelto, nello stesso punto.
// La nuova tabella si crea prima del primo paragrafo FUORI da ogni tabella che segue la vecchia,
// dopo aver tolto la vecchia. In Word LTSC i paragrafi "prima" o "dopo" la tabella dati da Word
// (insertParagraph, getParagraphAfter) possono cadere nella sua prima cella: la nuova tabella
// nasceva dentro la vecchia e spariva con lei. Per questo il paragrafo si cerca confrontando le posizioni.
async function sostituisciTabella(context, vecchia, foto, colonne, avanzamento) {
  const ancora = await paragrafoDopo(context, vecchia);

  vecchia.delete();
  await context.sync();

  return creaTabella(context, ancora, foto, colonne, avanzamento);
}

// Primo paragrafo del corpo, fuori dalle tabelle, che viene dopo la tabella (se non c'è, uno nuovo in fondo)
async function paragrafoDopo(context, tabella) {
  const ps = context.document.body.paragraphs;
  ps.load("items/tableNestingLevel");
  const zona = tabella.getRange("Whole");
  await context.sync();

  const fuori = ps.items.filter(p => p.tableNestingLevel === 0);
  const confronti = fuori.map(p => p.getRange("Whole").compareLocationWith(zona));
  await context.sync();

  const k = confronti.findIndex(c => c.value === "After" || c.value === "AdjacentAfter");
  if (k >= 0) return fuori[k];

  const nuovo = context.document.body.insertParagraph("", "End");
  nuovo.styleBuiltIn = Word.Style.normal;
  await context.sync();
  return nuovo;
}

// ===============================
// COMANDI
// ===============================

// Inserisce le foto [{ base64, nome }] in una nuova tabella, dove si trova il cursore
export async function inserisci(foto, colonne, avanzamento) {
  richiedi();
  return Word.run(async context => {
    const sel = context.document.getSelection();
    const t = sel.parentTableOrNullObject;
    const p = sel.paragraphs.getLast();
    t.load("rowCount");
    p.load("text");
    await context.sync();
    if (!t.isNullObject) throw new Error("Posiziona il cursore fuori da una tabella.");

    // su un paragrafo vuoto, così la tabella non spezza il testo
    const dopo = p.text.trim() ? p.insertParagraph("", "After") : p;
    const elenco = foto.map((f, i) => ({ base64: f.base64, didascalia: "Foto " + (i + 1) + " – " + nomeFoto(f.nome) }));
    await creaTabella(context, dopo, elenco, colonne, avanzamento);
    const n = await rinumeraTutto(context);
    return "Foto inserite: " + foto.length + ". Totale foto nel documento: " + n + ".";
  });
}

// Aggiunge foto in coda alla tabella sotto il cursore, con lo stesso layout
export async function aggiungi(foto, avanzamento) {
  richiedi();
  return Word.run(async context => {
    const t = await tabellaAlCursore(context);
    const { colonne, elementi } = await leggiTabella(context, t);
    if (!elementi.length) throw new Error("La tabella non contiene foto.");
    const esistenti = await immagini(context, elementi);
    const elenco = elementi.map((e, i) => ({ base64: esistenti[i], didascalia: e.didascalia }))
      .concat(foto.map(f => ({ base64: f.base64, didascalia: "Foto 0 – " + nomeFoto(f.nome) })));
    await sostituisciTabella(context, t, elenco, colonne, avanzamento);
    const n = await rinumeraTutto(context);
    return "Foto aggiunte: " + foto.length + ". Totale foto nel documento: " + n + ".";
  });
}

// Ridispone la tabella sotto il cursore con 1, 2 o 4 foto per riga
export async function cambiaLayout(colonne, avanzamento) {
  richiedi();
  return Word.run(async context => {
    const t = await tabellaAlCursore(context);
    const { elementi } = await leggiTabella(context, t);
    if (!elementi.length) throw new Error("La tabella non contiene foto.");
    const b64 = await immagini(context, elementi);
    const nuova = await sostituisciTabella(context, t, elementi.map((e, i) => ({ base64: b64[i], didascalia: e.didascalia })), colonne, avanzamento);
    await rinumeraTutto(context);
    nuova.getCell(0, 0).body.getRange("Start").select();
    await context.sync();
    return "Tabella ridisposta con " + colonne + (colonne === 1 ? " foto" : " foto") + " per riga.";
  });
}

// Dopo aver tolto delle foto: elimina i posti vuoti e le didascalie rimaste senza foto, poi rinumera
export async function riorganizza(avanzamento) {
  richiedi();
  return Word.run(async context => {
    const t = await tabellaAlCursore(context);
    const { colonne, elementi } = await leggiTabella(context, t);
    if (!elementi.length) {
      t.delete();
      await context.sync();
      return "Nella tabella non c'erano più foto: è stata eliminata. Foto nel documento: " + await rinumeraTutto(context) + ".";
    }
    const b64 = await immagini(context, elementi);
    await sostituisciTabella(context, t, elementi.map((e, i) => ({ base64: b64[i], didascalia: e.didascalia })), colonne, avanzamento);
    return "Foto riorganizzate e rinumerate: " + await rinumeraTutto(context) + ".";
  });
}

// Scambia la foto sotto il cursore (con la didascalia) con la precedente (-1) o la successiva (+1)
export async function sposta(direzione) {
  richiedi();
  return Word.run(async context => {
    const sel = context.document.getSelection();
    const cella = sel.parentTableCellOrNullObject;
    cella.load("rowIndex,cellIndex");
    const t = await tabellaAlCursore(context);
    const { elementi } = await leggiTabella(context, t);

    const i = elementi.findIndex(e => (e.fr === cella.rowIndex && e.fc === cella.cellIndex) ||
                                      (e.dr === cella.rowIndex && e.dc === cella.cellIndex));
    if (i < 0) throw new Error("Posiziona il cursore su una foto o sulla sua didascalia.");
    const j = i + direzione;
    if (j < 0 || j >= elementi.length) return direzione < 0 ? "È già la prima foto." : "È già l'ultima foto.";

    const [a, b] = [elementi[i], elementi[j]];
    const [ba, bb] = await immagini(context, [a, b]);
    const metti = (e, base64, larghezza, didascalia) => {
      const cf = t.getCell(e.fr, e.fc);
      cf.body.clear();
      const img = cf.body.insertInlinePictureFromBase64(base64, "Start");
      img.lockAspectRatio = true;
      img.width = larghezza;
      if (e.dr >= 0) {
        const cd = t.getCell(e.dr, e.dc);
        cd.body.clear();
        cd.body.insertText(didascalia, "Start");
        cd.body.font.italic = true;
        cd.body.font.size = 10;
      }
    };
    metti(a, bb, a.foto.width, b.didascalia);
    metti(b, ba, b.foto.width, a.didascalia);
    await context.sync();

    await rinumeraTutto(context);
    t.getCell(b.fr, b.fc).body.getRange("Start").select();          // il cursore segue la foto
    await context.sync();
    return "";
  });
}

// Rinumera in sequenza le didascalie "Foto N" di tutte le tabelle; restituisce il numero di foto
async function rinumeraTutto(context) {
  const tabelle = context.document.body.tables;
  tabelle.load("items/values");
  await context.sync();

  let n = 0;
  const cambi = [];
  for (const t of tabelle.items)
    t.values.forEach((riga, r) => riga.forEach((v, c) => {
      if (!RE_DIDASCALIA.test(v)) return;
      n++;
      const attuale = v.match(RE_FOTO_N)[0];
      if (attuale !== "Foto " + n) cambi.push({ t, r, c, attuale, nuovo: "Foto " + n });
    }));

  // solo il numero cambia: descrizione e formattazione restano
  const ricerche = cambi.map(x => { const rc = x.t.getCell(x.r, x.c).body.search(x.attuale, { matchCase: true }); rc.load("items"); return rc; });
  await context.sync();
  ricerche.forEach((rc, k) => { if (rc.items.length) rc.items[0].insertText(cambi[k].nuovo, "Replace"); });
  await context.sync();
  return n;
}

export async function rinumera() {
  richiedi();
  return Word.run(async context => "Foto rinumerate: " + await rinumeraTutto(context) + ".");
}

// Titolo "DOCUMENTAZIONE FOTOGRAFICA" (Titolo 1): sopra la tabella se il cursore è in una tabella,
// altrimenti dopo il paragrafo corrente
export async function capitoloFoto() {
  return Word.run(async context => {
    const sel = context.document.getSelection();
    const t = sel.parentTableOrNullObject;
    const p = sel.paragraphs.getLast();
    t.load("rowCount");
    p.load("text");
    await context.sync();

    if (!t.isNullObject) {
      const h = t.insertParagraph(TITOLO_CAPITOLO, "Before");
      h.styleBuiltIn = Word.Style.heading1;
      h.getRange("End").select();
      await context.sync();
      return "";
    }

    let h;
    if (!p.text.trim()) { p.insertText(TITOLO_CAPITOLO, "Replace"); h = p; }
    else h = p.insertParagraph(TITOLO_CAPITOLO, "After");
    h.styleBuiltIn = Word.Style.heading1;
    const dopo = h.insertParagraph("", "After");
    dopo.styleBuiltIn = Word.Style.normal;
    dopo.getRange("Start").select();
    await context.sync();
    return "";
  });
}

// Riduce le immagini del documento: lato lungo al massimo maxPx, JPEG; larghezza al massimo 15 cm.
// riduci(base64) → Promise<base64 | null> (null: lasciare com'è)
export async function comprimi(riduci, avanzamento) {
  return Word.run(async context => {
    const pics = context.document.body.inlinePictures;
    pics.load("items/width,items/height");
    await context.sync();

    let ridotte = 0, prima = 0, dopo = 0;
    const max = 15 * CM;

    for (let i = 0; i < pics.items.length; i += PER_SYNC) {
      if (avanzamento) avanzamento(i, pics.items.length);
      const gruppo = pics.items.slice(i, i + PER_SYNC);
      const dati = gruppo.map(p => p.getBase64ImageSrc());
      await context.sync();

      for (let k = 0; k < gruppo.length; k++) {
        const p = gruppo[k], b64 = dati[k].value;
        let w = p.width, h = p.height;
        if (w > max) { h = h * max / w; w = max; }
        const nuovo = await riduci(b64);
        prima += b64.length * 0.75;
        if (nuovo && nuovo.length < b64.length * 0.9) {
          const q = p.insertInlinePictureFromBase64(nuovo, "Replace");
          q.lockAspectRatio = false;
          q.width = w;
          q.height = h;
          q.lockAspectRatio = true;
          ridotte++;
          dopo += nuovo.length * 0.75;
        } else {
          if (w !== p.width) { p.lockAspectRatio = true; p.width = w; }
          dopo += b64.length * 0.75;
        }
      }
      await context.sync();
    }

    const mb = x => (x / 1048576).toFixed(1).replace(".", ",") + " MB";
    return "Immagini nel documento: " + pics.items.length + ", ridotte: " + ridotte +
           ".\nPeso delle immagini: da " + mb(prima) + " a " + mb(dopo) + ".";
  });
}

// PDF del documento (lo prepara Word), come Blob da scaricare
export function pdf() {
  return new Promise((ok, ko) => {
    Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 4194304 }, r => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { ko(new Error(r.error.message)); return; }
      const file = r.value, parti = [];
      const leggi = i => file.getSliceAsync(i, s => {
        if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); ko(new Error(s.error.message)); return; }
        parti.push(new Uint8Array(s.value.data));
        if (i + 1 < file.sliceCount) leggi(i + 1);
        else { file.closeAsync(); ok(new Blob(parti, { type: "application/pdf" })); }
      });
      leggi(0);
    });
  });
}
