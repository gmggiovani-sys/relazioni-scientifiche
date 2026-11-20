/* global Word, Office */
// Funzioni comuni per leggere e trovare il testo nel documento

export const supporta = (insieme, versione) => Office.context.requirements.isSetSupported(insieme, versione);

// Maiuscoletto leggibile e impostabile dagli add-in (Word per Mac 16.99+, Windows recente)
export const maiuscolettoDisponibile = () => supporta("WordApiDesktop", "1.3");

// Testo dei paragrafi con lo stile indicato, nell'ordine del documento
export async function vociConStile(context, stile) {
  const ps = context.document.body.paragraphs;
  ps.load("items/style,items/text");
  await context.sync();
  const paragrafi = ps.items.filter(p => p.style === stile && p.text.trim().length > 0);
  return { paragrafi, testi: paragrafi.map(p => p.text) };
}

// Note a piè di pagina: tutte (origine null) o quelle richiamate in un Range (es. la selezione),
// con il numero da mostrare
export async function note(context, origine) {
  if (!supporta("WordApi", "1.5"))
    throw new Error("Questa versione di Word non permette agli add-in di leggere le note a piè di pagina: aggiorna Word.");

  const fn = (origine || context.document.body).footnotes;

  // il numero visibile della nota, se Word lo dà; altrimenti la posizione nell'elenco
  let conNumero = true;
  try {
    fn.load("items/body/text,items/reference/text");
    await context.sync();
  } catch (e) {
    conNumero = false;
    fn.load("items/body/text");
    await context.sync();
  }

  return fn.items.map((f, i) => {
    const n = conNumero ? (f.reference.text || "").trim() : "";
    return {
      nota: f,
      corpo: f.body,
      testo: f.body.text || "",
      dove: /^\d+$/.test(n) ? "Nota " + n : (origine ? "Nota " + (i + 1) + " della selezione" : "Nota " + (i + 1))
    };
  });
}

// Note di chiusura (alla fine del documento)
export async function noteFinali(context) {
  if (!supporta("WordApi", "1.5")) return [];
  const en = context.document.body.endnotes;
  en.load("items/body/text");
  await context.sync();
  return en.items.map((e, i) => ({ nota: e, corpo: e.body, testo: e.body.text || "", dove: "Nota finale " + (i + 1) }));
}

// La selezione usata dagli elenchi "solo testo selezionato": si ricorda, perché "Vai" la cambia.
// nuova: si prende la selezione di adesso; altrimenti quella ricordata dall'ultimo controllo.
let selSalvata = null;

export async function conSelezione(nuova, lavoro) {
  if (nuova || !selSalvata) {
    if (selSalvata) {
      const vecchia = selSalvata;
      selSalvata = null;
      try { await Word.run(vecchia, async c => { vecchia.untrack(); await c.sync(); }); } catch (e) { /* già liberata */ }
    }
    return Word.run(async context => {
      const sel = context.document.getSelection();
      sel.track();
      await context.sync();
      selSalvata = sel;
      return lavoro(context, sel);
    });
  }
  const sel = selSalvata;
  return Word.run(sel, async context => lavoro(context, sel));
}

// Word.run sulla selezione (ricordata o nuova) oppure sull'intero documento
export function esegui(suSelezione, nuova, lavoro) {
  return suSelezione ? conSelezione(nuova, lavoro) : Word.run(async context => lavoro(context, null));
}

// Per la ricerca di Word: "^" è un carattere speciale, lo spazio unificatore si scrive ^s
export function perRicerca(s) {
  return s.replace(/\^/g, "^^").replace(/ /g, "^s");
}

// Quante volte s compare (senza sovrapposizioni, come la ricerca di Word) prima della posizione
export function occorrenzePrima(testo, s, posizione) {
  let n = 0, i = testo.indexOf(s);
  while (i >= 0 && i < posizione) { n++; i = testo.indexOf(s, i + s.length); }
  return n;
}

// Prepara la ricerca del tratto [idx, idx+len) del testo dentro contenitore (corpo della nota, paragrafo…).
// Dopo context.sync(), risolvi() restituisce il Range (o null se non trovato).
// I tratti lunghi (la ricerca di Word arriva a 255 caratteri) si trovano dall'inizio e dalla fine.
export function preparaTratto(contenitore, testo, idx, len) {
  const MAX = 200, PEZZO = 100;
  const opzioni = { matchCase: true };

  if (len <= MAX) {
    const seg = testo.substr(idx, len);
    const occ = occorrenzePrima(testo, seg, idx);
    const rc = contenitore.search(perRicerca(seg), opzioni);
    rc.load("items/text");
    return { risolvi: () => rc.items[occ] || null };
  }

  const a = testo.substr(idx, PEZZO), ib = idx + len - PEZZO, b = testo.substr(ib, PEZZO);
  const occA = occorrenzePrima(testo, a, idx), occB = occorrenzePrima(testo, b, ib);
  const rcA = contenitore.search(perRicerca(a), opzioni);
  const rcB = contenitore.search(perRicerca(b), opzioni);
  rcA.load("items/text");
  rcB.load("items/text");
  return {
    risolvi: () => {
      const ra = rcA.items[occA], rb = rcB.items[occB];
      return ra && rb ? ra.expandTo(rb) : null;
    }
  };
}

// Esegue lavoro(elementi di un blocco) a blocchi, per non mandare a Word migliaia di richieste in una volta
export async function aBlocchi(elementi, dimensione, lavoro, avanzamento) {
  for (let i = 0; i < elementi.length; i += dimensione) {
    if (avanzamento) avanzamento(i, elementi.length);
    await lavoro(elementi.slice(i, i + dimensione));
  }
}

// ===== Maiuscoletto nelle versioni di Word che non lo permettono agli add-in (es. Word LTSC 2024) =====
// Si usa uno stile carattere "Maiuscoletto" (aggiunto al documento se manca) e, per sapere quali parti sono
// già in maiuscoletto, si legge l'HTML che Word dà del testo.

export const STILE_MAIUSCOLETTO = "Maiuscoletto";

const PACCHETTO_STILE =
  '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
  '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships></pkg:xmlData></pkg:part>' +
  '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '</Relationships></pkg:xmlData></pkg:part>' +
  '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' +
  '<w:r><w:rPr><w:rStyle w:val="Maiuscoletto"/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>' +
  '</pkg:xmlData></pkg:part>' +
  '<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="character" w:customStyle="1" w:styleId="Maiuscoletto"><w:name w:val="Maiuscoletto"/>' +
  '<w:uiPriority w:val="1"/><w:qFormat/><w:rPr><w:smallCaps/></w:rPr></w:style>' +
  '</w:styles></pkg:xmlData></pkg:part></pkg:package>';

// Lo stile carattere "Maiuscoletto" esiste? Se no lo aggiunge (inserendo e togliendo un carattere in fondo al testo)
export async function assicuraStileMaiuscoletto(context) {
  const s = context.document.getStyles().getByNameOrNullObject(STILE_MAIUSCOLETTO);
  s.load("nameLocal");
  await context.sync();
  if (!s.isNullObject) return;

  const fine = context.document.body.paragraphs.getLast().getRange("End");
  const inserito = fine.insertOoxml(PACCHETTO_STILE, "Replace");
  await context.sync();
  inserito.delete();
  await context.sync();
}

// Maiuscoletto su un tratto: proprietà del carattere se la versione di Word lo permette, altrimenti lo stile
export function applicaMaiuscoletto(range, diretto) {
  if (diretto) range.font.smallCaps = true;
  else range.style = STILE_MAIUSCOLETTO;
}

// Per ogni carattere di testo: è in maiuscoletto? (dall'HTML di Word; null se l'HTML non si allinea al testo)
let cornice = null;

export function maiuscolettoDaHtml(html, testo) {
  if (!cornice) {
    cornice = document.createElement("iframe");
    cornice.setAttribute("aria-hidden", "true");
    cornice.style.cssText = "position:absolute;width:1px;height:1px;left:-9999px;border:0";
    document.body.append(cornice);
  }

  const doc = cornice.contentDocument;
  doc.open();
  doc.write(html);
  doc.close();

  // testo dell'HTML con il maiuscoletto di ogni carattere
  let t = "";
  const flag = [];
  const w = doc.createTreeWalker(doc.body || doc.documentElement, NodeFilter.SHOW_TEXT);
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    const el = n.parentElement;
    if (!el || /^(STYLE|SCRIPT|TITLE)$/.test(el.tagName)) continue;
    const cs = cornice.contentWindow.getComputedStyle(el);
    if (cs.display === "none") continue;
    const sc = /small-caps/.test(cs.fontVariant || "") || /small-caps/.test(cs.fontVariantCaps || "");
    for (const c of n.data) { t += c; flag.push(sc); }
  }

  return allinea(testo, t, flag);
}

// Allinea il testo della nota con quello dell'HTML (che può avere spazi in più, a capo, il numero della nota…)
function allinea(testo, html, flag) {
  const norm = c => (c === " " || c === "\r" || c === "\n" || c === "\t" ? " " : c);
  const res = new Array(testo.length).fill(false);
  let i = 0, j = 0, presi = 0;

  while (i < testo.length && j < html.length) {
    const a = norm(testo[i]), b = norm(html[j]);
    if (a === b) { res[i] = flag[j]; i++; j++; presi++; continue; }
    if (b === " ") { j++; continue; }
    if (a === " ") { i++; continue; }
    // carattere in più nell'HTML (numero di nota, simboli): si cerca il prossimo carattere uguale vicino
    const k = html.indexOf(testo[i], j);
    if (k >= 0 && k - j <= 6) { j = k; continue; }
    i++;
  }

  const significativi = testo.replace(/\s/g, "").length;
  return significativi === 0 || presi / significativi > 0.9 ? res : null;
}
