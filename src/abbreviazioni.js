/* global Word */
// Abbreviazioni secondo la tabella della cattedra, nel testo e nelle note.
// Non tocca le parti in corsivo, i titoli di opere citate e le citazioni tra «».
// Ambito: intero documento, testo selezionato (con le note richiamate), paragrafi con lo stile di quello selezionato.

import * as A from "./parser/abbreviazioni.js";
import { note, noteFinali, preparaTratto, aBlocchi, esegui, supporta } from "./word.js";

const BLOCCO = 250;

export const AMBITI = { documento: "intero documento", selezione: "testo selezionato", stile: "stile selezionato" };

// Zone di testo da esaminare: [{ contenitore (paragrafo, tratto o corpo della nota), testo, dove }]
async function zone(context, ambito, sel, conOrdine = false) {
  const res = [];

  if (ambito === "documento") {
    const ps = context.document.body.paragraphs;
    ps.load("items/text,items/styleBuiltIn");
    await context.sync();
    const zonaPar = new Map();
    ps.items.forEach((p, i) => {
      // il sommario automatico (indice) si rigenera dai titoli: le correzioni vanno fatte lì
      if (/^Toc/i.test(p.styleBuiltIn || "") || !p.text.trim()) return;
      zonaPar.set(i, res.length);
      res.push({ contenitore: p, testo: p.text, dove: "Testo, paragrafo " + (i + 1) });
    });

    const elencoNote = await note(context, null);
    const zonaNota = elencoNote.map(n => {
      if (!n.testo.trim()) return -1;
      res.push({ contenitore: n.corpo, testo: n.testo, dove: n.dove });
      return res.length - 1;
    });
    for (const n of await noteFinali(context))
      if (n.testo.trim()) res.push({ contenitore: n.corpo, testo: n.testo, dove: n.dove });

    return { zone: res, descrizione: AMBITI.documento, ordine: conOrdine ? await ordineLettura(context, ps, zonaPar, zonaNota, res.length) : null };
  }

  if (ambito === "selezione") {
    sel.load("text");
    const ps = sel.paragraphs;
    ps.load("items");
    await context.sync();
    if (!sel.text.trim())
      throw new Error("Seleziona il testo da controllare, oppure scegli \"Intero documento\".");

    // paragrafo per paragrafo, limitato alla selezione
    const tratti = ps.items.map(p => p.getRange("Whole").intersectWithOrNullObject(sel));
    tratti.forEach(t => t.load("text"));
    await context.sync();
    tratti.forEach((t, i) => {
      if (!t.isNullObject && t.text.trim()) res.push({ contenitore: t, testo: t.text, dove: "Selezione, paragrafo " + (i + 1) });
    });

    // le note richiamate nel testo selezionato
    try {
      for (const n of await note(context, sel)) if (n.testo.trim()) res.push({ contenitore: n.corpo, testo: n.testo, dove: n.dove });
    } catch (e) { /* selezione dentro una nota: niente richiami */ }
    return { zone: res, descrizione: AMBITI.selezione };
  }

  // stile del paragrafo dove si trova il cursore, nel testo e nelle note
  const primo = sel.paragraphs.getFirst();
  primo.load("style");
  await context.sync();
  const stile = primo.style;

  const ps = context.document.body.paragraphs;
  ps.load("items/text,items/style");
  await context.sync();
  ps.items.forEach((p, i) => {
    if (p.style === stile && p.text.trim()) res.push({ contenitore: p, testo: p.text, dove: "Testo, paragrafo " + (i + 1) });
  });

  const elencoNote = await note(context, null);
  const parNote = elencoNote.map(n => { const q = n.corpo.paragraphs; q.load("items/text,items/style"); return q; });
  await context.sync();
  parNote.forEach((q, k) => q.items.forEach(p => {
    if (p.style === stile && p.text.trim()) res.push({ contenitore: p, testo: p.text, dove: elencoNote[k].dove });
  }));

  if (res.length === 0) throw new Error("Nessun paragrafo con lo stile \"" + stile + "\".");
  return { zone: res, descrizione: "stile \"" + stile + "\"" };
}

// Ordine di lettura delle zone: ogni nota subito dopo il paragrafo che la richiama
// (per sapere quale citazione di una legge viene prima). Senza le note per paragrafo: testo, poi note.
async function ordineLettura(context, ps, zonaPar, zonaNota, quante) {
  const inOrdine = [...Array(quante).keys()];
  if (!supporta("WordApi", "1.5") || zonaNota.length === 0) return inOrdine;

  try {
    const perPar = ps.items.map(p => { const f = p.getRange("Whole").footnotes; f.load("items/type"); return f; });
    await context.sync();

    const ordine = [];
    let n = 0;
    ps.items.forEach((p, i) => {
      if (zonaPar.has(i)) ordine.push(zonaPar.get(i));
      for (let k = 0; k < perPar[i].items.length && n < zonaNota.length; k++, n++)
        if (zonaNota[n] >= 0) ordine.push(zonaNota[n]);
    });
    for (; n < zonaNota.length; n++) if (zonaNota[n] >= 0) ordine.push(zonaNota[n]);
    for (let z = 0; z < quante; z++) if (!ordine.includes(z)) ordine.push(z);   // note finali
    return ordine;
  } catch (e) {
    return inOrdine;
  }
}

// Tutte le forme trovate, con il Range di Word e il corsivo
async function analisi(context, ambito, sel, avanzamento) {
  const { zone: zz, descrizione, ordine } = await zone(context, ambito, sel, true);
  const trovate = [];

  // leggi: prima citazione per esteso, poi abbreviata (solo sull'intero documento)
  const leggi = zz.map(() => []);
  if (ordine) {
    const perOrdine = A.citazioniLeggi(ordine.map(k => zz[k].testo));
    ordine.forEach((k, j) => (leggi[k] = perOrdine[j]));
  }

  zz.forEach((z, iz) => {
    for (const o of A.cerca(z.testo).concat(leggi[iz]).sort((a, b) => a.idx - b.idx))
      trovate.push({
        chiave: iz + "|" + o.idx + "|" + o.len, iz, o, dove: z.dove,
        contesto: A.contesto(z.testo, o), range: null, corsivo: false
      });
  });

  await aBlocchi(trovate, BLOCCO, async blocco => {
    const ricerche = blocco.map(t => preparaTratto(zz[t.iz].contenitore, zz[t.iz].testo, t.o.idx, t.o.len));
    await context.sync();
    blocco.forEach((t, k) => {
      t.range = ricerche[k].risolvi();
      if (t.range) t.range.load("text,font/italic");
    });
    await context.sync();
    for (const t of blocco) if (t.range) t.corsivo = t.range.font.italic !== false;   // anche in parte
  }, avanzamento);

  return { trovate, descrizione };
}

// Correzione sicura: si applica con "Correggi" e parte spuntata nell'elenco
const applicabile = t => t.o.nuovo !== null && !t.o.incerta && !t.corsivo && t.range !== null;
// C'è una correzione proposta e il punto è individuato: la casella si può spuntare
const spuntabile = t => t.o.nuovo !== null && t.range !== null;

function perRegola(elenco) {
  const conta = new Map();
  for (const t of elenco) conta.set(t.o.regola, (conta.get(t.o.regola) || 0) + 1);
  return [...conta].sort((a, b) => b[1] - a[1]).map(([r, n]) => "• " + r + ": " + n).join("\n");
}

// Solo verifica: elenco di tutte le forme trovate
export async function controlla(ambito, avanzamento) {
  return esegui(ambito !== "documento", true, async (context, sel) => {
    const { trovate, descrizione } = await analisi(context, ambito, sel, avanzamento);
    return {
      descrizione,
      daCorreggere: trovate.filter(applicabile).length,
      trovate: trovate.map(t => ({
        chiave: t.chiave, dove: t.dove, trovato: t.o.trovato, nuovo: t.o.nuovo, regola: t.o.regola, contesto: t.contesto,
        stato: t.o.nuovo === null ? "da valutare" : !t.range ? "non individuato" : t.corsivo ? "in corsivo: decidi tu"
             : t.o.incerta ? "decidi tu" : "",
        spuntabile: spuntabile(t), applicabile: applicabile(t)
      }))
    };
  });
}

// Corregge le forme sicure (chiavi null) o quelle scelte nell'elenco
export async function correggi(ambito, chiavi, avanzamento) {
  return esegui(ambito !== "documento", chiavi === null, async (context, sel) => {
    const { trovate, descrizione } = await analisi(context, ambito, sel, avanzamento);
    const scelte = new Set(chiavi || []);
    const daFare = chiavi === null ? trovate.filter(applicabile)
                                   : trovate.filter(t => scelte.has(t.chiave) && spuntabile(t) && t.range.text === t.o.trovato);

    // da destra a sinistra dentro ogni zona
    daFare.sort((a, b) => a.iz - b.iz || b.o.idx - a.o.idx);
    await aBlocchi(daFare, BLOCCO, async blocco => {
      for (const t of blocco) t.range.insertText(t.o.nuovo, "Replace");
      await context.sync();
    });

    let esito = "Ambito: " + descrizione + "\n\n" + (daFare.length === 0
      ? "Nessuna abbreviazione da correggere."
      : "Abbreviazioni corrette: " + daFare.length + "\n" + perRegola(daFare));

    if (chiavi !== null) {
      const saltate = chiavi.length - daFare.length;
      if (saltate > 0) esito += "\n\nSaltate perché il testo è cambiato dopo il controllo: " + saltate;
      return esito;
    }

    const corsivo = trovate.filter(t => t.o.nuovo !== null && t.corsivo).length;
    const nonTrovate = trovate.filter(t => t.o.nuovo !== null && !t.range).length;
    const daValutare = trovate.filter(t => t.o.nuovo === null || t.o.incerta).length;
    if (corsivo) esito += "\n\nNon corrette perché in corsivo (titoli di opere): " + corsivo;
    if (nonTrovate) esito += "\nNon individuate nel documento: " + nonTrovate;
    if (daValutare)
      esito += "\n\nCorrezioni da decidere caso per caso (titoli di opere citate, \"cod. pen.\", nomi…): " + daValutare +
               "\nL'elenco completo è in \"Controlla\".";
    return esito;
  });
}

export async function vaiA(ambito, chiave) {
  return esegui(ambito !== "documento", false, async (context, sel) => {
    const { zone: zz } = await zone(context, ambito, sel);
    const [iz, idx, len] = chiave.split("|").map(Number);
    const z = zz[iz];
    if (!z) return false;
    const r = preparaTratto(z.contenitore, z.testo, idx, len);
    await context.sync();
    const range = r.risolvi();
    if (!range) return false;
    range.select();
    await context.sync();
    return true;
  });
}
