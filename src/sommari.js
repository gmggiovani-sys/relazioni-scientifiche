/* global Word, Office */
// Sommari di inizio capitolo, subito dopo il titolo (lo stile scelto, di solito "Titolo 1"),
// con i titoli dei paragrafi del capitolo, la loro numerazione e i corsivi.
// Un sommario già presente si aggiorna.

import {
  componi, stato, stilePredefinito, livello, eSommario, trattiCorsivi, pulisci, perRicerca, ETICHETTA
} from "./sommari-testo.js";
import { assicuraStileMaiuscoletto, applicaMaiuscoletto } from "./word.js";

export const STILE_SOMMARIO = "Sommario Capitolo";

const maiuscolettoDisponibile = () =>
  Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");

// Capitoli del documento, con i paragrafi di Word da usare nello stesso Word.run
// profondita: livelli di titoli nel sommario (1 = solo Titolo 2; 0 = tutti)
async function capitoli(context, stileTitolo, stileSommario, profondita = 1) {
  stileTitolo = (stileTitolo || "").trim();
  stileSommario = (stileSommario || "").trim() || STILE_SOMMARIO;

  // casella vuota: lo stile del paragrafo dove si trova il cursore
  if (!stileTitolo) {
    const p = context.document.getSelection().paragraphs.getFirst();
    p.load("style");
    await context.sync();
    stileTitolo = p.style;
  }
  const predefinito = stilePredefinito(stileTitolo);
  const eTitolo = p => p.style === stileTitolo || (predefinito !== null && p.styleBuiltIn === predefinito);

  const sezioni = context.document.sections;
  sezioni.load("items");
  await context.sync();

  // il livello di struttura c'è nelle versioni recenti; altrimenti basta lo stile predefinito (Titolo 2, 3…)
  let elenchi;
  try {
    elenchi = await paragrafi(context, sezioni, "items/style,items/styleBuiltIn,items/text,items/outlineLevel,items/tableNestingLevel");
  } catch (e) {
    elenchi = await paragrafi(context, sezioni, "items/style,items/styleBuiltIn,items/text,items/tableNestingLevel");
  }

  const res = [];
  let livelloTitolo = null;

  elenchi.forEach((ps, iSez) => {
    let cap = null, inTitolo = false, dopoTitolo = false;

    for (const p of ps.items) {
      if (/^Toc/i.test(p.styleBuiltIn || "")) continue;             // indice generale

      if (eTitolo(p)) {
        const testo = pulisci(p.text);
        if (!testo) continue;
        if (livelloTitolo === null) {
          livelloTitolo = livello(p.outlineLevel, p.styleBuiltIn);
          if (livelloTitolo > 9) livelloTitolo = 1;
        }
        if (cap && inTitolo) {                                          // "Capitolo I" + titolo
          cap.titolo += " – " + testo;
          cap.ancora = p;
        } else {
          cap = { sezione: iSez + 1, titolo: testo, ancora: p, esistente: null, testoEsistente: "", paragrafi: [] };
          res.push(cap);
        }
        inTitolo = true;
        dopoTitolo = true;
        continue;
      }

      inTitolo = false;
      if (!cap) continue;

      if (dopoTitolo) {
        dopoTitolo = false;
        if (eSommario(p.style, p.text, stileSommario)) {
          cap.esistente = p;
          cap.testoEsistente = pulisci(p.text);
          continue;
        }
      }

      const lv = livello(p.outlineLevel, p.styleBuiltIn);
      const ultimo = profondita <= 0 ? 9 : Math.min(9, livelloTitolo + profondita);
      if (lv <= livelloTitolo || lv > ultimo || p.tableNestingLevel > 0 || !pulisci(p.text)) continue;
      cap.paragrafi.push(p);
    }
  });

  // numerazione e corsivi dei titoli dei paragrafi
  const letture = [];
  for (const cap of res)
    for (const p of cap.paragrafi) {
      const li = p.listItemOrNullObject;
      li.load("listString");
      const parole = p.getTextRanges([" "], false);
      parole.load("items/text,items/font/italic");
      letture.push({ cap, p, li, parole });
    }
  await context.sync();

  for (const cap of res) cap.voci = [];
  for (const { cap, p, li, parole } of letture)
    cap.voci.push({
      numero: li.isNullObject ? "" : li.listString,
      titolo: p.text,
      corsivi: trattiCorsivi(p.text, parole.items.map(w => ({ text: w.text, italic: w.font.italic })))
    });

  for (const cap of res) {
    const c = componi(cap.voci);
    cap.testo = c.testo;
    cap.corsivi = c.corsivi;
    cap.voci = cap.voci.length;
    cap.stato = stato(cap);
    cap.chiave = cap.sezione + "|" + cap.titolo;
  }

  return { capitoli: res, stileTitolo, stileSommario };
}

async function paragrafi(context, sezioni, proprieta) {
  const elenchi = sezioni.items.map(s => {
    const ps = s.body.paragraphs;
    ps.load(proprieta);
    return ps;
  });
  await context.sync();
  return elenchi;
}

// Solo lettura: elenco per il pannello
export async function analizza(stileTitolo, stileSommario, profondita = 1) {
  return Word.run(async context => {
    const r = await capitoli(context, stileTitolo, stileSommario, profondita);
    return {
      stileTitolo: r.stileTitolo,
      stileSommario: r.stileSommario,
      capitoli: r.capitoli.map(c => ({
        chiave: c.chiave, sezione: c.sezione, titolo: c.titolo, voci: c.voci,
        stato: c.stato, testo: c.testo, testoEsistente: c.testoEsistente
      }))
    };
  });
}

export async function vaiA(chiave, stileTitolo, stileSommario, profondita = 1) {
  return Word.run(async context => {
    const r = await capitoli(context, stileTitolo, stileSommario, profondita);
    const cap = r.capitoli.find(c => c.chiave === chiave);
    if (!cap) return false;
    (cap.esistente || cap.ancora).select();
    await context.sync();
    return true;
  });
}

// Crea o aggiorna i sommari dei capitoli scelti (chiavi). Restituisce l'esito da mostrare.
export async function scrivi(chiavi, stileTitolo, stileSommario, profondita = 1) {
  return Word.run(async context => {
    const r = await capitoli(context, stileTitolo, stileSommario, profondita);
    const scelti = r.capitoli.filter(c => chiavi.includes(c.chiave) && c.voci > 0);
    if (scelti.length === 0) return "Nessun capitolo da scrivere.";

    const avvisi = [];
    const stile = await stileSommarioPronto(context, r.stileSommario, avvisi);
    const maiuscoletto = maiuscolettoDisponibile();
    if (!maiuscoletto) await assicuraStileMaiuscoletto(context);
    let creati = 0, aggiornati = 0;

    const scritti = scelti.map(c => {
      let par;
      if (c.esistente) {
        c.esistente.insertText(c.testo, "Replace");
        par = c.esistente;
        aggiornati++;
      } else {
        par = c.ancora.insertParagraph(c.testo, "After");
        creati++;
      }
      if (stile) par.style = stile;

      const tutto = par.getRange("Content");
      tutto.font.italic = false;
      tutto.font.bold = false;
      if (maiuscoletto) tutto.font.smallCaps = false;

      par.load("isListItem");
      const etichetta = par.search(ETICHETTA, { matchCase: true });
      etichetta.load("items");
      const corsivi = c.corsivi.map(k => {
        const trovati = par.search(perRicerca(k.testo), { matchCase: true });
        trovati.load("items");
        return { k, trovati };
      });
      return { par, etichetta, corsivi };
    });
    await context.sync();

    for (const { par, etichetta, corsivi } of scritti) {
      if (par.isListItem) par.detachFromList();                      // ereditato dal titolo numerato
      if (etichetta.items.length > 0) applicaMaiuscoletto(etichetta.items[0], maiuscoletto);
      for (const { k, trovati } of corsivi)
        if (trovati.items.length > k.occorrenza) trovati.items[k.occorrenza].font.italic = true;
    }
    await context.sync();


    return "Sommari creati: " + creati + ", aggiornati: " + aggiornati + ". Stile \"" + r.stileSommario + "\"." +
           (avvisi.length ? "\n\n" + avvisi.join("\n") : "");
  });
}

// Lo stile del sommario; se manca si crea come quello delle istruzioni:
// corpo 10, interlinea singola, 12 pt prima e dopo, giustificato, senza rientro
async function stileSommarioPronto(context, nome, avvisi) {
  const s = context.document.getStyles().getByNameOrNullObject(nome);
  s.load("nameLocal");
  await context.sync();
  if (!s.isNullObject) return nome;

  if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) {
    avvisi.push("Lo stile \"" + nome + "\" non esiste e questa versione di Word non permette di crearlo: " +
                "i sommari restano con lo stile del paragrafo. Crea lo stile e riprova.");
    return null;
  }

  const n = context.document.addStyle(nome, Word.StyleType.paragraph);
  n.font.size = 10;
  n.font.italic = false;
  n.font.bold = false;
  n.paragraphFormat.spaceBefore = 12;
  n.paragraphFormat.spaceAfter = 12;
  n.paragraphFormat.lineSpacing = 12;
  n.paragraphFormat.firstLineIndent = 0;
  n.paragraphFormat.leftIndent = 0;
  n.paragraphFormat.alignment = Word.Alignment.justified;
  await context.sync();

  // basato su Normale (il nome cambia con la lingua di Word: si prova, senza insistere)
  try {
    n.baseStyle = "Normal";
    n.nextParagraphStyle = "Normal";
    await context.sync();
  } catch (e) { /* resta basato sullo stile predefinito */ }

  return nome;
}
