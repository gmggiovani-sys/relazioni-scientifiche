/* global Word */
// Impaginazione secondo le istruzioni della cattedra (II.A):
// A4; margini 4 (superiore), 3,5 (inferiore), 3,5 (sinistro), 3 cm (destro); interlinea 1,5;
// rientro del paragrafo 1 cm; Times New Roman o Garamond 12 (o 14);
// note a piè di pagina con interlinea singola, corpo 10 (o 11 col 14).
// Elenca le differenze, raggruppate, e corregge quelle scelte. Come ImpaginazioneService.cs (versione Windows).

import { supporta, note } from "./word.js";
import { livello } from "./sommari-testo.js";

const CM = 28.3464567;                       // punti in un centimetro
const TOLLERANZA = 0.05;                     // cm
const CARATTERI = ["Times New Roman", "Garamond"];

const cm = pt => pt / CM;
const vicino = (a, b) => Math.abs(a - b) <= TOLLERANZA;
const fmtCm = v => (Math.round(v * 100) / 100).toLocaleString("it-IT");
const fmtPt = v => (Math.round(v * 10) / 10).toLocaleString("it-IT") + " pt";

// Interlinea in punti come la dà Word: 12 = singola, 18 = 1,5 righe, 24 = doppia
const interlinea15 = v => Math.abs(v - 18) < 0.5;
const interlineaSingola = v => Math.abs(v - 12) < 0.5;
const nomeInterlinea = v => (v === null || v === undefined ? "interlinea diversa nello stesso paragrafo"
  : interlineaSingola(v) ? "singola" : interlinea15(v) ? "1,5 righe" : Math.abs(v - 24) < 0.5 ? "doppia"
  : (Math.round(v / 12 * 100) / 100).toLocaleString("it-IT") + " righe (" + fmtPt(v) + ")");

// Stili che non sono corpo del testo
function escluso(stile) {
  const s = (stile || "").toLowerCase();
  return /^(sommario |toc |indice|didascalia|caption|titolo|heading|title|sottotitolo|subtitle|intestazione|header|piè di pagina|footer|testo nota|footnote|bibliografia|giurisprudenza|citazione|quote)/.test(s);
}


// Raggruppa per valore: [{ valore, elementi }]
function gruppi(elenco, valore) {
  const m = new Map();
  for (const x of elenco) {
    const k = valore(x);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(x);
  }
  return [...m].map(([k, v]) => ({ valore: k, elementi: v }));
}

// ===============================
// CONTROLLO
// ===============================

// Problemi trovati, con la correzione legata agli oggetti di Word di questo contesto:
// [{ firma, dove, cosa, attuale, atteso, dubbio, quanti, correggi(x) }]
async function analisi(context, stileBiblio, stileGiur, stileSommario, avanzamento) {
  {
    const problemi = [];
    const avvisi = [];

    // ----- pagina (margini e formato: solo nelle versioni di Word che li danno agli add-in)
    if (supporta("WordApiDesktop", "1.3")) {
      const sezioni = context.document.sections;
      sezioni.load("items");
      await context.sync();
      const ps = sezioni.items.map(s => { const p = s.pageSetup; p.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight,gutter"); return p; });
      await context.sync();

      const formato = new Map(), margini = new Map();
      ps.forEach((p, i) => {
        const w = cm(p.pageWidth), h = cm(p.pageHeight), oriz = w > h;
        if (!vicino(Math.min(w, h), 21) || !vicino(Math.max(w, h), 29.7) || oriz) {
          const ora = fmtCm(w) + " × " + fmtCm(h) + " cm" + (oriz ? " (orizzontale)" : "");
          if (!formato.has(ora))
            formato.set(ora, { dove: "Pagina", cosa: "Formato del foglio", attuale: ora, atteso: "A4, 21 × 29,7 cm", dubbio: oriz, sezioni: [],
                               correggi: x => { for (const k of x.sezioni) { ps[k].pageWidth = (oriz ? 29.7 : 21) * CM; ps[k].pageHeight = (oriz ? 21 : 29.7) * CM; } } });
          formato.get(ora).sezioni.push(i);
        }

        const su = cm(p.topMargin), giu = cm(p.bottomMargin), sx = cm(p.leftMargin), dx = cm(p.rightMargin), ril = cm(p.gutter || 0);
        if (!vicino(su, 4) || !vicino(giu, 3.5) || !vicino(sx, 3.5) || !vicino(dx, 3) || ril > TOLLERANZA) {
          const ora = "sup. " + fmtCm(su) + ", inf. " + fmtCm(giu) + ", sin. " + fmtCm(sx) + ", des. " + fmtCm(dx) +
                      (ril > TOLLERANZA ? ", rilegatura " + fmtCm(ril) : "");
          if (!margini.has(ora))
            margini.set(ora, { dove: "Pagina", cosa: "Margini", attuale: ora, atteso: "sup. 4, inf. 3,5, sin. 3,5, des. 3 cm", sezioni: [],
                               correggi: x => { for (const k of x.sezioni) Object.assign(ps[k], { topMargin: 4 * CM, bottomMargin: 3.5 * CM, leftMargin: 3.5 * CM, rightMargin: 3 * CM, gutter: 0 }); } });
          margini.get(ora).sezioni.push(i);
        }
      });
      problemi.push(...formato.values(), ...margini.values());
    } else
      avvisi.push("Margini e formato del foglio: questa versione di Word non li fa leggere agli add-in. " +
                  "Controllali a mano in Layout → Margini → Margini personalizzati (sup. 4, inf. 3,5, sin. 3,5, des. 3 cm; A4).");

    // ----- testo
    const ps = context.document.body.paragraphs;
    try {
      ps.load("items/text,items/style,items/styleBuiltIn,items/alignment,items/lineSpacing,items/firstLineIndent,items/leftIndent,items/tableNestingLevel,items/outlineLevel");
      await context.sync();
    } catch (e) {
      ps.load("items/text,items/style,items/styleBuiltIn,items/alignment,items/lineSpacing,items/firstLineIndent,items/leftIndent,items/tableNestingLevel");
      await context.sync();
    }

    const esclusi = new Set([stileBiblio, stileGiur, stileSommario].map(s => (s || "").toLowerCase()));
    const corpo = ps.items.filter(p =>
      /\p{L}/u.test(p.text) && p.tableNestingLevel === 0 && !/^Toc/i.test(p.styleBuiltIn || "") &&
      livello(p.outlineLevel, p.styleBuiltIn) > 9 && !esclusi.has((p.style || "").toLowerCase()) && !escluso(p.style) &&
      p.alignment !== "Centered" && p.alignment !== "Right");

    // carattere e corpo senza il segno di paragrafo
    const contenuti = corpo.map(p => { const r = p.getRange("Content"); r.load("font/name,font/size"); return r; });
    if (avanzamento) avanzamento(0, corpo.length);
    await context.sync();

    const testo = corpo.map((p, i) => ({ p, nome: contenuti[i].font.name || "", size: contenuti[i].font.size, ls: p.lineSpacing,
                                         fli: cm(p.firstLineIndent || 0), li: cm(p.leftIndent || 0) }));

    const conta = (v) => testo.filter(x => x.size === v).length;
    const corpoTesto = conta(14) > conta(12) ? 14 : 12;
    const corpoNote = corpoTesto === 14 ? 11 : 10;
    const ammessi = gruppi(testo.filter(x => CARATTERI.includes(x.nome)), x => x.nome).sort((a, b) => b.elementi.length - a.elementi.length);
    const carattere = ammessi.length ? ammessi[0].valore : "Times New Roman";

    for (const g of gruppi(testo.filter(x => !CARATTERI.includes(x.nome)), x => x.nome || "caratteri diversi nello stesso paragrafo"))
      problemi.push({ dove: "Testo", cosa: "Carattere", attuale: g.valore, atteso: "Times New Roman o Garamond (" + carattere + ")",
                      dubbio: g.valore.startsWith("caratteri diversi"), par: g.elementi.map(x => x.p),
                      correggi: x => { for (const p of x.par) p.font.name = carattere; } });

    for (const g of gruppi(testo.filter(x => x.size !== corpoTesto), x => (typeof x.size === "number" ? fmtPt(x.size) : "corpi diversi nello stesso paragrafo")))
      problemi.push({ dove: "Testo", cosa: "Corpo del carattere", attuale: g.valore, atteso: fmtPt(corpoTesto) + " (12 o 14)",
                      dubbio: g.valore.startsWith("corpi diversi"), par: g.elementi.map(x => x.p),
                      correggi: x => { for (const p of x.par) p.font.size = corpoTesto; } });

    for (const g of gruppi(testo.filter(x => !interlinea15(x.ls)), x => nomeInterlinea(x.ls)))
      problemi.push({ dove: "Testo", cosa: "Interlinea", attuale: g.valore, atteso: "1,5 righe", par: g.elementi.map(x => x.p),
                      correggi: x => { for (const p of x.par) p.lineSpacing = 18; } });

    // rientro della prima riga 1 cm (i paragrafi rientrati tutti, come le citazioni lunghe, restano da decidere)
    for (const g of gruppi(testo.filter(x => !vicino(x.fli, 1)), x => x.li > TOLLERANZA)) {
      const valori = gruppi(g.elementi, x => fmtCm(x.fli)).sort((a, b) => b.elementi.length - a.elementi.length);
      problemi.push({ dove: "Testo", cosa: g.valore ? "Rientro prima riga (paragrafi con rientro a sinistra)" : "Rientro prima riga",
                      attuale: valori.slice(0, 3).map(v => v.valore + " cm").join(", ") + (valori.length > 3 ? "…" : ""),
                      atteso: "1 cm", dubbio: g.valore, par: g.elementi.map(x => x.p),
                      correggi: x => { for (const p of x.par) p.firstLineIndent = CM; } });
    }

    // ----- note
    let elencoNote = [];
    try { elencoNote = await note(context, null); } catch (e) { avvisi.push("Note a piè di pagina: " + e.message); }
    if (elencoNote.length) {
      const parNote = elencoNote.map(n => { const q = n.corpo.paragraphs; q.load("items/text,items/lineSpacing"); return q; });
      await context.sync();
      const tutti = parNote.flatMap(q => q.items).filter(p => /[\p{L}\p{N}]/u.test(p.text));
      const cont = tutti.map(p => { const r = p.getRange("Content"); r.load("font/name,font/size"); return r; });
      await context.sync();
      const nn = tutti.map((p, i) => ({ p, nome: cont[i].font.name || "", size: cont[i].font.size, ls: p.lineSpacing }));

      for (const g of gruppi(nn.filter(x => !CARATTERI.includes(x.nome)), x => x.nome || "caratteri diversi nella stessa nota"))
        problemi.push({ dove: "Note", cosa: "Carattere", attuale: g.valore, atteso: "come il testo (" + carattere + ")",
                        dubbio: g.valore.startsWith("caratteri diversi"), par: g.elementi.map(x => x.p),
                        correggi: x => { for (const p of x.par) p.font.name = carattere; } });

      for (const g of gruppi(nn.filter(x => x.size !== corpoNote), x => (typeof x.size === "number" ? fmtPt(x.size) : "corpi diversi nella stessa nota")))
        problemi.push({ dove: "Note", cosa: "Corpo del carattere", attuale: g.valore, atteso: fmtPt(corpoNote),
                        dubbio: g.valore.startsWith("corpi diversi"), par: g.elementi.map(x => x.p),
                        correggi: x => { for (const p of x.par) p.font.size = corpoNote; } });

      for (const g of gruppi(nn.filter(x => !interlineaSingola(x.ls)), x => nomeInterlinea(x.ls)))
        problemi.push({ dove: "Note", cosa: "Interlinea", attuale: g.valore, atteso: "singola", par: g.elementi.map(x => x.p),
                        correggi: x => { for (const p of x.par) p.lineSpacing = 12; } });

      avvisi.push("Numerazione delle note (progressiva): questa versione dell'add-in non la legge; " +
                  "controllala in Riferimenti → Note a piè di pagina (freccia in basso a destra) → Numerazione: Continua.");
    }

    for (const x of problemi) {
      x.firma = x.dove + "|" + x.cosa + "|" + x.attuale;
      x.quanti = x.sezioni ? x.sezioni.length + (x.sezioni.length === 1 ? " sezione" : " sezioni")
                           : x.par.length + (x.par.length === 1 ? " paragrafo" : " paragrafi");
    }

    return {
      problemi, avvisi,
      riepilogo: "testo " + carattere + " " + fmtPt(corpoTesto) + ", note " + fmtPt(corpoNote) + "; " + corpo.length + " paragrafi di testo esaminati"
    };
  }
}

// Solo verifica: [{ chiave (firma), dove, cosa, attuale, atteso, dubbio, quanti }]
export async function controlla(stileBiblio, stileGiur, stileSommario, avanzamento) {
  return Word.run(async context => {
    const r = await analisi(context, stileBiblio, stileGiur, stileSommario, avanzamento);
    return {
      riepilogo: r.riepilogo, avvisi: r.avvisi,
      problemi: r.problemi.map(x => ({ chiave: x.firma, dove: x.dove, cosa: x.cosa, attuale: x.attuale, atteso: x.atteso, dubbio: !!x.dubbio, quanti: x.quanti }))
    };
  });
}

// ===============================
// CORREZIONE
// ===============================

// Rilegge il documento e corregge le voci scelte (chiavi: firme dell'ultimo controllo)
export async function correggi(stileBiblio, stileGiur, stileSommario, chiavi, avanzamento) {
  return Word.run(async context => {
    const r = await analisi(context, stileBiblio, stileGiur, stileSommario, avanzamento);
    const scelte = new Set(chiavi);
    const daFare = r.problemi.filter(x => scelte.has(x.firma));
    for (const x of daFare) x.correggi(x);
    await context.sync();

    const nonTrovate = chiavi.length - daFare.length;
    let esito = "Impaginazione: corrette " + daFare.length + " voci" +
                (daFare.length ? ":\n" + daFare.map(x => "• " + x.dove + ", " + x.cosa.toLowerCase() + ": " + x.attuale + " → " + x.atteso + " (" + x.quanti + ")").join("\n") : ".");
    if (nonTrovate > 0) esito += "\n\nVoci già cambiate dopo il controllo: " + nonTrovate + ".";
    esito += "\n\nLe correzioni valgono per i paragrafi segnalati; per i paragrafi nuovi conviene impostare " +
             "gli stessi valori anche negli stili (Normale e Testo nota a piè di pagina).";
    return esito;
  });
}
