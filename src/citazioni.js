/* global Word */
// Citazioni nelle note secondo le istruzioni della cattedra (solo formattazione, il testo non cambia):
// autore e curatore in MAIUSCOLETTO; titolo, rivista e volume in corsivo; cit. e op. cit. in tondo;
// ivi, ibidem in corsivo; ID. in maiuscoletto. Le opere si riconoscono confrontando le note con la bibliografia.

import * as C from "./parser/citazioni.js";
import {
  vociConStile, note, preparaTratto, aBlocchi, maiuscolettoDisponibile,
  assicuraStileMaiuscoletto, applicaMaiuscoletto, maiuscolettoDaHtml, esegui
} from "./word.js";

const BLOCCO = 250;     // tratti cercati per ogni richiesta a Word

export const NOME_FORMATO = {
  [C.CORSIVO]: "corsivo",
  [C.TONDO]: "tondo (non corsivo)",
  [C.MAIUSCOLETTO]: "MAIUSCOLETTO"
};

// Analisi dentro un Word.run: parti da formattare, con il Range di Word quando trovato
// origine: null (tutte le note) o il Range della selezione
async function analisi(context, stileBiblio, origine, avanzamento) {
  const { testi } = await vociConStile(context, stileBiblio);
  if (testi.length === 0)
    throw new Error("Per riconoscere le opere citate serve la bibliografia con stile \"" + stileBiblio +
                    "\". Controlla il nome dello stile nella casella \"Stile della bibliografia\".");

  const opere = C.opere(testi);
  const elencoNote = await note(context, origine);
  if (elencoNote.length === 0)
    throw new Error(origine ? "Nel testo selezionato non ci sono richiami di nota." : "Il documento non ha note a piè di pagina.");

  const maiusc = maiuscolettoDisponibile();
  const parti = [];
  const daControllare = [];

  elencoNote.forEach((n, iNota) => {
    if (!n.testo.trim()) return;
    const non = [];
    for (const f of C.analizza(n.testo, opere, non))
      parti.push({
        chiave: iNota + "|" + f.idx + "|" + f.len + "|" + f.tipo,
        iNota, dove: n.dove, f, testo: n.testo.substr(f.idx, f.len), range: null, conforme: false
      });
    for (const t of [...new Set(non)]) daControllare.push({ iNota, dove: n.dove, testo: t });
  });

  await aBlocchi(parti, BLOCCO, async blocco => {
    const ricerche = blocco.map(p => preparaTratto(elencoNote[p.iNota].corpo, elencoNote[p.iNota].testo, p.f.idx, p.f.len));
    await context.sync();

    blocco.forEach((p, k) => {
      p.range = ricerche[k].risolvi();
      if (p.range) p.range.load(maiusc ? "font/italic,font/smallCaps" : "font/italic");
    });
    await context.sync();

    for (const p of blocco) {
      if (!p.range) continue;
      const font = p.range.font;
      p.conforme = p.f.tipo === C.CORSIVO ? font.italic === true
                 : p.f.tipo === C.TONDO ? font.italic === false
                 : maiusc && font.smallCaps === true;
    }
  }, avanzamento);

  // senza il maiuscoletto tra le funzioni per gli add-in (es. Word LTSC 2024) si legge dall'HTML delle note
  if (!maiusc) await maiuscolettoDalleNote(context, parti, elencoNote, avanzamento);

  return { parti, daControllare, elencoNote, maiusc };
}

async function maiuscolettoDalleNote(context, parti, elencoNote, avanzamento) {
  const daVedere = [...new Set(parti.filter(p => p.f.tipo === C.MAIUSCOLETTO && p.range).map(p => p.iNota))];

  await aBlocchi(daVedere, 40, async blocco => {
    const html = blocco.map(i => elencoNote[i].corpo.getHtml());
    await context.sync();

    blocco.forEach((iNota, k) => {
      let flag = null;
      try { flag = maiuscolettoDaHtml(html[k].value, elencoNote[iNota].testo); } catch (e) { flag = null; }
      if (!flag) return;
      for (const p of parti)
        if (p.iNota === iNota && p.f.tipo === C.MAIUSCOLETTO && p.range)
          p.conforme = flag.slice(p.f.idx, p.f.idx + p.f.len).every(Boolean);
    });
  }, avanzamento);
}

function applica(p, maiusc) {
  if (p.f.tipo === C.CORSIVO) p.range.font.italic = true;
  else if (p.f.tipo === C.TONDO) p.range.font.italic = false;
  else applicaMaiuscoletto(p.range, maiusc);
}

function perRegola(parti) {
  const conta = new Map();
  for (const p of parti) conta.set(p.f.regola, (conta.get(p.f.regola) || 0) + 1);
  return [...conta].sort((a, b) => b[1] - a[1]).map(([r, n]) => "• " + r + ": " + n).join("\n");
}

// Solo verifica: elenco delle parti non conformi e delle citazioni da controllare a mano
export async function controlla(stileBiblio, soloSelezione, avanzamento) {
  return esegui(soloSelezione, true, async (context, sel) => {
    const r = await analisi(context, stileBiblio, sel, avanzamento);
    const nonConformi = r.parti.filter(p => !p.conforme);
    return {
      maiusc: r.maiusc,
      giaConformi: r.parti.length - nonConformi.length,
      parti: nonConformi.map(p => ({
        chiave: p.chiave, dove: p.dove, formato: NOME_FORMATO[p.f.tipo], tipo: p.f.tipo,
        testo: p.testo, regola: p.f.regola, trovato: p.range !== null
      })),
      daControllare: r.daControllare.map(d => ({ chiave: "nota|" + d.iNota, dove: d.dove, testo: d.testo }))
    };
  });
}

// Formatta le parti non conformi: tutte (chiavi null) o solo quelle scelte nell'elenco
// Con "solo selezione": dall'elenco vale la selezione del controllo, "Formatta tutto" usa quella di adesso
export async function formatta(stileBiblio, soloSelezione, chiavi, avanzamento) {
  return esegui(soloSelezione, chiavi === null, async (context, sel) => {
    const r = await analisi(context, stileBiblio, sel, avanzamento);
    const scelte = new Set(chiavi || []);
    const daFare = r.parti.filter(p => !p.conforme && p.range && (chiavi === null || scelte.has(p.chiave)))
                          // il tondo per ultimo: "cit." dentro un tratto in corsivo resta in tondo
                          .sort((a, b) => (a.f.tipo === C.TONDO) - (b.f.tipo === C.TONDO));

    if (!r.maiusc && daFare.some(p => p.f.tipo === C.MAIUSCOLETTO)) await assicuraStileMaiuscoletto(context);

    await aBlocchi(daFare, BLOCCO, async blocco => {
      for (const p of blocco) applica(p, r.maiusc);
      await context.sync();
    });

    const nonTrovate = r.parti.filter(p => !p.conforme && !p.range && (chiavi === null || scelte.has(p.chiave))).length;
    let esito = "Parti già formattate secondo le regole: " + r.parti.filter(p => p.conforme).length +
                "\nParti formattate ora: " + daFare.length + (daFare.length ? "\n" + perRegola(daFare) : "");

    if (nonTrovate) esito += "\n\nNon individuate nel documento (da fare a mano): " + nonTrovate;
    if (!r.maiusc && daFare.some(p => p.f.tipo === C.MAIUSCOLETTO))
      esito += "\n\nIl maiuscoletto è applicato con lo stile carattere \"Maiuscoletto\" (questa versione di Word non lo permette direttamente agli add-in).";
    if (chiavi === null && r.daControllare.length)
      esito += "\n\nCitazioni da controllare a mano: " + r.daControllare.length + " (elenco con \"Controlla\").";
    return esito;
  });
}

// Seleziona nel documento la parte (o la nota) indicata dalla chiave
export async function vaiA(chiave, testoParte, soloSelezione) {
  return esegui(soloSelezione, false, async (context, sel) => {
    const elencoNote = await note(context, sel);
    const [a, b, c] = chiave.split("|");

    if (a === "nota") {
      const n = elencoNote[+b];
      if (!n) return false;
      n.corpo.getRange("Start").select();
      await context.sync();
      return true;
    }

    const n = elencoNote[+a];
    if (!n || n.testo.substr(+b, +c) !== testoParte) return false;   // il documento è cambiato
    const ricerca = preparaTratto(n.corpo, n.testo, +b, +c);
    await context.sync();
    const range = ricerca.risolvi();
    if (!range) return false;
    range.select();
    await context.sync();
    return true;
  });
}
