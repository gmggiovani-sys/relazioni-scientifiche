/* global Word */
// Elenco della giurisprudenza: paragrafi con lo stile indicato (es. "Giurisprudenza").
// Sequenza: Ordina (per data, dentro ogni blocco) → Normalizza → Corsivi; Controlla non modifica nulla.
// Le regole sono in parser/giurisprudenza.js, come nella versione Windows.

import * as G from "./parser/giurisprudenza.js";
import { preparaTratto, aBlocchi } from "./word.js";
import { voci, sostituisci, ordina as ordinaVoci } from "./elenchi.js";

const BLOCCO = 40;
const MAX_ELENCO = 15;
const breve = t => (t.length > 60 ? t.slice(0, 60) + "…" : t);

// ===============================
// PASSAGGI (dentro un Word.run)
// ===============================

const normalizza = (context, stile, avanzamento) => sostituisci(context, stile, G.normalizzazioni, avanzamento);

// Per data della decisione dentro ogni blocco (Cassazione, Corte cost., Tribunali…),
// lasciando i blocchi nell'ordine in cui sono nel documento
function chiavi(testi) {
  const parsed = testi.map(G.parse);
  const gruppi = [];
  for (const p of parsed) if (!gruppi.includes(p.gruppo)) gruppi.push(p.gruppo);
  return parsed.map(p => G.chiaveOrdinamento(p, gruppi));
}

const ordina = (context, stile) => ordinaVoci(context, stile, chiavi, "la giurisprudenza");

// Azzera il corsivo e lo applica ai nomi delle riviste e delle banche dati
async function corsivi(context, stile, avanzamento) {
  const { paragrafi, testi } = await voci(context, stile);
  let n = 0;

  await aBlocchi(paragrafi.map((p, i) => ({ p, t: testi[i] })), BLOCCO, async blocco => {
    const lavori = [];
    for (const { p, t } of blocco) {
      p.font.italic = false;
      const v = G.parse(t);

      // le riviste stanno dopo "in": si cerca da lì in avanti, altrimenti l'ultima occorrenza
      const inizioFonte = v.fonte.length > 0 ? Math.max(t.lastIndexOf(v.fonte), 0) : 0;
      for (const nome of v.riviste) {
        let pos = t.indexOf(nome, inizioFonte);
        if (pos < 0) pos = t.lastIndexOf(nome);
        if (pos >= 0) lavori.push(preparaTratto(p, t, pos, nome.length));
      }
    }
    await context.sync();

    for (const l of lavori) {
      const r = l.risolvi();
      if (r) { r.font.italic = true; n++; }
    }
    await context.sync();
  }, avanzamento);

  return n;
}

// Voci da completare o correggere a mano
function rapporto(testi) {
  const vv = testi.map(G.parse);
  const problemi = vv.filter(v => v.problemi.length > 0);
  const senzaPagina = vv.filter(G.senzaPaginaRivista).length;

  let s = "Decisioni nell'elenco: " + vv.length + "\n";

  // abbreviazioni non conformi alla tabella, che "Normalizza" corregge
  const daUniformare = vv.map(v => G.normalizzazioni(v.testo));
  const nUniformare = daUniformare.filter(j => j.length > 0).length;

  if (nUniformare > 0) {
    const conta = new Map();
    const pulisci = x => x.replace(/^[, ]+|[, ]+$/g, "");
    for (const [a, b] of daUniformare.flat()) {
      const k = "\"" + pulisci(a) + "\" → \"" + pulisci(b) + "\"";
      conta.set(k, (conta.get(k) || 0) + 1);
    }
    const tipi = [...conta].sort((x, y) => y[1] - x[1]).slice(0, MAX_ELENCO).map(([k, c]) => "• " + k + " (" + c + ")");
    s += "\nAbbreviazioni da uniformare alla tabella in " + nUniformare + " voci (le corregge \"Normalizza\"):\n" +
         tipi.join("\n") + "\n";
  }

  if (problemi.length === 0) s += "Nessuna voce da correggere.\n";
  else
    s += "\nDa controllare a mano (" + problemi.length + "):\n" +
         problemi.slice(0, MAX_ELENCO).map(v => "• " + breve(v.testo) + " → " + v.problemi.join("; ")).join("\n") +
         (problemi.length > MAX_ELENCO ? "\n… (elenco completo con \"Estrai campi\")" : "") + "\n";

  if (senzaPagina > 0)
    s += "\nVoci su rivista senza pagina (\"p.\" o \"c.\"): " + senzaPagina +
         ". Le istruzioni chiedono rivista, anno, parte e pagina.";

  return s;
}

// ===============================
// COMANDI
// ===============================

const fase = (avanzamento, cosa) => avanzamento && ((f, t) => avanzamento(cosa, f, t));

export async function sistema(stile, avanzamento) {
  return Word.run(async context => {
    // prima si ordina, poi si modifica il testo, infine i corsivi
    const spostate = await ordina(context, stile);
    const norm = await normalizza(context, stile, fase(avanzamento, "Normalizzazione"));
    const c = await corsivi(context, stile, fase(avanzamento, "Corsivi"));
    const { testi } = await voci(context, stile);
    return "Giurisprudenza sistemata (stile \"" + stile + "\").\n\n" +
           "Voci normalizzate (abbreviazioni, date, riviste): " + norm + "\n" +
           "Voci spostate in ordine cronologico: " + spostate + "\n" +
           "Riviste in corsivo: " + c + "\n\n" + rapporto(testi);
  });
}

export async function soloNormalizza(stile, avanzamento) {
  return Word.run(async context => "Voci normalizzate: " + await normalizza(context, stile, fase(avanzamento, "Normalizzazione")));
}

export async function soloOrdina(stile) {
  return Word.run(async context => {
    const n = await ordina(context, stile);
    if (n === 0) return "Le voci sono già in ordine cronologico.";
    await corsivi(context, stile, null);          // le voci spostate sono state riscritte
    return "Voci spostate: " + n + " (con i corsivi delle riviste riapplicati).";
  });
}

export async function soloCorsivi(stile, avanzamento) {
  return Word.run(async context => "Riviste in corsivo: " + await corsivi(context, stile, fase(avanzamento, "Corsivi")));
}

// Solo verifica: non modifica il documento
export async function controlla(stile) {
  return Word.run(async context => rapporto((await voci(context, stile)).testi));
}

// Campi di ogni decisione, per l'elenco del pannello
export async function campi(stile) {
  return Word.run(async context => {
    const { testi } = await voci(context, stile);
    return testi.map((t, i) => {
      const v = G.parse(t);
      return {
        i, testo: t, autorita: v.autorita, sezione: v.sezione, provvedimento: v.provvedimento, data: v.data,
        deposito: v.deposito, numero: v.numero, parti: v.parti, fonte: v.fonte, problemi: v.problemi
      };
    });
  });
}

export async function vaiAVoce(stile, indice) {
  return Word.run(async context => {
    const { paragrafi } = await voci(context, stile);
    if (!paragrafi[indice]) return false;
    paragrafi[indice].select();
    await context.sync();
    return true;
  });
}
