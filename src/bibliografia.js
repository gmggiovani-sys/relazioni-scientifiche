/* global Word */
// Bibliografia (paragrafi con lo stile indicato, di solito "Biblio"):
// Editore, Città e iniziali; ordine alfabetico per cognome; corsivi e maiuscoletto; campi; voci dalle note.
// Le regole sono in parser/bibliografia.js, come nella versione Windows.

import * as B from "./parser/bibliografia.js";
import * as NB from "./parser/note-bibliografia.js";
import * as I from "./parser/iniziali.js";
import {
  note, preparaTratto, aBlocchi, maiuscolettoDisponibile, assicuraStileMaiuscoletto, applicaMaiuscoletto
} from "./word.js";
import { voci, sostituisci, ordina as ordinaVoci } from "./elenchi.js";

const BLOCCO = 40;      // voci lavorate per ogni richiesta a Word

const breve = t => (t.length > 60 ? t.slice(0, 60) + "…" : t);

// ===============================
// PASSAGGI (dentro un Word.run)
// ===============================

// "Milano, Giuffrè" → "Giuffrè, Milano"; "fasc. II, 2024" → "2024, fasc. II"; "Paolo Frosina" → "P. Frosina"
const ordineCampi = (context, stile, avanzamento) => sostituisci(context, stile, B.sostituzioniOrdineCampi, avanzamento);

// Ordine alfabetico per cognome (poi anno): si spostano solo le voci fuori posto
const ordina = (context, stile) => ordinaVoci(context, stile, testi => testi.map(B.chiaveOrdinamento), "la bibliografia");

// Corsivi e maiuscoletto delle voci indicate (tutte se paragrafi è null): si azzerano e si riapplicano
async function corsivi(context, stile, avanzamento, paragrafi = null, testi = null) {
  if (!paragrafi) ({ paragrafi, testi } = await voci(context, stile));

  const maiusc = maiuscolettoDisponibile();
  if (!maiusc) await assicuraStileMaiuscoletto(context);

  const res = { corsivi: 0, maiuscoletto: 0, mancanti: [], perTipo: {} };
  const elenco = paragrafi.map((p, i) => ({ p, t: testi[i] }));

  await aBlocchi(elenco, BLOCCO, async blocco => {
    const lavori = [];
    for (const { p, t } of blocco) {
      p.font.italic = false;
      if (maiusc) p.font.smallCaps = false;

      // autori (e curatori dei saggi in volume) in MAIUSCOLETTO, cercati in ordine:
      // lo stesso nome può comparire due volte ("Catenacci, …, in Catenacci (a cura di)")
      let da = 0;
      for (const nome of B.maiuscoletto(t)) {
        const pos = t.indexOf(nome, da);
        if (pos < 0) { res.mancanti.push(breve(t)); continue; }
        lavori.push({ tipo: "m", cerca: preparaTratto(p, t, pos, nome.length), t });
        da = pos + nome.length;
      }

      const { tipo, parti } = B.corsivi(t);
      for (const parte of parti) {
        const pos = t.indexOf(parte);
        if (pos < 0) { res.mancanti.push(breve(t)); continue; }
        lavori.push({ tipo: "c", cerca: preparaTratto(p, t, pos, parte.length), t });
      }
      res.perTipo[tipo] = (res.perTipo[tipo] || 0) + 1;
    }
    await context.sync();

    for (const l of lavori) {
      const r = l.cerca.risolvi();
      if (!r) { res.mancanti.push(breve(l.t)); continue; }
      if (l.tipo === "c") { r.font.italic = true; res.corsivi++; }
      else { applicaMaiuscoletto(r, maiusc); res.maiuscoletto++; }
    }
    await context.sync();
  }, avanzamento);

  return res;
}

function rapportoCorsivi(r) {
  const tipi = Object.keys(r.perTipo).sort().map(k => k + " " + r.perTipo[k]).join(", ");
  let s = "Corsivi applicati: " + r.corsivi + (tipi ? " (" + tipi + ")" : "") +
          "\nNomi in maiuscoletto (autori e curatori): " + r.maiuscoletto;
  const mancanti = [...new Set(r.mancanti)];
  if (mancanti.length)
    s += "\n\nDa controllare a mano (" + mancanti.length + "):\n" +
         mancanti.slice(0, 15).map(x => "• " + x).join("\n") + (mancanti.length > 15 ? "\n…" : "");
  return s;
}

// ===============================
// COMANDI
// ===============================

export async function sistema(stile, avanzamento) {
  return Word.run(async context => {
    // prima si ordina, poi si modifica il testo, infine corsivi e maiuscoletto
    const spostate = await ordina(context, stile);
    const campi = await ordineCampi(context, stile, avanzamento && ((f, t) => avanzamento("Editore, Città", f, t)));
    const c = await corsivi(context, stile, avanzamento && ((f, t) => avanzamento("Corsivi e maiuscoletto", f, t)));
    return "Bibliografia sistemata (stile \"" + stile + "\").\n\n" +
           "Voci portate a Editore, Città (e iniziali dei nomi): " + campi + "\n" +
           "Voci spostate in ordine alfabetico: " + spostate + "\n" + rapportoCorsivi(c);
  });
}

export async function soloOrdineCampi(stile, avanzamento) {
  return Word.run(async context =>
    "Voci riordinate (Editore, Città / Anno, fascicolo / iniziali dei nomi): " +
    await ordineCampi(context, stile, avanzamento && ((f, t) => avanzamento("Editore, Città", f, t))));
}

export async function soloOrdina(stile) {
  return Word.run(async context => {
    const n = await ordina(context, stile);
    if (n === 0) return "Le voci sono già in ordine alfabetico.";
    // le voci spostate sono state riscritte: si rimettono corsivi e maiuscoletto
    await corsivi(context, stile, null);
    return "Voci spostate: " + n + " (con corsivi e maiuscoletto riapplicati).";
  });
}

export async function soloCorsivi(stile, avanzamento) {
  return Word.run(async context =>
    rapportoCorsivi(await corsivi(context, stile, avanzamento && ((f, t) => avanzamento("Corsivi e maiuscoletto", f, t)))));
}

// Campi estratti da ogni voce, per la tabella del pannello
export async function campi(stile) {
  return Word.run(async context => {
    const { testi } = await voci(context, stile);
    return testi.map((t, i) => {
      const v = B.parse(t);
      return { i, testo: t, autore: v.autore, titolo: v.titolo, fonte: v.fonte, editore: v.editore, anno: v.anno, citta: v.citta, tipo: v.tipo };
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

// ===============================
// NOMI DEGLI AUTORI: iniziale del nome ("F. Antolisei") e "e" tra gli ultimi due autori
// ===============================

// Elenco delle iniziali ("Cognome = F."): si tiene nella memoria del pannello, modificabile dall'utente
const CHIAVE_ELENCO = "rs.inizialiElenco";

export function elencoIniziali() {
  try {
    const v = localStorage.getItem(CHIAVE_ELENCO);
    if (v) return v;
  } catch (e) { /* senza memoria locale: l'elenco predefinito */ }
  return I.INTESTAZIONE_ELENCO + I.PREDEFINITE;
}

export function salvaElencoIniziali(testo) {
  try {
    if (testo && testo.trim()) localStorage.setItem(CHIAVE_ELENCO, testo);
    else localStorage.removeItem(CHIAVE_ELENCO);
    return true;
  } catch (e) { return false; }
}

export const elencoPredefinito = () => I.INTESTAZIONE_ELENCO + I.PREDEFINITE;

// Proposte per tutte le voci: [{ i, t (testo letto), p (proposta) }]
export async function nomiAutori(stile) {
  return Word.run(async context => {
    const { testi } = await voci(context, stile);

    // iniziali usate nelle note per gli stessi cognomi ("P. Frosina")
    let testiNote = [];
    try { testiNote = (await note(context, null)).map(n => n.testo); } catch (e) { /* note non leggibili */ }
    const dalleNote = I.dalleNote(testiNote, testi.flatMap(B.maiuscoletto));

    const regole = I.carica(elencoIniziali());
    const righe = [];
    testi.forEach((t, i) => { for (const p of I.proposte(t, regole, dalleNote)) righe.push({ i, t, p }); });
    return righe;
  });
}

// Applica le proposte scelte, poi rimette corsivi e maiuscoletto alle voci cambiate
export async function applicaNomi(stile, scelte) {
  return Word.run(async context => {
    const { paragrafi, testi } = await voci(context, stile);
    let iniziali = 0, congiunzioni = 0, saltate = 0;
    const cambiate = new Set();

    // la voce deve essere ancora quella letta dal controllo
    const valide = scelte.filter(x => {
      const ok = testi[x.i] === x.t;
      if (!ok) saltate++;
      return ok;
    });

    const lavori = valide.map(x => ({ x, cerca: preparaTratto(paragrafi[x.i], x.t, x.p.idx, x.p.nome.length) }));
    await context.sync();

    // dentro ogni voce dall'ultimo nome al primo
    lavori.sort((a, b) => a.x.i - b.x.i || b.x.p.idx - a.x.p.idx);
    for (const { x, cerca } of lavori) {
      const r = cerca.risolvi();
      if (!r) { saltate++; continue; }
      if (x.p.sostituto) { r.insertText(x.p.sostituto, "Replace"); congiunzioni++; }
      else { r.insertText(x.p.iniziale + " ", "Start"); iniziali++; }
      cambiate.add(x.i);
    }
    await context.sync();

    // il maiuscoletto comprende anche le iniziali
    const indici = [...cambiate];
    const dopo = await voci(context, stile);
    const c = await corsivi(context, stile, null, indici.map(i => dopo.paragrafi[i]), indici.map(i => dopo.testi[i]));

    let esito = "Iniziali aggiunte: " + iniziali + "\n\"e\" tra gli ultimi due autori o curatori: " + congiunzioni +
                "\nVoci cambiate: " + indici.length + " (con corsivi e maiuscoletto riapplicati: " + c.maiuscoletto + " nomi).";
    if (saltate) esito += "\n\nNon applicate perché la voce è cambiata dopo il controllo: " + saltate + ". Ripeti il controllo.";
    return esito;
  });
}

// Opere citate per esteso nelle note e assenti dalla bibliografia
export async function proposteDalleNote(stile) {
  return Word.run(async context => {
    const { testi } = await voci(context, stile);
    const elenco = await note(context, null);
    return NB.proposte(elenco.map(n => [n.dove, n.testo]), testi);
  });
}

// Inserisce le voci scelte al loro posto in ordine alfabetico, con stile, corsivi e maiuscoletto
export async function aggiungiVoci(stile, scelte) {
  return Word.run(async context => {
    const { paragrafi, testi } = await voci(context, stile);
    const elenco = paragrafi.map((p, i) => ({ p, chiave: B.chiaveOrdinamento(testi[i]) }));
    const nuove = [], testiNuove = [];

    for (const v of [...scelte].sort((a, b) => B.ordinale(a.chiave, b.chiave))) {
      const idx = elenco.findIndex(e => B.ordinale(e.chiave, v.chiave) > 0);
      const p = idx >= 0 ? elenco[idx].p.insertParagraph(v.voce, "Before")
                         : elenco[elenco.length - 1].p.insertParagraph(v.voce, "After");
      p.style = stile;
      elenco.splice(idx >= 0 ? idx : elenco.length, 0, { p, chiave: v.chiave });
      nuove.push(p);
      testiNuove.push(v.voce);
    }
    await context.sync();
    await corsivi(context, stile, null, nuove, testiNuove);

    return "Voci aggiunte alla bibliografia: " + scelte.length +
           "\nSono al loro posto in ordine alfabetico, con stile \"" + stile + "\", corsivi e maiuscoletto." +
           "\n\nControllale una per una: sono ricavate dalle citazioni nelle note, " +
           "dove possono mancare dati (pagine, fascicolo) o comparire refusi.";
  });
}

