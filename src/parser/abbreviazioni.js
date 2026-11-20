// Abbreviazioni secondo la tabella delle istruzioni della cattedra (nessun accesso a Word).
// Una forma trovata: { idx, len, trovato, nuovo (null: solo da segnalare), regola, incerta }.
// Incerta: c'è una correzione proposta, ma la scelta spetta a chi scrive (casella non spuntata).
// Porting di AbbreviazioniParser.cs (versione Windows).

import * as B from "./bibliografia.js";

// inizio di parola come \b di .NET (lettere accentate comprese)
const W = "(?<![\\p{L}\\p{N}_])";
const re = s => new RegExp(s.replace(/\\b/g, W), "gu");

// L'occorrenza sta nel titolo di un'opera citata: nello stesso tratto di citazione
// (fino a ";" o alla fine della frase) segue ", in …" o ", cit."
const RE_FINE_TRATTO = /;|\.\s+(?=[A-ZÀ-Ý«])/g;
const RE_SEGNO_CITAZIONE = /,\s(?:in\s|cit\.|op\.\scit\.)/;

function inTitoloCitato(t, fine) {
  RE_FINE_TRATTO.lastIndex = fine;
  const f = RE_FINE_TRATTO.exec(t);
  return RE_SEGNO_CITAZIONE.test(t.slice(fine, f ? f.index : t.length));
}

const C = (nome, s, nuovo) => ({ nome, re: re(s), nuovo: () => nuovo });
const L = (nome, s, nuovo) => ({ nome, re: re(s), nuovo, prudenzaTitoli: true });   // nei titoli diventa incerta
const D = (nome, s, nuovo) => ({ nome, re: re(s), nuovo, dubbia: true });          // sempre da decidere

const ELENCO_DOPO = B.EDITORI.concat(B.CITTA);
const iniziaMaiuscola = c => c !== c.toLowerCase() && c === c.toUpperCase();

// Parole che, subito prima, indicano un nome dentro un titolo ("nel pensiero di Francesco Carrara",
// "Studi in onore di Antonio Fiorella"): lì il nome resta per esteso
const PRIMA_DI_NOME_NEL_TITOLO = new Set([
  "di", "del", "della", "dello", "degli", "dei", "delle", "a", "al", "alla", "per", "su", "sul",
  "onore", "memoria", "ricordo", "studi", "Studi", "centro", "Centro", "premio", "Premio", "Oscar", "on."
]);

const parolaPrima = (t, i) => { const m = /(\S+)\s*$/.exec(t.slice(0, i)); return m ? m[1] : null; };
const dopoDi = (t, m) => t.slice(m.index + m[0].length).replace(/^[, ]+/, "");

// "Rossi, Mario, Titolo": 0 non è un autore; 1 da decidere (mai automatica: "Canosa, Federico" sono due autori)
function cognomeNomeInPosizioneDiAutore(t, m) {
  const cognome = m[1], nome = m[2];
  if (!B.NOMI_PROPRI.has(nome) || B.NOMI_PROPRI.has(cognome)) return 0;

  const dopo = dopoDi(t, m);
  if (dopo.length === 0 || !(iniziaMaiuscola(dopo[0]) || "“\"«‘'".includes(dopo[0]))) return 0;
  if (ELENCO_DOPO.some(x => dopo.startsWith(x))) return 0;

  // nome dentro un titolo, oppure (di solito) due autori: in ogni caso solo proposta
  return 1;
}

// 0: non è un autore; 1: forse (nome dentro un titolo: si decide a mano); 2: autore citato
function nomeInPosizioneDiAutore(t, m) {
  if (!B.NOMI_PROPRI.has(m[1])) return 0;

  // dopo "Nome Cognome, " deve iniziare un titolo (o un altro autore): "Scalfaro, il quale" no
  const dopo = dopoDi(t, m);
  if (dopo.length === 0) return 0;
  const c = dopo[0];
  if (!(iniziaMaiuscola(c) || /\p{N}/u.test(c) || "“\"«‘'".includes(c)) && !dopo.startsWith("in ") &&
      !dopo.startsWith("cit.") && !dopo.startsWith("vol."))
    return 0;

  // "nel pensiero di Francesco Carrara", "Studi in onore di Antonio Fiorella": forse un titolo
  if (PRIMA_DI_NOME_NEL_TITOLO.has(parolaPrima(t, m.index))) return 1;

  // seguito da editore, città, "in", "cit." o numero: è la fine di un titolo
  if (/^(?:in\s|cit\.|vol\.|[IVXL]+,|[0-9])/.test(dopo)) return 1;
  if (ELENCO_DOPO.some(x => dopo.startsWith(x))) return 1;

  return 2;
}

const MESI = "gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre";

// L'ordine conta: a parità di posizione vince la regola che viene prima
const REGOLE = [
  // magistrature
  C("Cass. pen.", "\\bCass\\.\\s?Pen\\.", "Cass. pen."),
  C("Cass. civ.", "\\bCass\\.\\s?Civ\\.", "Cass. civ."),
  C("Corte cost.", "\\bCorte\\s+Cost\\.", "Corte cost."),
  C("Cons. Stato", "\\bCons\\.\\s?St\\.(?!ato)", "Cons. Stato"),
  C("Sez. Un.", "(?<=Cass\\.(?:\\s?(?:pen|civ|Pen|Civ)\\.)?,?\\s)(?:S\\.\\s?U\\.|SS\\.\\s?UU\\.|[Ss]ez\\.\\s?un\\.|sez\\.\\s?Un\\.|Sezioni\\s[Uu]nite)", "Sez. Un."),
  C("Sez. + numero romano", "(?<=,\\s)sez\\.(?=\\s[IVXL]+(?![\\p{L}\\p{N}_]))", "Sez."),

  // codici "cod. pen." ecc.: da decidere, perché compaiono anche nei titoli dei libri
  D("cod. proc. pen. → c.p.p. (può essere un titolo)", "\\bcod\\.\\s?proc\\.\\s?pen\\.", () => "c.p.p."),
  D("cod. proc. civ. → c.p.c. (può essere un titolo)", "\\bcod\\.\\s?proc\\.\\s?civ\\.", () => "c.p.c."),
  D("cod. pen. → c.p. (può essere un titolo)", "\\bcod\\.\\s?pen\\.", () => "c.p."),
  D("cod. civ. → c.c. (può essere un titolo)", "\\bcod\\.\\s?civ\\.", () => "c.c."),

  // lettere singole senza spazio (c.c., non c. c.)
  C("c.p.p.", "\\bc\\.\\s+p\\.\\s+p\\.|\\bc\\.p\\.\\s+p\\.|\\bc\\.\\s+p\\.p\\.", "c.p.p."),
  C("c.p.c.", "\\bc\\.\\s+p\\.\\s+c\\.|\\bc\\.p\\.\\s+c\\.|\\bc\\.\\s+p\\.c\\.", "c.p.c."),
  C("c.p.", "\\bc\\.\\s+p\\.(?!\\s?[pc]\\.)", "c.p."),
  C("c.c.", "\\bc\\.\\s+c\\.", "c.c."),
  C("c.d.", "\\bc\\.\\s+d\\.", "c.d."),
  C("d.lgs.", "\\bd\\.\\s+lgs\\.", "d.lgs."),
  C("D.Lgs.", "\\bD\\.\\s+Lgs\\.", "D.Lgs."),
  C("d.p.r.", "\\bd\\.\\s+p\\.\\s+r\\.", "d.p.r."),
  C("D.P.R.", "\\bD\\.\\s+P\\.\\s+R\\.", "D.P.R."),
  C("d.l.", "\\bd\\.\\s+l\\.(?!\\s?gs)", "d.l."),
  C("G.U.", "\\bG\\.\\s+U\\.", "G.U."),

  // comma per esteso (art. 10, comma 1, c.c.)
  C("comma", "\\bco\\.\\s?(?=[0-9])", "comma "),

  // autore con il nome per esteso: solo l'iniziale puntata ("Paolo Frosina, Titolo" → "P. Frosina, …")
  {
    nome: "nome dell'autore → iniziale",
    re: re("(?<![\\p{L}\\p{N}_’'.])([A-ZÀ-Ý][a-zà-ÿ]+)\\s((?:D’|D'|De |Di |Della |Del )?[A-ZÀ-Ý][a-zà-ÿ’'\\-]+)(?=,\\s)"),
    nuovo: m => m[1][0] + ". " + m[2],
    valida: nomeInPosizioneDiAutore
  },

  // forse "Cognome, Nome" ("Rossi, Mario" → "M. Rossi"): solo proposta, perché di solito sono due autori
  {
    nome: "forse cognome e nome → iniziale (di solito sono due autori)",
    re: re("(?<![\\p{L}\\p{N}_’'.])((?:D’|D'|De |Di |Della |Del )?[A-ZÀ-Ý][a-zà-ÿ’'\\-]+),\\s([A-ZÀ-Ý][a-zà-ÿ]+)(?=,\\s)"),
    nuovo: m => m[2][0] + ". " + m[1],
    valida: cognomeNomeInPosizioneDiAutore
  },

  // un solo punto dopo un'abbreviazione (cit., mai cit..)
  C("punto doppio", "(?<![.])\\.\\.(?![.])", "."),

  D("pag./pagg. → p./pp. (le istruzioni usano p./pp.)", "\\bpagg?\\.(?=\\s?[0-9])",
    m => (m[0].startsWith("pagg") ? "pp." : "p.")),

  // leggi: forma della tabella ("legge n. 10/1994"; con l'iniziale maiuscola se era maiuscola)
  L("legge n.", "(?<![.\\p{L}\\p{N}_])([lL])\\.\\s?(?=n\\.\\s?[0-9])", m => (m[1] === "L" ? "Legge " : "legge ")),
  L("legge + data", "(?<![.\\p{L}\\p{N}_])([lL])\\.\\s?(?=[0-9]{1,2}°?\\s(?:" + MESI + ")\\s[0-9]{4})",
    m => (m[1] === "L" ? "Legge " : "legge ")),
  L("legge n. (numero/anno)", "(?<![.\\p{L}\\p{N}_])([lL])\\.\\s?(?=[0-9]+/[0-9]{4})", m => (m[1] === "L" ? "Legge n. " : "legge n. "))
];

// Forme da correggere o segnalare, senza sovrapposizioni e fuori dal virgolettato («…», “…”, "…")
export function cerca(testo) {
  const tutte = [];
  if (!testo) return tutte;

  const citazione = zoneCitazione(testo);

  for (const r of REGOLE)
    for (const m of testo.matchAll(r.re)) {
      if (m[0].length === 0 || citazione[m.index]) continue;

      const esito = r.valida ? r.valida(testo, m) : 2;
      if (esito === 0) continue;

      const nuovo = r.nuovo ? r.nuovo(m) : null;
      if (nuovo === m[0]) continue;

      let nome = r.nome;
      let incerta = !!r.dubbia || esito === 1;

      // prudenza: dentro il titolo di un'opera altrui la scelta spetta a chi scrive
      if (r.prudenzaTitoli && inTitoloCitato(testo, m.index + m[0].length)) {
        incerta = true;
        nome = r.nome + " (sembra il titolo di un'opera citata)";
      } else if (esito === 1) nome = r.nome + " (forse dentro un titolo)";

      tutte.push({ idx: m.index, len: m[0].length, trovato: m[0], nuovo, regola: nome, incerta });
    }

  // a parità di inizio vince la regola elencata prima (ordinamento stabile)
  const scelte = [];
  for (const o of tutte.map((o, k) => ({ o, k })).sort((a, b) => a.o.idx - b.o.idx || a.k - b.k).map(x => x.o))
    if (!scelte.some(x => o.idx < x.idx + x.len && x.idx < o.idx + o.len)) scelte.push(o);

  return scelte;
}

// Caratteri dentro un virgolettato: «…», “…” e "…" (citazioni testuali o espressioni riportate),
// da non toccare. Gli apici singoli ‘…’ no: lo stesso carattere fa anche da apostrofo.
function zoneCitazione(t) {
  const z = new Array(t.length + 1).fill(false);
  let basse = 0, alte = 0, dritte = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (c === "«") basse++;
    else if (c === "“") alte++;
    else if (c === "\"") dritte = !dritte;
    z[i] = basse > 0 || alte > 0 || dritte || c === "\"";
    if (c === "»" && basse > 0) basse--;
    else if (c === "”" && alte > 0) alte--;
  }
  return z;
}

// Testo intorno all'occorrenza, per l'elenco
export function contesto(testo, o, ampiezza = 35) {
  const a = Math.max(0, o.idx - ampiezza);
  const b = Math.min(testo.length, o.idx + o.len + ampiezza);
  const s = testo.slice(a, b).replace(/[\r\n\u0007]/g, " ");
  return (a > 0 ? "…" : "") + s + (b < testo.length ? "…" : "");
}
