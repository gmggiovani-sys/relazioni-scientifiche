// Analisi del testo delle decisioni (nessun accesso a Word), secondo le istruzioni:
// "Organo, sez., giorno mese anno, n. X, in Rivista (corsivo), anno, parte, pagina".
// Porting di GiurisprudenzaParser.cs (versione Windows).

const MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno",
              "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"];

// Nome per esteso della rivista → formula breve d'uso
export const RIVISTE = [
  ["Rivista italiana di diritto e procedura penale", "Riv. it. dir. proc. pen."],
  ["Rivista trimestrale di diritto penale dell'economia", "Riv. trim. dir. pen. econ."],
  ["Rivista trimestrale di diritto penale dell’economia", "Riv. trim. dir. pen. econ."],
  ["Nuova giurisprudenza civile commentata", "Nuova giur. civ. comm."],
  ["Rivista trimestrale di diritto e procedura civile", "Riv. trim. dir. proc. civ."],
  ["Diritto penale contemporaneo", "Dir. pen. cont."],
  ["Diritto penale e processo", "Dir. pen. proc."],
  ["Giurisprudenza italiana", "Giur. it."],
  ["Giurisprudenza penale", "Giur. pen."],
  ["Giurisprudenza costituzionale", "Giur. cost."],
  ["Rivista di diritto civile", "Riv. dir. civ."],
  ["Il Foro italiano", "Foro it."],
  ["Foro italiano", "Foro it."],
  ["Giustizia civile", "Giust. civ."],
  ["Giustizia penale", "Giust. pen."],
  ["Cassazione penale", "Cass. pen."],
  ["Rivista penale", "Riv. pen."],
  ["Archivio penale", "Arch. pen."],
  ["Sistema penale", "Sist. pen."],
  ["Guida al diritto", "Guida dir."],
  ["Diritto e Giustizia", "Dir. e giust."],
  ["Diritto e giustizia", "Dir. e giust."],
  ["Gazzetta giuridica", "Gazz. giur."],
  ["Gazzetta Ufficiale della Repubblica italiana", "G.U."],
  ["Gazzetta Ufficiale della Repubblica Italiana", "G.U."],
  ["Gazzetta Ufficiale della Repubblica", "G.U."],
  ["Gazzetta Ufficiale della Regione siciliana", "G.U. Si."],
  ["Gazzetta Ufficiale della Regione Siciliana", "G.U. Si."],
  ["Gazzetta Ufficiale delle Comunità europee", "G.U.C.E."],
  ["Gazzetta Ufficiale della Comunità Europea", "G.U.C.E."],
  ["Gazzetta Ufficiale della Comunità europea", "G.U.C.E."],
  ["Gazzetta Ufficiale", "G.U."],
  ["Famiglia e diritto", "Fam. dir."]
];

// Banche dati e Gazzetta: nessuna pagina da indicare
const SENZA_PAGINA = ["Dejure", "One LEGALE", "C.E.D. Cass.", "Mass. Uff.", "G.U.", "giurcost.org", "Giur. pen. web"];

// Gruppi dell'elenco, riconosciuti dall'inizio della voce (abbreviazioni della tabella della cattedra)
const GRUPPI = [
  "Corte cost.", "Cass.", "Cons. Stato", "Cons. St.", "Corte conti", "C. conti",
  "Corte ass. App.", "Corte. ass. App.", "Corte ass.", "Corte. ass.",
  "App.", "Corte app.", "Trib.", "Pret.", "T.A.R.", "TAR", "Comm. trib.", "G.i.p.", "G.u.p.",
  "Tribunale di primo grado CE", "Corte giust. CE", "Corte giust.", "C. giust.", "CGUE",
  "Corte EDU", "C. EDU"
];

// Abbreviazioni ammesse per l'organo (inizio della voce)
const ORGANI_AMMESSI = [
  "Corte cost.", "Cass. pen.", "Cass. civ.", "Cass. lav.", "Cass.", "Corte conti", "Cons. Stato",
  "Corte ass. App.", "Corte. ass. App.", "Corte ass.", "Corte. ass.", "App.", "Trib.", "Pret.",
  "T.A.R.", "Tribunale di primo grado CE", "Comm. trib.", "Corte giust. CE", "Corte EDU"
];

// Organo per esteso all'inizio della voce → abbreviazione della tabella
const ORGANI = [
  ["Corte costituzionale,", "Corte cost.,"],
  ["Corte Costituzionale,", "Corte cost.,"],
  ["Cassazione penale,", "Cass. pen.,"],
  ["Cassazione civile,", "Cass. civ.,"],
  ["Corte di cassazione,", "Cass.,"],
  ["Corte di Cassazione,", "Cass.,"],
  ["Corte dei conti", "Corte conti"],
  ["Corte dei Conti", "Corte conti"],
  ["Consiglio di Stato", "Cons. Stato"],
  ["Corte d'Assise d'Appello di ", "Corte ass. App. "],
  ["Corte d’Assise d’Appello di ", "Corte ass. App. "],
  ["Corte d'Assise di ", "Corte ass. "],
  ["Corte d’Assise di ", "Corte ass. "],
  ["Corte d'Appello di ", "App. "],
  ["Corte d’Appello di ", "App. "],
  ["Corte d'appello di ", "App. "],
  ["Corte d’appello di ", "App. "],
  ["Tribunale amministrativo regionale", "T.A.R."],
  ["TAR ", "T.A.R. "],
  ["Tribunale di primo grado CE", "Tribunale di primo grado CE"],
  ["Tribunale di ", "Trib. "],
  ["Pretura di ", "Pret. "],
  ["Commissione tributaria", "Comm. trib."],
  ["Corte di giustizia CE", "Corte giust. CE"],
  ["Corte di Giustizia CE", "Corte giust. CE"]
];

const RE_MESE = "(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)";

const RE_DATA = new RegExp("\\b([0-9]{1,2})°? " + RE_MESE + " ((?:18|19|20)[0-9]{2})\\b");
const RE_DATA_NUMERICA = /(?<![0-9/.])([0-9]{1,2})[./]([0-9]{1,2})[./]((?:18|19|20)[0-9]{2})\b/g;
const RE_MESE_ABBREVIATO =
  /\b([0-9]{1,2})°? (genn|febbr|febb|feb|mar|apr|magg|mag|giu|lug|ag|sett|set|ott|nov|dic)\. ((?:18|19|20)[0-9]{2})\b/g;
const RE_IN = /,\s*in\b\s*(.*)$/s;
const RE_NUMERO = /\bn\. ?([0-9][^\s,;]*)/;
const RE_DEPOSITO = /\((dep\.|ud\.) ([^)]*)\)/;
const RE_ANNO = /\b((?:18|19|20)[0-9]{2})\b/g;
const RE_PARTE_ANNO = /, ([IVX]+), ((?:18|19|20)[0-9]{2})\b/g;
const RE_SEPARA_FONTI = / e (?=[A-Z])/;

const giorniNelMese = (anno, mese) => new Date(Date.UTC(anno, mese, 0)).getUTCDate();   // mese 1–12

// ---------- estrazione dei campi ----------

export function parse(raw) {
  const t = (raw || "").trim();
  const o = {
    testo: t, autorita: "", gruppo: gruppo(t), sezione: "", provvedimento: "", data: "", dataDecisione: null,
    deposito: "", numero: "", parti: "", fonte: "", riviste: [], problemi: []
  };

  // parte prima di ", in": organo, sezione, data, numero, parti
  const mi = RE_IN.exec(t);
  const testa = mi ? t.slice(0, mi.index) : t.replace(/\.+$/, "");
  o.fonte = mi ? mi[1].trim() : "";

  const md = RE_DATA.exec(testa);
  if (md) {
    o.data = md[0];
    const { data, valida } = dataDi(md);
    o.dataDecisione = data;
    if (!valida) o.problemi.push("data impossibile (" + md[0] + ")");
  }

  const mdep = RE_DEPOSITO.exec(testa);
  if (mdep) o.deposito = mdep[1] + " " + mdep[2];

  // segmenti separati da virgola, escluse le parentesi del deposito
  const senzaDep = testa.replace(new RegExp(RE_DEPOSITO.source, "g"), "");
  const seg = senzaDep.split(", ").map(s => s.trim());

  const idxData = seg.findIndex(s => RE_DATA.test(s));
  const prima = idxData >= 0 ? seg.slice(0, idxData) : seg.slice(0, 1);
  const dopo = idxData >= 0 ? seg.slice(idxData + 1) : seg.slice(1);

  if (prima.length > 0) o.autorita = prima[0];
  o.sezione = prima.slice(1).join(", ");

  const segData = idxData >= 0 ? seg[idxData] : "";
  if (segData.startsWith("ord.") || dopo.some(s => s.startsWith("ord."))) o.provvedimento = "ord.";
  else if (segData.startsWith("sent.")) o.provvedimento = "sent.";

  for (const s of dopo) {
    const mn = RE_NUMERO.exec(s);
    if (mn && o.numero.length === 0) {
      o.numero = mn[1];
      if (!/^[0-9]+(?:[/-][0-9]+)*$/.test(o.numero)) o.problemi.push("numero anomalo (n. " + o.numero + ")");
    } else if (!mn && s.length > 0 && isUpper(s[0]))
      o.parti = o.parti.length === 0 ? s : o.parti + ", " + s;
  }

  if (!ORGANI_AMMESSI.some(x => o.autorita.startsWith(x)))
    o.problemi.push("organo non abbreviato come in tabella (" + o.autorita + ")");

  // fonti: "in C.E.D. Cass., n. 181834 e Cass. pen., 1990, I"
  o.riviste = nomiRiviste(o.fonte);
  if (o.riviste.length === 0) o.problemi.push("fonte mancante");

  // pubblicazione anteriore alla decisione
  if (o.dataDecisione) {
    for (const ma of o.fonte.matchAll(RE_ANNO))
      if (parseInt(ma[0], 10) < o.dataDecisione.anno) {
        o.problemi.push("anno della rivista (" + ma[0] + ") anteriore alla decisione");
        break;
      }
  } else o.problemi.push("data non riconosciuta (serve giorno, mese per esteso e anno)");

  return o;
}

const isUpper = c => c !== c.toLowerCase() && c === c.toUpperCase();

// "Dejure." → il punto chiude la voce; "Cass. pen." → il punto è dell'abbreviazione
function ultimaParolaLunga(nome) {
  const s = nome.replace(/\.+$/, "");
  const sp = s.lastIndexOf(" ");
  const w = sp >= 0 ? s.slice(sp + 1) : s;
  return w.length > 4 || w.includes(".") || w.toUpperCase() === w;
}

// Nomi delle riviste o banche dati citate dopo "in" (le parti da mettere in corsivo)
export function nomiRiviste(fonte) {
  const nomi = [];
  const f = (fonte || "").trim();

  for (const parte of f.split(RE_SEPARA_FONTI)) {
    const p = parte.trim();
    const virgola = p.indexOf(", ");
    let nome = virgola >= 0 ? p.slice(0, virgola) : p;

    // il nome finisce prima di due punti, virgolette o parentesi ("One LEGALE: «…»")
    const stop = nome.search(/[:«“(]/);
    if (stop > 0) nome = nome.slice(0, stop).trimEnd();

    // il punto finale della voce non fa parte del nome
    if (virgola < 0 && nome.endsWith(".") && ultimaParolaLunga(nome)) nome = nome.slice(0, -1);

    // Gazzetta Ufficiale: in corsivo solo la sigla
    if (nome.startsWith("G.U.C.E.")) nome = "G.U.C.E.";
    else if (nome.startsWith("G.U. Si.")) nome = "G.U. Si.";
    else if (nome.startsWith("G.U")) nome = "G.U.";

    nome = nome.trim();
    if (nome.replace(/^[. ]+|[. ]+$/g, "").length > 1) nomi.push(nome);
  }

  return nomi;
}

// Blocco dell'elenco a cui appartiene la voce
export function gruppo(t) {
  let g = GRUPPI.find(x => t.startsWith(x)) || null;
  if (g === "Cons. St.") g = "Cons. Stato";
  if (g === "C. conti") g = "Corte conti";
  if (g === "C. EDU") g = "Corte EDU";
  if (g === "C. giust." || g === "CGUE" || g === "Corte giust.") g = "Corte giust. CE";
  if (g === "Corte app.") g = "App.";
  if (g === "Corte. ass. App.") g = "Corte ass. App.";
  if (g === "Corte. ass.") g = "Corte ass.";
  if (g === "TAR") g = "T.A.R.";
  if (g !== null) return g;

  const v = t.indexOf(",");
  return v > 0 ? t.slice(0, v) : t;
}

function dataDi(m) {
  const g = parseInt(m[1], 10);
  const mese = MESI.indexOf(m[2]) + 1;
  const anno = parseInt(m[3], 10);
  const max = giorniNelMese(anno, mese);
  return { valida: g >= 1 && g <= max, data: { anno, mese, giorno: Math.min(Math.max(g, 1), max) } };
}

// ---------- normalizzazione ----------

// Sostituzioni nel testo della voce [vecchio, nuovo] (sempre sull'ultima occorrenza):
// organi e sezioni con le abbreviazioni della tabella ("Cass. pen., Sez. VI", "Sez. Un.", "Cons. Stato");
// date in cifre o con mese abbreviato → mese per esteso;
// rivista per esteso → formula breve; "in Rivista, I, 1965" → "in Rivista, 1965, I"
export function normalizzazioni(raw) {
  const t = (raw || "").trimEnd();
  const jobs = [];

  // punteggiatura doppia (anche lasciata da esecuzioni precedenti)
  const mp = /,\s*,/.exec(t);
  if (mp) jobs.push([mp[0], ","]);

  // solo nella parte prima di "in" (organo, sezione, data, numero)
  const mHead = RE_IN.exec(t);
  const head = mHead ? t.slice(0, mHead.index) : t;
  const addHead = (a, b) => { if (a !== b && head.includes(a)) jobs.push([a, b]); };

  const organo = ORGANI.find(o => t.startsWith(o[0]));
  if (organo) addHead(organo[0], organo[1]);

  addHead("Cass. Pen.,", "Cass. pen.,");
  addHead("Cass. Civ.,", "Cass. civ.,");
  addHead("Cass. Lav.,", "Cass. lav.,");
  addHead("Cons. St.,", "Cons. Stato,");

  for (const su of [", sez. un.,", ", Sez. un.,", ", sez. Un.,", ", Sez. Unite,", ", sez. unite,",
                    ", Sezioni Unite,", ", Sezioni unite,", ", S.U.,", ", SS.UU.,"])
    addHead(su, ", Sez. Un.,");

  const ms = /, (?:sez\.|sezione|Sezione) (?=[IVXL]+\b|fer\.|pen\.|civ\.|lav\.|trib\.|GIP|GUP)/.exec(head);
  if (ms) jobs.push([ms[0], ", Sez. "]);

  for (const [a, b] of [["Ad. plen.", "Ad. Plen."], ["Adunanza plenaria", "Ad. Plen."],
                        ["Ad. gen.", "Ad. Gen."], ["Adunanza generale", "Ad. Gen."]])
    addHead(a, b);

  for (const m of t.matchAll(RE_DATA_NUMERICA)) {
    const mese = parseInt(m[2], 10);
    if (mese >= 1 && mese <= 12)
      jobs.push([m[0], parseInt(m[1], 10) + " " + MESI[mese - 1] + " " + m[3]]);
  }

  for (const m of t.matchAll(RE_MESE_ABBREVIATO)) {
    const esteso = MESI.find(x => x.startsWith(m[2]));
    jobs.push([m[0], m[1] + " " + esteso + " " + m[3]]);
  }

  const mi = RE_IN.exec(t);
  if (mi) {
    let fonte = mi[1];
    for (const r of RIVISTE)
      if (fonte.includes(r[0]) && !fonte.includes(r[1])) {
        jobs.push(r);
        fonte = fonte.split(r[0]).join(r[1]);
      }
    for (const m of fonte.matchAll(RE_PARTE_ANNO))
      jobs.push([m[0], ", " + m[2] + ", " + m[1]]);
  }

  return jobs;
}

// ---------- ordinamento ----------

// Chiave: blocco (nell'ordine in cui compare nel documento), poi data della decisione, poi numero;
// soloData: data della decisione e numero, senza blocchi (a parità di data conta il blocco)
export function chiaveOrdinamento(v, ordineGruppi, soloData = false) {
  let g = ordineGruppi.indexOf(v.gruppo);
  if (g < 0) g = 999;

  const d = v.dataDecisione;
  const data = d ? String(d.anno).padStart(4, "0") + String(d.mese).padStart(2, "0") + String(d.giorno).padStart(2, "0")
                 : "99999999";

  const mn = /^[0-9]+/.exec(v.numero);
  const num = mn ? mn[0].padStart(10, "0") : "9999999999";

  return soloData
    ? data + "#" + String(g).padStart(3, "0") + "#" + num
    : String(g).padStart(3, "0") + "#" + data + "#" + num;
}

// Voce da controllare per la pagina: rivista cartacea senza "p." o "c."
export function senzaPaginaRivista(v) {
  if (v.riviste.length === 0) return false;
  if (v.riviste.every(r => SENZA_PAGINA.some(x => r.startsWith(x)))) return false;
  return !/\b(p|pp|c|cc)\. ?[0-9]/.test(v.fonte);
}
