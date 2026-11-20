// Analisi del testo delle voci di bibliografia (nessun accesso a Word):
// estrazione dei campi, parti in corsivo e in maiuscoletto, ordine dei campi, chiave di ordinamento.
// Porting di BibliografiaParser.cs (versione Windows).

export const EDITORI = [
  "Giuffrè", "CEDAM", "Cedam", "Utet", "Giappichelli", "Jovene", "Laterza", "Il Mulino",
  "Zanichelli", "Pacini Giuridica", "Ciampi Editore", "Pisa University Press", "Dike Giuridica",
  "Edizioni ETS", "Monduzzi", "CUEC", "Tipografia Giachetti", "Roma Tre-Press",
  "Istituto della Enciclopedia Italiana", "Edizioni Scientifiche Italiane",
  "Editoria Scientifica di Ateneo dell’Università degli Studi di Trento"
];

export const CITTA = [
  "Milano", "Padova", "Torino", "Roma", "Napoli", "Bologna", "Bari", "Pisa", "Prato",
  "Cagliari", "Trento", "Firenze", "Venezia", "Genova", "Palermo", "Castelfranco Veneto"
];

const MESI = "gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre";

// Segmento con forma di cognome (anche "D’Amico", "De Vero", "F. Galgano", "Rossi (a cura di)", "Aa. Vv.")
export const RE_AUTORE = new RegExp(
  "^(?:[A-Z]\\.\\s)?(?:(?:D’|D'|De |Di |Della |Del |La |Lo )(?=[A-Za-zÀ-ÿ]))?" +
  "[A-Za-zÀ-ÿ][a-zà-ÿ’'\\-]+(?:[\\s\\-][A-ZÀ-Ý][a-zà-ÿ’'\\-]+)?(?:\\s\\(a cura di\\))?$" +
  "|^Aa\\. ?Vv\\.$");

const RE_ANNO = /\b(1[89][0-9]{2}|20[0-9]{2})\b/g;
const RE_CURATELA = /\((a cura di|diretta da|coordinato da)\)/;
const RE_SEPARATORE = /,\s+/;
const RE_STRUTTURA = /^(vol\.|fasc\.|n\.\s?[0-9]|ed\.|edd\.|Parte|Relazione|Tomo|[IVXL]+$|[0-9])/i;
const RE_DATA = new RegExp("^[0-9]{1,2} (" + MESI + ")");
const RE_VOCE_SUB = /^(voce|sub)\b/i;
const RE_INIZIO_FONTE = /^(in|nota a|disponibile su)\b/;
const RE_MINUSCOLA = /^[a-zà-ÿ]/;
const RE_IN = /^in\b/;

const RE_SOLO_IN = /^in\.?\s+(.*)$/s;
const RE_TITOLO_IN = /^(.*?),\s+(?:(nota a .*?),\s+)?in\.?\s+(.*)$/s;
const RE_WEB = /^(.*?),\s+disponibile su\s+(.*)$/s;
const RE_WEB_INIZIO = /^(.*?),\s+disponibile su\s/s;

// "fasc. II, 2024" in fondo alle voci di rivista
const RE_FASCICOLO_ANNO = /, ((?:fasc\. |vol\. )?[IVX]+), ((?:19|20)[0-9]{2})\.?$/;

const RE_INIZIALE = /^[A-Z]\.\s+/;
const RE_CARATTERI_CHIAVE = /[’'\s\-.]/g;

export function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
}

function anni(t) {
  return [...t.matchAll(RE_ANNO)].map(m => m[0]);
}

// ---------- funzioni di supporto ----------

const isYear = p => /\b(1[89][0-9]{2}|20[0-9]{2})\b/.test(p);
const isStruct = p => RE_STRUTTURA.test(p) || RE_DATA.test(p);
export const editoreDi = p => EDITORI.find(e => p === e || p.startsWith(e + " ")) || null;
const isLibro = p => CITTA.includes(p) || editoreDi(p) !== null;
export const dividi = s => s.split(RE_SEPARATORE);

function pulisci(raw) {
  const t = (raw || "").trim();
  return t.endsWith(".") ? t.slice(0, -1) : t;
}

// Parte iniziale fino al primo elemento "di libro" (città, editore), strutturale o anno
function testa(parts) {
  const s = parts.findIndex(p => isLibro(p) || isStruct(p) || isYear(p));
  const n = s < 0 ? parts.length : Math.max(s, 1);
  return parts.slice(0, n).join(", ");
}

// Toglie gli autori iniziali da segs (che viene modificato) e li restituisce.
// Un coautore non è valido se il segmento successivo inizia in minuscolo
// (es. "Gambardella, Peculato, abuso d'ufficio…": "Peculato" fa parte del titolo)
export function stripAutori(segs) {
  const a = [];
  while (segs.length > 1 && RE_AUTORE.test(segs[0]) && !RE_VOCE_SUB.test(segs[0])) {
    if (a.length > 0 && (RE_INIZIO_FONTE.test(segs[1]) || RE_MINUSCOLA.test(segs[1]))) break;
    a.push(segs.shift());
    if (RE_IN.test(segs[0])) break;
  }
  return a;
}

// ---------- estrazione dei campi ----------

export function parse(raw) {
  const t = pulisci(raw);
  const o = { testo: (raw || "").trim(), autore: "", titolo: "", fonte: "", editore: "", anno: "", citta: "", nota: "", tipo: "" };

  const ys = anni(t);
  if (ys.length) o.anno = ys[ys.length - 1];

  const seg = dividi(t);
  o.autore = stripAutori(seg).join(", ");

  const rest = seg.join(", ");
  let tail = "";
  let m;

  if ((m = RE_SOLO_IN.exec(rest))) tail = m[1];
  else if ((m = RE_TITOLO_IN.exec(rest))) {
    o.titolo = m[1];
    o.nota = m[2] || "";
    tail = m[3];
  } else if ((m = RE_WEB.exec(rest))) {
    o.titolo = m[1];
    o.fonte = dividi(m[2])[0];
    o.tipo = "web";
  } else tail = rest;

  const parts = tail.length > 0 ? dividi(tail).map(s => s.trim()) : [];

  const ci = parts.findIndex(p => CITTA.includes(p));
  if (ci >= 0) o.citta = parts[ci];

  const ei = parts.findIndex(p => editoreDi(p) !== null);
  if (ei >= 0) o.editore = editoreDi(parts[ei]);
  else if (ci >= 0) {
    // editore non in elenco: subito dopo la città, oppure subito prima
    let j = -1;
    if (ci + 1 < parts.length && !isYear(parts[ci + 1]) && !isStruct(parts[ci + 1])) j = ci + 1;
    else if (ci > 0 && !isStruct(parts[ci - 1]) && !/[0-9]/.test(parts[ci - 1])) j = ci - 1;
    if (j >= 0) o.editore = parts[j];
  }

  if (o.titolo.length === 0 && !RE_IN.test(rest)) {
    o.titolo = testa(parts);
    o.tipo = "monografia";
  } else if (o.fonte.length === 0) o.fonte = testa(parts);

  if (o.tipo.length === 0)
    o.tipo = (parts.some(isLibro) || RE_CURATELA.test(tail)) ? "collettanea" : "rivista";

  return o;
}

// ---------- corsivi ----------

// Parti della voce da mettere in corsivo, secondo il tipo:
// monografia → titolo; rivista → titolo e rivista; collettanea → saggio e opera; web → titolo
export function corsivi(raw) {
  const seg = dividi(pulisci(raw));
  stripAutori(seg);

  const rest = seg.join(", ");
  let titolo = "", fonte = "";
  let m;

  if ((m = RE_WEB_INIZIO.exec(rest))) return { tipo: "web", parti: [m[1]] };

  if ((m = RE_SOLO_IN.exec(rest))) fonte = m[1];
  else if ((m = /^(.*?),\s+(?:nota a .*?,\s+)?in\.?\s+(.*)$/s.exec(rest))) {
    titolo = m[1];
    fonte = m[2];
  } else return { tipo: "monografia", parti: [testa(dividi(rest))] };

  const parts = dividi(fonte).map(s => s.trim());

  if (!parts.some(isLibro) && !RE_CURATELA.test(fonte))
    return { tipo: "rivista", parti: [titolo, testa(parts)].filter(s => s.length > 0) };

  let op;
  const cu = parts.findIndex(p => RE_CURATELA.test(p));
  if (cu >= 0) op = parts.slice(cu + 1);
  else {
    op = [...parts];
    stripAutori(op);
  }

  return { tipo: "collettanea", parti: [titolo, testa(op)].filter(s => s.length > 0) };
}

// ---------- nomi propri degli autori ----------

// Nomi propri: distinguono "Paolo Frosina" (nome e cognome) da "Cerulli Irelli" (cognome doppio)
const ELENCO_NOMI =
  "Adriano Agostino Alberto Alessandro Alessio Alfonso Alfredo Amedeo Andrea Angelo Antonio Arturo " +
  "Augusto Aurelio Bartolomeo Benedetto Bernardo Bruno Carlo Cesare Claudio Corrado Cristiano Cristian " +
  "Daniele Dario Davide Domenico Edoardo Emanuele Emilio Enrico Enzo Ernesto Ettore Eugenio " +
  "Fabio Fabrizio Federico Felice Ferdinando Filippo Flavio Francesco Franco Gabriele Gaetano Giacomo " +
  "Giampaolo Giancarlo Gianfranco Gianluca Gianluigi Gianmarco Gian Giorgio Giovanni Girolamo Giulio Giuliano " +
  "Giuseppe Gregorio Guglielmo Guido Ignazio Italo Jacopo Iacopo Lamberto Leonardo Lorenzo Luca " +
  "Luciano Lucio Luigi Manfredi Marcello Marco Mario Massimiliano Massimo Matteo Mattia Maurizio " +
  "Mauro Michele Nicola Nicolò Niccolò Oreste Orazio Ottavio Pasquale Paolo Pier Piergiorgio " +
  "Pierluigi Piero Pietro Raffaele Renato Renzo Riccardo Roberto Rocco Rodolfo Romano Rosario " +
  "Salvatore Sandro Santi Santo Saverio Sergio Silvio Simone Stefano Teodoro Tiziano Tommaso " +
  "Ugo Umberto Valerio Vincenzo Vittorio Adele Agnese Alessandra Alice Angela Anna Annalisa " +
  "Antonella Barbara Beatrice Benedetta Carla Carlotta Caterina Cecilia Chiara Claudia Cristina Daniela " +
  "Donatella Elena Eleonora Elisa Elisabetta Emanuela Emma Federica Francesca Gabriella Giada Giorgia " +
  "Giovanna Giulia Giuliana Ilaria Irene Laura Letizia Lucia Luisa Marcella Margherita Maria " +
  "Marina Marta Martina Michela Monica Nadia Paola Patrizia Raffaella Rita Roberta Rosa " +
  "Rossella Sabrina Sara Serena Silvia Simona Sofia Stefania Teresa Valentina Valeria Vanessa " +
  "Veronica Viviana ";

export const NOMI_PROPRI = new Set(ELENCO_NOMI.split(" ").filter(Boolean));

// Nomi propri che sono anche cognomi frequenti: "Fiandaca, Romano" può essere una coppia di autori
export const NOMI_AMBIGUI = new Set([
  "Romano", "Bruno", "Santi", "Santo", "Felice", "Sandro", "Marino", "Fiore", "Gentile", "Rocco", "Valerio"
]);

// "Rossi, Mario": il secondo potrebbe essere il nome proprio del primo. Di solito però due nomi
// separati da virgola sono due autori ("Canosa, Federico"): si propone soltanto, senza correggere
export function nomeDopoCognome(cognome, nome) {
  return NOMI_PROPRI.has(nome) && !NOMI_PROPRI.has(cognome) && !NOMI_AMBIGUI.has(nome);
}

export const RE_NOME_COGNOME =
  /^([A-ZÀ-Ý][a-zà-ÿ]+)\s((?:D’|D'|De |Di |Della |Del )?[A-ZÀ-Ý][a-zà-ÿ’'\-]+(?:\s[A-ZÀ-Ý][a-zà-ÿ’'\-]+)?)$/;

// "Paolo Frosina" → "P. Frosina"; null se non è nome + cognome
export function conIniziale(autore) {
  const m = RE_NOME_COGNOME.exec(autore || "");
  if (!m || !NOMI_PROPRI.has(m[1])) return null;
  return m[1][0] + ". " + m[2];
}

// ---------- maiuscoletto ----------

const RE_SUFFISSO_CURATELA = /\s*\((a cura di|diretta da|coordinato da)\)\s*$/;

// Nomi da comporre in MAIUSCOLETTO (istruzioni della cattedra): l'autore o gli autori
// all'inizio della voce e, nei saggi in volume, il curatore ("in F. Galgano (a cura di)")
export function maiuscoletto(raw) {
  const nomi = [];
  const seg = dividi(pulisci(raw));

  for (const a of stripAutori(seg)) {
    const n = a.replace(RE_SUFFISSO_CURATELA, "").trim();
    if (n.length > 0) nomi.push(n);
  }

  // curatore dopo "in": i segmenti fino a quello con "(a cura di)"
  const rest = seg.join(", ");
  const m = /(?:^|,\s+)in\.?\s+(.*)$/s.exec(rest);
  if (m) {
    const parts = dividi(m[1]);
    const cu = parts.findIndex(x => RE_CURATELA.test(x));
    if (cu >= 0 && cu <= 3)
      for (let i = 0; i <= cu; i++) {
        const n = parts[i].replace(RE_SUFFISSO_CURATELA, "").trim();
        if (n.length > 0 && RE_AUTORE.test(n)) nomi.push(n);
      }
  }

  return nomi;
}

// ---------- ordine dei campi ----------

// Sostituzioni da fare nel testo della voce [vecchio, nuovo] (sull'ultima occorrenza),
// secondo le istruzioni della cattedra (Editore, seguito da virgola – Luogo):
// "Milano, Giuffrè" e "Giuffrè Milano" → "Giuffrè, Milano"; nelle riviste "fasc. II, 2024" → "2024, fasc. II";
// autori con il nome per esteso → iniziale puntata ("Paolo Frosina" → "P. Frosina")
export function sostituzioniOrdineCampi(t) {
  const jobs = [];

  for (const a of stripAutori(dividi(pulisci(t)))) {
    const nome = a.replace(RE_SUFFISSO_CURATELA, "").trim();
    const ini = conIniziale(nome);
    if (ini !== null && t.includes(nome + ",")) jobs.push([nome + ",", ini + ","]);
    else if (ini !== null && t.includes(nome + " (")) jobs.push([nome + " (", ini + " ("]);
  }

  let ed = null;
  for (const e of [...EDITORI].sort((x, y) => y.length - x.length)) {
    // "Città, Editore" (anche seguito da altro: "Milano, Giuffrè, 2008")
    let c = CITTA.find(x => new RegExp("(?<![\\p{L}\\p{N}_’'])" + escapeRe(x + ", " + e) + "(?![\\p{L}\\p{N}_’'])", "u").test(t));
    if (c) {
      jobs.push([c + ", " + e, e + ", " + c]);
      ed = e;
      break;
    }

    // "Editore Città" senza virgola ("CEDAM Padova")
    c = CITTA.find(x => t.includes(e + " " + x));
    if (c) {
      jobs.push([e + " " + c, e + ", " + c]);
      ed = e;
      break;
    }

    if (t.includes(e + ", ") && CITTA.some(x => t.includes(e + ", " + x))) {
      ed = e;   // già nell'ordine giusto
      break;
    }
  }

  const m = RE_FASCICOLO_ANNO.exec(t);
  if (m && t.includes(", in ") && ed === null)
    jobs.push([m[1] + ", " + m[2], m[2] + ", " + m[1]]);

  return jobs;
}

// ---------- ordinamento ----------

function normalizzaChiave(s) {
  let r = s.normalize("NFD").replace(/\p{Mn}/gu, "").toLowerCase();
  r = r.replace(new RegExp(RE_CURATELA.source, "g"), "");
  return r.replace(RE_CARATTERI_CHIAVE, "");
}

// Cognome per l'ordinamento: senza iniziale ("B. Romano") né nome per esteso ("Paolo Frosina")
export function soloCognome(autore) {
  const a = (autore || "").replace(RE_INIZIALE, "");
  const c = conIniziale(a);
  return c !== null ? c.slice(3) : a;
}

// Chiave "cognome1|cognome2#anno": le opere di un solo autore precedono quelle con coautori
export function chiaveOrdinamento(t) {
  const seg = dividi(t);
  let a = stripAutori(seg);
  if (a.length === 0) a = [seg[0]];

  const aut = a.map(x => normalizzaChiave(soloCognome(x))).join("|");
  const ys = anni(t);
  return aut + "#" + (ys.length ? ys[ys.length - 1] : "9999");
}

// Confronto ordinale (come string.CompareOrdinal in C#)
export const ordinale = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

// Indici delle voci che restano ferme: la sottosequenza già in ordine più lunga
export function vociInOrdine(keys) {
  const n = keys.length;
  const len = new Array(n).fill(1);
  const prev = new Array(n).fill(-1);

  for (let i = 0; i < n; i++)
    for (let j = 0; j < i; j++)
      if (ordinale(keys[j], keys[i]) <= 0 && len[j] + 1 > len[i]) {
        len[i] = len[j] + 1;
        prev[i] = j;
      }

  const keep = new Set();
  if (n === 0) return keep;

  const max = Math.max(...len);
  let b = len.lastIndexOf(max);
  for (; b >= 0; b = prev[b]) keep.add(b);
  return keep;
}
