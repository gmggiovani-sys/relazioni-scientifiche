// Formattazione delle citazioni nelle note secondo le istruzioni della cattedra:
// autore e curatore in MAIUSCOLETTO; titolo, rivista e volume in corsivo; cit. e op. cit. in tondo;
// ivi, ibidem in corsivo; ID. in maiuscoletto. Il testo non cambia; le citazioni testuali tra «» non si toccano.
// Porting di CitazioniParser.cs (versione Windows).

import * as B from "./bibliografia.js";
import { nomiRiviste } from "./giurisprudenza.js";

export const CORSIVO = "corsivo", MAIUSCOLETTO = "maiuscoletto", TONDO = "tondo";   // tondo: senza corsivo

const MIN_TITOLO_ABBREVIATO = 12;
const isLetter = c => /\p{L}/u.test(c);
const isLetterOrDigit = c => /[\p{L}\p{N}]/u.test(c);

// Opere della bibliografia (dalle voci con stile Biblio)
export function opere(vociBibliografia) {
  const res = [];
  for (const raw of vociBibliografia) {
    const v = B.parse(raw);
    if (v.autore.length === 0 || v.titolo.length === 0) continue;

    const it = B.corsivi(raw).parti;
    const autori = v.autore.split(", ");

    res.push({
      autore: v.autore.replace(/\s*\((a cura di|diretta da|coordinato da)\)/g, ""),
      titolo: v.titolo,
      corsivi: it.filter(x => x !== v.titolo),
      curatori: B.maiuscoletto(raw).filter(n => !autori.includes(n))
    });
  }
  return res;
}

// Rinvii: "cit." e "op. cit." in tondo (scelta della tesi), ivi e ibidem in corsivo, ID. in maiuscoletto
const RE_CIT = /(?<=,\s?)(?:op\.\s?cit\.|cit\.)/g;
const RE_RINVIO = /(?:(?<=^\s*)|(?<=[;.]\s)|(?<=,\s))(?:Ibidem|ibidem|Ibid\.|ibid\.|Ivi|ivi)(?=[\s,.;:]|$)/g;
const RE_ID = /(?:(?<=^\s*)|(?<=[;.]\s))(?:Id\.|ID\.|Idem|IDEM)(?=[\s,])/g;

const MESI = "gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre";

// Sentenze: "Cass. pen., …, in Dejure" → rivista o banca dati in corsivo
const RE_DECISIONE = new RegExp(
  "\\b(?:Cass\\.|Corte cost\\.|Cons\\.\\s?St(?:ato)?\\.?|Trib\\.|App\\.|Pret\\.|T\\.A\\.R\\.|Corte conti|Corte giust\\.|Corte ass\\.|Corte EDU|C\\.\\s?Cost\\.)" +
  "[^;«»]*?\\b[0-9]{1,2}°?\\s(?:" + MESI + ")\\s[0-9]{4}[^;«»]*?,\\s(in)\\s", "dg");

const RE_DATA_ESTESA = new RegExp("\\b[0-9]{1,2}°?\\s(?:" + MESI + ")\\s[0-9]{4}");

const RE_DOPO_AUTORE = /^(?:\s\((?:a cura di|diretta da|coordinato da)\))?,\s/;

const RE_FINE_TRATTO = /;|\.\s+(?=[A-ZÀ-Ý«(])|$/g;

// Tratti da formattare in una nota: [{ idx, len, tipo, regola }].
// nonTrovate (facoltativo): citazioni con "cit." che non corrispondono a nessuna opera della bibliografia
export function analizza(testo, elencoOpere, nonTrovate = null) {
  const res = [];
  if (!testo) return res;

  const citazione = zoneCitazione(testo);
  const norm = normalizza(testo);
  const riconosciute = [];

  // ---- dottrina: "Autore, Titolo" di un'opera della bibliografia
  // il cognome si cerca senza distinguere maiuscole e minuscole ("Di martino" in bibliografia,
  // "Di Martino" o "DI MARTINO" nella nota): stessa lunghezza, quindi gli indici restano validi
  const normMin = minuscolo(norm);

  const gruppi = new Map();
  for (const o of elencoOpere) {
    const k = minuscolo(o.autore);
    if (!gruppi.has(k)) gruppi.set(k, []);
    gruppi.get(k).push(o);
  }

  for (const [chiave, gruppo] of gruppi) {
    const a = normalizza(chiave);
    let da = 0;

    while ((da = normMin.indexOf(a, da)) >= 0) {
      const inizio = da;
      da += a.length;

      // l'autore deve iniziare una parola
      if (inizio > 0 && isLetter(norm[inizio - 1])) continue;
      if (citazione[inizio]) continue;

      // dopo il nome: ", " oppure " (a cura di), "
      const sep = RE_DOPO_AUTORE.exec(norm.slice(inizio + a.length));
      if (!sep) continue;

      const dopo = inizio + a.length + sep[0].length;

      // tra le opere dell'autore, quella col titolo che combacia più a lungo
      let scelta = null, lunghezza = 0;
      for (const o of gruppo) {
        const l = titoloCombacia(norm, dopo, normalizza(o.titolo));
        if (l > lunghezza) { lunghezza = l; scelta = o; }
      }
      if (scelta === null) continue;

      aggiungiAutori(res, inizio, chiave);
      res.push({ idx: dopo, len: lunghezza, tipo: CORSIVO, regola: "titolo" });

      // resto della stessa citazione: rivista o volume in corsivo, curatori in maiuscoletto
      const fineTitolo = dopo + lunghezza;
      const fine = fineCitazione(norm, fineTitolo);

      let cerca = fineTitolo;
      for (const c of scelta.curatori)
        cerca = aggiungi(res, norm, normalizza(c), cerca, fine, MAIUSCOLETTO, "curatore", c.length);

      cerca = fineTitolo;
      for (const c of scelta.corsivi)
        cerca = aggiungi(res, norm, normalizza(c), cerca, fine, CORSIVO, "rivista o volume", c.length);

      // curatori e volume scritti nella nota ("…, in Bondi, Di Martino, Fornasari (a cura di), Reati …"),
      // anche quando la voce di bibliografia non li riporta nello stesso modo
      volumeCurato(res, norm, fineTitolo, fine);

      riconosciute.push([inizio, fine]);
    }
  }

  // ---- giurisprudenza: rivista o banca dati dopo "in"
  for (const m of testo.matchAll(RE_DECISIONE)) {
    const idxIn = m.indices[1][0] + 3;
    if (citazione[m.index]) continue;

    const fine = fineCitazione(norm, idxIn);
    const fonte = testo.slice(idxIn, fine);

    let cerca = idxIn;
    for (const nome of nomiRiviste(fonte)) {
      const i = testo.indexOf(nome, cerca);
      if (i < 0 || i >= fine) continue;
      res.push({ idx: i, len: nome.length, tipo: CORSIVO, regola: "rivista (giurisprudenza)" });
      cerca = i + nome.length;
    }
  }

  // ---- rinvii
  for (const m of testo.matchAll(RE_CIT))
    if (!citazione[m.index])
      res.push({ idx: m.index, len: m[0].length, tipo: TONDO, regola: m[0].startsWith("op") ? "op. cit. (tondo)" : "cit. (tondo)" });

  for (const m of testo.matchAll(RE_RINVIO))
    if (!citazione[m.index]) res.push({ idx: m.index, len: m[0].length, tipo: CORSIVO, regola: "ivi / ibidem" });

  for (const m of testo.matchAll(RE_ID))
    if (!citazione[m.index]) res.push({ idx: m.index, len: m[0].length, tipo: MAIUSCOLETTO, regola: "ID." });

  // ---- "Autore, Titolo, cit." senza corrispondenza in bibliografia
  const autoriNoti = [...new Set(elencoOpere.flatMap(o => o.autore.split(", ")))]
    .sort((x, y) => y.length - x.length);

  for (const m of testo.matchAll(RE_CIT)) {
    if (citazione[m.index]) continue;
    if (riconosciute.some(c => m.index >= c[0] && m.index <= c[1])) continue;

    const ini = inizioCitazione(testo, m.index);
    const frammento = testo.slice(ini, m.index);

    // sentenze e rinvii non sono dottrina
    if (RE_DATA_ESTESA.test(frammento) || /^(?:Ibidem|Ivi|ivi|Id\.|ID\.)\b/.test(frammento.trimStart())) continue;
    if (!/[A-ZÀ-Ý][a-zà-ÿ’'\-]+,\s[^,]/.test(frammento)) continue;

    // autore presente in bibliografia: si formatta comunque e si segnala il titolo diverso
    let posA = -1, autore = null;
    for (const au of autoriNoti) {
      const ma = new RegExp("(?<![\\p{L}\\p{N}_’'])" + B.escapeRe(au) + ",\\s", "u").exec(frammento);
      if (ma && (posA < 0 || ma.index < posA)) { posA = ma.index; autore = au; }
    }

    const mostra = testo.slice(ini, m.index + m[0].length).trim();

    if (autore !== null) {
      const a0 = ini + posA, t0 = a0 + autore.length + 2;
      let t1 = m.index;
      while (t1 > t0 && (testo[t1 - 1] === " " || testo[t1 - 1] === ",")) t1--;

      res.push({ idx: a0, len: autore.length, tipo: MAIUSCOLETTO, regola: "autore" });
      if (t1 > t0) res.push({ idx: t0, len: t1 - t0, tipo: CORSIVO, regola: "titolo (non in bibliografia)" });
      if (nonTrovate) nonTrovate.push("titolo diverso dalla bibliografia: " + mostra);
    } else if (nonTrovate) nonTrovate.push("autore non in bibliografia: " + mostra);
  }

  return pulisci(res);
}

// Minuscole carattere per carattere: la lunghezza non cambia (gli indici restano validi)
function minuscolo(s) {
  let r = "";
  for (const c of s) {
    const l = c.toLowerCase();
    r += l.length === c.length ? l : c;
  }
  return r;
}

const compatta = s => [...s].filter(isLetterOrDigit).map(c => c.toLowerCase()).join("");

// Lunghezza del titolo riconosciuto in norm da pos: titolo completo, oppure le sue prime parole
// seguite da "cit.", "in", editore, città o numero (citazioni successive abbreviate).
// Il confronto ignora maiuscole, punteggiatura e spazi ("Introduzione. Tra" = "Introduzione: tra").
function titoloCombacia(norm, pos, titolo) {
  const tc = compatta(titolo);
  if (tc.length === 0) return 0;

  let i = pos, k = 0, fineBuona = pos;
  while (i < norm.length && k < tc.length) {
    const ch = norm[i];
    if (!isLetterOrDigit(ch)) { i++; continue; }
    if (ch.toLowerCase() !== tc[k]) break;
    k++;
    i++;
    fineBuona = i;
  }

  if (k === tc.length) {
    // titolo completo: si includono la parentesi o le virgolette finali, se ci sono
    let l = fineBuona - pos;
    let coda = titolo.length;
    while (coda > 0 && !isLetterOrDigit(titolo[coda - 1])) coda--;
    const fineTitolo = titolo.slice(coda).replace(/[. ]+$/, "");
    if (fineTitolo.length > 0 && norm.slice(fineBuona).startsWith(fineTitolo)) l += fineTitolo.length;
    return l;
  }

  if (k < MIN_TITOLO_ABBREVIATO) return 0;

  // titolo abbreviato: deve finire a fine parola
  let lun = fineBuona - pos;
  if (fineBuona < norm.length && isLetterOrDigit(norm[fineBuona])) {
    while (lun > 0 && isLetterOrDigit(norm[pos + lun - 1])) lun--;
    while (lun > 0 && !isLetterOrDigit(norm[pos + lun - 1])) lun--;
    if (compatta(norm.slice(pos, pos + lun)).length < MIN_TITOLO_ABBREVIATO) return 0;
  }

  const seguito = norm.slice(pos + lun);
  const chiude = /^(?:,\s?(?:op\.\s?)?cit\.|\s?cit\.|,\s(?:in\s|[0-9])|\.?\s*$)/.test(seguito) ||
                 B.EDITORI.concat(B.CITTA).some(x => seguito.startsWith(", " + x));

  return chiude ? lun : 0;
}

const RE_NOME = "(?:[A-Z]\\.\\s)?(?:D’|D'|De |Di |Della |Del )?[A-ZÀ-Ý][a-zà-ÿ’'\\-]+(?:\\s[A-ZÀ-Ý][a-zà-ÿ’'\\-]+)?";
const RE_VOLUME_CURATO = new RegExp(
  ",\\s+in\\s+(" + RE_NOME + "(?:,\\s" + RE_NOME + ")*)\\s\\((?:a cura di|diretta da|coordinato da)\\),\\s", "yd");
const RE_NOME_SINGOLO = new RegExp(RE_NOME, "g");

function volumeCurato(res, norm, da, fine) {
  RE_VOLUME_CURATO.lastIndex = da;
  const m = RE_VOLUME_CURATO.exec(norm);
  if (!m || m.index > fine) return;

  // il titolo del volume può contenere punti ("… amministrazione. Riscrivere l'abuso …"):
  // si arriva fino al ";" o alla fine della nota, e ci si ferma a editore, città o anno
  const pv = norm.indexOf(";", m.index + m[0].length);
  fine = pv >= 0 ? pv : norm.length;

  // curatori: un tratto per nome
  const g = m[1];
  const gIdx = m.indices[1][0];
  for (const n of g.matchAll(RE_NOME_SINGOLO))
    res.push({ idx: gIdx + n.index, len: n[0].length, tipo: MAIUSCOLETTO, regola: "curatore" });

  // titolo del volume: fino a editore, città, anno, volume o pagina
  const inizio = m.index + m[0].length;
  let fineVol = inizio;
  for (const p of norm.slice(inizio, fine).split(", ")) {
    const q = p.trim();
    if (q.length === 0 || B.EDITORI.concat(B.CITTA).some(x => q.startsWith(x)) ||
        /^(?:[0-9]|vol\.|t\.|tomo|[IVXL]+(?:$|\s)|ed\.|(?:op\.\s?)?cit\.)/.test(q)) break;
    fineVol += (fineVol === inizio ? 0 : 2) + p.length;
  }

  if (fineVol > inizio) res.push({ idx: inizio, len: fineVol - inizio, tipo: CORSIVO, regola: "volume" });
}

// Un tratto per ogni cognome ("Cantone, Gatta" → "Cantone" e "Gatta"): la virgola e lo spazio
// in mezzo non contano, altrimenti il maiuscoletto risulterebbe sempre "misto"
function aggiungiAutori(res, inizio, autori) {
  let pos = 0;
  for (const a of autori.split(", ")) {
    if (a.length > 0) res.push({ idx: inizio + pos, len: a.length, tipo: MAIUSCOLETTO, regola: "autore" });
    pos += a.length + 2;
  }
}

function aggiungi(res, norm, cosa, da, fine, tipo, regola, len) {
  if (cosa.length === 0) return da;
  const i = norm.indexOf(cosa, da);
  if (i < 0 || i + cosa.length > fine) return da;
  res.push({ idx: i, len, tipo, regola });
  return i + cosa.length;
}

// Fine della citazione: ";" o fine della frase (punto seguito da maiuscola) o fine della nota
function fineCitazione(t, da) {
  RE_FINE_TRATTO.lastIndex = Math.min(da, t.length);
  const m = RE_FINE_TRATTO.exec(t);
  return m ? m.index : t.length;
}

function inizioCitazione(t, prima) {
  let i = Math.max(t.lastIndexOf(";", Math.max(prima - 1, 0)), 0);
  const m = /^.*(?:\.\s+(?=[A-ZÀ-Ý]))/s.exec(t.slice(i, prima));
  if (m) i += m[0].length;
  if (i < t.length && t[i] === ";") i++;
  return i;
}

// Stesso testo con apostrofi e virgolette uniformi (stessa lunghezza: gli indici restano validi)
export function normalizza(s) {
  return s.replace(/’/g, "'").replace(/‘/g, "'").replace(/“/g, "\"").replace(/”/g, "\"").replace(/ /g, " ");
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

// Toglie i doppioni (stesso tratto e stesso tipo) e i tratti vuoti
function pulisci(res) {
  const visti = new Set();
  return res.filter(f => {
    if (f.len <= 0) return false;
    const k = f.idx + "|" + f.len + "|" + f.tipo;
    if (visti.has(k)) return false;
    visti.add(k);
    return true;
  }).sort((a, b) => a.idx - b.idx);
}
