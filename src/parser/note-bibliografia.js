// Ricava dalle note le citazioni complete (autore, titolo, anno e editore/città o "in Rivista")
// e propone le voci che mancano in bibliografia. Le citazioni brevi ("cit.") non bastano.
// Porting di NoteBibliografiaParser.cs (versione Windows).

import * as B from "./bibliografia.js";

const RE_ANNO = /\b(1[89][0-9]{2}|20[0-9]{2})\b/;
const RE_CIT = /\b(?:op\.\s?)?cit\.|\bibidem\b|\bIbidem\b|\b[Ii]vi\b/;
const RE_DECISIONE =
  /\b(?:Cass\.|Corte cost\.|Cons\.\s?St|Trib\.|App\.|Pret\.|T\.A\.R\.|Corte conti|Corte giust\.|Corte EDU|C\.\s?Cost\.)/;

// solo nomi separati da virgole ("Bondi, ", "B. Romano, Marandola, ")
const RE_SOLO_NOMI =
  /^\s*(?:(?:[A-Z]\.\s)?(?:D’|D'|De |Di |Della |Del )?[A-ZÀ-Ý][a-zà-ÿ’'\-]+(?:\s[A-ZÀ-Ý][a-zà-ÿ’'\-]+)?,\s)*(?:[A-Z]\.\s|D’|D'|De |Di |Della |Del )?$/;

const ATTACCHI = new Set([
  "Secondo", "Così", "Anche", "Cfr", "Vedi", "Come", "Contra", "Inoltre", "Ancora", "Sempre",
  "Già", "Ora", "Più", "Si", "In", "Per", "Sul", "Sulla", "Sui", "Nel", "Nella", "Nello", "Tra",
  "Fra", "Da", "Ex", "Al", "Alla", "Infatti", "Peraltro", "Invece", "Cosi", "V", "Sebbene", "Benché",
  "Pur", "Proprio", "Emblematica", "Emblematico", "Analogamente", "Diversamente", "Sullo", "Mentre", "Ove"
]);

const PARTICELLE = new Set([
  "de", "di", "del", "della", "d’", "d'", "da", "van", "von", "(a", "cura", "di)", "(diretta", "da)", "(coordinato"
]);

// note: [[dove, testo], …] nell'ordine del documento.
// Restituisce [{ voce, autore, titolo, anno, citazione, chiave, dove }] in ordine alfabetico
export function proposte(note, bibliografia) {
  const esistenti = bibliografia.map(B.parse);
  const autoriNoti = new Set(esistenti.flatMap(e => e.autore.split(", ")));
  const res = [];

  for (const [dove, testo] of note)
    for (const tratto of tratti(testo))
      for (const v of daTratto(tratto, autoriNoti)) {
        v.dove = dove;
        if (esistenti.some(e => stessaOpera(e.autore, e.titolo, v.autore, v.titolo))) continue;
        if (res.some(p => stessaOpera(p.autore, p.titolo, v.autore, v.titolo))) continue;
        res.push(v);
      }

  return res.sort((a, b) => B.ordinale(a.chiave, b.chiave));
}

// Citazioni di una nota: tratti separati da ";" (non dal punto: i titoli contengono punti
// e le iniziali dei nomi pure); le citazioni testuali tra «» si tolgono
function tratti(nota) {
  return togliCitazioniTestuali(nota || "").split(";").map(x => x.trim()).filter(x => x.length > 0);
}

function togliCitazioniTestuali(t) {
  let r = "", basse = 0;
  for (const c of t) {
    if (c === "«") basse++;
    r += basse > 0 ? " " : c;
    if (c === "»" && basse > 0) basse--;
  }
  return r;
}

// Voci complete in un tratto di nota (anche più di una)
function daTratto(tratto, autoriNoti) {
  const res = [];
  if (tratto.length < 15 || RE_DECISIONE.test(tratto)) return res;

  let minimo = 0;

  // l'autore può essere preceduto da un attacco ("Cfr.", "In tal senso, v.", "osserva …"):
  // si prova da ogni parola con l'iniziale maiuscola
  for (const m of tratto.matchAll(/(?<![\p{L}\p{N}_’'.])[A-ZÀ-Ý]/gu)) {
    if (m.index < minimo) continue;

    // dopo "in" (senza un anno in mezzo) si è dentro il volume che contiene un saggio:
    // non è un'opera a sé ("…, in B. Romano (a cura di), Il “nuovo” abuso d'ufficio, …")
    const prima = tratto.slice(0, m.index);
    if (/,\s?in$/.test(prima.trimEnd())) continue;
    const ultimoIn = prima.lastIndexOf(", in ");
    if (ultimoIn >= 0 && RE_SOLO_NOMI.test(prima.slice(ultimoIn + 5))) continue;

    // "Id.", "Ead.": l'autore è quello della citazione precedente
    if (/^(?:Id|ID|Ead|EAD)\./.test(tratto.slice(m.index))) continue;

    const s = tratto.slice(m.index).trim();
    const anno = RE_ANNO.exec(s);
    if (!anno) break;

    // fino all'anno compreso; le pagine citate (", 45 ss.", ", p. 12") non vanno in bibliografia
    let voce = s.slice(0, anno.index + anno[0].length).replace(/[, ]+$/, "");
    if (RE_CIT.test(voce)) continue;

    let v = B.parse(voce + ".");
    if (v.autore.length === 0 || v.titolo.length < 4) continue;
    if (v.titolo.split(" ").length > 40) continue;

    // il titolo inizia con la maiuscola (o con virgolette); il nome non ha parole minuscole
    const t0 = v.titolo[0];
    if (!(isUpper(t0) || /\p{N}/u.test(t0) || "“\"«‘'".includes(t0))) continue;

    const parole = v.autore.split(/[ ,]/).filter(Boolean);
    if (parole.some(w => isLower(w[0]) && !PARTICELLE.has(w.toLowerCase()))) continue;
    if (ATTACCHI.has(parole[0].replace(/\.+$/, ""))) continue;

    // "Sebbene Contento": se il secondo nome da solo è un autore noto, si riparte da lì
    if (parole.length >= 2 && v.autore.startsWith(parole[0] + " " + parole[1]) &&
        !parole[0].endsWith(".") && !PARTICELLE.has(parole[0].toLowerCase()) && autoriNoti.has(parole[1]))
      continue;

    const completa = v.editore.length > 0 || v.citta.length > 0 ||
                     (v.tipo === "rivista" && v.fonte.length > 0) || v.tipo === "collettanea" || v.tipo === "web";
    if (!completa) continue;

    voce = voce.replace(/\s+/g, " ") + ".";

    // autori con il nome per esteso: solo l'iniziale puntata ("Paolo Frosina" → "P. Frosina")
    for (const [vecchio, nuovo] of B.sostituzioniOrdineCampi(voce).filter(j => /^[A-ZÀ-Ý]\. /.test(j[1]))) {
      const i = voce.indexOf(vecchio);
      if (i >= 0) voce = voce.slice(0, i) + nuovo + voce.slice(i + vecchio.length);
    }

    v = B.parse(voce);
    res.push({
      voce, autore: v.autore, titolo: v.titolo, anno: v.anno,
      citazione: tratto, chiave: B.chiaveOrdinamento(voce), dove: ""
    });

    // la prossima citazione nello stesso tratto inizia dopo questa
    minimo = m.index + s.indexOf(anno[0]) + anno[0].length;
  }

  return res;
}

const isUpper = c => c !== c.toLowerCase() && c === c.toUpperCase();
const isLower = c => c !== c.toUpperCase() && c === c.toLowerCase();

// Stessa opera: stesso autore e titoli che coincidono (ignorando maiuscole, punteggiatura e spazi)
// almeno nella parte iniziale
export function stessaOpera(autoreA, titoloA, autoreB, titoloB) {
  // solo i cognomi: "Paolo Frosina", "P. Frosina" e "Frosina" sono lo stesso autore
  if (cognomi(autoreA) !== cognomi(autoreB)) return false;

  const a = compatta(titoloA), b = compatta(titoloB);
  if (a.length === 0 || b.length === 0) return false;

  const n = Math.min(a.length, b.length, 25);
  return a.slice(0, n) === b.slice(0, n);
}

function cognomi(autori) {
  // senza separatori: "De Luca, Segreto" e "De Luca-Segreto" coincidono
  return (autori || "").split(", ").map(x => compatta(B.soloCognome(x))).join("");
}

function compatta(s) {
  return [...(s || "").normalize("NFD")].filter(c => /[\p{L}\p{N}]/u.test(c)).map(c => c.toLowerCase()).join("");
}
