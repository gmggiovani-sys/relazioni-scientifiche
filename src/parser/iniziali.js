// Iniziali dei nomi degli autori (istruzioni della cattedra: "Iniziale puntata del nome dell'Autore, Cognome per esteso").
// Le iniziali stanno in un elenco di testo modificabile ("Cognome = F."), con eventuali righe
// "Cognome | inizio del titolo = L." per autori diversi con lo stesso cognome.
// Proposta: { nome, idx, iniziale (null se sconosciuta), dubbia, fonte, sostituto, autori, inizioAutori }
// (con sostituto = " e ": la virgola tra gli ultimi due autori, nome = ", ").
// Porting di InizialiAutori.cs (versione Windows).

import * as B from "./bibliografia.js";
import { analizza, MAIUSCOLETTO } from "./citazioni.js";

const RE_HA_INIZIALE = new RegExp("^" + B.INIZIALI + "\\s");
const RE_INIZIA_CON_INIZIALE = new RegExp("^" + B.INIZIALI);
const RE_RIGA = new RegExp("^\\s*([^=|#]+?)\\s*(?:\\|\\s*([^=]+?)\\s*)?=\\s*(" + B.INIZIALI + ")\\s*(\\?)?\\s*$");
const RE_AAVV = /^Aa\.\s?Vv\./;

export const INTESTAZIONE_ELENCO = "# Iniziali dei nomi degli autori (Relazioni Scientifiche → Bibliografia → Iniziali dei Nomi)\n#\n#   Cognome = F.                      iniziale sicura: la voce è già spuntata\n#   Cognome = F. ?                    da verificare: la voce compare senza spunta\n#   Cognome | inizio del titolo = L.  per autori diversi con lo stesso cognome\n#\n# Si può modificare e aggiungere righe: vale alla prossima esecuzione del comando.\n# Le righe che iniziano con # sono commenti.\n\n";

function leggi(testo) {
  const res = [];
  for (const riga of (testo || "").split("\n")) {
    const m = RE_RIGA.exec(riga.replace(/\r$/, ""));
    if (!m) continue;
    res.push({ cognome: m[1].trim(), titolo: m[2] !== undefined ? m[2].trim() : null, iniziale: m[3].replace(/\s+/g, ""), dubbia: m[4] !== undefined });
  }
  return res;
}

// Regole dell'elenco dell'utente, poi quelle predefinite per i cognomi che l'elenco non ha
export function carica(testoUtente) {
  const utente = leggi(testoUtente);
  const conosciuti = new Set(utente.map(r => chiave(r.cognome) + "|" + chiave(r.titolo || "")));
  return utente.concat(leggi(PREDEFINITE).filter(r => !conosciuti.has(chiave(r.cognome) + "|" + chiave(r.titolo || ""))));
}

function chiave(s) {
  return [...(s || "").normalize("NFD")].filter(c => /[\p{L}\p{N}]/u.test(c)).map(c => c.toLowerCase()).join("");
}

const escapeRe = s => s.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");

// ---------- proposte per una voce ----------

// Autori e curatori della voce senza iniziale, con l'iniziale trovata (elenco o note),
// e la "e" tra gli ultimi due autori o curatori
export function proposte(voce, rr, dalleNote) {
  const res = [];
  const titolo = chiave(B.parse(voce).titolo);
  let da = 0;

  for (const nome of B.maiuscoletto(voce)) {
    // il nome come elemento della voce (non dentro un titolo: "la proposta “Castaldo-Naddeo”")
    const re = new RegExp("(?<=^|,\\s|\\se\\s|(?<![\\p{L}\\p{N}_])in\\.?\\s)" + escapeRe(nome) + "(?=,|\\s\\(|\\se\\s)", "gu");
    re.lastIndex = da;
    const pos = re.exec(voce);
    if (!pos) continue;
    const idx = pos.index;
    da = idx + nome.length;

    if (RE_INIZIA_CON_INIZIALE.test(nome) || /^Aa\.\s?Vv\.$/.test(nome)) continue;
    if (NON_PERSONE.has(nome)) continue;

    const k = chiave(nome);
    const r = rr.find(x => chiave(x.cognome) === k && x.titolo !== null && titolo.startsWith(chiave(x.titolo)))
           || rr.find(x => chiave(x.cognome) === k && x.titolo === null);

    const p = { nome, idx, iniziale: null, dubbia: false, fonte: "", sostituto: null };
    if (r) {
      p.iniziale = r.iniziale;
      p.dubbia = r.dubbia;
      p.fonte = r.dubbia ? "elenco (da verificare)" : "elenco";
    } else if (dalleNote && dalleNote.has(k)) {
      p.iniziale = dalleNote.get(k);
      p.dubbia = true;          // ricavata dalle note: da confermare
      p.fonte = "note";
    } else p.fonte = "sconosciuta";

    res.push(p);
  }

  // "e" tra gli ultimi due autori o curatori: sicura se i due nomi sono noti (con l'iniziale o nell'elenco),
  // altrimenti da decidere ("Giovenale, Saturae": il secondo può essere il titolo)
  for (const [pos, x, y] of B.virgoleTraAutori(voce)) {
    const noto = n => RE_HA_INIZIALE.test(n) || rr.some(r => chiave(r.cognome) === chiave(B.soloCognome(n)));
    const sicura = noto(x) && noto(y) && !B.nomeDopoCognome(B.soloCognome(x), y);
    res.push({
      nome: ", ", idx: pos, sostituto: " e ", iniziale: "e",
      autori: x + ", " + y, inizioAutori: pos - x.length,
      dubbia: !sicura,
      fonte: sicura ? "istruzioni: \"e\" tra gli ultimi due" : "\"e\" tra gli ultimi due: verifica che siano autori"
    });
  }

  return res;
}

// Testo con le correzioni scelte ("Antolisei, …" → "F. Antolisei, …"; "Cantone, Gatta" → "Cantone e Gatta")
export function applica(voce, scelte) {
  let t = voce;
  for (const p of scelte.filter(x => x.iniziale !== null).sort((a, b) => b.idx - a.idx))
    t = p.sostituto !== null && p.sostituto !== undefined
      ? t.slice(0, p.idx) + p.sostituto + t.slice(p.idx + p.nome.length)
      : t.slice(0, p.idx) + p.iniziale + " " + t.slice(p.idx);
  return t;
}

// ---------- proposte per le note ----------

// Iniziali dei nomi che hanno già l'iniziale in bibliografia: cognome → Set di iniziali trovate
export function dallaBibliografia(voci) {
  const res = new Map();
  const re = new RegExp("^(" + B.INIZIALI + ")\\s*(.+)$");
  for (const v of voci)
    for (const n of B.maiuscoletto(v)) {
      const m = re.exec(n);
      if (!m) continue;
      const k = chiave(m[2]);
      if (!res.has(k)) res.set(k, new Set());
      res.get(k).add(m[1].replace(/\s+/g, ""));
    }
  return res;
}

const RE_INIZIALE_PRIMA = new RegExp("(?<![\\p{L}\\p{N}_’'.])" + B.INIZIALI + "\\s?$", "u");

// Autori e curatori citati in una nota (riconosciuti come in "Controlla Citazioni") senza iniziale.
// L'iniziale viene dalla bibliografia (se lì c'è già) o dall'elenco; con più autori dello stesso
// cognome in bibliografia resta da scegliere a mano. In più, la "e" tra gli ultimi due.
export function nelleNote(nota, opere, daBiblio, regole) {
  const res = [];
  if (!nota) return res;

  const formati = analizza(nota, opere, null);
  congiunzioneE(nota, formati, res);

  for (const f of formati) {
    if (f.tipo !== MAIUSCOLETTO || (f.regola !== "autore" && f.regola !== "curatore")) continue;

    const nome = nota.slice(f.idx, f.idx + f.len);
    if (RE_INIZIA_CON_INIZIALE.test(nome) || RE_AAVV.test(nome) || NON_PERSONE.has(nome)) continue;
    if (RE_INIZIALE_PRIMA.test(nota.slice(0, f.idx))) continue;   // iniziale già scritta
    if (res.some(x => x.idx === f.idx)) continue;

    const k = chiave(nome);
    const p = { nome, idx: f.idx, iniziale: null, dubbia: false, fonte: "", sostituto: null };
    const b = daBiblio.get(k);
    const r = regole.find(x => chiave(x.cognome) === k && x.titolo === null);

    if (b && b.size === 1) {
      p.iniziale = [...b][0];
      p.fonte = "bibliografia";
    } else if (b && b.size > 1) p.fonte = "omonimi in bibliografia (" + [...b].join(", ") + "): a mano";
    else if (r) {
      p.iniziale = r.iniziale;
      p.dubbia = r.dubbia;
      p.fonte = r.dubbia ? "elenco (da verificare)" : "elenco";
    } else if (regole.some(x => chiave(x.cognome) === k)) p.fonte = "omonimi nell'elenco: a mano";
    else p.fonte = "sconosciuta";

    res.push(p);
  }

  return res;
}

// Due o più autori (o curatori) di fila separati solo da virgole: tra gli ultimi due va "e"
// ("Cantone, Gatta, Abuso…" → "Cantone e Gatta, Abuso…")
function congiunzioneE(nota, formati, res) {
  for (const regola of ["autore", "curatore"]) {
    const nomi = formati.filter(f => f.tipo === MAIUSCOLETTO && f.regola === regola).sort((a, b) => a.idx - b.idx);
    let i = 0;

    while (i < nomi.length) {
      // catena di nomi: tra uno e l'altro ", " oppure " e " (o lo spazio fra iniziale e cognome)
      let j = i;
      const separatori = [];
      let conE = false;

      while (j + 1 < nomi.length) {
        const a = nomi[j].idx + nomi[j].len, b = nomi[j + 1].idx;
        if (b < a) break;
        const tra = nota.slice(a, b);
        if (tra === ", ") separatori.push(a);
        else if (tra === " e ") conE = true;
        else if (tra !== " ") break;
        j++;
      }

      if (separatori.length > 0 && !conE) {
        const s = separatori[separatori.length - 1];
        const inizio = nomi[i].idx, fine = nomi[j].idx + nomi[j].len;
        const x = B.soloCognome(nota.slice(inizio, s).split(", ").pop());
        const y = nota.slice(s + 2, fine);
        res.push({
          nome: ", ", idx: s, sostituto: " e ", autori: nota.slice(inizio, fine), inizioAutori: inizio,
          iniziale: "e",
          dubbia: B.nomeDopoCognome(x, y),   // "Canosa, Federico": forse nome e cognome
          fonte: "istruzioni: \"e\" tra gli ultimi due " + (regola === "autore" ? "autori" : "curatori")
        });
      }

      i = j + 1;
    }
  }
}

// Iniziali usate nelle note per un cognome ("P. Frosina", "Paolo Frosina"): solo se sempre la stessa
export function dalleNote(note, cognomi) {
  const trovate = new Map();
  const tutto = (note || []).join("\n");

  for (const c of [...new Set(cognomi)]) {
    const esc = escapeRe(c);
    const set = new Set();

    // solo in posizione di autore: "P. Frosina, Titolo" (non "vol. I. Gatta" o "Sez. V. Rossi")
    const re1 = new RegExp("(?<=(?:^|[;:(]|\\.)\\s*|,\\s|(?<![\\p{L}\\p{N}_])(?:[Cc]fr|[Vv]|[Ii]n)\\.?\\s)((?:[A-Z]\\.\\s?)+)\\s?" + esc + "(?=,\\s)", "gu");
    for (const m of tutto.matchAll(re1)) {
      // "cap. I. Gatta", "vol. V. Manes": numero romano, non iniziale
      const prima = tutto.slice(Math.max(0, m.index - 8), m.index);
      if (/(?:cap|vol|[Ss]ez|tomo|parte|par|art|fasc|t)\.\s*$|§\s*$/.test(prima)) continue;
      set.add(m[1].replace(/\s+/g, ""));
    }

    const re2 = new RegExp("(?<![\\p{L}\\p{N}_’'.])([A-ZÀ-Ý][a-zà-ÿ]+)\\s" + esc + "(?=,\\s)", "gu");
    for (const m of tutto.matchAll(re2))
      if (B.NOMI_PROPRI.has(m[1])) set.add(m[1][0] + ".");

    if (set.size > 0) trovate.set(chiave(c), set);
  }

  return new Map([...trovate].filter(([, v]) => v.size === 1).map(([k, v]) => [k, [...v][0]]));
}

// Nomi che non sono persone (opere antiche, titoli presi per autori)
const NON_PERSONE = new Set([
  "Annales", "Giustiniano", "Institutiones", "Liquidated Damages", "Marathon Woman", "Saturae",
  "Tacito", "Giovenale", "Montesquieu"
]);

// ---------- iniziali predefinite ----------
// Cercate per gli autori delle tesi in lavorazione (riviste, cataloghi, opere che li citano).
// "?" = non verificate su una fonte: proposte senza spunta.
export const PREDEFINITE = `Acconciamessa = L.
Alibrandi = L. ?
Altavilla = E.
Amarelli = G.
Amato = G.
Ambrosetti = E.M.
Amore = N. ?
Antolisei = F.
Ballesi = F.
Bartoli = R.
Basso = C.
Battelli = E.
Benussi = C.
Bevilacqua = B.
Biondi = G.
Bondi = A.
Borsari = R.
Bricola = F.
Brunelli = D.
Buford = K.
Cadoppi = A.
Califano = G.S.
Camaioni = S.
Canestrari = S.
Canosa = R.
Cantone = R.
Canzian = N.
Caprioli = F.
Caputo = M.
Cardenal Carro = M.
Carmona = A.
Carnelutti = F.
Carrara = F.
Carraro = L.
Carroll = L.
Castaldo = A.R.
Catenacci = M.
Cazzaniga = M.
Cerulli Irelli = V.
Cesarini Sforza = W.
Coase = R.H.
Cocco = G.
Contento = G.
Coppi = F.
Corrado = D.
Crisafulli = V. ?
Cupelli = C.
Cutolo = M.
D’Andrea = R.
D’Ascola = V.N.
D’Avirro = A.
Dagnino = E.
Dalia = A.A.
Davigo = P.
De Blasis = S.
De Francesco = G.
Della Monica = G.
Della Ragione = L.
Di Giovine = O.
Di Martino = A.
Di Vetta = G.
Dolcini = E.
Domenicali = C.
Donini = M.
Elster = J.
Federico = P.
Ferraioli = M.
Ferrara = V.
Fiandaca = G.
Finocchiaro = S.
Fiore = S.
Fiorella = A.
Fornasari = G.
Franchini = C.
Frattarolo = V.
Frosina = P.
Gambardella = M.
Gargani = A.
Garofoli = R.
Gatta = G.L.
Gatto = C.E.
Giannini = M.S.
Goetz = C.J.
Gramsci = A.
Grossi | Mitologie giuridiche = P.
Grossi | I “nuovi” obblighi = L.
Grosso = C.F.
Guidi = D.
Holmes = O.W.
Huizinga = J.
Infante = E.
Infantini = F.
Insolera = G.
Irti = N.
Kahn-Freund = O.
Kessler = F.
Lanzi = A.
Lattanzi = G.
Leoni = M.
Maine = H.S.
Manes = V.
Manna = A.
Mannozzi = G.
Mantovani = F.
Manzini = V.
Marandola = A.
Marconi = G.
Marinucci = G.
Massi = S.
Mattevi = E.
Maugeri = A.M.
Merlo = A.
Mezzetti = E.
Micheletti = D.
Michieli = P.
Miele = G.
Miettinen = S.
Miller = M.
Milone = A.
Modugno = F.
Mongillo = V.
Morbidelli = G.
Mortati = C.
Musco = E.
Naddeo = M.
Napoleoni = V.
Neale = W.C.
Nisco = A.
Nuvolone = P.
Padovani = T.
Pagella = C.
Pagliaro = A.
Palazzo = F.
Paliero = C.E.
Paonessa = C.
Papa = M.
Parodi Giusino = M.
Parrish = R.
Pavei = A.
Pelissero = M.
Perin = A.
Perongini = S.
Picotti = L.
Pisa = P.
Pisani = N.
Piva = D.
Posner = R.A.
Pulitanò = D.
Rampioni = R.
Ravenna = M.
Rispoli = A.G.
Rocco = A. ?
Roiati = A.
Romanelli = M.
Romano | L’ordinamento giuridico = S.
Rotondo = F.
Rottenberg = S.
Ruggiero = G.
Ruotolo = G.M.
Salcuni = G.
Saleilles = R.
Sassi = S.
Satta = A.
Scalfati = A.
Schiavon = M.
Scoletta = M.
Scordamaglia = V.
Scott = R.E.
Seminara = S.
Severino = P.
Sicignano = G.J.
Silva = C.
Smith = A.
Spadafora = M.T.
Stile = A.M.
Stortoni = L.
Switzer = K.
Tagliarini = F.
Tesauro = A.
Thiene = G.
Titmuss = R.M.
Trimboli = M.C.
Valentini = V.
Vallini = A.
Vicoli = D.
Vinciguerra = S.
Weatherill = S.
Zagrebelsky = V.
Zanobini = G.
Zuliani = D.
`;
