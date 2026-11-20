// Testo dei sommari di capitolo, senza Word: si può provare anche fuori dal documento.
// "Sommario: 1. Primo paragrafo. – 2. Secondo. – 2.1. Sottoparagrafo."

export const ETICHETTA = "Sommario:";
export const SEPARATORE = " – ";

// Senza richiami di nota, campi e spazi doppi
export function pulisci(s) {
  let r = "";
  for (const c of s || "") {
    if (c === "\t" || c === " ") r += " ";
    else if (c >= " ") r += c;
  }
  return r.replace(/ {2,}/g, " ").trim();
}

// Per confrontare il sommario esistente con quello nuovo
export function normalizza(s) {
  return pulisci(s).replace(/\s+/g, " ").replace(/-/g, "–");
}

// "2.1" + "La riscrittura" → "2.1. La riscrittura." (con il punto finale se manca)
export function voce(numero, titolo) {
  numero = (numero || "").trim();
  titolo = pulisci(titolo);

  if (numero && /[\p{L}\p{N}]$/u.test(numero)) numero += ".";
  if (numero && titolo.startsWith(numero)) numero = "";   // numero già scritto a mano nel titolo

  let testo = (numero ? numero + " " : "") + titolo;
  if (!/[.?!…]$/.test(testo)) testo += ".";

  return { testo, inizioTitolo: numero ? numero.length + 1 : 0 };
}

// Tratti in corsivo del titolo, dalle parole lette in Word (testo e corsivo di ciascuna):
// le parole in corsivo vicine (separate solo da spazi) formano un unico tratto
export function trattiCorsivi(testoTitolo, parole) {
  const t = testoTitolo || "";
  const corsivo = new Array(t.length).fill(false);
  let cursore = 0;

  for (const p of parole || []) {
    const w = (p.text || "").trim();
    if (!w) continue;
    const i = t.indexOf(w, cursore);
    if (i < 0) continue;
    if (p.italic === true) for (let k = i; k < i + w.length; k++) corsivo[k] = true;
    cursore = i + w.length;
  }

  // gli spazi tra due parole in corsivo sono in corsivo anche loro
  for (let k = 0; k < t.length; k++) {
    if (corsivo[k] || !/\s/.test(t[k])) continue;
    let a = k - 1, b = k;
    while (b < t.length && /\s/.test(t[b])) b++;
    if (a >= 0 && corsivo[a] && b < t.length && corsivo[b])
      for (let j = k; j < b; j++) corsivo[j] = true;
  }

  const tratti = [];
  let inizio = -1;
  for (let k = 0; k <= t.length; k++) {
    if (k < t.length && corsivo[k]) { if (inizio < 0) inizio = k; continue; }
    if (inizio >= 0) {
      const s = pulisci(t.slice(inizio, k));
      if (s) tratti.push(s);
      inizio = -1;
    }
  }
  return tratti;
}

// Sommario completo. voci: [{ numero, titolo, corsivi: [stringhe] }]
// Restituisce il testo e, per ogni tratto in corsivo, quale occorrenza è nel testo
// (serve per ritrovarlo con la ricerca di Word)
export function componi(voci) {
  if (!voci || voci.length === 0) return { testo: "", corsivi: [] };

  let testo = ETICHETTA + " ";
  const posizioni = [];

  voci.forEach((v, n) => {
    if (n > 0) testo += SEPARATORE;
    const inizioVoce = testo.length;
    const { testo: tv, inizioTitolo } = voce(v.numero, v.titolo);
    testo += tv;

    let da = inizioVoce + inizioTitolo;
    for (const c of v.corsivi || []) {
      const i = testo.indexOf(c, da);
      if (i < 0) continue;
      posizioni.push({ inizio: i, fine: i + c.length, testo: c });
      da = i + c.length;
    }
  });

  const corsivi = posizioni.map(p => ({ ...p, occorrenza: occorrenzePrima(testo, p.testo, p.inizio) }));
  return { testo, corsivi };
}

// Quante volte s compare (senza sovrapposizioni) prima della posizione indicata
function occorrenzePrima(testo, s, posizione) {
  let n = 0, i = testo.indexOf(s);
  while (i >= 0 && i < posizione) { n++; i = testo.indexOf(s, i + s.length); }
  return n;
}

export function stato(cap) {
  if (cap.voci === 0) return "nessun paragrafo";
  if (!cap.esistente) return "da creare";
  return normalizza(cap.testoEsistente) === normalizza(cap.testo) ? "aggiornato" : "da aggiornare";
}

// "Titolo 1", "Heading 1", "Intestazione 1" → "Heading1" (nome dello stile predefinito in Office.js)
export function stilePredefinito(nome) {
  const m = /^(?:titolo|heading|intestazione)\s*([1-9])$/i.exec((nome || "").trim());
  return m ? "Heading" + m[1] : null;
}

// Livello di struttura di un paragrafo (1–9; 10 = corpo del testo)
export function livello(outlineLevel, styleBuiltIn) {
  if (typeof outlineLevel === "number" && outlineLevel >= 1 && outlineLevel <= 9) return outlineLevel;
  const m = /(\d)$/.exec(typeof outlineLevel === "string" ? outlineLevel : "");
  if (m && !/body/i.test(outlineLevel)) return +m[1];
  const h = /^Heading([1-9])$/.exec(styleBuiltIn || "");
  return h ? +h[1] : 10;
}

// Il paragrafo subito dopo il titolo è già un sommario?
export function eSommario(stile, testo, stileSommario) {
  return stile === stileSommario || /^Sommario\s*[:.]/i.test(pulisci(testo));
}

// Per la ricerca di Word: "^" è un carattere speciale; massimo 255 caratteri
export function perRicerca(s) {
  return s.slice(0, 200).replace(/\^/g, "^^");
}
