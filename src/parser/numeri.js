// Numeri in lettere (importi in euro), come nella versione Windows (NumberToItalianWords.cs e WordUtils.cs).
// Gli importi si tengono in centesimi interi, per evitare gli arrotondamenti dei numeri in virgola mobile.

const UNITA = ["", "uno", "due", "tre", "quattro", "cinque", "sei", "sette", "otto", "nove",
  "dieci", "undici", "dodici", "tredici", "quattordici", "quindici",
  "sedici", "diciassette", "diciotto", "diciannove"];
const DECINE = ["", "", "venti", "trenta", "quaranta", "cinquanta", "sessanta", "settanta", "ottanta", "novanta"];

// Valore massimo convertibile: 999 miliardi (in centesimi)
export const MAX_CENTESIMI = 99999999999999;

// Numero con separatore delle migliaia, sempre lo stesso in tutto il numero:
// punto, virgola, spazio, spazio non divisibile o apostrofo ("1.234.567,89", "1 234,56", "1'234,56", "1,234.56"),
// oppure numero semplice con decimali opzionali ("1234,5", "1234.5")
export const RE_NUMERO = /\d{1,3}(?<sep>[.,'’   ])\d{3}(?:\k<sep>\d{3})*(?!\d)(?:[.,]\d+)?|\d+(?:[.,]\d+)?/g;
const SEPARATORI_SPAZIO = /['’   ]/g;
const SOLO_MIGLIAIA = /^\d{1,3}(?:\.\d{3})+$/;

// Testo in lettere " (…)" lasciato da una conversione precedente
export const RE_LETTERE_PRECEDENTI = /^ \([^()\r\u0007]*\)/;

// Individua il separatore decimale:
// - con punto e virgola insieme, è decimale l'ultimo dei due ("1.234,56", "1,234.56")
// - con sole virgole, una virgola è decimale ("1234,5"), più virgole sono migliaia
// - con soli punti, un punto è decimale ("1.5") tranne se seguito da tre cifre ("1.500" = millecinquecento)
// Restituisce i centesimi (arrotondati come in .NET: a metà verso il pari) o null
export function leggiNumero(testo) {
  const clean = testo.replace(SEPARATORI_SPAZIO, "");
  const lastDot = clean.lastIndexOf("."), lastComma = clean.lastIndexOf(",");
  let dec = -1;

  if (lastDot >= 0 && lastComma >= 0) dec = Math.max(lastDot, lastComma);
  else if (lastComma >= 0 && lastComma === clean.indexOf(",")) dec = lastComma;
  else if (lastDot >= 0 && lastDot === clean.indexOf(".") && !SOLO_MIGLIAIA.test(clean)) dec = lastDot;

  const intero = (dec >= 0 ? clean.slice(0, dec) : clean).replace(/[.,]/g, "");
  const decimali = dec >= 0 ? clean.slice(dec + 1) : "";
  if (!/^\d+$/.test(intero) || !/^\d*$/.test(decimali)) return null;

  // centesimi = intero * 100 + prime due cifre decimali, arrotondando con le cifre successive
  const d = (decimali + "00").slice(0, 2);
  let cent = BigInt(intero) * 100n + BigInt(d);
  const resto = decimali.slice(2);
  if (resto.length > 0) {
    const primo = +resto[0];
    const oltre = /[1-9]/.test(resto.slice(1));
    if (primo > 5 || (primo === 5 && (oltre || cent % 2n === 1n))) cent += 1n;
  }
  return cent > BigInt(Number.MAX_SAFE_INTEGER) ? Infinity : Number(cent);
}

// "1.234,50"
export function formatta(centesimi) {
  const euro = Math.floor(centesimi / 100), cent = centesimi % 100;
  const gruppi = String(euro).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return gruppi + "," + String(cent).padStart(2, "0");
}

function convertiIntero(n) {
  if (n === 0) return "zero";
  if (n < 20) return UNITA[n];
  if (n < 100) {
    const t = Math.floor(n / 10), u = n % 10;
    let decina = DECINE[t];
    if (u === 1 || u === 8) decina = decina.slice(0, -1);         // ventuno, ventotto
    return decina + UNITA[u];
  }
  if (n < 1000) {
    const h = Math.floor(n / 100), r = n % 100;
    let prefisso = h === 1 ? "cento" : UNITA[h] + "cento";
    if (r === 8 || Math.floor(r / 10) === 8) prefisso = prefisso.slice(0, -1);   // centotto, centottanta
    return prefisso + (r > 0 ? convertiIntero(r) : "");
  }
  if (n < 1000000) {
    const th = Math.floor(n / 1000), r = n % 1000;
    return (th === 1 ? "mille" : convertiIntero(th) + "mila") + (r > 0 ? convertiIntero(r) : "");
  }
  if (n < 1000000000) {
    const m = Math.floor(n / 1000000), r = n % 1000000;
    return (m === 1 ? "unmilione" : convertiIntero(m) + "milioni") + (r > 0 ? convertiIntero(r) : "");
  }
  const b = Math.floor(n / 1000000000), r = n % 1000000000;
  return (b === 1 ? "unmiliardo" : convertiIntero(b) + "miliardi") + (r > 0 ? convertiIntero(r) : "");
}

// Accento sul "tre" finale dei composti: ventitré, centotré, milletré
function inLettere(n) {
  let w = convertiIntero(n);
  if (n > 3 && w.endsWith("tre")) w = w.slice(0, -3) + "tré";
  return w;
}

function controlla(centesimi) {
  if (centesimi < 0 || centesimi > MAX_CENTESIMI) throw new Error("Importo fuori intervallo (da 0 a 999 miliardi).");
}

// "milleduecentotrentaquattro euro e cinquanta centesimi"
export function euro(centesimi) {
  controlla(centesimi);
  const e = Math.floor(centesimi / 100), c = centesimi % 100;
  // "di" davanti alla valuta solo per milioni e miliardi tondi: "due milioni di euro"
  const testoEuro = e === 1 ? "un euro" : e >= 1000000 && e % 1000000 === 0 ? inLettere(e) + " di euro" : inLettere(e) + " euro";
  if (c === 0) return testoEuro;
  return testoEuro + " e " + (c === 1 ? "un centesimo" : inLettere(c) + " centesimi");
}

// "milleduecentotrentaquattro/50"
export function finanziario(centesimi) {
  controlla(centesimi);
  return inLettere(Math.floor(centesimi / 100)) + "/" + String(centesimi % 100).padStart(2, "0");
}

// "Euro milleduecentotrentaquattro/50"
export function amministrativo(centesimi) {
  return "Euro " + finanziario(centesimi);
}

export const FORMATI = { euro, finanziario, amministrativo };

// Numero sotto il cursore (posizione nel testo) o, se il cursore non è su un numero, il primo
export function trovaNumero(testo, posizione) {
  let scelto = null;
  for (const m of testo.matchAll(RE_NUMERO)) {
    if (!scelto) scelto = m;
    if (posizione >= m.index && posizione <= m.index + m[0].length) { scelto = m; break; }
  }
  return scelto ? { idx: scelto.index, testo: scelto[0] } : null;
}
