/* global Word */
// Passaggi comuni agli elenchi di paragrafi con uno stile (bibliografia, giurisprudenza):
// sostituzioni nel testo delle voci e riordino spostando solo le voci fuori posto.

import { vociConStile, preparaTratto, aBlocchi, supporta } from "./word.js";

const BLOCCO = 40;      // voci lavorate per ogni richiesta a Word

export async function voci(context, stile) {
  const v = await vociConStile(context, stile);
  if (v.paragrafi.length === 0)
    throw new Error("Nessun paragrafo con stile \"" + stile + "\". Controlla il nome dello stile.");
  return v;
}

// Sostituzioni [vecchio, nuovo] date da regole(testo), sempre sull'ultima occorrenza, da destra a sinistra.
// Quelle sovrapposte a un'altra passano al giro successivo (al massimo 3 giri, come nella versione Windows).
// Restituisce il numero di voci modificate.
export async function sostituisci(context, stile, regole, avanzamento) {
  const modificate = new Set();

  for (let giro = 0; giro < 3; giro++) {
    const { paragrafi, testi } = await voci(context, stile);
    const lavori = [];
    let saltate = 0;

    paragrafi.forEach((p, i) => {
      const t = testi[i];
      const presi = [];
      for (const [vecchio, nuovo] of regole(t)) {
        const pos = t.lastIndexOf(vecchio);
        if (pos < 0) continue;
        if (presi.some(x => pos < x.pos + x.vecchio.length && x.pos < pos + vecchio.length)) { saltate++; continue; }
        presi.push({ i, p, t, pos, vecchio, nuovo });
      }
      if (presi.length) lavori.push(presi.sort((a, b) => b.pos - a.pos));
    });

    await aBlocchi(lavori, BLOCCO, async blocco => {
      const tutti = blocco.flat().map(x => ({ ...x, cerca: preparaTratto(x.p, x.t, x.pos, x.vecchio.length) }));
      await context.sync();
      for (const x of tutti) {
        const r = x.cerca.risolvi();
        if (r) { r.insertText(x.nuovo, "Replace"); modificate.add(x.i); }
      }
      await context.sync();
    }, avanzamento);

    if (lavori.length === 0 || saltate === 0) break;
  }

  return modificate.size;
}

// Spostare le voci con le revisioni attive lascerebbe nel testo le voci cancellate: meglio fermarsi
export async function revisioniDisattivate(context, cosa) {
  if (!supporta("WordApi", "1.4")) return;
  context.document.load("changeTrackingMode");
  await context.sync();
  if (context.document.changeTrackingMode !== "Off")
    throw new Error("Le revisioni sono attive: per riordinare " + cosa + " disattivale (Revisione → Revisioni) " +
                    "e accetta o rifiuta quelle in sospeso nell'elenco, poi riprova. " +
                    "Gli altri passaggi funzionano anche con le revisioni attive.");
}

const ordinale = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

// Indici delle voci che restano ferme: la sottosequenza già in ordine più lunga
export function vociInOrdine(keys) {
  const n = keys.length;
  const len = new Array(n).fill(1);
  const prev = new Array(n).fill(-1);
  for (let i = 0; i < n; i++)
    for (let j = 0; j < i; j++)
      if (ordinale(keys[j], keys[i]) <= 0 && len[j] + 1 > len[i]) { len[i] = len[j] + 1; prev[i] = j; }
  const keep = new Set();
  if (n === 0) return keep;
  let b = len.lastIndexOf(Math.max(...len));
  for (; b >= 0; b = prev[b]) keep.add(b);
  return keep;
}

// Riordina le voci secondo chiavi(testi) (confronto ordinale) spostando solo quelle fuori posto.
// Le voci spostate sono riscritte: chi chiama rimette poi la formattazione. Restituisce quante sono.
export async function ordina(context, stile, chiavi, cosa) {
  await revisioniDisattivate(context, cosa);
  const { paragrafi, testi } = await voci(context, stile);
  const k = chiavi(testi);
  const fermi = vociInOrdine(k);
  if (fermi.size === paragrafi.length) return 0;

  const ordinate = [...fermi].sort((a, b) => a - b).map(i => ({ p: paragrafi[i], chiave: k[i] }));
  let spostate = 0;

  for (let i = 0; i < paragrafi.length; i++) {
    if (fermi.has(i)) continue;
    const idx = ordinate.findIndex(e => ordinale(e.chiave, k[i]) > 0);
    const nuova = idx >= 0
      ? ordinate[idx].p.insertParagraph(testi[i], "Before")
      : ordinate[ordinate.length - 1].p.insertParagraph(testi[i], "After");
    nuova.style = stile;
    ordinate.splice(idx >= 0 ? idx : ordinate.length, 0, { p: nuova, chiave: k[i] });
    paragrafi[i].delete();
    spostate++;
  }

  await context.sync();
  return spostate;
}
