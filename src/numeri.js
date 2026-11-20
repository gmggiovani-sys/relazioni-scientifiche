/* global Word */
// Numeri in lettere: il numero sotto il cursore (o il primo del testo selezionato) diventa
// "1.234,50 (milleduecentotrentaquattro euro e cinquanta centesimi)", con la parte in lettere in corsivo.
// Se il comando era già stato eseguito, la vecchia parte in lettere si sostituisce.

import * as N from "./parser/numeri.js";
import { preparaTratto } from "./word.js";

export async function converti(formato) {
  const inLettere = N.FORMATI[formato];
  return Word.run(async context => {
    const sel = context.document.getSelection();
    sel.load("text");
    await context.sync();

    // con il solo cursore: il paragrafo, e la posizione del cursore dentro di esso
    let contenitore = sel, posizione = 0;
    if (!sel.text) {
      contenitore = sel.paragraphs.getFirst();
      const prima = contenitore.getRange("Start").expandTo(sel.getRange("Start"));
      prima.load("text");
      await context.sync();
      posizione = prima.text.length;
    }
    contenitore.load("text");
    await context.sync();

    const testo = contenitore.text;
    const num = N.trovaNumero(testo, posizione);
    if (!num) throw new Error("Nessun numero trovato.");

    const cent = N.leggiNumero(num.testo);
    if (cent === null) throw new Error("Numero non valido.");
    if (cent > N.MAX_CENTESIMI) throw new Error("Importo troppo grande (massimo 999 miliardi).");

    // la parte in lettere di una conversione precedente, subito dopo il numero
    const fine = num.idx + num.testo.length;
    const vecchie = N.RE_LETTERE_PRECEDENTI.exec(testo.slice(fine));

    const rNum = preparaTratto(contenitore, testo, num.idx, num.testo.length);
    const rVecchie = vecchie ? preparaTratto(contenitore, testo, fine, vecchie[0].length) : null;
    await context.sync();

    const numero = rNum.risolvi();
    if (!numero) throw new Error("Impossibile individuare il numero nel documento.");
    const vecchio = rVecchie && rVecchie.risolvi();
    if (vecchio) {
      vecchio.load("font/italic");
      await context.sync();
      // solo testo interamente in corsivo, come quello inserito dall'add-in
      if (vecchio.font.italic === true) vecchio.delete();
    }

    // sostituisce solo il numero, lasciando intatti il resto del testo e la formattazione
    const nuovo = numero.insertText(N.formatta(cent), "Replace");
    const lettere = nuovo.insertText(" (" + inLettere(cent) + ")", "After");
    lettere.font.italic = true;
    await context.sync();
    return "";
  });
}
