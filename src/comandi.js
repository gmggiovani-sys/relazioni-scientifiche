/* global Office */
// Comandi eseguiti direttamente dai pulsanti della scheda Relazioni Tecniche (senza pannello).
// Un messaggio compare solo per gli errori e per gli esiti da leggere.

import * as Foto from "./foto.js";
import * as Numeri from "./numeri.js";

function avviso(testo) {
  if (!testo) return Promise.resolve();
  return new Promise(ok => {
    const url = new URL("avviso.html", location.href);
    url.searchParams.set("t", testo);
    Office.context.ui.displayDialogAsync(url.toString(), { height: 30, width: 35, displayInIframe: true }, r => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { ok(); return; }
      const d = r.value;
      d.addEventHandler(Office.EventType.DialogMessageReceived, () => { d.close(); ok(); });
      d.addEventHandler(Office.EventType.DialogEventReceived, () => ok());
    });
  });
}

function comando(lavoro, mostraEsito = true) {
  return async event => {
    try {
      const esito = await lavoro();
      if (mostraEsito) await avviso(esito);
    } catch (e) {
      await avviso("Operazione non riuscita: " + (e && e.message ? e.message : e));
    } finally {
      event.completed();
    }
  };
}

const COMANDI = {
  spostaIndietro: comando(() => Foto.sposta(-1)),
  spostaAvanti: comando(() => Foto.sposta(+1)),
  riorganizza: comando(() => Foto.riorganizza()),
  rinumera: comando(() => Foto.rinumera()),
  layout1: comando(() => Foto.cambiaLayout(1), false),
  layout2: comando(() => Foto.cambiaLayout(2), false),
  layout4: comando(() => Foto.cambiaLayout(4), false),
  capitoloFoto: comando(() => Foto.capitoloFoto()),
  numeriEuro: comando(() => Numeri.converti("euro")),
  numeriFinanziario: comando(() => Numeri.converti("finanziario")),
  numeriAmministrativo: comando(() => Numeri.converti("amministrativo"))
};

// nomi globali (FunctionName del manifest) e, dove c'è, anche Office.actions
for (const [nome, f] of Object.entries(COMANDI)) window[nome] = f;
Office.onReady(() => {
  if (Office.actions && Office.actions.associate)
    for (const [nome, f] of Object.entries(COMANDI)) Office.actions.associate(nome, f);
});
