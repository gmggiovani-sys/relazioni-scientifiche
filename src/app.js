/* global Office */
// Pannello laterale: sceglie lo strumento e collega i pulsanti ai comandi

import * as Sommari from "./sommari.js";
import * as Citazioni from "./citazioni.js";
import * as Elenco from "./elenco.js";
import * as Bibliografia from "./bibliografia.js";
import * as Giurisprudenza from "./giurisprudenza.js";
import * as Abbreviazioni from "./abbreviazioni.js";
import * as Foto from "./foto.js";
import * as Numeri from "./numeri.js";
import { preparaFoto, riduci } from "./immagini.js";

const $ = id => document.getElementById(id);

// ----- messaggi e preferenze -----

function messaggio(testo, errore = false) {
  const m = $("messaggio");
  m.hidden = !testo;
  m.textContent = testo || "";
  m.classList.toggle("errore", errore);
}

function ricorda(id) {
  const el = $(id);
  try {
    const v = localStorage.getItem("rs." + id);
    if (v !== null) el.value = v;
  } catch (e) { /* senza memoria locale si usa il valore predefinito */ }
  el.addEventListener("change", () => {
    try { localStorage.setItem("rs." + id, el.value); } catch (e) { /* niente */ }
  });
}

// Esegue un comando con i pulsanti disattivati e l'errore in chiaro
async function esegui(pulsanti, azione) {
  pulsanti.forEach(b => (b.disabled = true));
  if (pulsanti.length) messaggio("");       // "Vai" lascia il messaggio com'è
  try {
    await azione();
  } catch (e) {
    console.error(e);
    messaggio("Operazione non riuscita: " + (e && e.message ? e.message : e), true);
  } finally {
    pulsanti.forEach(b => (b.disabled = false));
  }
}

// ----- Sommari dei capitoli -----

let ultimi = null;     // ultima analisi

function stiliSommari() {
  return [$("stileTitolo").value, $("stileSommario").value];
}

function disegnaSommari() {
  const ul = $("elenco");
  ul.textContent = "";
  if (!ultimi) return;

  const solo = $("soloDaFare").checked;
  for (const c of ultimi.capitoli) {
    const daFare = c.voci > 0 && c.stato !== "aggiornato";
    if (solo && !daFare) continue;

    const li = document.createElement("li");
    li.className = daFare ? "daFare" : c.stato === "aggiornato" ? "fatto" : "";

    const testa = document.createElement("label");
    testa.className = "testa";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.disabled = c.voci === 0;
    cb.checked = c.scelto;
    cb.addEventListener("change", () => (c.scelto = cb.checked));
    const t = document.createElement("span");
    t.className = "titolo";
    const strong = document.createElement("strong");
    strong.textContent = c.titolo;
    const small = document.createElement("small");
    small.textContent = "Sezione " + c.sezione + " · " + c.voci + " paragrafi · " + c.stato;
    t.append(strong, small);
    testa.append(cb, t);

    const azioni = document.createElement("div");
    azioni.className = "azioni";
    const vai = document.createElement("a");
    vai.textContent = "Vai al capitolo";
    vai.addEventListener("click", () =>
      esegui([], async () => {
        if (!(await Sommari.vaiA(c.chiave, ...stiliSommari()))) messaggio("Capitolo non trovato: il documento è cambiato, cerca di nuovo.");
      }));
    const mostra = document.createElement("a");
    mostra.textContent = "Mostra il sommario";
    azioni.append(vai);
    if (c.voci > 0) azioni.append(mostra);

    const anteprima = document.createElement("div");
    anteprima.className = "anteprima";
    anteprima.hidden = true;
    anteprima.textContent = c.testo;
    if (c.testoEsistente && c.stato === "da aggiornare") {
      const v = document.createElement("div");
      v.className = "vecchio";
      v.textContent = "Ora: " + c.testoEsistente;
      anteprima.append(v);
    }
    mostra.addEventListener("click", () => {
      anteprima.hidden = !anteprima.hidden;
      mostra.textContent = anteprima.hidden ? "Mostra il sommario" : "Nascondi il sommario";
    });

    li.append(testa, azioni, anteprima);
    ul.append(li);
  }

  if (!ul.children.length) {
    const li = document.createElement("li");
    li.textContent = "Tutti i sommari sono già aggiornati.";
    ul.append(li);
  }
}

async function analizzaSommari() {
  const r = await Sommari.analizza(...stiliSommari());
  for (const c of r.capitoli) c.scelto = c.voci > 0 && c.stato !== "aggiornato";
  ultimi = r;
  $("risultati").hidden = r.capitoli.length === 0;
  disegnaSommari();

  if (r.capitoli.length === 0)
    messaggio("Nessun paragrafo con stile \"" + r.stileTitolo + "\". Indica lo stile dei titoli dei capitoli " +
              "(es. Titolo 1), oppure lascia vuota la casella e metti il cursore su un titolo di capitolo.");
  else {
    const daFare = r.capitoli.filter(c => c.scelto).length;
    messaggio(r.capitoli.length + " capitoli dopo lo stile \"" + r.stileTitolo + "\": " +
              (daFare ? daFare + " da creare o aggiornare." : "sommari tutti aggiornati."));
  }
}

async function scriviSommari() {
  if (!ultimi) return;
  const chiavi = ultimi.capitoli.filter(c => c.scelto).map(c => c.chiave);
  if (!chiavi.length) { messaggio("Nessun capitolo spuntato."); return; }

  const esito = await Sommari.scrivi(chiavi, ...stiliSommari());
  await analizzaSommari();
  messaggio(esito);
}

function preparaSommari() {
  ricorda("stileTitolo");
  ricorda("stileSommario");
  const pulsanti = [$("analizza"), $("scrivi")];
  $("analizza").addEventListener("click", () => esegui(pulsanti, analizzaSommari));
  $("scrivi").addEventListener("click", () => esegui(pulsanti, scriviSommari));
  $("soloDaFare").addEventListener("change", disegnaSommari);
  const tutte = v => () => {
    if (!ultimi) return;
    for (const c of ultimi.capitoli) if (c.voci > 0) c.scelto = v;
    disegnaSommari();
  };
  $("tutte").addEventListener("click", tutte(true));
  $("nessuna").addEventListener("click", tutte(false));
}

// ----- Citazioni nelle note -----

let righeCit = [];

function opzioniCit() {
  return [$("stileBiblio").value.trim() || "Biblio", $("ambitoCit").value === "selezione"];
}

const avanzamento = cosa => (fatti, totale) => messaggio(cosa + ": " + fatti + " di " + totale + "…");

async function controllaCitazioni() {
  const [stile, sel] = opzioniCit();
  messaggio("Lettura delle note…");
  const r = await Citazioni.controlla(stile, sel, avanzamento("Ricerca nel documento"));

  const vai = (chiave, testo) => () =>
    esegui([], async () => {
      if (!(await Citazioni.vaiA(chiave, testo, sel))) messaggio("Non trovato: il documento è cambiato, controlla di nuovo.");
    });

  righeCit = [
    ...r.daControllare.map(d => ({
      chiave: d.chiave, titolo: d.dove + " · da controllare a mano", testo: d.testo,
      classe: "dubbio", spuntabile: false, vai: vai(d.chiave, null)
    })),
    ...r.parti.map(p => ({
      chiave: p.chiave, titolo: p.testo, classe: p.trovato ? "" : "dubbio",
      dettaglio: p.dove + " · " + p.formato + " · " + p.regola + (p.trovato ? "" : " · non individuato: da fare a mano"),
      spuntabile: p.trovato, spuntato: true,
      vai: p.trovato ? vai(p.chiave, p.testo) : null
    }))
  ];

  $("risultatiCit").hidden = righeCit.length === 0;
  Elenco.disegna($("elencoCit"), righeCit);

  messaggio(righeCit.length === 0
    ? "Niente da formattare: le " + r.giaConformi + " parti riconosciute (autori, titoli, riviste, rinvii) sono già a posto."
    : r.giaConformi + " parti già a posto, " + r.parti.length + " da formattare, " +
      r.daControllare.length + " citazioni da controllare a mano (in giallo)." +
      (r.maiusc ? "" : "\nIl maiuscoletto si applica con lo stile carattere \"Maiuscoletto\"."));
}

async function formattaCitazioni(soloScelte) {
  const [stile, sel] = opzioniCit();
  const chiavi = soloScelte ? Elenco.scelte(righeCit) : null;
  if (soloScelte && chiavi.length === 0) { messaggio("Nessuna parte spuntata."); return; }

  messaggio("Lettura delle note…");
  const esito = await Citazioni.formatta(stile, sel, chiavi, avanzamento("Ricerca nel documento"));
  if (soloScelte) await controllaCitazioni();
  else { righeCit = []; $("risultatiCit").hidden = true; }
  messaggio(esito);
}

function preparaCitazioni() {
  ricorda("stileBiblio");
  ricorda("ambitoCit");
  const pulsanti = [$("controllaCit"), $("formattaCit"), $("formattaScelteCit")];
  $("controllaCit").addEventListener("click", () => esegui(pulsanti, controllaCitazioni));
  $("formattaCit").addEventListener("click", () => esegui(pulsanti, () => formattaCitazioni(false)));
  $("formattaScelteCit").addEventListener("click", () => esegui(pulsanti, () => formattaCitazioni(true)));
  $("tutteCit").addEventListener("click", () => { Elenco.spunta(righeCit, true); Elenco.disegna($("elencoCit"), righeCit); });
  $("nessunaCit").addEventListener("click", () => { Elenco.spunta(righeCit, false); Elenco.disegna($("elencoCit"), righeCit); });
}

// ----- Bibliografia -----

let righeBib = [];
let modoBib = null;          // "campi" o "note"
let campiBib = [];

const stileVoci = () => $("stileVoci").value.trim() || "Biblio";
const progresso = (cosa, f, t) => messaggio(cosa + ": voce " + f + " di " + t + "…");

function mostraRisultatiBib(modo, titolo) {
  modoBib = modo;
  $("risultatiBib").hidden = false;
  $("titoloBib").textContent = titolo;
  const note = modo === "note";
  $("tutteBib").hidden = !note;
  $("nessunaBib").hidden = !note;
  $("aggiungiBib").hidden = !note;
  $("filtroBibLabel").hidden = note;
  $("copiaBib").hidden = note;
  Elenco.disegna($("elencoBib"), righeBib, {
    soloEvidenziate: !note && $("soloDubbieBib").checked,
    vuoto: note ? "Nessuna voce da aggiungere." : "Tutte le voci hanno autore e titolo."
  });
}

function nascondiRisultatiBib() {
  $("risultatiBib").hidden = true;
  righeBib = [];
}

async function eseguiBib(azione) {
  nascondiRisultatiBib();
  messaggio("In corso…");
  messaggio(await azione());
}

async function estraiCampiBib() {
  nascondiRisultatiBib();
  campiBib = await Bibliografia.campi(stileVoci());
  const vai = i => () => esegui([], async () => { await Bibliografia.vaiAVoce(stileVoci(), i); });
  righeBib = campiBib.map(c => {
    const dubbia = !c.autore || !c.titolo;
    return {
      chiave: String(c.i), classe: dubbia ? "dubbio" : "",
      titolo: (c.autore || "(autore non riconosciuto)") + " – " + (c.titolo || "(titolo non riconosciuto)"),
      dettaglio: [c.tipo, c.fonte, c.editore, c.citta, c.anno].filter(Boolean).join(" · "),
      spuntabile: false, vai: vai(c.i)
    };
  });
  const dubbie = righeBib.filter(r => r.classe).length;
  mostraRisultatiBib("campi", campiBib.length + " voci, " + dubbie + " senza autore o titolo");
  messaggio("Campi estratti da " + campiBib.length + " voci. In giallo quelle senza autore o titolo riconosciuto. " +
            "\"Copia per Excel\" copia la tabella completa da incollare in un foglio di calcolo.");
}

async function copiaCampiBib() {
  await copiaTabella(["Autore", "Titolo", "Fonte", "Editore", "Anno", "Città", "Tipo", "Voce"],
    campiBib.map(c => [c.autore, c.titolo, c.fonte, c.editore, c.anno, c.citta, c.tipo, c.testo]));
}

// Tabella negli appunti, separata da tabulazioni: si incolla così com'è in Excel
async function copiaTabella(col, righe) {
  const pulito = x => String(x || "").replace(/[\t\r\n]+/g, " ");
  const tsv = [col, ...righe].map(r => r.map(pulito).join("\t")).join("\r\n");
  try {
    await navigator.clipboard.writeText(tsv);
  } catch (e) {
    const ta = document.createElement("textarea");
    ta.value = tsv;
    document.body.append(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }
  messaggio("Tabella copiata: incollala in Excel (Ctrl+V).");
}

async function proposteBib() {
  nascondiRisultatiBib();
  messaggio("Lettura delle note…");
  const proposte = await Bibliografia.proposteDalleNote(stileVoci());
  righeBib = proposte.map((p, i) => ({
    chiave: String(i), dati: p, titolo: p.voce, dettaglio: p.dove + " · nella nota: " + p.citazione,
    spuntabile: true, spuntato: true
  }));
  if (!proposte.length) {
    messaggio("Tutte le opere citate per esteso nelle note sono già in bibliografia.\n\n" +
              "Le opere citate solo in forma breve (\"cit.\") non bastano per creare una voce: le segnala il controllo delle citazioni.");
    return;
  }
  mostraRisultatiBib("note", proposte.length + " opere citate nelle note e assenti dalla bibliografia");
  messaggio("Controlla ogni voce e togli la spunta a quelle sbagliate (si possono correggere dopo nel documento).");
}

async function aggiungiBib() {
  const scelte = righeBib.filter(r => r.spuntabile && r.spuntato).map(r => r.dati);
  if (!scelte.length) { messaggio("Nessuna voce spuntata."); return; }
  messaggio("Inserimento delle voci…");
  const esito = await Bibliografia.aggiungiVoci(stileVoci(), scelte);
  nascondiRisultatiBib();
  messaggio(esito);
}

function preparaBibliografia() {
  ricorda("stileVoci");
  const pulsanti = ["sistemaBib", "campiBib", "ordinaBib", "corsiviBib", "estraiBib", "daNoteBib", "aggiungiBib"].map($);
  const su = (id, f) => $(id).addEventListener("click", () => esegui(pulsanti, f));
  su("sistemaBib", () => eseguiBib(() => Bibliografia.sistema(stileVoci(), progresso)));
  su("campiBib", () => eseguiBib(() => Bibliografia.soloOrdineCampi(stileVoci(), progresso)));
  su("ordinaBib", () => eseguiBib(() => Bibliografia.soloOrdina(stileVoci())));
  su("corsiviBib", () => eseguiBib(() => Bibliografia.soloCorsivi(stileVoci(), progresso)));
  su("estraiBib", estraiCampiBib);
  su("daNoteBib", proposteBib);
  su("aggiungiBib", aggiungiBib);
  $("copiaBib").addEventListener("click", () => esegui([], copiaCampiBib));
  $("soloDubbieBib").addEventListener("change", () => mostraRisultatiBib(modoBib, $("titoloBib").textContent));
  $("tutteBib").addEventListener("click", () => { Elenco.spunta(righeBib, true); mostraRisultatiBib(modoBib, $("titoloBib").textContent); });
  $("nessunaBib").addEventListener("click", () => { Elenco.spunta(righeBib, false); mostraRisultatiBib(modoBib, $("titoloBib").textContent); });
}

// ----- Giurisprudenza -----

let righeGiur = [];
let campiGiur = [];
const stileGiur = () => $("stileGiur").value.trim() || "Giurisprudenza";

function disegnaGiur() {
  Elenco.disegna($("elencoGiur"), righeGiur, {
    soloEvidenziate: $("soloDubbieGiur").checked, vuoto: "Nessuna voce da controllare."
  });
}

async function eseguiGiur(azione) {
  $("risultatiGiur").hidden = true;
  messaggio("In corso…");
  messaggio(await azione());
}

async function estraiCampiGiur() {
  $("risultatiGiur").hidden = true;
  campiGiur = await Giurisprudenza.campi(stileGiur());
  const vai = i => () => esegui([], async () => { await Giurisprudenza.vaiAVoce(stileGiur(), i); });
  righeGiur = campiGiur.map(c => ({
    chiave: String(c.i), classe: c.problemi.length ? "dubbio" : "",
    titolo: c.testo,
    dettaglio: [c.autorita, c.sezione, c.provvedimento, c.data, c.deposito, c.numero ? "n. " + c.numero : "", c.parti]
      .filter(Boolean).join(" · ") + (c.problemi.length ? " — " + c.problemi.join("; ") : ""),
    spuntabile: false, vai: vai(c.i)
  }));
  const dubbie = righeGiur.filter(r => r.classe).length;
  $("titoloGiur").textContent = campiGiur.length + " decisioni, " + dubbie + " da controllare";
  $("risultatiGiur").hidden = false;
  disegnaGiur();
  messaggio("In giallo le decisioni da controllare, con il motivo. \"Copia per Excel\" copia la tabella completa.");
}

async function copiaCampiGiur() {
  await copiaTabella(["Autorità", "Sezione", "Provv.", "Data", "Deposito", "Numero", "Parti", "Fonte", "Note", "Voce"],
    campiGiur.map(c => [c.autorita, c.sezione, c.provvedimento, c.data, c.deposito, c.numero, c.parti, c.fonte, c.problemi.join("; "), c.testo]));
}

function preparaGiurisprudenza() {
  ricorda("stileGiur");
  const pulsanti = ["sistemaGiur", "normalizzaGiur", "ordinaGiur", "corsiviGiur", "controllaGiur", "estraiGiur"].map($);
  const su = (id, f) => $(id).addEventListener("click", () => esegui(pulsanti, f));
  su("sistemaGiur", () => eseguiGiur(() => Giurisprudenza.sistema(stileGiur(), progresso)));
  su("normalizzaGiur", () => eseguiGiur(() => Giurisprudenza.soloNormalizza(stileGiur(), progresso)));
  su("ordinaGiur", () => eseguiGiur(() => Giurisprudenza.soloOrdina(stileGiur())));
  su("corsiviGiur", () => eseguiGiur(() => Giurisprudenza.soloCorsivi(stileGiur(), progresso)));
  su("controllaGiur", () => eseguiGiur(() => Giurisprudenza.controlla(stileGiur())));
  su("estraiGiur", estraiCampiGiur);
  $("copiaGiur").addEventListener("click", () => esegui([], copiaCampiGiur));
  $("soloDubbieGiur").addEventListener("change", disegnaGiur);
}

// ----- Abbreviazioni -----

let righeAbbr = [];
const ambitoAbbr = () => $("ambitoAbbr").value;

function disegnaAbbr() {
  Elenco.disegna($("elencoAbbr"), righeAbbr, { soloEvidenziate: $("soloDubbieAbbr").checked, vuoto: "Nessuna voce da decidere." });
}

async function controllaAbbr() {
  messaggio("Lettura del testo e delle note…");
  const ambito = ambitoAbbr();
  const r = await Abbreviazioni.controlla(ambito, avanzamento("Ricerca nel documento"));
  const vai = chiave => () => esegui([], async () => {
    if (!(await Abbreviazioni.vaiA(ambito, chiave))) messaggio("Non trovato: il documento è cambiato, controlla di nuovo.");
  });

  righeAbbr = r.trovate.map(t => ({
    chiave: t.chiave, classe: t.applicabile ? "" : "dubbio",
    titolo: t.trovato + "  →  " + (t.nuovo === null ? "(da valutare)" : t.nuovo.trim()) + (t.stato ? "   (" + t.stato + ")" : ""),
    dettaglio: t.dove + " · " + t.regola, testo: t.contesto,
    spuntabile: t.spuntabile, spuntato: t.applicabile, vai: t.stato === "non individuato" ? null : vai(t.chiave)
  }));

  $("risultatiAbbr").hidden = righeAbbr.length === 0;
  disegnaAbbr();
  messaggio(righeAbbr.length === 0
    ? "Nessuna abbreviazione da uniformare (" + r.descrizione + ")."
    : "Abbreviazioni (" + r.descrizione + "): " + righeAbbr.length + " forme trovate, " + r.daCorreggere +
      " sicure (già spuntate). In giallo, senza spunta, quelle da decidere: spunta quelle che vuoi correggere.");
}

async function correggiAbbr(soloScelte) {
  const chiavi = soloScelte ? Elenco.scelte(righeAbbr) : null;
  if (soloScelte && !chiavi.length) { messaggio("Nessuna voce spuntata."); return; }
  messaggio("Correzione in corso…");
  const esito = await Abbreviazioni.correggi(ambitoAbbr(), chiavi, avanzamento("Ricerca nel documento"));
  righeAbbr = [];
  $("risultatiAbbr").hidden = true;
  messaggio(esito);
}

function preparaAbbreviazioni() {
  ricorda("ambitoAbbr");
  const pulsanti = ["controllaAbbr", "correggiAbbr", "correggiScelteAbbr"].map($);
  $("controllaAbbr").addEventListener("click", () => esegui(pulsanti, controllaAbbr));
  $("correggiAbbr").addEventListener("click", () => esegui(pulsanti, () => correggiAbbr(false)));
  $("correggiScelteAbbr").addEventListener("click", () => esegui(pulsanti, () => correggiAbbr(true)));
  $("soloDubbieAbbr").addEventListener("change", disegnaAbbr);
  $("tutteAbbr").addEventListener("click", () => { Elenco.spunta(righeAbbr, true); disegnaAbbr(); });
  $("nessunaAbbr").addEventListener("click", () => { Elenco.spunta(righeAbbr, false); disegnaAbbr(); });
}

// ----- Foto (Relazioni Tecniche) -----

async function fotoScelte() {
  const files = $("fileFoto").files;
  if (!files || !files.length) throw new Error("Scegli prima le foto con \"Foto da inserire\".");
  const foto = await preparaFoto(files, $("riduciFoto").checked, (f, t) => messaggio("Lettura delle foto: " + f + " di " + t + "…"));
  if (!foto.length) throw new Error("Tra i file scelti non ci sono immagini.");
  return foto;
}

const fotoInserite = (f, t) => messaggio("Inserimento: foto " + f + " di " + t + "…");

async function esportaPdf() {
  messaggio("Word prepara il PDF…");
  const blob = await Foto.pdf();
  const a = $("scaricaPdf");
  if (a.href) URL.revokeObjectURL(a.href);
  a.href = URL.createObjectURL(blob);
  let nome = "documento";
  try { nome = decodeURIComponent((Office.context.document.url || "").split(/[\\/]/).pop()).replace(/\.[^.]+$/, "") || nome; } catch (e) { /* senza nome */ }
  a.download = nome + ".pdf";
  a.hidden = false;
  a.click();
  messaggio("PDF pronto (" + (blob.size / 1048576).toFixed(1).replace(".", ",") + " MB). Se il download non parte, clicca \"Scarica il PDF\"; " +
            "in alternativa usa File → Esporta → PDF.");
}

function preparaFotoPannello() {
  const ids = ["ins1", "ins2", "ins4", "aggiungiFoto", "lay1", "lay2", "lay4", "indietro", "avanti", "riorganizza", "rinumera", "capitolo", "comprimi", "esportaPdf"];
  const pulsanti = ids.map($);
  const su = (id, f) => $(id).addEventListener("click", () => esegui(pulsanti, async () => { const r = await f(); messaggio(r || "Fatto."); }));
  for (const n of [1, 2, 4]) {
    su("ins" + n, async () => Foto.inserisci(await fotoScelte(), n, fotoInserite));
    su("lay" + n, () => Foto.cambiaLayout(n, fotoInserite));
  }
  su("aggiungiFoto", async () => Foto.aggiungi(await fotoScelte(), fotoInserite));
  su("indietro", () => Foto.sposta(-1));
  su("avanti", () => Foto.sposta(+1));
  su("riorganizza", () => Foto.riorganizza(fotoInserite));
  su("rinumera", () => Foto.rinumera());
  su("capitolo", () => Foto.capitoloFoto());
  su("comprimi", () => Foto.comprimi(b64 => riduci(b64), (f, t) => messaggio("Compressione: immagine " + f + " di " + t + "…")));
  $("esportaPdf").addEventListener("click", () => esegui(pulsanti, esportaPdf));
}

function preparaNumeri() {
  const pulsanti = ["numEuro", "numFin", "numAmm"].map($);
  const su = (id, formato) => $(id).addEventListener("click", () => esegui(pulsanti, async () => {
    await Numeri.converti(formato);
    messaggio("Fatto.");
  }));
  su("numEuro", "euro");
  su("numFin", "finanziario");
  su("numAmm", "amministrativo");
}

// ----- avvio -----

function mostraPannello(nome) {
  document.querySelectorAll(".pannello").forEach(p => (p.hidden = p.id !== nome));
  $("pannello").value = nome;
  messaggio("");
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) {
    messaggio("Questo add-in funziona solo in Word.", true);
    return;
  }
  if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
    messaggio("Questa versione di Word è troppo vecchia per l'add-in: aggiorna Word (Microsoft 365).", true);
    return;
  }

  preparaSommari();
  preparaCitazioni();
  preparaBibliografia();
  preparaGiurisprudenza();
  preparaAbbreviazioni();
  preparaFotoPannello();
  preparaNumeri();

  // due add-in con gli stessi file: Relazioni Scientifiche e Relazioni Tecniche, ciascuno con i suoi strumenti
  const parametri = new URLSearchParams(location.search);
  const tecniche = parametri.get("scheda") === "tecniche";
  for (const g of document.querySelectorAll("#pannello optgroup"))
    if ((g.label === "Relazioni Tecniche") !== tecniche) g.remove();
  const ammessi = [...document.querySelectorAll("#pannello option")].map(o => o.value);
  const richiesto = parametri.get("pannello");
  mostraPannello(ammessi.includes(richiesto) ? richiesto : tecniche ? "foto" : "citazioni");
  $("pannello").addEventListener("change", e => mostraPannello(e.target.value));
});
