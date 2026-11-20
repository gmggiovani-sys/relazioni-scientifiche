# Relazioni Scientifiche – add-in per Word (Mac e Windows)

Versione "Office Add-in" degli strumenti della scheda Relazioni Scientifiche. Funziona in Word per Mac,
Word per Windows e Word sul web. I file del programma stanno su Netlify; i documenti restano sul
computer e non passano da lì.

Contiene: **Bibliografia**, **Giurisprudenza**, **Abbreviazioni**, **Citazioni nelle note** e **Sommari dei capitoli**, come la scheda Relazioni Scientifiche della versione Windows.
abbreviazioni, citazioni) si aggiungono man mano.

## Due add-in, stessi file

- `manifest.xml` → scheda **Relazioni Scientifiche** (bibliografia, giurisprudenza, abbreviazioni, citazioni, sommari).
- `manifest-tecniche.xml` → scheda **Relazioni Tecniche** (tabelle di foto, compressione, PDF, numeri in lettere).

Si installano allo stesso modo (anche tutti e due); usano gli stessi file pubblicati su Netlify.

## 1. Pubblicare su Netlify

Il sito è `https://effulgent-moonbeam-a555b4.netlify.app/` (gli indirizzi sono già in `manifest.xml`).

- **Primo caricamento / aggiornamento**: su <https://app.netlify.com>, apri il sito, scheda **Deploys**,
  e trascina la cartella `RelazioniScientificheWeb` nel riquadro in fondo alla pagina.
- **Prova**: `https://effulgent-moonbeam-a555b4.netlify.app/taskpane.html` deve aprirsi nel browser
  (il pannello senza Word mostra solo i campi, è normale).
- Se cambi il nome del sito (Site configuration → Change site name), sostituisci il vecchio indirizzo
  in `manifest.xml` e reinstallalo.

## 2. Installare in Word per Mac

Serve Word di Microsoft 365 aggiornato (per il maiuscoletto almeno la versione 16.99).

1. Chiudi Word.
2. Nel Finder: **Vai → Vai alla cartella…** e incolla
   `~/Library/Containers/com.microsoft.Word/Data/Documents/`
3. Se non c'è, crea la cartella `wef`; copia lì dentro `manifest.xml`.
4. Apri Word e un documento: **Home → Componenti aggiuntivi** (o **Inserisci → Componenti aggiuntivi →
   I miei componenti aggiuntivi**) e scegli **Relazioni Scientifiche**. Compare la scheda
   **Relazioni Scientifiche**.

## 3. Installare in Word per Windows (facoltativo)

Su Windows c'è già la versione completa (VSTO). Se vuoi provare anche questa:

1. Metti `manifest.xml` in una cartella condivisa, ad esempio `\\DESKTOP-4H670LG\Install\RelazioniScientificheWeb`.
2. Word → **File → Opzioni → Centro protezione → Impostazioni Centro protezione → Cataloghi
   componenti aggiuntivi attendibili**: aggiungi il percorso della cartella, spunta **Mostra nel menu**, OK.
3. Riavvia Word: **Home → Componenti aggiuntivi → Altri componenti aggiuntivi → CARTELLA CONDIVISA**.

Con entrambe le versioni installate vedrai due schede con lo stesso nome: va bene, ma tienine presente.

## Aggiornamenti

- Se cambiano i file del programma (`.js`, `.html`, `.css`): basta ricaricarli su Netlify (scheda Deploys).
  Word può tenere in memoria la versione precedente per un po'; chiudere e riaprire il pannello di solito basta.
  Su Mac, se resta la vecchia: chiudi Word e svuota
  `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`.
- Se cambia `manifest.xml` (nuovi pulsanti nella scheda): ricopialo nella cartella `wef` e riavvia Word.

## Differenze rispetto alla versione Windows (VSTO)

- Le finestre con le spunte sono nel pannello laterale.
- Non c'è un unico "Annulla" per tutto il comando: ogni modifica si annulla a parte.
  Con le revisioni attive tutto resta comunque rivedibile.
- Il maiuscoletto richiede Word per Mac 16.99 o successivo; nelle versioni precedenti il comando avvisa.

## File

| File | A cosa serve |
|---|---|
| `manifest.xml` | descrive l'add-in a Word: scheda, pulsanti, indirizzi |
| `taskpane.html`, `taskpane.css` | il pannello laterale |
| `src/app.js` | collega il pannello ai comandi |
| `src/sommari.js` | Sommari dei capitoli (lettura e scrittura nel documento) |
| `src/sommari-testo.js` | composizione del testo del sommario (senza Word) |
| `src/bibliografia.js` | Bibliografia (ordine, Editore/Città, corsivi e maiuscoletto, campi, voci dalle note) |
| `src/giurisprudenza.js` | Giurisprudenza (normalizza, ordina per data, corsivi, controlla, campi) |
| `src/abbreviazioni.js` | Abbreviazioni (controlla, correggi le sicure o le scelte) |
| `src/elenchi.js` | passaggi comuni agli elenchi (sostituzioni nel testo, riordino) |
| `src/foto.js`, `src/immagini.js` | Relazioni Tecniche: tabelle di foto, riduzione delle immagini, PDF |
| `src/numeri.js`, `src/parser/numeri.js` | Numeri in lettere |
| `src/comandi.js`, `avviso.html` | pulsanti della scheda che lavorano senza pannello, e i loro messaggi |
| `src/citazioni.js` | Citazioni nelle note (lettura delle note, controllo, formattazione) |
| `src/parser/` | regole di bibliografia, citazioni e giurisprudenza (senza Word), come nella versione Windows |
| `src/word.js`, `src/elenco.js` | funzioni comuni: ricerca nel documento, elenco con le spunte |
| `commands.html` | richiesto da Word per i pulsanti della scheda |
| `index.html` | pagina di presentazione del sito |
| `assets/` | icone |
