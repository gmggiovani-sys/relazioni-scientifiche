// Elenco con caselle di spunta per il pannello (come le finestre della versione Windows)
// righe: [{ chiave, titolo, dettaglio, testo, classe, spuntabile, spuntato, vai }]

export function disegna(ul, righe, { soloEvidenziate = false, vuoto = "Niente da mostrare." } = {}) {
  ul.textContent = "";
  let mostrate = 0;

  for (const r of righe) {
    if (soloEvidenziate && r.classe !== "daFare" && r.classe !== "dubbio") continue;
    mostrate++;

    const li = document.createElement("li");
    if (r.classe) li.className = r.classe;

    const testa = document.createElement("label");
    testa.className = "testa";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.disabled = !r.spuntabile;
    cb.checked = !!r.spuntato && !!r.spuntabile;
    cb.addEventListener("change", () => (r.spuntato = cb.checked));
    if (!r.spuntabile) cb.style.visibility = "hidden";

    const t = document.createElement("span");
    t.className = "titolo";
    const strong = document.createElement("strong");
    strong.textContent = r.titolo;
    t.append(strong);
    if (r.dettaglio) {
      const small = document.createElement("small");
      small.textContent = r.dettaglio;
      t.append(small);
    }
    testa.append(cb, t);
    li.append(testa);

    if (r.testo) {
      const d = document.createElement("div");
      d.className = "anteprima";
      d.textContent = r.testo;
      li.append(d);
    }

    if (r.vai) {
      const azioni = document.createElement("div");
      azioni.className = "azioni";
      const a = document.createElement("a");
      a.textContent = "Vai";
      a.addEventListener("click", r.vai);
      azioni.append(a);
      li.append(azioni);
    }

    ul.append(li);
  }

  if (!mostrate) {
    const li = document.createElement("li");
    li.textContent = vuoto;
    ul.append(li);
  }
}

export function spunta(righe, valore) {
  for (const r of righe) if (r.spuntabile) r.spuntato = valore;
}

export const scelte = righe => righe.filter(r => r.spuntabile && r.spuntato).map(r => r.chiave);
