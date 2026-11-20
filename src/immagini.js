// Foto scelte dal computer: lettura e riduzione nel browser (prima di mandarle a Word)

export const LATO_MASSIMO = 1600;      // pixel del lato lungo dopo la riduzione
const QUALITA = 0.85;

const comeBase64 = dataUrl => dataUrl.slice(dataUrl.indexOf(",") + 1);

function leggiFile(file) {
  return new Promise((ok, ko) => {
    const r = new FileReader();
    r.onload = () => ok(r.result);
    r.onerror = () => ko(new Error("Impossibile leggere " + file.name));
    r.readAsDataURL(file);
  });
}

function carica(src) {
  return new Promise((ok, ko) => {
    const img = new Image();
    img.onload = () => ok(img);
    img.onerror = () => ko(new Error("immagine non leggibile"));
    img.src = src;
  });
}

// Riduce un'immagine (data URL o base64 JPEG/PNG): lato lungo al massimo maxPx, JPEG.
// Restituisce il base64 ridotto, oppure null se non serve o non è possibile (es. TIFF)
export async function riduci(src, maxPx = LATO_MASSIMO) {
  if (!src.startsWith("data:")) {
    const tipo = src.startsWith("/9j/") ? "image/jpeg" : src.startsWith("iVBOR") ? "image/png" : null;
    if (!tipo) return null;
    // le immagini PNG piccole (loghi, schemi) restano come sono: hanno trasparenze e linee nette
    if (tipo === "image/png" && src.length < 400000) return null;
    src = "data:" + tipo + ";base64," + src;
  }

  let img;
  try { img = await carica(src); } catch (e) { return null; }

  const lato = Math.max(img.naturalWidth, img.naturalHeight);
  const k = lato > maxPx ? maxPx / lato : 1;
  const w = Math.round(img.naturalWidth * k), h = Math.round(img.naturalHeight * k);

  const c = document.createElement("canvas");
  c.width = w;
  c.height = h;
  const g = c.getContext("2d");
  g.fillStyle = "#ffffff";                // le trasparenze diventano bianche
  g.fillRect(0, 0, w, h);
  g.drawImage(img, 0, 0, w, h);
  return comeBase64(c.toDataURL("image/jpeg", QUALITA));
}

// File scelti → [{ nome, base64 }] in ordine di nome, ridotti se richiesto
export async function preparaFoto(files, ridurre, avanzamento) {
  const elenco = [...files].filter(f => /^image\//.test(f.type) || /\.(jpe?g|png|bmp|gif|tiff?)$/i.test(f.name))
    .sort((a, b) => a.name.localeCompare(b.name, "it", { numeric: true }));
  const res = [];
  for (let i = 0; i < elenco.length; i++) {
    if (avanzamento) avanzamento(i, elenco.length);
    const f = elenco[i];
    const dataUrl = await leggiFile(f);
    const originale = comeBase64(dataUrl);
    const ridotta = ridurre ? await riduci(dataUrl) : null;
    res.push({ nome: f.name, base64: ridotta && ridotta.length < originale.length ? ridotta : originale });
  }
  return res;
}
