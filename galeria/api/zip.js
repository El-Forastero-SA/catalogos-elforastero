// Descarga masiva: recibe una lista de SKUs y arma un ZIP al vuelo con los
// originales de Magento (streaming). El media de Magento es público (sin token).
const archiver = require("archiver");
const manifest = require("../manifest.json");

const BY_SKU = new Map(manifest.productos.map((p) => [p.sku, p]));
const MEDIA_BASE = manifest.media_base;
const ZIP_MAX = 400;        // guardrail servidor
const CONCURRENCY = 6;      // descargas Magento en paralelo

function slug(s) {
  return (s || "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 70);
}

async function fetchBuffer(p) {
  try {
    const r = await fetch(MEDIA_BASE + p.img, { headers: { "User-Agent": "ForasteroGaleria/1.0" } });
    if (!r.ok) return null;
    return Buffer.from(await r.arrayBuffer());
  } catch (e) {
    return null;
  }
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.statusCode = 405;
    return res.end("Método no permitido");
  }
  let body = req.body;
  if (typeof body === "string") { try { body = JSON.parse(body); } catch { body = {}; } }
  const skus = (body && Array.isArray(body.skus)) ? body.skus.map(String) : [];
  const nombre = slug((body && body.nombre) || "Imagenes") || "Imagenes";

  const productos = skus.map((s) => BY_SKU.get(s)).filter(Boolean).slice(0, ZIP_MAX);
  if (!productos.length) {
    res.statusCode = 400;
    return res.end("Sin imágenes válidas");
  }

  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${nombre}.zip"`);

  const archive = archiver("zip", { zlib: { level: 6 } });
  archive.on("error", () => { try { res.destroy(); } catch {} });
  archive.pipe(res);

  // Descarga con concurrencia limitada; agrega al zip a medida que llegan.
  let i = 0;
  const usados = new Set();
  async function worker() {
    while (i < productos.length) {
      const p = productos[i++];
      const buf = await fetchBuffer(p);
      if (!buf) continue;
      let name = `${p.sku}_${slug(p.nombre)}${p.ext}`;
      while (usados.has(name)) name = `${p.sku}_${Math.random().toString(36).slice(2, 6)}${p.ext}`;
      usados.add(name);
      archive.append(buf, { name });
    }
  }
  await Promise.all(Array.from({ length: CONCURRENCY }, worker));
  await archive.finalize();
};
