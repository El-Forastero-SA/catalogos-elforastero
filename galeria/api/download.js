// Descarga de 1 imagen original: proxy a Magento con nombre de archivo prolijo.
// El media de Magento (/pub/media/...) es público; no requiere token.
const manifest = require("../manifest.json");

const BY_SKU = new Map(manifest.productos.map((p) => [p.sku, p]));
const MEDIA_BASE = manifest.media_base;

const MIME = { ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp" };

function slug(s) {
  return (s || "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 70);
}

module.exports = async (req, res) => {
  const sku = (req.query && req.query.sku) || "";
  const p = BY_SKU.get(String(sku));
  if (!p) {
    res.statusCode = 404;
    return res.end("Producto no encontrado");
  }
  try {
    const upstream = await fetch(MEDIA_BASE + p.img, {
      headers: { "User-Agent": "ForasteroGaleria/1.0" },
    });
    if (!upstream.ok) {
      res.statusCode = 502;
      return res.end("No se pudo obtener la imagen");
    }
    const buf = Buffer.from(await upstream.arrayBuffer());
    const filename = `${p.sku}_${slug(p.nombre)}${p.ext}`;
    res.setHeader("Content-Type", MIME[p.ext] || "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Cache-Control", "public, max-age=86400");
    res.statusCode = 200;
    res.end(buf);
  } catch (e) {
    res.statusCode = 500;
    res.end("Error al descargar la imagen");
  }
};
