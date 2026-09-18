#!/usr/bin/env python3
"""
generar_manifest.py — Genera el manifest.json + miniaturas de la galería pública
de imágenes de producto (imagenes.elforastero.com.ar).

Qué hace:
  1. Por cada categoría del catálogo, baja la lista de productos de
     listas.elforastero.com.ar (mismos productos que el catálogo PDF).
  2. Consulta Magento por SKU para obtener la ruta de imagen (/pub/media/...),
     la especie (atributo tipomascota) y la marca (atributo tipomarca).
  3. Descarga la imagen original de Magento (media público, sin auth) y genera
     una miniatura webp en thumbs/<sku>.webp.
  4. Escribe manifest.json con todos los productos que tienen imagen.

Inputs:
  - catalogo-template/.env  (MAGENTO_BASE_URL, MAGENTO_ACCESS_TOKEN)  [ver --env]
  - listas.elforastero.com.ar  (público)

Outputs (en el directorio del script):
  - manifest.json
  - thumbs/<sku>.webp

Uso:
  python3 generar_manifest.py                      # todas las categorías
  python3 generar_manifest.py --solo "Bebidas"     # una sola (iterar rápido)
  python3 generar_manifest.py --limit 10           # tope por categoría (test)
  python3 generar_manifest.py --force              # regenera thumbs existentes
"""

import argparse
import io
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CATALOGO_TEMPLATE = os.path.join(SCRIPT_DIR, "..", "catalogo-template")
sys.path.insert(0, CATALOGO_TEMPLATE)

# Reutiliza utilidades del pipeline de catálogos
from generar_desde_lista import load_env, magento_get_product  # noqa: E402

LISTAS_BASE_URL = "https://listas.elforastero.com.ar"
SUCURSAL_LISTA = "Cervantes|Deposito"
THUMB_MAX = 400          # lado mayor de la miniatura, px
THUMB_QUALITY = 80

# Categorías del catálogo (mismas 7 que actualizar_catalogos.py, sin Accesorios).
# slug + color para la UI de la galería.
CATEGORIAS = [
    {"nombre": "Mascotas - Alimentos Super Premium",            "slug": "super-premium",   "color": "#BC1221"},
    {"nombre": "Mascotas - Alimentos Húmedos Snack Absorbentes", "slug": "humedos-snack",   "color": "#BC1221"},
    {"nombre": "Mascotas - Alimentos Premium y Standard",       "slug": "premium-standard", "color": "#BC1221"},
    {"nombre": "Equinos",                                       "slug": "equinos",         "color": "#C96A2B"},
    {"nombre": "Alimentos Saludables",                          "slug": "saludables",      "color": "#7AAD8A"},
    {"nombre": "Bebidas",                                       "slug": "bebidas",         "color": "#2E7D9A"},
    {"nombre": "Comestibles e Higiene",                         "slug": "comestibles",     "color": "#F5B800"},
]


# =============================================================================
# Descarga de listas (idéntico a actualizar_catalogos.descargar_lista)
# =============================================================================

def descargar_lista(categoria, dest_dir):
    """POST /generate + GET /preview → xlsx local. Devuelve (path, total)."""
    boundary = "----galeriamanifest"
    parts = [
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"categorias\"\r\n\r\n{categoria}\r\n",
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"sl_list\"\r\n\r\n{SUCURSAL_LISTA}\r\n",
        f"--{boundary}--\r\n",
    ]
    body = "".join(parts).encode("utf-8")
    req = urllib.request.Request(f"{LISTAS_BASE_URL}/generate", data=body, method="POST")
    req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if not data.get("ok"):
        raise RuntimeError(f"/generate ok=False para '{categoria}': {data}")
    res = (data.get("results") or [{}])[0]
    if res.get("error") or not res.get("archivo_xlsx"):
        raise RuntimeError(f"/generate sin archivo para '{categoria}': {res}")
    xlsx_remote = res["archivo_xlsx"]
    url = f"{LISTAS_BASE_URL}/preview?path={urllib.parse.quote(xlsx_remote)}"
    local = os.path.join(dest_dir, os.path.basename(xlsx_remote))
    with urllib.request.urlopen(url, timeout=120) as resp, open(local, "wb") as f:
        f.write(resp.read())
    return local, res.get("total_productos")


def leer_productos_xlsx(path):
    """[{sku, nombre}] desde el xlsx de lista (col A = SKU, col B = descripción)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    out = []
    for row in ws.iter_rows(min_row=6, max_col=2):
        va = row[0].value
        if va is None:
            continue
        try:
            sku = str(int(float(str(va).strip()))).zfill(5)
        except (ValueError, TypeError):
            continue
        nombre = str(row[1].value or "").strip()
        if nombre:
            out.append({"sku": sku, "nombre": nombre})
    wb.close()
    return out


# =============================================================================
# Magento: opciones de atributos (id → label) e imagen
# =============================================================================

def fetch_attribute_options(base, token, code):
    url = f"{base.rstrip('/')}/rest/V1/products/attributes/{code}/options"
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        opts = json.loads(r.read().decode())
    return {str(o["value"]): o["label"] for o in opts if o.get("value")}


def product_attrs(product):
    return {a["attribute_code"]: a.get("value") for a in product.get("custom_attributes", [])}


def image_path_from_product(product):
    attrs = product_attrs(product)
    for k in ("image", "small_image", "thumbnail"):
        v = attrs.get(k)
        if v and v != "no_selection":
            return v
    return None


# =============================================================================
# Miniaturas
# =============================================================================

def descargar_bytes(url, timeout=60):
    req = urllib.request.Request(url, headers={"User-Agent": "ForasteroGaleria/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def generar_thumb(img_bytes, dest_path):
    """Redimensiona a THUMB_MAX (lado mayor) sobre fondo blanco y guarda webp."""
    from PIL import Image
    im = Image.open(io.BytesIO(img_bytes))
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im).convert("RGB")
    else:
        im = im.convert("RGB")
    im.thumbnail((THUMB_MAX, THUMB_MAX), Image.LANCZOS)
    im.save(dest_path, "WEBP", quality=THUMB_QUALITY, method=6)


# =============================================================================
# Utilidades
# =============================================================================

def slug(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")[:70]


# =============================================================================
# Main
# =============================================================================

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", default=os.path.join(CATALOGO_TEMPLATE, ".env"))
    ap.add_argument("--solo", help="Procesar solo esta categoría (nombre exacto)")
    ap.add_argument("--limit", type=int, help="Tope de productos por categoría (test)")
    ap.add_argument("--force", action="store_true", help="Regenerar thumbs existentes")
    args = ap.parse_args()

    env = load_env(args.env)
    base = env["MAGENTO_BASE_URL"].rstrip("/")
    token = env["MAGENTO_ACCESS_TOKEN"]

    thumbs_dir = os.path.join(SCRIPT_DIR, "thumbs")
    os.makedirs(thumbs_dir, exist_ok=True)
    listas_dir = os.path.join(SCRIPT_DIR, ".listas_cache")
    os.makedirs(listas_dir, exist_ok=True)

    print("Bajando diccionarios de atributos (especie, marca)...")
    especies = fetch_attribute_options(base, token, "tipomascota")   # 45 Perro, 44 Gato...
    marcas = fetch_attribute_options(base, token, "tipomarca")

    cats = [c for c in CATEGORIAS if not args.solo or c["nombre"] == args.solo]
    if args.solo and not cats:
        print(f"Categoría no encontrada: {args.solo}", file=sys.stderr)
        sys.exit(1)

    productos = []
    resumen_cat = []
    total_sin_img = 0
    seen_sku = set()

    for cat in cats:
        print(f"\n=== {cat['nombre']} ===")
        xlsx = None
        for intento in range(1, 4):
            try:
                xlsx, total = descargar_lista(cat["nombre"], listas_dir)
                break
            except Exception as e:
                print(f"  Intento {intento}/3 bajando lista falló: {e}", file=sys.stderr)
                time.sleep(5 * intento)
        if xlsx is None:
            print(f"  ERROR: no se pudo bajar la lista de '{cat['nombre']}' tras 3 intentos",
                  file=sys.stderr)
            raise SystemExit(f"Abortado: falta la categoría '{cat['nombre']}'")
        items = leer_productos_xlsx(xlsx)
        if args.limit:
            items = items[:args.limit]
        print(f"  {len(items)} productos en lista")

        cat_count = 0
        for it in items:
            sku, nombre = it["sku"], it["nombre"]
            if sku in seen_sku:
                continue
            try:
                prod = magento_get_product(base, token, sku)
                if prod is None:  # probar sin padding
                    prod = magento_get_product(base, token, str(int(sku)))
            except Exception as e:
                print(f"  ERROR Magento {sku}: {e}", file=sys.stderr)
                continue
            if prod is None:
                continue
            img_path = image_path_from_product(prod)
            if not img_path:
                total_sin_img += 1
                continue

            attrs = product_attrs(prod)
            especie = especies.get(str(attrs.get("tipomascota")), "")
            marca = marcas.get(str(attrs.get("tipomarca")), "")
            if marca in ("0", ""):
                marca = "Sin marca"

            ext = os.path.splitext(img_path)[1].lower() or ".jpg"
            if ext not in (".jpg", ".jpeg", ".png", ".webp"):
                ext = ".jpg"

            thumb_path = os.path.join(thumbs_dir, f"{sku}.webp")
            if args.force or not os.path.exists(thumb_path):
                try:
                    raw = descargar_bytes(f"{base}/pub/media/catalog/product{img_path}")
                    generar_thumb(raw, thumb_path)
                except Exception as e:
                    print(f"  ERROR thumb {sku}: {e}", file=sys.stderr)
                    continue

            seen_sku.add(sku)
            cat_count += 1
            productos.append({
                "sku": sku,
                "nombre": nombre,
                "categoria": cat["nombre"],
                "categoria_slug": cat["slug"],
                "especie": especie,
                "marca": marca,
                "img": img_path,   # ruta en /pub/media/catalog/product<img>
                "ext": ext,
            })
        print(f"  → {cat_count} con imagen")
        resumen_cat.append({**cat, "count": cat_count})

    manifest = {
        "generado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "media_base": f"{base}/pub/media/catalog/product",
        "categorias": resumen_cat,
        "productos": productos,
    }
    out_path = os.path.join(SCRIPT_DIR, "manifest.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)

    print("\n========== RESUMEN ==========")
    for c in resumen_cat:
        print(f"  {c['count']:>4}  {c['nombre']}")
    print(f"  Total con imagen: {len(productos)}")
    print(f"  Sin imagen (omitidos): {total_sin_img}")
    print(f"  manifest.json: {out_path}")


if __name__ == "__main__":
    main()
