#!/usr/bin/env python3
"""
actualizar_catalogos.py — Regenera los 8 catálogos PDF y deploya a Vercel via git push.

Uso:
    python3 actualizar_catalogos.py

Flujo:
    1. Descarga las 7 listas Nexion (Cervantes Depósito) desde
       listas.elforastero.com.ar — flujo automático, sin Excel manuales.
    2. Para Accesorios (Cocooning/Hunter): usa el Excel local fijo
       (~/Downloads/lista precios 24-02-26.xlsx) — manual por ahora.
    3. Regenera los PDFs desde cada lista.
    4. Sincroniza los contadores del index.html con el count real de cada PDF.
    5. Commit + push → Vercel re-deploya automáticamente.

Requisitos:
    - generar_desde_lista.py + generar_accesorios.py en ../catalogo-template/
    - .env con credenciales Magento en ../catalogo-template/
    - playwright instalado (pip3 install playwright && python3 -m playwright install chromium)
    - git configurado con remote origin
    - listas.elforastero.com.ar accesible (Nexion base CSV ya cargado en el sitio)
"""

import os
import re
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
from datetime import datetime

# =============================================================================
# Configuración
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CATALOGO_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "catalogo-template")
PDFS_DIR = os.path.join(SCRIPT_DIR, "pdfs")
INDEX_HTML = os.path.join(SCRIPT_DIR, "index.html")
GENERATOR = os.path.join(CATALOGO_DIR, "generar_desde_lista.py")
GENERATOR_ACCESORIOS = os.path.join(CATALOGO_DIR, "generar_accesorios.py")

# Endpoint de listas.elforastero.com.ar (genera + descarga listas Cervantes)
LISTAS_BASE_URL = "https://listas.elforastero.com.ar"
SUCURSAL_LISTA = "Cervantes|Deposito"

# Regex que matchea la línea "Productos: N" del resumen del generador
PRODUCT_COUNT_RE = re.compile(r"Productos:\s+(\d+)")

# Configuración de catálogos.
# Para los 7 Nexion: "categoria" = nombre exacto de la categoría en
# listas.elforastero.com.ar. Lista descargada al vuelo.
# Para Accesorios: "excel" = path local fijo (Cocooning/Hunter, manual).
CATALOGOS = [
    {
        "categoria": "Mascotas - Alimentos Super Premium",
        "pdf_name": "alimentos-super-premium.pdf",
    },
    {
        "categoria": "Mascotas - Alimentos Húmedos Snack Absorbentes",
        "pdf_name": "alimentos-humedos-snack-absorbentes.pdf",
    },
    {
        "categoria": "Mascotas - Alimentos Premium y Standard",
        "pdf_name": "alimentos-premium-y-standard.pdf",
    },
    {
        "categoria": "Equinos",
        "pdf_name": "equinos.pdf",
    },
    {
        "categoria": "Alimentos Saludables",
        "pdf_name": "alimentos-saludables.pdf",
    },
    {
        "categoria": "Bebidas",
        "pdf_name": "bebidas.pdf",
    },
    {
        "categoria": "Comestibles e Higiene",
        "pdf_name": "comestibles-e-higiene.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/lista precios 24-02-26.xlsx"),
        "pdf_name": "accesorios.pdf",
        "generator": GENERATOR_ACCESORIOS,
    },
]


# =============================================================================
# Descarga automática de listas desde listas.elforastero.com.ar
# =============================================================================

def descargar_lista(categoria, dest_dir):
    """
    Genera la lista en listas.elforastero.com.ar (POST /generate) y descarga
    el XLSX resultante (GET /preview?path=...) a dest_dir. Devuelve el path local.

    Eleva RuntimeError si algo falla.
    """
    # 1) POST /generate con multipart/form-data manual
    boundary = "----actualizarcatalogos"
    parts = []
    for cat in [categoria]:
        parts.append(f"--{boundary}\r\n"
                     f'Content-Disposition: form-data; name="categorias"\r\n\r\n'
                     f"{cat}\r\n")
    parts.append(f"--{boundary}\r\n"
                 f'Content-Disposition: form-data; name="sl_list"\r\n\r\n'
                 f"{SUCURSAL_LISTA}\r\n")
    parts.append(f"--{boundary}--\r\n")
    body = "".join(parts).encode("utf-8")

    req = urllib.request.Request(
        f"{LISTAS_BASE_URL}/generate",
        data=body,
        method="POST",
    )
    req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            import json
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        raise RuntimeError(f"POST /generate falló para '{categoria}': {e}")

    if not data.get("ok"):
        raise RuntimeError(f"/generate respondió ok=False para '{categoria}': {data}")

    results = data.get("results") or []
    if not results:
        raise RuntimeError(f"/generate no devolvió results para '{categoria}'")

    res = results[0]
    if res.get("error"):
        raise RuntimeError(f"/generate error para '{categoria}': {res['error']}")

    xlsx_remote = res.get("archivo_xlsx")
    if not xlsx_remote:
        raise RuntimeError(f"/generate no devolvió archivo_xlsx para '{categoria}'")

    # 2) GET /preview?path=... y guardar en dest_dir
    url = f"{LISTAS_BASE_URL}/preview?path={urllib.parse.quote(xlsx_remote)}"
    local_name = os.path.basename(xlsx_remote)
    local_path = os.path.join(dest_dir, local_name)
    try:
        with urllib.request.urlopen(url, timeout=120) as resp, \
             open(local_path, "wb") as f:
            f.write(resp.read())
    except Exception as e:
        raise RuntimeError(f"GET /preview falló para '{categoria}': {e}")

    return local_path, res.get("total_productos")


# =============================================================================
# Auto-actualización de contadores en index.html
# =============================================================================

def update_index_counts(counts):
    """
    Actualiza el texto "<N> productos" de cada card en index.html según el dict
    {pdf_name: count}. Usa el href="pdfs/<pdf_name>" como ancla estable para
    encontrar la card correcta — así no depende del orden ni del título.

    Devuelve lista de (pdf_name, old_count, new_count) con los cambios efectivos.
    """
    if not os.path.exists(INDEX_HTML):
        print(f"  Aviso: no existe {INDEX_HTML}, salteo actualización de contadores",
              file=sys.stderr)
        return []

    with open(INDEX_HTML, "r", encoding="utf-8") as f:
        html = f.read()

    changes = []
    for pdf_name, count in counts.items():
        # Buscar la card que contiene este pdf y, mirando hacia atrás desde el href,
        # encontrar el <div class="card-meta">N productos</div> más cercano.
        pattern = re.compile(
            r'(<div class="card-meta">)(\d+)( productos</div>)'
            r'(?P<between>(?:(?!<div class="card-meta">).)*?)'
            r'href="pdfs/' + re.escape(pdf_name) + r'"',
            re.DOTALL,
        )
        m = pattern.search(html)
        if not m:
            print(f"  Aviso: no encontré card para {pdf_name} en index.html",
                  file=sys.stderr)
            continue
        old_count = int(m.group(2))
        if old_count == count:
            continue
        new_card_meta = f'{m.group(1)}{count}{m.group(3)}'
        html = html[:m.start()] + new_card_meta + m.group("between") + \
               f'href="pdfs/{pdf_name}"' + html[m.end():]
        changes.append((pdf_name, old_count, count))

    if changes:
        with open(INDEX_HTML, "w", encoding="utf-8") as f:
            f.write(html)

    return changes


# =============================================================================
# Main
# =============================================================================

def main():
    print("=" * 60)
    print("  Actualización de Catálogos El Forastero")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)
    print()

    # Verificar que los generadores existen
    for g in (GENERATOR, GENERATOR_ACCESORIOS):
        if not os.path.exists(g):
            print(f"Error: no se encontró {g}", file=sys.stderr)
            sys.exit(1)

    # Verificar Excel locales (solo Accesorios u otros con "excel" hardcodeado)
    missing = []
    for cat in CATALOGOS:
        if "excel" in cat and not os.path.exists(cat["excel"]):
            missing.append(cat["excel"])
    if missing:
        print("Error: faltan archivos Excel locales:", file=sys.stderr)
        for m in missing:
            print(f"  - {m}", file=sys.stderr)
        sys.exit(1)

    # Carpeta temp para listas descargadas (se borra al final del proceso)
    tmpdir = tempfile.mkdtemp(prefix="listas_forastero_")
    print(f"Carpeta temp para listas descargadas: {tmpdir}\n")

    # Regenerar cada catálogo
    errors = []
    counts = {}
    for i, cat in enumerate(CATALOGOS, 1):
        name = cat["pdf_name"].replace(".pdf", "").replace("-", " ").title()
        print(f"\n[{i}/{len(CATALOGOS)}] Generando: {name}")

        # Resolver path del Excel: descarga remota si tiene "categoria",
        # path local fijo si tiene "excel"
        if "categoria" in cat:
            print(f"  Descargando lista Nexion: {cat['categoria']} (Cervantes Depósito)...")
            try:
                excel_path, total_remoto = descargar_lista(cat["categoria"], tmpdir)
                print(f"  OK — {os.path.basename(excel_path)} ({total_remoto} prod. en lista)")
            except RuntimeError as e:
                print(f"  ERROR descarga: {e}", file=sys.stderr)
                errors.append(cat["pdf_name"])
                continue
        else:
            excel_path = cat["excel"]
            print(f"  Excel local: {os.path.basename(excel_path)}")

        # Nombre temporal del HTML/PDF que genera el script
        safe_name = cat["pdf_name"].replace(".pdf", "").replace("-", "_")
        temp_html = os.path.join(CATALOGO_DIR, f"catalogo_{safe_name}.html")
        temp_pdf = os.path.join(CATALOGO_DIR, f"catalogo_{safe_name}.pdf")

        generator = cat.get("generator", GENERATOR)
        cmd = [
            sys.executable, generator,
            "--excel", excel_path,
            "--output", temp_html,
            "--pdf",
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, cwd=CATALOGO_DIR)

        if result.returncode != 0:
            print(f"  ERROR generador: {result.stderr[-200:]}", file=sys.stderr)
            errors.append(cat["pdf_name"])
            continue

        # Capturar count de productos del stdout del generador
        m = PRODUCT_COUNT_RE.search(result.stdout or "")
        if m:
            counts[cat["pdf_name"]] = int(m.group(1))

        # Copiar PDF a destino
        dest = os.path.join(PDFS_DIR, cat["pdf_name"])
        if os.path.exists(temp_pdf):
            os.replace(temp_pdf, dest)
            size_mb = os.path.getsize(dest) / 1024 / 1024
            count_str = f" — {counts.get(cat['pdf_name'], '?')} productos"
            print(f"  OK — {size_mb:.1f} MB{count_str}")
        else:
            print(f"  ERROR: PDF no generado", file=sys.stderr)
            errors.append(cat["pdf_name"])

        # Limpiar HTML temporal
        if os.path.exists(temp_html):
            os.remove(temp_html)

    print()
    print("=" * 60)

    if errors:
        print(f"  Errores en {len(errors)} catálogo(s):")
        for e in errors:
            print(f"    - {e}")
        print("  Los catálogos sin error se actualizaron.")
        print("=" * 60)
        sys.exit(1)

    print(f"  {len(CATALOGOS)} catálogos regenerados exitosamente.")
    print("=" * 60)

    # Sincronizar contadores del index.html con el conteo real de cada PDF
    print("\nSincronizando contadores de index.html...")
    changes = update_index_counts(counts)
    if changes:
        for pdf_name, old, new in changes:
            print(f"  {pdf_name}: {old} → {new}")
    else:
        print("  Sin cambios en contadores (todo ya estaba sincronizado).")

    # Git commit + push
    print("\nPublicando en GitHub (Vercel re-deploya automáticamente)...")
    os.chdir(SCRIPT_DIR)

    # git add explícito por archivo — evita arrastrar PDFs huérfanos / copias
    # accidentales (ej: "alimentos-saludables 2.pdf") que estén sueltos en pdfs/.
    files_to_add = ["index.html"] + [
        os.path.join("pdfs", c["pdf_name"]) for c in CATALOGOS
    ]
    subprocess.run(["git", "add", "--"] + files_to_add, check=True)
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    subprocess.run([
        "git", "commit", "-m", f"update: catálogos actualizados — {date_str}",
    ], check=True)
    subprocess.run(["git", "push"], check=True)

    print("\nListo. Vercel va a re-deployar en ~1 minuto.")
    print("URL: https://catalogos.elforastero.com.ar")


if __name__ == "__main__":
    main()
