#!/usr/bin/env python3
"""
generar_accesorios.py — Catalogo PDF de Accesorios (Cocooning + Hunter).

Variante del pipeline de catalogos de El Forastero, adaptada al Excel de
Cocooning (formato distinto al Excel Nexion):
  - Fila 1: headers
  - Filas con solo columna A rellena = subcategoria (ej: "COCOONING CAMAS")
  - Resto: productos (columna A = SKU, columna B = descripcion)

Comportamiento:
  - Busca imagenes en Magento API via SKU (mismo catalogo que otros 7)
  - Productos sin imagen en Magento se descartan
  - Page-break automatico al cambiar de subcategoria
  - Header bar muestra la subcategoria actual
  - Portada con logo Cocooning + distribuidor El Forastero
  - Color de portada: configurable via --category-color (default #2c3b54)

Uso:
    python3 generar_accesorios.py --excel "/path/to/lista_cocooning.xlsx" --pdf

Reusa utilidades de generar_desde_lista.py (Magento API, image processing,
CSS base, SVG icons, footer, back page).
"""

import argparse
import math
import os
import sys
from html import escape

# Reusar utilidades del script base
from generar_desde_lista import (
    load_env,
    fetch_product_image,
    get_processed_image_data_uri,
    split_product_name,
    SVG_ICON_WEB,
    SVG_ICON_INSTAGRAM,
    SVG_ICON_WHATSAPP,
    SVG_ICON_WHATSAPP_GREEN,
    SVG_BG_PATTERN,
    LOGO_WHITE_B64,
    LOGO_BLACK_B64,
    _load_b64_file,
    generate_css,
    render_back_page,
)


LOGO_COCOONING_B64 = _load_b64_file("logo_cocooning_b64.txt")


# =============================================================================
# Excel reader — formato Cocooning
# =============================================================================

def read_excel_accesorios(excel_path):
    """Lee Excel de Cocooning y retorna lista de productos con subcategoria.

    Estructura del Excel:
        - Fila 1: headers (Articulo, Descripcion, EAN, ...)
        - Fila donde solo col A esta rellena y col B vacia -> subcategoria
        - Fila con col A numerico + col B texto -> producto

    Retorna: [{sku, name, subcategory}, ...] en orden del Excel.
    """
    try:
        import openpyxl
    except ImportError:
        print("Error: openpyxl no esta instalado. pip3 install openpyxl",
              file=sys.stderr)
        sys.exit(1)

    if not os.path.exists(excel_path):
        print(f"Error: no se encontro el archivo: {excel_path}", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    products = []
    current_subcategory = None

    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            continue  # header

        val_a = row[0]
        val_b = row[1] if len(row) > 1 else None

        if val_a is None:
            continue

        # Subcategoria: col A con texto, col B vacia
        if val_b is None:
            val_a_str = str(val_a).strip()
            if val_a_str:
                current_subcategory = val_a_str
            continue

        # Producto: col A numerico + col B texto
        val_a_str = str(val_a).strip()
        try:
            numeric_val = int(float(val_a_str))
            sku = str(numeric_val).zfill(5)
        except (ValueError, TypeError):
            continue

        name = str(val_b).strip()
        if not name:
            continue

        products.append({
            "sku": sku,
            "name": name,
            "subcategory": current_subcategory or "Accesorios",
        })

    wb.close()
    return products


# =============================================================================
# Rendering — portada con logo Cocooning, header variable por subcategoria
# =============================================================================

def render_cover_page_accesorios(category_name):
    """Portada con logo Cocooning grande + logo El Forastero abajo."""
    if LOGO_COCOONING_B64:
        cocooning_logo = (
            f'<img class="cover-logo" '
            f'src="data:image/png;base64,{LOGO_COCOONING_B64}" '
            f'alt="Cocooning Pet &amp; Co" '
            f'style="width:480px;background:#fff;border-radius:50%;padding:20px;">'
        )
    else:
        cocooning_logo = (
            '<div style="font-family:Playfair Display,serif;font-size:56px;'
            'font-weight:800;color:white;letter-spacing:6px;'
            'margin-bottom:40px;">Cocooning Pet &amp; Co</div>'
        )

    # Logo El Forastero blanco abajo (chico)
    if LOGO_WHITE_B64:
        distributor_html = (
            '<div style="position:absolute;bottom:60px;left:50%;'
            'transform:translateX(-50%);opacity:0.9;">'
            f'<img src="data:image/png;base64,{LOGO_WHITE_B64}" '
            'alt="El Forastero" style="height:55px;width:auto;">'
            '</div>'
        )
    else:
        distributor_html = ""

    return f"""
  <div class="page page-cover" style="position:relative;">
    {cocooning_logo}
    <div class="cover-divider"></div>
    <div class="cover-category">{escape(category_name)}</div>
    <div class="cover-subcategory">Portfolio de Productos</div>
    <div class="cover-contact">
      <span>{SVG_ICON_WEB} www.elforastero.com.ar</span>
      <span>{SVG_ICON_INSTAGRAM} @elforasterosa</span>
    </div>
    {distributor_html}
  </div>"""


def render_product_page_accesorios(products_chunk, subcategory_label):
    """Pagina de productos con subcategoria en el header bar."""
    from generar_desde_lista import render_product_card

    cards_html = "\n".join(render_product_card(p) for p in products_chunk)

    cocooning_header_logo = ""
    if LOGO_COCOONING_B64:
        cocooning_header_logo = (
            f'<img class="header-logo-img" style="height:108px;margin-left:24px;" '
            f'src="data:image/png;base64,{LOGO_COCOONING_B64}" alt="Cocooning">'
        )

    # Font-size dinamico: si el label es largo (Hunter subcats), achicar.
    title_style = ""
    if len(subcategory_label) > 22:
        title_style = ' style="font-size:22px;letter-spacing:2px;"'

    return f"""
  <div class="page">
    <div class="page-header">
      <div class="header-logo">
        <img class="header-logo-img" src="data:image/png;base64,{LOGO_BLACK_B64}" alt="El Forastero">
        {cocooning_header_logo}
      </div>
      <div class="header-bar">
        <span class="header-bar-title"{title_style}>{escape(subcategory_label)}</span>
      </div>
    </div>

    <div class="content-area">
      <div class="product-grid">
{cards_html}
      </div>
    </div>

    <div class="page-footer">
      <img class="footer-logo-img" src="data:image/png;base64,{LOGO_WHITE_B64}" alt="El Forastero">
      <div class="footer-contact-row">
        <span class="footer-contact-item">
          {SVG_ICON_WEB} www.elforastero.com.ar
        </span>
        <span class="footer-contact-item">
          {SVG_ICON_INSTAGRAM} @elforasterosa
        </span>
      </div>
      <div class="footer-divider"></div>
      <div class="footer-phones">
        <div class="footer-phone-item">
          {SVG_ICON_WHATSAPP}
          <span>298 4222 800</span>
          <span class="footer-phone-branch">CERVANTES</span>
        </div>
        <div class="footer-phone-item">
          {SVG_ICON_WHATSAPP}
          <span>298 4222 800</span>
          <span class="footer-phone-branch">SANTA ROSA</span>
        </div>
        <div class="footer-phone-item">
          {SVG_ICON_WHATSAPP}
          <span>294 441-0174</span>
          <span class="footer-phone-branch">BARILOCHE</span>
        </div>
        <div class="footer-phone-item">
          {SVG_ICON_WHATSAPP}
          <span>2396 44-0420</span>
          <span class="footer-phone-branch">PEHUAJO</span>
        </div>
      </div>
    </div>
  </div>"""


def clean_subcategory_label(raw):
    """Formatea la subcategoria para mostrar en el header bar.

    "COCOONING CAMAS" -> "Cocooning · Camas"
    "HUNTER JUGUETES PARA GATOS" -> "Hunter · Juguetes para Gatos"
    """
    if not raw:
        return "Accesorios"
    parts = raw.strip().split(None, 1)
    if len(parts) == 2:
        brand, rest = parts
        return f"{brand.title()} · {rest.title()}"
    return raw.title()


def chunk_products_by_subcategory(products, products_per_page):
    """Agrupa por subcategoria y chunks de N productos.

    Garantiza page-break al cambiar de subcategoria (ninguna pagina mezcla).

    Retorna: [(subcategory_label, products_chunk), ...]
    """
    pages = []
    current_sub = None
    buffer = []

    def flush():
        if not buffer:
            return
        label = clean_subcategory_label(current_sub)
        for i in range(0, len(buffer), products_per_page):
            pages.append((label, buffer[i:i + products_per_page]))

    for p in products:
        sub = p.get("subcategory", "")
        if sub != current_sub:
            flush()
            buffer = []
            current_sub = sub
        buffer.append(p)
    flush()

    return pages


# =============================================================================
# HTML Generation
# =============================================================================

def generate_html(category_name, category_color, products, products_per_page):
    """Genera el HTML completo del catalogo de accesorios."""
    css = generate_css(category_color)

    page_groups = chunk_products_by_subcategory(products, products_per_page)
    pages_html = [
        render_product_page_accesorios(chunk, label)
        for label, chunk in page_groups
    ]

    cover = render_cover_page_accesorios(category_name)
    back = render_back_page()
    all_pages = cover + "\n" + "\n".join(pages_html) + "\n" + back

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=1080">
  <title>El Forastero - {escape(category_name)} - Portfolio de Productos</title>
  <style>
    @font-face {{ font-family: 'Montserrat'; font-weight: 400; font-style: normal;
      src: url('fonts/Montserrat-400.woff2') format('woff2'); }}
    @font-face {{ font-family: 'Montserrat'; font-weight: 500; font-style: normal;
      src: url('fonts/Montserrat-500.woff2') format('woff2'); }}
    @font-face {{ font-family: 'Montserrat'; font-weight: 600; font-style: normal;
      src: url('fonts/Montserrat-600.woff2') format('woff2'); }}
    @font-face {{ font-family: 'Montserrat'; font-weight: 700; font-style: normal;
      src: url('fonts/Montserrat-700.woff2') format('woff2'); }}
    @font-face {{ font-family: 'Montserrat'; font-weight: 800; font-style: normal;
      src: url('fonts/Montserrat-800.woff2') format('woff2'); }}
    @font-face {{ font-family: 'Playfair Display'; font-weight: 700; font-style: normal;
      src: url('fonts/PlayfairDisplay-700.woff2') format('woff2'); }}
    @font-face {{ font-family: 'Playfair Display'; font-weight: 800; font-style: normal;
      src: url('fonts/PlayfairDisplay-800.woff2') format('woff2'); }}
{css}
  </style>
</head>
<body>
{all_pages}
</body>
</html>"""

    return html, len(pages_html)


# =============================================================================
# Main
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Genera catalogo de Accesorios (Cocooning + Hunter) desde Excel."
    )
    parser.add_argument("--excel", required=True, help="Ruta al Excel Cocooning")
    parser.add_argument("--category-name", default="Accesorios para Mascotas",
                        help="Titulo de portada")
    parser.add_argument("--category-color", default="#2c3b54",
                        help="Color hex de portada/header/footer")
    parser.add_argument("--output", default=None,
                        help="Archivo de salida (default: catalogo_accesorios.html)")
    parser.add_argument("--products-per-page", type=int, default=12)
    parser.add_argument("--pdf", action="store_true", help="Generar tambien PDF")

    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(script_dir, ".env")
    env_vars = load_env(env_path)

    base_url = env_vars.get("MAGENTO_BASE_URL", "").rstrip("/")
    token = env_vars.get("MAGENTO_ACCESS_TOKEN", "")

    if not base_url or not token:
        print("Error: MAGENTO_BASE_URL y MAGENTO_ACCESS_TOKEN requeridos en .env",
              file=sys.stderr)
        sys.exit(1)

    # --- Leer Excel ---
    excel_path = os.path.abspath(args.excel)
    print(f"Leyendo Excel: {excel_path}")
    products = read_excel_accesorios(excel_path)
    print(f"Productos encontrados: {len(products)}")

    # Resumen subcategorias
    subs_count = {}
    for p in products:
        s = p["subcategory"]
        subs_count[s] = subs_count.get(s, 0) + 1
    print(f"Subcategorias: {len(subs_count)}")
    for s, c in subs_count.items():
        print(f"  {s}: {c}")

    if not products:
        print("No se encontraron productos. Abortando.", file=sys.stderr)
        sys.exit(1)

    # --- Magento: imagenes ---
    print(f"\nObteniendo imagenes de Magento API ({len(products)} SKUs)...")
    not_found = 0
    for idx, product in enumerate(products, 1):
        sys.stdout.write(f"\r  {idx}/{len(products)}...")
        sys.stdout.flush()
        image_url = fetch_product_image(base_url, token, product["sku"])
        product["image_url"] = image_url
        if image_url is None:
            not_found += 1
    print(f"\r  {len(products)}/{len(products)} - listo.")
    print(f"  Sin imagen en Magento: {not_found}")

    # --- Procesar imagenes ---
    print(f"\nProcesando imagenes (crop + JPG)...")
    products_with_images = [p for p in products if p.get("image_url")]
    proc_failed = 0
    for idx, product in enumerate(products_with_images, 1):
        sys.stdout.write(f"\r  {idx}/{len(products_with_images)}...")
        sys.stdout.flush()
        data_uri = get_processed_image_data_uri(product["image_url"], script_dir)
        if data_uri:
            product["image_url"] = data_uri
        else:
            product["image_url"] = None
            proc_failed += 1
    print(f"\r  {len(products_with_images)}/{len(products_with_images)} - listo.")
    if proc_failed:
        print(f"  Fallos: {proc_failed}")

    # Filtrar productos sin imagen
    products = [p for p in products if p.get("image_url")]
    print(f"\nProductos finales (con imagen): {len(products)}")

    if not products:
        print("Ningun producto tiene imagen. Abortando.", file=sys.stderr)
        sys.exit(1)

    # --- Generar HTML ---
    html_content, num_pages = generate_html(
        args.category_name,
        args.category_color,
        products,
        args.products_per_page,
    )

    output_path = args.output or "catalogo_accesorios.html"
    if not os.path.isabs(output_path):
        output_path = os.path.join(script_dir, output_path)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    # --- PDF ---
    pdf_path = None
    if args.pdf:
        pdf_path = output_path.replace(".html", ".pdf")
        print(f"\nGenerando PDF...")
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page()
                page.goto(f"file://{output_path}", wait_until="networkidle",
                          timeout=120000)
                page.wait_for_timeout(3000)
                page.pdf(
                    path=pdf_path,
                    width="1080px",
                    height="1920px",
                    margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
                    print_background=True,
                    prefer_css_page_size=True,
                )
                browser.close()
            print(f"PDF: {pdf_path}")
        except Exception as e:
            print(f"Error generando PDF: {e}", file=sys.stderr)
            pdf_path = None

    # --- Resumen ---
    print("")
    print("=" * 55)
    print(f"  Catalogo Accesorios generado")
    print(f"  Titulo:     {args.category_name}")
    print(f"  Color:      {args.category_color}")
    print(f"  Productos:  {len(products)}")
    print(f"  Sin imagen: {not_found}")
    print(f"  Paginas:    {num_pages + 2} (portada + {num_pages} productos + contra)")
    print(f"  HTML:       {output_path}")
    if pdf_path:
        print(f"  PDF:        {pdf_path}")
    print("=" * 55)


if __name__ == "__main__":
    main()
