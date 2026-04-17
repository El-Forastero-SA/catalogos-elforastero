#!/usr/bin/env python3
"""
actualizar_catalogos.py — Regenera los 7 catálogos PDF y deploya a Vercel via git push.

Uso:
    python3 actualizar_catalogos.py

Requisitos:
    - generar_desde_lista.py en ../catalogo-template/
    - Archivos Excel de listas de precios en las rutas configuradas
    - .env con credenciales Magento en ../catalogo-template/
    - playwright instalado (pip3 install playwright && python3 -m playwright install chromium)
    - git configurado con remote origin

Flujo:
    1. Regenera los 7 PDFs desde los Excel de listas de precios
    2. Copia los PDFs a pdfs/
    3. Commit + push a GitHub
    4. Vercel re-deploya automáticamente
"""

import os
import subprocess
import sys
from datetime import datetime

# =============================================================================
# Configuración
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CATALOGO_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "catalogo-template")
PDFS_DIR = os.path.join(SCRIPT_DIR, "pdfs")
GENERATOR = os.path.join(CATALOGO_DIR, "generar_desde_lista.py")
GENERATOR_ACCESORIOS = os.path.join(CATALOGO_DIR, "generar_accesorios.py")

# Mapeo: nombre de archivo PDF destino -> ruta del Excel fuente.
# Si un catalogo usa un script distinto (ej: Accesorios con formato Cocooning),
# se indica con "generator" y los argumentos extra con "extra_args".
CATALOGOS = [
    {
        "excel": os.path.expanduser("~/Downloads/Mascotas___Alimentos_Super_Premium_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "alimentos-super-premium.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/Mascotas___Alimentos_Húmedos_Snack_Absorbentes_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "alimentos-humedos-snack-absorbentes.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/Mascotas___Alimentos_Premium_y_Standard_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "alimentos-premium-y-standard.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/Equinos_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "equinos.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/Alimentos_Saludables_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "alimentos-saludables.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/Bebidas_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "bebidas.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/Comestibles_e_Higiene_Cervantes_ListaDeposito.xlsx"),
        "pdf_name": "comestibles-e-higiene.pdf",
    },
    {
        "excel": os.path.expanduser("~/Downloads/lista precios 24-02-26.xlsx"),
        "pdf_name": "accesorios.pdf",
        "generator": GENERATOR_ACCESORIOS,
    },
]


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

    # Verificar Excel sources
    missing = []
    for cat in CATALOGOS:
        if not os.path.exists(cat["excel"]):
            missing.append(cat["excel"])
    if missing:
        print("Error: faltan archivos Excel:", file=sys.stderr)
        for m in missing:
            print(f"  - {m}", file=sys.stderr)
        print("\nActualizá las rutas en CATALOGOS si los archivos cambiaron de ubicación.",
              file=sys.stderr)
        sys.exit(1)

    # Regenerar cada catálogo
    errors = []
    for i, cat in enumerate(CATALOGOS, 1):
        name = cat["pdf_name"].replace(".pdf", "").replace("-", " ").title()
        print(f"\n[{i}/{len(CATALOGOS)}] Generando: {name}")
        print(f"  Excel: {os.path.basename(cat['excel'])}")

        # Nombre temporal del HTML/PDF que genera el script
        safe_name = cat["pdf_name"].replace(".pdf", "").replace("-", "_")
        temp_html = os.path.join(CATALOGO_DIR, f"catalogo_{safe_name}.html")
        temp_pdf = os.path.join(CATALOGO_DIR, f"catalogo_{safe_name}.pdf")

        generator = cat.get("generator", GENERATOR)
        cmd = [
            sys.executable, generator,
            "--excel", cat["excel"],
            "--output", temp_html,
            "--pdf",
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, cwd=CATALOGO_DIR)

        if result.returncode != 0:
            print(f"  ERROR: {result.stderr[-200:]}", file=sys.stderr)
            errors.append(cat["pdf_name"])
            continue

        # Copiar PDF a destino
        dest = os.path.join(PDFS_DIR, cat["pdf_name"])
        if os.path.exists(temp_pdf):
            os.replace(temp_pdf, dest)
            size_mb = os.path.getsize(dest) / 1024 / 1024
            print(f"  OK — {size_mb:.1f} MB")
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

    # Git commit + push
    print("\nPublicando en GitHub (Vercel re-deploya automáticamente)...")
    os.chdir(SCRIPT_DIR)

    subprocess.run(["git", "add", "pdfs/"], check=True)
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    subprocess.run([
        "git", "commit", "-m", f"update: catálogos actualizados — {date_str}",
    ], check=True)
    subprocess.run(["git", "push"], check=True)

    print("\nListo. Vercel va a re-deployar en ~1 minuto.")
    print("URL: https://catalogos.elforastero.com.ar")


if __name__ == "__main__":
    main()
