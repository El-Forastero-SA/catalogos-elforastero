"""Sube los 8 catálogos de El Forastero a Heyzine como flipbooks estilo revista.

Heyzine deduplica por URL del PDF: re-correr este script mantiene los mismos
flipbook IDs (y por lo tanto los mismos links que ya circularon con clientes),
mientras refresca la metadata y, según la API, el contenido del PDF asociado.

Lee la API key de ../.env (HEYZINE_API_KEY) — el .env vive un nivel más arriba
porque catalogos-web es un repo separado y los secretos no van en él.

Uso:
  python heyzine_upload.py            # sube los 8
  python heyzine_upload.py --test     # solo el primero (validación)
"""

import argparse
import json
import os
import sys
import time
from pathlib import Path

import urllib.request


API_URL = "https://heyzine.com/api1/rest"
TEMPLATE_ID = "8a18001ddf"
BASE_PDF_URL = "https://catalogos.elforastero.com.ar/pdfs"

CATALOGOS = [
    ("alimentos-super-premium.pdf",             "Alimentos Super Premium"),
    ("alimentos-humedos-snack-absorbentes.pdf", "Húmedos, Snack y Absorbentes"),
    ("alimentos-premium-y-standard.pdf",        "Alimentos Premium y Standard"),
    ("accesorios.pdf",                           "Accesorios para Mascotas"),
    ("equinos.pdf",                              "Equinos"),
    ("alimentos-saludables.pdf",                 "Alimentos Saludables"),
    ("bebidas.pdf",                              "Bebidas"),
    ("comestibles-e-higiene.pdf",                "Comestibles e Higiene"),
]


def load_api_key() -> str:
    # Busca primero en ../.env (Claude-Code/.env), después env var
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            if line.startswith("HEYZINE_API_KEY="):
                value = line.split("=", 1)[1].strip()
                if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
                    value = value[1:-1]
                return value
    key = os.environ.get("HEYZINE_API_KEY")
    if not key:
        sys.exit(f"ERROR: HEYZINE_API_KEY no definida (revisar {env_path})")
    return key


def upload_one(api_key_full: str, pdf_filename: str, title: str) -> dict:
    bearer, client_id = api_key_full.split(".", 1)
    pdf_url = f"{BASE_PDF_URL}/{pdf_filename}"
    payload = {
        "pdf": pdf_url,
        "client_id": client_id,
        "title": f"Catálogo {title}",
        "subtitle": "El Forastero S.A.",
        "tags": "catalogo,2026,forastero",
        "template": TEMPLATE_ID,
        "page_effect": "magazine",
        "download": True,
        "full_screen": True,
        "share": True,
        "prev_next": True,
    }
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(API_URL, data=body, method="POST")
    req.add_header("Authorization", f"Bearer {bearer}")
    req.add_header("Content-Type", "application/json")

    status_code = None
    response_data = None
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            status_code = resp.getcode()
            ctype = resp.headers.get("content-type", "")
            raw = resp.read().decode("utf-8")
            response_data = json.loads(raw) if ctype.startswith("application/json") else raw
    except urllib.error.HTTPError as e:
        status_code = e.code
        try:
            response_data = json.loads(e.read().decode("utf-8"))
        except Exception:
            response_data = str(e)
    except Exception as e:
        status_code = -1
        response_data = str(e)

    return {
        "title": title,
        "pdf_url": pdf_url,
        "status_code": status_code,
        "response": response_data,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true", help="Subir solo el primero")
    args = parser.parse_args()

    api_key = load_api_key()
    targets = CATALOGOS[:1] if args.test else CATALOGOS

    results = []
    failures = 0
    for i, (pdf, title) in enumerate(targets, 1):
        print(f"[{i}/{len(targets)}] Subiendo: {title} ...", flush=True)
        r = upload_one(api_key, pdf, title)
        results.append(r)
        ok = r["status_code"] == 200 and isinstance(r["response"], dict) and r["response"].get("url")
        if ok:
            pages = (r["response"].get("meta") or {}).get("num_pages", "?")
            print(f"  OK  -> {r['response']['url']} ({pages} pág.)")
        else:
            failures += 1
            resp_short = str(r["response"])[:200]
            print(f"  FAIL ({r['status_code']}): {resp_short}")
        time.sleep(1)

    out_path = Path(__file__).resolve().parent / "heyzine_results.json"
    out_path.write_text(json.dumps(results, indent=2, ensure_ascii=False))
    print(f"\nResultados guardados: {out_path}")

    if failures:
        sys.exit(1)


if __name__ == "__main__":
    main()
