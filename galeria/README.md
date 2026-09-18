# Galería de imágenes de producto — `imagenes.elforastero.com.ar`

Galería pública autoservicio para que los clientes descarguen las imágenes de
producto (las mismas que el catálogo, fuente Magento). Se mantiene sola con el
cron que regenera los catálogos (días 1 y 15).

## Cómo funciona

- **`generar_manifest.py`** (corre en CI): baja las listas de
  `listas.elforastero.com.ar`, consulta Magento por SKU (especie = `tipomascota`,
  marca = `tipomarca`, ruta de imagen), genera miniaturas `thumbs/<sku>.webp` y
  escribe `manifest.json`. Productos sin imagen se omiten.
- **Front-end estático** (`index.html` + `app.js` + `style.css`): grid con
  búsqueda y filtros (categoría / especie / marca). Miniaturas servidas estáticas.
- **`api/download.js`**: descarga 1 imagen — proxy al original **vivo** de Magento
  con nombre de archivo prolijo (`CodART_nombre.ext`).
- **`api/zip.js`**: descarga masiva — arma un ZIP al vuelo con los originales del
  set filtrado.

> Los thumbs son un **preview cacheado**. La descarga (individual o ZIP) siempre
> trae el **original vivo** de Magento, así que el archivo entregado nunca queda
> desactualizado aunque el thumb sea viejo.

## Correr local

```bash
python3 generar_manifest.py                 # todas las categorías
python3 generar_manifest.py --solo Bebidas  # una sola (iterar rápido)
node devserver.mjs                          # sirve estáticos + funciones en :8899
```

## Deploy (Vercel) — pasos manuales, una sola vez

1. **Vercel → Add New → Project** → importar `El-Forastero-SA/catalogos-elforastero`.
2. **Root Directory** = `galeria/`  ← clave, así no se mezcla con el sitio de catálogos.
3. Framework preset: **Other**. Build command: vacío. Install: `npm install` (default).
4. **Variables de entorno: ninguna** (el `media_base` viaja dentro del `manifest.json`).
5. Deploy. Verificar en la URL `*.vercel.app` que carga el grid y que descarga anda.
6. **Domains** → agregar `imagenes.elforastero.com.ar` → seguir la instrucción de
   DNS (CNAME `imagenes` → `cname.vercel-dns.com`, mismo proveedor que
   `catalogos.elforastero.com.ar`).

El cron ya existente (`.github/workflows/regenerar-catalogos.yml`) regenera el
manifest + thumbs y los commitea; cada push redeploya el proyecto de Vercel.
