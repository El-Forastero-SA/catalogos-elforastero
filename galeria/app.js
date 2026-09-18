/* Galería de imágenes de producto — El Forastero
   Carga manifest.json, filtra/busca y descarga (individual + ZIP). */

const BATCH = 60;           // cards por tanda (render incremental)
const ZIP_MAX = 300;        // guardrail de descarga masiva

const state = { texto: "", especie: "", categoria: "", marca: "" };
let PRODUCTOS = [];
let filtrados = [];
let rendered = 0;

const $ = (id) => document.getElementById(id);
const norm = (s) => (s || "").toLowerCase()
  .normalize("NFD").replace(/[̀-ͯ]/g, "");

function toast(msg, ms = 2500) {
  const t = $("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast._t);
  if (ms) toast._t = setTimeout(() => t.classList.remove("show"), ms);
}

async function init() {
  let manifest;
  try {
    const r = await fetch("manifest.json", { cache: "no-cache" });
    manifest = await r.json();
  } catch (e) {
    $("grid").innerHTML = '<div class="vacio">No se pudo cargar el catálogo de imágenes.</div>';
    return;
  }
  PRODUCTOS = manifest.productos || [];

  // Especies presentes (orden fijo)
  const espSet = new Set(PRODUCTOS.map(p => p.especie).filter(Boolean));
  const orden = ["Perro", "Gato", "Equino"];
  const especies = orden.filter(e => espSet.has(e));
  const chips = $("chips-especie");
  especies.forEach(e => {
    const b = document.createElement("button");
    b.className = "chip";
    b.textContent = e;
    b.onclick = () => {
      state.especie = state.especie === e ? "" : e;
      [...chips.children].forEach(c => c.classList.toggle("activo", c.textContent === state.especie));
      aplicar();
    };
    chips.appendChild(b);
  });

  // Categorías (según manifest.categorias, con count)
  const selCat = $("filtro-categoria");
  (manifest.categorias || []).forEach(c => {
    if (!c.count) return;
    const o = document.createElement("option");
    o.value = c.nombre;
    o.textContent = `${c.nombre.replace(/^Mascotas - /, "")} (${c.count})`;
    selCat.appendChild(o);
  });
  selCat.onchange = () => { state.categoria = selCat.value; aplicar(); };

  // Marcas presentes (alfabético)
  const selMarca = $("filtro-marca");
  [...new Set(PRODUCTOS.map(p => p.marca).filter(m => m && m !== "Sin marca"))]
    .sort((a, b) => a.localeCompare(b, "es"))
    .forEach(m => {
      const o = document.createElement("option");
      o.value = m; o.textContent = m;
      selMarca.appendChild(o);
    });
  selMarca.onchange = () => { state.marca = selMarca.value; aplicar(); };

  $("buscar").oninput = (e) => { state.texto = norm(e.target.value); aplicar(); };
  $("btn-zip").onclick = descargarZip;

  // Render incremental al hacer scroll
  new IntersectionObserver((entries) => {
    if (entries[0].isIntersecting) renderMas();
  }).observe($("sentinel"));

  aplicar();
}

function aplicar() {
  const t = state.texto;
  filtrados = PRODUCTOS.filter(p => {
    if (state.especie && p.especie !== state.especie) return false;
    if (state.categoria && p.categoria !== state.categoria) return false;
    if (state.marca && p.marca !== state.marca) return false;
    if (t && !norm(p.nombre + " " + p.marca).includes(t)) return false;
    return true;
  });
  rendered = 0;
  $("grid").innerHTML = "";
  renderMas();

  $("conteo").textContent = filtrados.length
    ? `${filtrados.length} ${filtrados.length === 1 ? "imagen" : "imágenes"}`
    : "";
  if (!filtrados.length) {
    $("grid").innerHTML = '<div class="vacio">No hay imágenes para ese filtro.</div>';
  }

  const btn = $("btn-zip");
  btn.disabled = filtrados.length === 0;
  $("btn-zip-label").textContent = filtrados.length
    ? `Descargar todo (${filtrados.length})` : "Descargar todo";
}

function renderMas() {
  const grid = $("grid");
  const next = filtrados.slice(rendered, rendered + BATCH);
  const frag = document.createDocumentFragment();
  next.forEach(p => frag.appendChild(card(p)));
  grid.appendChild(frag);
  rendered += next.length;
}

function card(p) {
  const el = document.createElement("div");
  el.className = "card";
  el.innerHTML = `
    <div class="thumb"><img loading="lazy" src="thumbs/${p.sku}.webp" alt="${escapeHtml(p.nombre)}"></div>
    <div class="info">
      ${p.marca && p.marca !== "Sin marca" ? `<div class="marca">${escapeHtml(p.marca)}</div>` : ""}
      <div class="nombre">${escapeHtml(p.nombre)}</div>
      <a class="btn-desc" href="api/download?sku=${p.sku}">
        <svg viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
        Descargar
      </a>
    </div>`;
  return el;
}

async function descargarZip() {
  if (!filtrados.length) return;
  if (filtrados.length > ZIP_MAX) {
    if (!confirm(`Son ${filtrados.length} imágenes y puede tardar. `
      + `¿Descargar igual? (tip: filtrá por marca o especie para bajar menos)`)) return;
  }
  const btn = $("btn-zip");
  const prev = $("btn-zip-label").textContent;
  btn.disabled = true;
  $("btn-zip-label").textContent = "Preparando ZIP…";
  toast("Armando el ZIP, aguantá un toque…", 0);
  try {
    const skus = filtrados.map(p => p.sku);
    const nombre = nombreZip();
    const r = await fetch("api/zip", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ skus, nombre }),
    });
    if (!r.ok) throw new Error("HTTP " + r.status);
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = nombre + ".zip";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    toast("Listo ✓");
  } catch (e) {
    toast("Hubo un problema armando el ZIP. Probá de nuevo.");
  } finally {
    btn.disabled = false;
    $("btn-zip-label").textContent = prev;
  }
}

function nombreZip() {
  const partes = ["Imagenes"];
  if (state.especie) partes.push(state.especie);
  if (state.categoria) partes.push(state.categoria.replace(/^Mascotas - /, ""));
  if (state.marca) partes.push(state.marca);
  return partes.join("_").replace(/[^A-Za-z0-9_]+/g, "_") || "Imagenes";
}

function escapeHtml(s) {
  return (s || "").replace(/[&<>"']/g, c => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

init();
