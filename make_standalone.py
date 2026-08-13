""""
make_standalone.py  v4
Incrusta todos los archivos locales que Puppy Pages no sirve:
  - chart.min.js             -> <script> inline
  - datalabels.min.js        -> <script> inline
  - sw_data.json             -> var JS inline
  - bigquery_results/cd_chart.json -> var JS inline (reemplaza lazy fetch)
Tambien oculta el panel Actualizar (requiere servidor FastAPI, no aplica en Puppy Pages).
Tailwind CDN se deja intacto (el navegador de Walmart lo carga via proxy autenticado).
"""
import re
import json
from pathlib import Path

BASE = Path(__file__).parent
src  = BASE / "tablero_los_proveedores.html"
dst  = BASE / "tablero_standalone.html"

print("Leyendo HTML fuente...")
html = src.read_text(encoding="utf-8")

# ── 1. Incrustar chart.min.js ────────────────────────────────────────────────
chart_path = BASE / "chart.min.js"
if chart_path.exists():
    print(f"Incrustando chart.min.js ({chart_path.stat().st_size // 1024} KB)...")
    chart_js = chart_path.read_text(encoding="utf-8")
    html = html.replace(
        '<script src="chart.min.js"></script>',
        f'<script>{chart_js}</script>'
    )
    print("  OK")
else:
    print("  ADVERTENCIA: chart.min.js no encontrado")

# ── 2. Incrustar datalabels.min.js ──────────────────────────────────────────
dl_path = BASE / "datalabels.min.js"
if dl_path.exists():
    print(f"Incrustando datalabels.min.js ({dl_path.stat().st_size // 1024} KB)...")
    dl_js = dl_path.read_text(encoding="utf-8")
    html = html.replace(
        '<script src="datalabels.min.js"></script>',
        f'<script>{dl_js}</script>'
    )
    print("  OK")
else:
    print("  ADVERTENCIA: datalabels.min.js no encontrado")

# ── 3. Incrustar sw_data.json ────────────────────────────────────────────────
sw_path = BASE / "sw_data.json"
if sw_path.exists():
    print(f"Incrustando sw_data.json ({sw_path.stat().st_size // 1024} KB)...")
    sw = json.loads(sw_path.read_text(encoding="utf-8"))
    sw_inline = json.dumps(sw, ensure_ascii=False, separators=(",", ":"))

    # El HTML puede tener cache-busting o no:
    #   sin cache:  fetch('sw_data.json')
    #   con cache:  fetch('sw_data.json?v='+Date.now())
    # Nota: la comilla de cierre esta ANTES del + en la version con cache.
    _FETCH_RE = re.compile(
        r"fetch\("
        r"(?:\'sw_data\.json\'\)|'sw_data\.json\?v=\'\+Date\.now\(\)\))"
        r"\.then\(function\(r\)\{return r\.json\(\);\}\)\.then\(function\(d\)\{\n  SW_DATA = d;"
    )
    NEW_INLINE = "(function(){\n  var d = " + sw_inline + ";\n  SW_DATA = d;"
    OLD_END = "}).catch(function(){console.warn('sw_data.json no encontrado');});"
    NEW_END  = "})();"

    if _FETCH_RE.search(html):
        # Lambda evita que re.sub interprete backslashes del JSON como backreferences
        html = _FETCH_RE.sub(lambda _: NEW_INLINE, html, count=1)
        html = html.replace(OLD_END, NEW_END, 1)
        print("  OK")
    else:
        print("  ADVERTENCIA: patron fetch sw_data.json no encontrado en el HTML")
else:
    print("  ADVERTENCIA: sw_data.json no encontrado")

# ── 4. Incrustar bigquery_results/cd_chart.json ─────────────────────────────
cd_path = BASE / "bigquery_results" / "cd_chart.json"
if cd_path.exists():
    print(f"Incrustando cd_chart.json ({cd_path.stat().st_size // 1024} KB)...")
    cd_data = cd_path.read_text(encoding="utf-8").strip()
    OLD_CD = "var DATA_CHART_CD   = null; // cargado lazy via fetch"
    NEW_CD = f"var DATA_CHART_CD   = {cd_data}; // incrustado por make_standalone"
    if OLD_CD in html:
        html = html.replace(OLD_CD, NEW_CD, 1)
        print("  OK")
    else:
        print("  ADVERTENCIA: patron var DATA_CHART_CD no encontrado")
else:
    print("  ADVERTENCIA: bigquery_results/cd_chart.json no encontrado")

# ── 5. Ocultar panel Actualizar (requiere servidor, no funciona en Puppy Pages) ─
OLD_UPDATEBAR = 'id="updateBar"'
NEW_UPDATEBAR = 'id="updateBar" style="display:none"'
if OLD_UPDATEBAR in html:
    html = html.replace(OLD_UPDATEBAR, NEW_UPDATEBAR, 1)
    print("Panel Actualizar ocultado (no aplica en Puppy Pages)")
else:
    print("ADVERTENCIA: updateBar no encontrado")

# ── 6. Guardar ───────────────────────────────────────────────────────────────
dst.write_text(html, encoding="utf-8")
size_kb = dst.stat().st_size // 1024
print(f"\ntablero_standalone.html generado: {size_kb} KB")
print(f"Ruta: {dst}")

# Verificacion final
remaining_scripts = re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', html)
remaining_fetch   = re.findall(r"fetch\(['\"]([^'\"]+)['\"]", html)
local_srcs = [s for s in remaining_scripts if not s.startswith("http")]
cdn_srcs   = [s for s in remaining_scripts if s.startswith("http")]

print("\n=== Verificacion ===")
print(f"Scripts CDN (el navegador los carga): {cdn_srcs}")
if local_srcs:
    print(f"PROBLEMA - scripts locales pendientes: {local_srcs}")
else:
    print("OK: ningun script local pendiente")
print(f"fetch() restantes: {remaining_fetch}  (solo API interna, OK)")
