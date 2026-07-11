"""Generate the LOS dashboard HTML with embedded data."""
import json
import os
import subprocess

BASE = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP'

with open(os.path.join(BASE, 'dashboard_data.json'), encoding='utf-8') as f:
    DATA = json.load(f)

DATA_JS = json.dumps(DATA, ensure_ascii=False)

cedis_options_html = '\n'.join(
    f'<option value="{c["cedis_code"]}">{c["cedis_code"]} \u2014 {c["cedis_name"]}</option>'
    for c in DATA['cedis_list']
)
mes_options_html = '\n'.join(
    f'<option value="{m}">{m}</option>'
    for m in DATA['meses']
)

# Pre-build all CEDIS codes per category for the matrix header
AUTO_CEDIS = ['CUU', 'CLN', 'MXL', 'MTY', 'CUAU', 'STB', 'CHL', 'GDL', 'MER', 'VHSA']
SAMS_CEDIS = ['CUU', 'CLN', 'MTY', 'SMO', 'CHL', 'GDL', 'MER', 'VHSA']

HTML = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Tablero LOS Proveedores 2026</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
  <style>
    body {{ font-family: 'Segoe UI', Arial, sans-serif; }}
    select {{ appearance: auto; }}
    .chart-wrapper {{ position: relative; height: 340px; }}

    /* Matrix table */
    .mat-table {{ border-collapse: collapse; font-size: 0.72rem; width: 100%; }}
    .mat-table th, .mat-table td {{
      border: 1px solid #d1d5db;
      padding: 4px 7px;
      text-align: center;
      white-space: nowrap;
    }}
    .mat-table td.vendor-name {{
      text-align: left;
      font-weight: 500;
      min-width: 160px;
      background: #f9fafb;
    }}
    .mat-table tr:hover td {{ filter: brightness(0.96); }}
    /* header rows */
    .mat-table .th-group {{
      background: #0053e2; color: white; font-weight: 700; font-size: 0.75rem;
    }}
    .mat-table .th-norte {{ background: #1d4ed8; color: white; font-weight: 600; }}
    .mat-table .th-centro {{ background: #7c3aed; color: white; font-weight: 600; }}
    .mat-table .th-sur {{ background: #059669; color: white; font-weight: 600; }}
    .mat-table .th-prom  {{ background: #d97706; color: white; font-weight: 700; }}
    .mat-table .th-obj-group {{ background: #1e40af; color: white; font-weight: 700; }}
    .mat-table .th-mes-group {{ background: #0053e2; color: white; font-weight: 700; }}
    .mat-table .th-cedis {{ background: #1e3a5f; color: #bfdbfe; font-size: 0.68rem; font-weight: 700; }}
    /* conditional cell colors */
    .cell-over  {{ background: #fed7aa; color: #92400e; font-weight: 700; }} /* over obj */
    .cell-ok    {{ background: #bbf7d0; color: #14532d; font-weight: 600; }} /* ok */
    .cell-empty {{ color: #9ca3af; }}
    .cell-total {{ background: #f3f4f6; font-weight: 700; color: #1e3a5f; }}
    /* section label */
    .section-label {{
      background: #1e3a5f; color: white;
      padding: 6px 12px; font-weight: 700; font-size: 0.78rem;
      letter-spacing: 0.05em;
    }}
    /* tab button */
    .tab-btn {{
      padding: 6px 18px; border-radius: 6px; font-size: 0.82rem;
      font-weight: 600; cursor: pointer; border: 2px solid transparent;
      transition: all 0.15s;
    }}
    .tab-btn.active {{ background: #0053e2; color: white; }}
    .tab-btn:not(.active) {{
      background: white; color: #374151; border-color: #d1d5db;
    }}
    .tab-btn:not(.active):hover {{ border-color: #0053e2; color: #0053e2; }}
  </style>
</head>
<body class="bg-gray-50 text-gray-800 min-h-screen">

<!-- HEADER -->
<header class="bg-[#0053e2] text-white shadow">
  <div class="max-w-screen-2xl mx-auto px-6 py-4 flex items-center gap-4">
    <div>
      <h1 class="text-xl font-bold leading-tight">Tablero LOS Proveedores 2026</h1>
      <p class="text-blue-200 text-sm">Level of Service &mdash; Tiempo total en CEDIS por proveedor y locacion</p>
    </div>
  </div>
</header>

<!-- FILTER BAR -->
<div class="bg-white border-b border-gray-200 shadow-sm sticky top-0 z-30">
  <div class="max-w-screen-2xl mx-auto px-6 py-3 flex flex-wrap items-center gap-5">

    <!-- view toggle -->
    <div class="flex gap-2">
      <button id="btn-graficas" class="tab-btn active" onclick="setView('graficas')">Graficas</button>
      <button id="btn-tabla"    class="tab-btn"        onclick="setView('tabla')">Tabla Matricial</button>
    </div>

    <div class="w-px h-6 bg-gray-200"></div>

    <!-- CEDIS filter (charts only) -->
    <label id="label-cedis" class="flex items-center gap-2 text-sm font-medium text-gray-700">
      Filtrar por CEDIS:
      <select id="sel-cedis"
        class="border border-gray-300 rounded-md px-3 py-1.5 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
        {cedis_options_html}
      </select>
    </label>

    <!-- Month filter (both views) -->
    <label class="flex items-center gap-2 text-sm font-medium text-gray-700">
      Mes:
      <select id="sel-mes"
        class="border border-gray-300 rounded-md px-3 py-1.5 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
        {mes_options_html}
      </select>
    </label>

    <span id="info-subtitle" class="ml-auto text-xs text-gray-500 italic hidden sm:block"></span>
  </div>
</div>

<!-- STAT CARDS -->
<div class="max-w-screen-2xl mx-auto px-6 mt-5 grid grid-cols-2 sm:grid-cols-4 gap-4">
  <div class="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
    <p class="text-xs text-gray-500 mb-1">Total citas (Autoserv.)</p>
    <p id="stat-citas-auto" class="text-2xl font-bold text-[#0053e2]">--</p>
  </div>
  <div class="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
    <p class="text-xs text-gray-500 mb-1">LOS Prom. Autoserv.</p>
    <p id="stat-los-auto" class="text-2xl font-bold text-[#2a8703]">--</p>
  </div>
  <div class="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
    <p class="text-xs text-gray-500 mb-1">Total citas (SAM'S)</p>
    <p id="stat-citas-sams" class="text-2xl font-bold text-[#0053e2]">--</p>
  </div>
  <div class="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
    <p class="text-xs text-gray-500 mb-1">LOS Prom. SAM'S</p>
    <p id="stat-los-sams" class="text-2xl font-bold text-[#2a8703]">--</p>
  </div>
</div>

<!-- ======= CHARTS VIEW ======= -->
<main id="view-graficas" class="max-w-screen-2xl mx-auto px-6 mt-5 space-y-6 pb-12">

  <section class="bg-white rounded-2xl border border-gray-200 shadow-sm p-5">
    <div class="flex items-center gap-3 mb-4">
      <h2 class="text-base font-bold text-gray-800">Autoservicios &mdash; LOS por Proveedor</h2>
    </div>
    <div class="chart-wrapper"><canvas id="chart-auto"></canvas></div>
    <p class="text-xs text-gray-400 mt-2">Barras = etapas (Llegada / Recibo / Salida). Verde = Total LOS. Azul = Prom. historico. Rojo = Objetivo 2026.</p>
  </section>

  <section class="bg-white rounded-2xl border border-gray-200 shadow-sm p-5">
    <div class="flex items-center gap-3 mb-4">
      <h2 class="text-base font-bold text-gray-800">SAM'S Club &mdash; LOS por Proveedor</h2>
    </div>
    <div class="chart-wrapper"><canvas id="chart-sams"></canvas></div>
  </section>

  <section class="bg-white rounded-2xl border border-gray-200 shadow-sm p-5">
    <div class="flex items-center gap-3 mb-4">
      <h2 class="text-base font-bold text-gray-800">Tendencia Mensual &mdash; LOS Promedio Global</h2>
    </div>
    <div class="chart-wrapper"><canvas id="chart-trend"></canvas></div>
  </section>
</main>

<!-- ======= MATRIX TABLE VIEW ======= -->
<main id="view-tabla" class="hidden max-w-screen-2xl mx-auto px-4 mt-5 pb-12 space-y-6">

  <!-- Legend -->
  <div class="bg-white rounded-xl border border-gray-200 shadow-sm px-5 py-3 flex flex-wrap gap-5 text-xs font-semibold">
    <span class="flex items-center gap-2">
      <span class="inline-block w-5 h-4 rounded cell-ok border border-green-300"></span>
      Dentro de objetivo (Actual &le; Objetivo)
    </span>
    <span class="flex items-center gap-2">
      <span class="inline-block w-5 h-4 rounded cell-over border border-orange-300"></span>
      Fuera de objetivo (Actual &gt; Objetivo)
    </span>
    <span class="flex items-center gap-2">
      <span class="inline-block w-5 h-4 rounded bg-gray-100 border border-gray-300"></span>
      Sin datos / sin objetivo
    </span>
  </div>

  <!-- AUTO Secos -->
  <div class="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden">
    <div class="section-label">AUTO Secos</div>
    <div class="overflow-x-auto">
      <table class="mat-table" id="mat-auto"></table>
    </div>
  </div>

  <!-- SAM'S Secos -->
  <div class="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden">
    <div class="section-label">SAM'S Secos</div>
    <div class="overflow-x-auto">
      <table class="mat-table" id="mat-sams"></table>
    </div>
  </div>
</main>

<script>
// ─── Data ─────────────────────────────────────────────────────────────────────
const DATA = {DATA_JS};

// ─── Colors ──────────────────────────────────────────────────────────────────
const C = {{
  llegada:  '#9ca3af',
  recibo:   '#f97316',
  salida:   '#ffc220',
  total:    '#2a8703',
  histProm: '#0053e2',
  objetivo: '#ea1100',
  gridLine: '#f3f4f6',
}};

// ─── State ───────────────────────────────────────────────────────────────────
let currentView = 'graficas';
let chartAuto = null, chartSams = null, chartTrend = null;

const $ = id => document.getElementById(id);

// ─── View toggle ─────────────────────────────────────────────────────────────
function setView(v) {{
  currentView = v;
  $('view-graficas').classList.toggle('hidden', v !== 'graficas');
  $('view-tabla').classList.toggle('hidden', v !== 'tabla');
  $('label-cedis').classList.toggle('hidden', v !== 'graficas');
  $('btn-graficas').classList.toggle('active', v === 'graficas');
  $('btn-tabla').classList.toggle('active', v === 'tabla');
  if (v === 'tabla') renderMatrix();
  else updateCharts();
}}

// ─── Chart helpers ───────────────────────────────────────────────────────────
function avg(arr) {{
  const vals = arr.filter(v => v != null && !isNaN(v));
  return vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) : null;
}}
function fmtH(h) {{ return h == null ? '--' : (+h).toFixed(1) + 'h'; }}

function buildLOSChart(canvasId, vendors) {{
  const ctx = $(canvasId).getContext('2d');
  return new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: vendors.map(v => v.vendor),
      datasets: [
        {{ type:'bar',  label:'Llegada',       data: vendors.map(v=>v.llegada),    backgroundColor:C.llegada,  stack:'los', order:3 }},
        {{ type:'bar',  label:'Recibo',        data: vendors.map(v=>v.recibo),     backgroundColor:C.recibo,   stack:'los', order:3 }},
        {{ type:'bar',  label:'Salida CD',     data: vendors.map(v=>v.salida),     backgroundColor:C.salida,   stack:'los', order:3 }},
        {{ type:'line', label:'Total',         data: vendors.map(v=>v.total),      borderColor:C.total,        backgroundColor:C.total, pointStyle:'circle', pointRadius:6, borderWidth:2, showLine:false, order:1 }},
        {{ type:'line', label:'Prom. historico', data: vendors.map(v=>v.hist_prom), borderColor:C.histProm, backgroundColor:'transparent', borderWidth:2, borderDash:[6,3], pointRadius:3, tension:0.3, order:2 }},
        {{ type:'line', label:'Objetivo 2026', data: vendors.map(v=>v.objetivo),   borderColor:C.objetivo, backgroundColor:'transparent', borderWidth:1.5, borderDash:[4,4], pointRadius:3, pointStyle:'rectRot', order:2 }},
      ],
    }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{ position:'bottom', labels:{{ usePointStyle:true, padding:14, font:{{size:11}} }} }},
        tooltip:{{ callbacks:{{ label(ctx) {{
          const v = vendors[ctx.dataIndex];
          if (ctx.dataset.label==='Total') return `Total: ${{fmtH(v.total)}} (${{v.citas}} citas)`;
          if (ctx.dataset.label==='Prom. historico') return `Prom: ${{fmtH(v.hist_prom)}}`;
          if (ctx.dataset.label==='Objetivo 2026') return `Objetivo: ${{fmtH(v.objetivo)}}`;
          return `${{ctx.dataset.label}}: ${{fmtH(ctx.parsed.y)}}`;
        }} }} }},
      }},
      scales:{{
        x:{{ stacked:true, grid:{{display:false}}, ticks:{{font:{{size:11,weight:'600'}},maxRotation:30}} }},
        y:{{ stacked:true, title:{{display:true,text:'Horas'}}, grid:{{color:C.gridLine}}, ticks:{{callback:v=>v+'h'}}, beginAtZero:true }},
      }},
    }},
  }});
}}

function updateCharts() {{
  const cedis = $('sel-cedis').value;
  const mes   = $('sel-mes').value;
  const cInfo = DATA.cedis_list.find(c=>c.cedis_code===cedis);
  $('info-subtitle').textContent = `${{cInfo ? cInfo.cedis_name : cedis}} | ${{mes}}`;

  const chartData = (DATA.charts[cedis]||{{}})[mes]||{{}};
  const autoV = chartData['Autoservicios']||[];
  const samsV = chartData["SAM'S Club"]||[];

  // Stats
  const autoCitas = autoV.reduce((s,v)=>s+v.citas,0);
  const samsCitas = samsV.reduce((s,v)=>s+v.citas,0);
  $('stat-citas-auto').textContent = autoCitas.toLocaleString('es-MX');
  $('stat-citas-sams').textContent = samsCitas.toLocaleString('es-MX');
  $('stat-los-auto').textContent   = autoV.length ? (avg(autoV.map(v=>v.total))||0).toFixed(1)+'h' : '--';
  $('stat-los-sams').textContent   = samsV.length ? (avg(samsV.map(v=>v.total))||0).toFixed(1)+'h' : '--';

  if (chartAuto) chartAuto.destroy();
  if (chartSams) chartSams.destroy();
  chartAuto = autoV.length ? buildLOSChart('chart-auto', autoV) : null;
  chartSams = samsV.length ? buildLOSChart('chart-sams', samsV) : null;
}}

// ─── Matrix Table ─────────────────────────────────────────────────────────────
function cellClass(actual, objetivo) {{
  if (actual == null) return 'cell-empty';
  if (objetivo == null) return '';  // no objective data
  return actual > objetivo ? 'cell-over' : 'cell-ok';
}}

function buildMatrixHTML(cat, tableId) {{
  const mes = $('sel-mes').value;
  const mx  = DATA.matrix[cat];
  if (!mx) return;

  const regions   = mx.regions;
  const vendors   = mx.vendors;
  const objetivos = mx.objectives;
  const mesData   = mx.meses[mes] || {{}};
  const allCedis  = regions.flatMap(r => r.cedis);

  // Count cols: 1 (vendor) + allCedis + 1 (prom) | allCedis + 1 (prom)
  const nCedis = allCedis.length;
  const nObjCols = nCedis + 1;   // cedis + prom
  const nMesCols = nCedis + 1;

  let html = '<thead>';

  // Row 1: group headers
  html += '<tr>';
  html += `<th rowspan="3" style="min-width:170px;text-align:left;background:#f9fafb;color:#1e3a5f;font-weight:700;">Proveedores TOP</th>`;
  html += `<th colspan="${{nObjCols}}" class="th-obj-group">Objetivos LOS 2026 por Cedis</th>`;
  html += `<th colspan="${{nMesCols}}" class="th-mes-group">${{mes}}</th>`;
  html += '</tr>';

  // Row 2: region headers
  html += '<tr>';
  for (const rg of regions) {{
    html += `<th colspan="${{rg.cedis.length}}" class="th-${{rg.name.toLowerCase()}}">${{rg.name}}</th>`;
  }}
  html += `<th class="th-prom">2026</th>`;
  for (const rg of regions) {{
    html += `<th colspan="${{rg.cedis.length}}" class="th-${{rg.name.toLowerCase()}}">${{rg.name}}</th>`;
  }}
  html += `<th class="th-prom">2026</th>`;
  html += '</tr>';

  // Row 3: CEDIS codes
  html += '<tr>';
  for (const c of allCedis) html += `<th class="th-cedis">${{c}}</th>`;
  html += `<th class="th-cedis">Prom</th>`;
  for (const c of allCedis) html += `<th class="th-cedis">${{c}}</th>`;
  html += `<th class="th-cedis">Prom</th>`;
  html += '</tr>';

  html += '</thead><tbody>';

  // Vendor rows
  for (const v of vendors) {{
    const vname = v.name;
    const obj   = objetivos[vname] || {{}};
    const actual = mesData[vname]  || {{}};

    // Compute objective prom
    const objVals = allCedis.map(c => obj[c]).filter(x=>x!=null);
    const objProm = objVals.length ? (objVals.reduce((a,b)=>a+b,0)/objVals.length).toFixed(1) : null;

    html += '<tr>';
    // Vendor name — truncate for readability
    const displayName = vname.length > 28 ? vname.slice(0,26)+'...' : vname;
    html += `<td class="vendor-name" title="${{vname}}">${{displayName}}</td>`;

    // Objetivo columns
    for (const c of allCedis) {{
      const val = obj[c];
      html += `<td>${{val != null ? val.toFixed(1) : '<span class="cell-empty">--</span>'}}</td>`;
    }}
    html += `<td style="font-weight:700;background:#fff3cd;">${{objProm != null ? objProm : '--'}}</td>`;

    // Actual columns
    for (const c of allCedis) {{
      const val = actual[c];
      const cls = cellClass(val, obj[c]);
      html += `<td class="${{cls}}">${{val != null ? val.toFixed(1) : '<span class="cell-empty">--</span>'}}</td>`;
    }}
    const actProm = actual['prom'];
    html += `<td style="font-weight:700;background:#dbeafe;">${{actProm != null ? actProm.toFixed(1) : '--'}}</td>`;

    html += '</tr>';
  }}

  // Total row
  const totObj  = mesData['__total__'] || {{}};
  const totObjVals = allCedis.map(c=>totObj[c]).filter(x=>x!=null);
  const totProm = totObjVals.length ? (totObjVals.reduce((a,b)=>a+b,0)/totObjVals.length).toFixed(1) : '--';

  html += '<tr style="border-top:2px solid #374151;">';
  html += `<td class="vendor-name cell-total">Total</td>`;
  // Objective total = blank (show em dash)
  for (const c of allCedis) html += `<td class="cell-total">--</td>`;
  html += `<td class="cell-total">--</td>`;
  // Actual totals
  for (const c of allCedis) {{
    const val = totObj[c];
    html += `<td class="cell-total">${{val != null ? val.toFixed(1) : '--'}}</td>`;
  }}
  html += `<td class="cell-total">${{totProm}}</td>`;
  html += '</tr>';

  html += '</tbody>';
  $(tableId).innerHTML = html;
}}

function renderMatrix() {{
  const mes = $('sel-mes').value;
  buildMatrixHTML('Autoservicios', 'mat-auto');
  buildMatrixHTML("SAM'S Club",   'mat-sams');

  // Update stat cards from matrix totals too
  let autoCitas = 0, samsCitas = 0;
  const mxAuto = DATA.matrix['Autoservicios'];
  const mxSams = DATA.matrix["SAM'S Club"];
  if (mxAuto) {{
    const mesD = mxAuto.meses[mes] || {{}};
    for (const v of mxAuto.vendors) {{
      // Citas from charts if available (approximate)
    }}
  }}
}}

// ─── Stats update (shared) ───────────────────────────────────────────────────
function updateStats() {{
  const cedis = $('sel-cedis').value;
  const mes   = $('sel-mes').value;
  const chartData = (DATA.charts[cedis]||{{}})[mes]||{{}};
  const autoV = chartData['Autoservicios']||[];
  const samsV = chartData["SAM'S Club"]||[];
  $('stat-citas-auto').textContent = autoV.reduce((s,v)=>s+v.citas,0).toLocaleString('es-MX');
  $('stat-citas-sams').textContent = samsV.reduce((s,v)=>s+v.citas,0).toLocaleString('es-MX');
  $('stat-los-auto').textContent   = autoV.length ? (avg(autoV.map(v=>v.total))||0).toFixed(1)+'h' : '--';
  $('stat-los-sams').textContent   = samsV.length ? (avg(samsV.map(v=>v.total))||0).toFixed(1)+'h' : '--';
}}

// ─── Trend chart (static) ────────────────────────────────────────────────────
function buildTrendChart() {{
  const ctx = $('chart-trend').getContext('2d');
  const meses = DATA.meses;
  const auto  = DATA.monthly_trend['Autoservicios']||[];
  const sams  = DATA.monthly_trend["SAM'S Club"]||[];
  return new Chart(ctx, {{
    type: 'line',
    data: {{
      labels: meses,
      datasets: [
        {{ label:'Autoservicios', data:auto.map(d=>d.avg_total), borderColor:C.recibo, backgroundColor:C.recibo+'33', fill:true, tension:0.4, pointRadius:5 }},
        {{ label:"SAM'S Club",   data:sams.map(d=>d.avg_total), borderColor:C.histProm, backgroundColor:C.histProm+'22', fill:true, tension:0.4, pointRadius:5, borderDash:[5,3] }},
      ],
    }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{ position:'bottom', labels:{{ usePointStyle:true }} }} }},
      scales:{{
        y:{{ title:{{display:true,text:'LOS Prom (h)'}}, grid:{{color:C.gridLine}}, ticks:{{callback:v=>v+'h'}} }},
        x:{{ grid:{{display:false}} }},
      }},
    }},
  }});
}}

// ─── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {{
  $('sel-cedis').value = 'CLN';
  $('sel-mes').value   = DATA.meses[0] || 'Enero';

  chartTrend = buildTrendChart();
  updateStats();
  updateCharts();

  $('sel-cedis').addEventListener('change', () => {{ updateStats(); if (currentView==='graficas') updateCharts(); else renderMatrix(); }});
  $('sel-mes').addEventListener('change',   () => {{ updateStats(); if (currentView==='graficas') updateCharts(); else renderMatrix(); }});
}});
</script>
</body>
</html>
"""

out_path = os.path.join(BASE, 'tablero_los_2026.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(HTML)

size_kb = os.path.getsize(out_path) // 1024
print(f'Dashboard saved: {out_path}  ({size_kb} KB)')
subprocess.Popen(['start', '', out_path], shell=True)
print('Opened in browser.')
