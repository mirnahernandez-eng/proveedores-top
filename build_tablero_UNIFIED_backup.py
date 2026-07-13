#!/usr/bin/env python3
"""
build_tablero.py
Generates tablero_los_proveedores.html from matrix_FINAL.json + matrix_PEREC.json
"""
import json
import os

BASE = r"C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP"
BQ   = os.path.join(BASE, "bigquery_results")
OUT  = os.path.join(BASE, "tablero_los_proveedores.html")


def j(obj):
    """Compact JSON safe for inline JS embedding."""
    return json.dumps(obj, ensure_ascii=False)


def load_data():
    with open(os.path.join(BQ, "matrix_FINAL.json"), encoding="utf-8") as f:
        final = json.load(f)
    with open(os.path.join(BQ, "matrix_PEREC.json"), encoding="utf-8") as f:
        perec = json.load(f)
    return final, perec


# ── HTML SECTIONS ──────────────────────────────────────────────────────────────

HEAD = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tablero LOS Proveedores TOP 2026</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="chart.min.js"></script>
<script src="datalabels.min.js"></script>
<style>
body { font-family: sans-serif; background:#f8fafc; }
.tbl { border-collapse: collapse; width: 100%; }
.tbl th {
  position: sticky; top: 0; z-index: 1;
  background: #1e293b; color: #f8fafc;
  font-size: 0.70rem; padding: 4px 7px;
  text-align: center; white-space: nowrap;
}
.tbl td {
  font-size: 0.70rem; padding: 3px 7px;
  text-align: center; border-bottom: 1px solid #e2e8f0;
}
.tbl tr:hover td { background: #f1f5f9; }
.v-name { text-align:left !important; max-width:200px; white-space:nowrap;
          overflow:hidden; text-overflow:ellipsis; }
.cell-ok   { background:#d1fae5; color:#065f46; font-weight:700; }
.cell-bad  { background:#fed7aa; color:#7c2d12; font-weight:700; }
.cell-null { color:#94a3b8; }
.cell-obj  { background:#dbeafe; color:#1e40af; font-weight:700; }
.cell-tot  { background:#f1f5f9; font-weight:700; }
.reg-hdr   { background:#334155; color:#e2e8f0; font-size:0.63rem;
             letter-spacing:.05em; text-transform:uppercase; }
.chart-wrap { position:relative; height:420px; }
.ms-wrap  { position:relative; display:inline-block; }
.ms-btn   { display:flex;align-items:center;gap:5px;padding:5px 10px;border:1px solid #cbd5e1;
            border-radius:6px;background:#fff;font-size:0.85rem;color:#334155;
            cursor:pointer;white-space:nowrap;min-width:100px;justify-content:space-between; }
.ms-btn:hover { border-color:#94a3b8; }
.ms-panel { position:absolute;top:calc(100% + 4px);left:0;z-index:999;background:#fff;
            border:1px solid #cbd5e1;border-radius:8px;box-shadow:0 8px 24px #0003;
            padding:4px 0;min-width:170px;max-height:260px;overflow-y:auto;display:none; }
.ms-panel.open { display:block; }
.ms-item  { display:flex;align-items:center;gap:8px;padding:5px 12px;font-size:0.8rem;cursor:pointer; }
.ms-item:hover { background:#f1f5f9; }
.ms-item input[type=checkbox] { accent-color:#2563eb;cursor:pointer; }
select { border:1px solid #cbd5e1; border-radius:6px;
         padding:6px 10px; font-size:0.85rem; background:white; cursor:pointer; }
</style>
</head>
<body class="p-4">
<div class="mb-4">
  <h1 class="text-2xl font-bold text-slate-800">&#128230; Tablero LOS Proveedores TOP 2026</h1>
  <p class="text-slate-500 text-sm mt-1">Level of Service &#8212; Tiempo total en CEDIS por proveedor y locaci&#243;n</p>
</div>

<!-- Panel Actualizar datos -->
<div id="updateBar" class="flex flex-wrap items-center gap-3 mb-4 bg-blue-50 border border-blue-200 rounded-xl px-4 py-2 text-sm">
  <span class="font-semibold text-blue-700">Actualizar datos:</span>
  <div class="flex items-center gap-1">
    <label class="text-slate-600 text-xs">Inicio:</label>
    <input type="date" id="fIni" class="border border-slate-300 rounded px-2 py-1 text-xs" value="2026-01-01">
  </div>
  <div class="flex items-center gap-1">
    <label class="text-slate-600 text-xs">Fin:</label>
    <input type="date" id="fFin" class="border border-slate-300 rounded px-2 py-1 text-xs">
  </div>
  <button onclick="lanzarActualizacion()" id="btnUpdate"
    class="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-4 py-1.5 rounded-lg transition">
    Actualizar
  </button>
  <div id="updateStatus" class="flex items-center gap-2 text-xs text-slate-600">
    <span id="statusMsg">Listo</span>
    <div id="spinnerBox" style="display:none">
      <svg class="animate-spin h-4 w-4 text-blue-600" viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" stroke-dasharray="31.4" stroke-dashoffset="10"/>
      </svg>
    </div>
  </div>
  <div id="updateBar-progress" style="display:none;width:120px;height:6px;background:#dbeafe;border-radius:3px;overflow:hidden;">
    <div id="progressFill" style="height:100%;width:0;background:#2563eb;transition:width .4s;"></div>
  </div>
</div>

<div class="flex flex-wrap gap-3 mb-6 bg-white p-3 rounded-xl shadow-sm border border-slate-200 items-center">
  <div class="flex items-center gap-2">
    <label class="text-sm font-semibold text-slate-600">Locaci&#243;n:</label>
    <select id="locSel" onchange="onLocChange()" class="text-sm border border-slate-300 rounded-md px-2 py-1">
      <option value="__all__">(Todas)</option>
      <option value="CUU">Chihuahua</option>
      <option value="CLN">Culiac&#225;n</option>
      <option value="MXL">Mexicali</option>
      <option value="MTY">Monterrey</option>
      <option value="CUAU">Cuautitl&#225;n</option>
      <option value="STB">Sta. B&#225;rbara</option>
      <option value="SMO">S. Mart&#237;n Obispo</option>
      <option value="CHL">Chalco</option>
      <option value="GDL">Guadalajara</option>
      <option value="MER">M&#233;rida</option>
      <option value="VHSA">Villahermosa</option>
    </select>
  </div>
  <div class="flex items-center gap-2">
    <label class="text-sm font-semibold text-slate-600">CD:</label>
    <div class="ms-wrap">
      <button class="ms-btn" onclick="toggleCdPanel()">
        <span id="cdBtnLbl">(Todos)</span><span style="font-size:.65rem">&#9660;</span>
      </button>
      <div id="cdPanel" class="ms-panel"></div>
    </div>
  </div>
  <div class="flex items-center gap-2">
    <label class="text-sm font-semibold text-slate-600">Mes:</label>
    <select id="mesSel" onchange="renderAll()" class="text-sm border border-slate-300 rounded-md px-2 py-1">
      <option value="prom">Prom Ene-Jun</option>
      <option value="Enero">Enero</option>
      <option value="Febrero">Febrero</option>
      <option value="Marzo">Marzo</option>
      <option value="Abril">Abril</option>
      <option value="Mayo">Mayo</option>
      <option value="Junio">Junio</option>
      <option value="Julio">Julio (1-3)</option>
      <option value="ytd">YTD Ene-Jul</option>
    </select>
  </div>
  <div id="badge" class="ml-auto text-xs text-slate-400 italic"></div>
  <div class="flex items-center gap-2">
    <label class="text-sm font-semibold text-slate-600">Semana:</label>
    <div class="ms-wrap">
      <button class="ms-btn" onclick="toggleSwPanel()">
        <span id="swBtnLbl">Todas SW</span><span style="font-size:.65rem">&#9660;</span>
      </button>
      <div id="swPanel" class="ms-panel"></div>
    </div>
  </div>
</div>

<div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4 mb-6">
  <h2 id="chartTitle" class="font-bold text-slate-700 mb-2 text-sm">&#128202; LOS por Proveedor</h2>
  <div class="chart-wrap"><canvas id="chartMain"></canvas></div>
</div>

<div id="tableMain" class="space-y-8"></div>

<div class="mt-8 text-center text-xs text-slate-400">
  Walmart Supply Chain &middot; Datos YMS 2026 &middot; Generado autom&#225;ticamente
</div>
"""

OBJECTIVES_JS = """
var OBJ_AUTO = {
  "KIMBERLY CLARK DE MEX SA B CV":  {CUU:7.9,CLN:5.4,MXL:5.0,MTY:9.7,CUAU:10.2,STB:6.7,CHL:10.2,GDL:10.2,MER:6.0,VHSA:4.6,"2026":10.2},
  "ENBOTELLAD NIAGARA D MX":        {CUU:5.5,CLN:5.5,MXL:5.5,MTY:5.5,CUAU:9.0,STB:5.5,CHL:5.5,GDL:5.5,MER:5.5,VHSA:4.1,"2026":5.9},
  "JUGOS DEL VALLE":                 {CUU:6.6,CLN:5.1,MXL:4.4,MTY:8.4,CUAU:8.6,STB:4.7,CHL:11.2,GDL:14.7,MER:5.2,VHSA:4.1,"2026":8.8},
  "SANTA CLARA MERC PACHU S RL CV": {CUU:7.4,CLN:7.4,MXL:7.4,MTY:7.4,CUAU:8.8,STB:7.4,CHL:7.4,GDL:7.4,MER:7.4,VHSA:7.4,"2026":7.4},
  "PROCTER AND GAMBLE MEXICO INC":  {CUU:9.6,CLN:7.2,MXL:6.9,MTY:10.0,CUAU:10.0,STB:7.7,CHL:10.0,GDL:10.0,MER:9.4,VHSA:5.3,"2026":10.0},
  "MARCAS NESTLE":                   {CUU:8.3,CLN:6.9,MXL:5.0,MTY:10.0,CUAU:10.0,STB:5.4,CHL:10.0,GDL:10.0,MER:7.3,VHSA:4.8,"2026":10.0},
  "COLGATE PALMOLIVE SA CV":         {CUU:11.1,CLN:7.4,MXL:8.6,MTY:10.4,CUAU:11.5,STB:7.1,CHL:11.5,GDL:11.5,MER:9.3,VHSA:6.3,"2026":11.5},
  "COMERC PEPSICO MEXICO S RL CV":   {CUU:10.8,CLN:8.7,MXL:7.1,MTY:10.2,CUAU:9.3,STB:6.1,CHL:10.8,GDL:10.8,MER:8.9,VHSA:5.6,"2026":10.8},
  "BONAFONT + ENVASASORA":           {CUU:6.8,CLN:8.0,MXL:4.5,MTY:8.0,CUAU:8.0,STB:4.4,CHL:8.0,GDL:8.0,MER:5.7,VHSA:6.0,"2026":8.0},
  "UNILEVER DE MEXICO S RL CV":      {CUU:13.5,CLN:13.0,MXL:10.1,MTY:13.4,CUAU:9.1,STB:5.3,CHL:13.5,GDL:13.5,MER:13.5,VHSA:8.8,"2026":13.5},
  "HERDEZ SA DE CV":                 {CUU:10.1,CLN:8.5,MXL:6.6,MTY:10.7,CUAU:10.1,STB:6.4,CHL:11.0,GDL:11.0,MER:8.6,VHSA:5.3,"2026":11.0},
  "CERVEZA CANAL MO S D":            {CUU:7.4,CLN:6.8,MXL:5.0,MTY:8.0,CUAU:8.0,STB:4.4,CHL:8.0,GDL:8.0,MER:5.4,VHSA:5.0,"2026":8.0},
  "FRABEL SA DE CV":                 {CUU:12.0,CLN:12.0,MXL:11.4,MTY:12.0,CUAU:12.0,STB:8.4,CHL:12.0,GDL:12.0,MER:8.1,VHSA:10.2,"2026":12.0},
  "MONDELEZ MEXICO S DE RL DE CV":   {CUU:9.5,CLN:9.5,MXL:9.5,MTY:9.5,CUAU:12.3,STB:9.5,CHL:9.5,GDL:9.5,MER:9.5,VHSA:9.5,"2026":9.5},
  "KELLOGG COMPANY MEXICO SRL CV":   {CUU:8.9,CLN:9.9,MXL:6.8,MTY:12.5,CUAU:12.5,STB:7.5,CHL:12.5,GDL:12.5,MER:10.4,VHSA:5.8,"2026":12.5},
  "__total__":                        {CUU:8.8,CLN:7.1,MXL:6.2,MTY:10.1,CUAU:10.0,STB:6.1,CHL:9.6,GDL:10.3,MER:7.3,VHSA:5.2,"2026":10.0}
};
// Sub-cedis CUAU heredan el mismo objetivo que CUAU
Object.keys(OBJ_AUTO).forEach(function(v){
  var o=OBJ_AUTO[v]; if(!o||o.CUAU===undefined) return;
  o.CUAU7494=o.CUAU; o.CUAU7464=o.CUAU; o.CUAU7492=o.CUAU;
});
var OBJ_SAMS = {
  "KIMBERLY CLARK DE MEX SA B CV":  {CUU:4.8,CLN:3.6,MTY:4.5,SMO:5.2,CHL:5.2,GDL:5.2,MER:4.6,VHSA:1.6,"2026":5.2},
  "ENBOTELLAD NIAGARA D MX":        null,
  "JUGOS DEL VALLE":                 {CUU:4.7,CLN:3.5,MTY:4.2,SMO:5.1,CHL:5.5,GDL:5.5,MER:4.1,VHSA:1.8,"2026":5.5},
  "SANTA CLARA MERC PACHU S RL CV": null,
  "PROCTER AND GAMBLE MEXICO INC":  {CUU:4.4,CLN:2.8,MTY:4.1,SMO:5.0,CHL:5.0,GDL:5.0,MER:5.0,VHSA:1.7,"2026":5.0},
  "MARCAS NESTLE":                   {CUU:4.7,CLN:4.1,MTY:4.5,SMO:4.9,CHL:5.5,GDL:5.5,MER:4.5,VHSA:1.8,"2026":5.5},
  "COLGATE PALMOLIVE SA CV":         {CUU:4.2,CLN:3.1,MTY:4.3,SMO:4.7,CHL:4.7,GDL:4.7,MER:4.7,VHSA:2.0,"2026":4.7},
  "COMERC PEPSICO MEXICO S RL CV":   {CUU:5.3,CLN:3.8,MTY:4.3,SMO:5.5,CHL:5.6,GDL:5.6,MER:5.6,VHSA:2.1,"2026":5.6},
  "BONAFONT + ENVASASORA":           {CUU:5.0,CLN:1.8,MTY:3.8,SMO:5.3,CHL:6.0,GDL:6.0,MER:6.0,VHSA:2.0,"2026":6.0},
  "UNILEVER DE MEXICO S RL CV":      {CUU:5.7,CLN:3.8,MTY:5.6,SMO:7.6,CHL:6.9,GDL:6.9,MER:4.6,VHSA:1.7,"2026":6.9},
  "HERDEZ SA DE CV":                 {CUU:6.0,CLN:4.2,MTY:5.2,SMO:5.5,CHL:7.1,GDL:6.4,MER:5.4,VHSA:1.8,"2026":6.4},
  "CERVEZA CANAL MO S D":            {CUU:5.8,CLN:3.9,MTY:4.7,SMO:4.8,CHL:5.9,GDL:5.9,MER:5.0,VHSA:3.0,"2026":5.9},
  "FRABEL SA DE CV":                 {CUU:3.5,CLN:2.8,MTY:3.6,SMO:5.0,CHL:5.0,GDL:5.0,MER:3.3,VHSA:1.5,"2026":5.0},
  "MONDELEZ MEXICO S DE RL DE CV":   null,
  "KELLOGG COMPANY MEXICO SRL CV":   {CUU:4.6,CLN:3.1,MTY:4.2,SMO:4.8,CHL:4.8,GDL:4.8,MER:4.2,VHSA:2.6,"2026":4.8},
  "__total__":                        {CUU:4.7,CLN:3.6,MTY:4.4,SMO:5.1,CHL:5.6,GDL:6.3,MER:4.3,VHSA:1.6,"2026":5.4}
};
var OBJ_PEREC_AUTO = {
  "DRISCOLL S OPERACIONES SA C":     {MTY:4.1,SMO:7.1,CHL:5.8,GDL:5.9,VHSA:4.9,"2026":5.8},
  "PILGRIMS PRIDE S DE RL DE C":     {MTY:4.1,SMO:5.2,CHL:6.0,GDL:3.6,VHSA:4.4,"2026":4.5},
  "LANDEROS PALAZUELOS EDUARDO":      {MTY:4.5,SMO:6.6,CHL:6.0,GDL:5.0,VHSA:4.3,"2026":5.6},
  "MJ INTERNATIONAL MARKETIN S":     {MTY:4.6,SMO:8.3,CHL:6.0,GDL:5.0,VHSA:4.9,"2026":6.5},
  "FRUTAS Y LEGUMBRES ALPHA SA CV":  {MTY:5.8,SMO:8.0,CHL:6.0,GDL:5.4,VHSA:5.0,"2026":6.1},
  "__total__":                        {MTY:4.4,SMO:6.6,CHL:6.0,GDL:5.0,VHSA:4.8,"2026":5.4}
};
"""

LOGIC_JS = """
// ── Locacion → CDs mapping (AUTO + SAM'S) ───────────────────────────────
var LOC_CDS = {
  CUU:  [{n:"4640 AUTOSERVICIOS",k:"CUU"},  {n:"5780 SAM'S",        k:"CUU"}],
  CLN:  [{n:"7487 AUTOSERVICIOS",k:"CLN"},  {n:"7455 BAE",          k:"CLN"}, {n:"4971 SAM'S",      k:"CLN"}],
  MXL:  [{n:"4924 AUTOSERVICIOS",k:"MXL"},  {n:"6140 SAM'S",        k:"MXL"}],
  MTY:  [{n:"7490 AUTOSERVICIOS",k:"MTY"},  {n:"7498 AUTOSERVICIOS",k:"MTY"}, {n:"7461 BAE",        k:"MTY"}, {n:"8806 BAE SUR",    k:"MTY"}, {n:"4995 SAM'S",k:"MTY"}, {n:"7502 SAM'S",k:"MTY"}],
  CUAU: [{n:"7494 NAVE 1",       k:"CUAU7494"},{n:"7464 NAVE 2",     k:"CUAU7464"},{n:"7492 NAVE 3",    k:"CUAU7492"},{n:"CUAU Combinado",k:"CUAU"}],
  STB:  [{n:"7482 AUTOSERVICIOS",k:"STB"},  {n:"7457 BAE",          k:"STB"}],
  SMO:  [{n:"7466 AUTOSERVICIOS",k:"SMO"},  {n:"4996 SAM'S",        k:"SMO"}, {n:"6388 SAM'S",      k:"SMO"}],
  CHL:  [{n:"7471 AUTOSERVICIOS",k:"CHL"},  {n:"7459 BAE",          k:"CHL"}, {n:"7505 SAM'S",      k:"CHL"}],
  GDL:  [{n:"7493 AUTOSERVICIOS",k:"GDL"},  {n:"5907 BAE",          k:"GDL"}, {n:"7460 BAE",        k:"GDL"}, {n:"6238 SAM'S",    k:"GDL"}],
  MER:  [{n:"4188 AUTOSERVICIOS",k:"MER"},  {n:"7103 BAE",          k:"MER"}, {n:"7506 SAM'S",      k:"MER"}],
  VHSA: [{n:"7468 AUTOSERVICIOS",k:"VHSA"}, {n:"7453 BAE",          k:"VHSA"},{n:"6550 SAM'S",      k:"VHSA"}]
};

function toggleCdPanel() {
  var p = document.getElementById('cdPanel');
  var isOpen = p.classList.contains('open');
  closeAllPanels();
  if (!isOpen) { p.classList.add('open'); gBackdrop.style.display='block'; }
}
function toggleSwPanel() {
  var p = document.getElementById('swPanel');
  var isOpen = p.classList.contains('open');
  closeAllPanels();
  if (!isOpen) { p.classList.add('open'); gBackdrop.style.display='block'; }
}
function closeAllPanels() {
  document.getElementById('swPanel').classList.remove('open');
  document.getElementById('cdPanel').classList.remove('open');
  gBackdrop.style.display = 'none';
}
var gBackdrop = (function(){
  var el = document.createElement('div');
  el.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:998;display:none;';
  el.onclick = closeAllPanels;
  document.body.appendChild(el);
  return el;
})();

var gSelCD = [];

var LOC_LABELS = {
  CUU:"Chihuahua",CLN:"Culiac\u00e1n",MXL:"Mexicali",MTY:"Monterrey",
  CUAU:"Cuautitl\u00e1n",STB:"Sta. B\u00e1rbara",SMO:"S. Mart\u00edn Obispo",
  CHL:"Chalco",GDL:"Guadalajara",MER:"M\u00e9rida",VHSA:"Villahermosa"
};

function buildCdPanel(loc) {
  var panel = document.getElementById('cdPanel');
  panel.innerHTML = '';
  // Botón limpiar
  var clear = document.createElement('div');
  clear.className = 'ms-item';
  clear.style.cssText = 'color:#64748b;font-style:italic;border-bottom:1px solid #e2e8f0;';
  clear.innerHTML = '&#10005; (Todos)';
  clear.onclick = function(e){ e.stopPropagation(); gSelCD=[]; syncCdPanel(); renderAll(); };
  panel.appendChild(clear);

  // CDs a mostrar: los de la locación específica, o TODOS si es __all__
  var locKeys = (loc !== '__all__' && LOC_CDS[loc]) ? [loc] : Object.keys(LOC_CDS);

  locKeys.forEach(function(lk) {
    var cds = LOC_CDS[lk]; if (!cds || !cds.length) return;
    // Separador de grupo cuando se muestran todos
    if (loc === '__all__') {
      var sep = document.createElement('div');
      sep.className = 'ms-item';
      sep.style.cssText = 'color:#1e3a8a;font-weight:700;font-size:0.72rem;background:#f0f4ff;pointer-events:none;padding:4px 12px;';
      sep.textContent = LOC_LABELS[lk] || lk;
      panel.appendChild(sep);
    }
    cds.forEach(function(d) {
      var item = document.createElement('div'); item.className = 'ms-item';
      item.dataset.val = d.k; item.dataset.lbl = d.n;
      var cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = false;
      var sp = document.createElement('span'); sp.textContent = d.n;
      item.appendChild(cb); item.appendChild(sp);
      item.addEventListener('click', function(e){ e.stopPropagation(); toggleCD(d.k, d.n); });
      panel.appendChild(item);
    });
  });
}

function onLocChange() {
  gSelCD = [];
  var loc = document.getElementById('locSel').value;
  buildCdPanel(loc);
  syncCdPanel(); renderAll();
}

function toggleCD(key, lbl) {
  var idx = gSelCD.findIndex(function(x){ return x.n===lbl; });
  if (idx >= 0) gSelCD.splice(idx,1); else gSelCD.push({k:key,n:lbl});
  syncCdPanel(); renderAll();
}
function syncCdPanel() {
  document.querySelectorAll('#cdPanel .ms-item').forEach(function(item){
    var chk = item.querySelector('input');
    if (!chk) return;
    chk.checked = gSelCD.some(function(x){ return x.n === item.dataset.lbl; });
  });
  var lbl = document.getElementById('cdBtnLbl');
  lbl.textContent = gSelCD.length === 0 ? '(Todos)' :
                    gSelCD.length === 1 ? gSelCD[0].n : gSelCD.length + ' CDs';
}

function getLocKey() {
  if (gSelCD.length === 1) return gSelCD[0].k;  // un CD exacto
  if (gSelCD.length > 1) {
    // todos los keys son iguales (ej: 3 CDs AUTO de MTY -> todos k:"MTY")
    var first = gSelCD[0].k;
    var allSame = gSelCD.every(function(x){ return x.k === first; });
    return allSame ? first : document.getElementById('locSel').value;
  }
  var locVal = document.getElementById('locSel').value;
  return locVal !== '__all__' ? locVal : '__all__';
}

// Returns "auto", "bae", "sams", or "all" based on selected CDs
function getChannelFilter() {
  if (!gSelCD.length) return "all";
  var hasAuto = gSelCD.some(function(x){
    return x.n.indexOf("AUTOSERVICIOS") >= 0 || x.n.indexOf("NAVE") >= 0 || x.n.indexOf("Combinado") >= 0;
  });
  var hasBae  = gSelCD.some(function(x){ return x.n.indexOf("BAE") >= 0; });
  var hasSams = gSelCD.some(function(x){ return x.n.indexOf("SAM") >= 0; });
  var count = (hasAuto?1:0) + (hasBae?1:0) + (hasSams?1:0);
  if (count > 1) return "all";
  if (hasAuto) return "auto";
  if (hasBae)  return "bae";
  if (hasSams) return "sams";
  return "all";
}

var SHORT_NAMES = {
  "KIMBERLY CLARK DE MEX SA B CV":"KIMBERLY","ENBOTELLAD NIAGARA D MX":"NIAGARA",
  "JUGOS DEL VALLE":"JUGOS","SANTA CLARA MERC PACHU S RL CV":"STA CLARA",
  "PROCTER AND GAMBLE MEXICO INC":"PROCTER","MARCAS NESTLE":"NESTLE",
  "COLGATE PALMOLIVE SA CV":"COLGATE","COMERC PEPSICO MEXICO S RL CV":"PEPSICO",
  "BONAFONT + ENVASASORA":"BONAFONT","UNILEVER DE MEXICO S RL CV":"UNILEVER",
  "HERDEZ SA DE CV":"HERDEZ","CERVEZA CANAL MO S D":"CANAL MO",
  "FRABEL SA DE CV":"FRABEL","MONDELEZ MEXICO S DE RL DE CV":"MONDELEZ",
  "KELLOGG COMPANY MEXICO SRL CV":"KELLOGG"
};

var REGIONS_AUTO  = [{n:"NORTE",l:["CUU","CLN","MXL","MTY"]},{n:"CENTRO",l:["CUAU","STB"]},{n:"SUR",l:["CHL","GDL","MER","VHSA"]}];
var REGIONS_BAE   = [{n:"NORTE",l:["CLN","MTY"]},{n:"CENTRO",l:["STB"]},{n:"SUR",l:["CHL","GDL","MER","VHSA"]}];
var CUAU_SUBS     = ["CUAU7494","CUAU7464","CUAU7492"];
var REGIONS_SAMS  = [{n:"NORTE",l:["CUU","CLN","MTY"]},{n:"CENTRO",l:["SMO"]},{n:"SUR",l:["CHL","GDL","MER","VHSA"]}];
var SW_DATA = null;
var gSelSW  = [];

var ALL_PER = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","ytd"];
var PER_LBL = {prom:"Prom",Enero:"Enero",Febrero:"Febrero",Marzo:"Marzo",Abril:"Abril",Mayo:"Mayo",Junio:"Junio",Julio:"Jul (1-3)",ytd:"YTD"};
var PER_BG  = {prom:"#1e3a8a",Enero:"#1d4ed8",Febrero:"#1d4ed8",Marzo:"#1d4ed8",Abril:"#1d4ed8",Mayo:"#1d4ed8",Junio:"#1d4ed8",Julio:"#d97706",ytd:"#065f46"};
var REG_BG  = {NORTE:"#1d4ed8",CENTRO:"#6d28d9",SUR:"#065f46"};

var gChartMain = null;

function fmt(v) {
  if (v === null || v === undefined) return "\u2014";
  return parseFloat(v).toFixed(1);
}

function cellCls(val, obj) {
  if (val === null || val === undefined) return "cell-null";
  if (obj === null || obj === undefined) return "";
  return val <= obj ? "cell-ok" : "cell-bad";
}

function flatLocs(regions) {
  var out = [];
  for (var i = 0; i < regions.length; i++) {
    for (var k = 0; k < regions[i].l.length; k++) out.push(regions[i].l[k]);
  }
  return out;
}

function toggleSwPanel() {
  var p = document.getElementById('swPanel');
  p.classList.toggle('open');
}
document.addEventListener('click', function(e) {
  if (!e.target.closest('.ms-wrap')) document.getElementById('swPanel').classList.remove('open');
});
function toggleSW(key) {
  var idx = gSelSW.indexOf(key);
  if (idx >= 0) gSelSW.splice(idx,1); else gSelSW.push(key);
  syncSwPanel(); renderAll();
}
function syncSwPanel() {
  document.querySelectorAll('#swPanel .ms-item').forEach(function(item){
    var chk = item.querySelector('input');
    if (chk) chk.checked = gSelSW.indexOf(item.dataset.val) >= 0;
  });
  var lbl = document.getElementById('swBtnLbl');
  lbl.textContent = gSelSW.length === 0 ? 'Todas SW' :
                    gSelSW.length === 1 ? gSelSW[0] : gSelSW.length + ' semanas';
}
function mergeSWChartData(base, keys) {
  var out = {};
  Object.keys(base).forEach(function(v) {
    out[v] = {};
    Object.keys(base[v]).forEach(function(cedis) {
      var lS=0,rS=0,sS=0,tS=0,n=0;
      keys.forEach(function(k){ var d=base[v][cedis][k]; if(!d) return; lS+=d.l||0;rS+=d.r||0;sS+=d.s||0;tS+=d.t||0;n++; });
      out[v][cedis] = Object.assign({}, base[v][cedis]);
      if (n) out[v][cedis]['__sel__'] = {l:lS/n,r:rS/n,s:sS/n,t:tS/n};
    });
  });
  return out;
}
function mergeSWTblData(base, keys) {
  var out = {};
  Object.keys(base).forEach(function(v) {
    out[v] = Object.assign({}, base[v]);
    var merged={}, cnt={};
    keys.forEach(function(k){ var d=base[v][k]; if(!d) return; Object.keys(d).forEach(function(c){ merged[c]=(merged[c]||0)+(d[c]||0); cnt[c]=(cnt[c]||0)+1; }); });
    var avg={}; Object.keys(merged).forEach(function(c){ avg[c]=cnt[c]?merged[c]/cnt[c]:null; });
    out[v]['__sel__'] = avg;
  });
  return out;
}

function renderAll() {
  var loc    = getLocKey();
  var period = document.getElementById("mesSel").value;
  var locEl  = document.getElementById("locSel");
  var perLbl = document.getElementById("mesSel").options[document.getElementById("mesSel").selectedIndex].text;
  var swLbl  = gSelSW.length ? (" \u00b7 " + (gSelSW.length===1 ? gSelSW[0] : gSelSW.length+" semanas")) : "";

  // ── Badge ─────────────────────────────────────────────────────────────
  var badgeLoc;
  if (gSelCD.length === 1) {
    var cityKey = gSelCD[0].k.replace(/[0-9]+$/,'').replace('CUAU7494','CUAU').replace('CUAU7464','CUAU').replace('CUAU7492','CUAU');
    var cityLbl = LOC_LABELS[gSelCD[0].k] || LOC_LABELS[cityKey] || gSelCD[0].k;
    badgeLoc = gSelCD[0].n + " (" + cityLbl + ")";
  } else if (gSelCD.length > 1) {
    badgeLoc = gSelCD.length + " CDs";
  } else if (locEl.value !== "__all__") {
    badgeLoc = locEl.options[locEl.selectedIndex].text;
  } else {
    badgeLoc = "Todas las locaciones";
  }
  document.getElementById("badge").textContent = "Mostrando: " + badgeLoc + " \u00b7 " + perLbl + swLbl;

  // ── Resolver fuente de datos (SW o mensual) ───────────────────────────
  var useSW = gSelSW.length > 0 && SW_DATA;
  var swPer = gSelSW.length === 1 ? gSelSW[0] : '__sel__';
  var autoChartData, baeChartData, samsChartData, autoBaeChartData,
      autoTblData, baeTblData, autoBaeTblData, samsTblData, chartPeriod, tblPeriod;
  if (useSW) {
    autoChartData    = gSelSW.length > 1 ? mergeSWChartData(SW_DATA.auto, gSelSW) : SW_DATA.auto;
    baeChartData     = gSelSW.length > 1 ? mergeSWChartData(SW_DATA.bae,  gSelSW) : SW_DATA.bae;
    samsChartData    = gSelSW.length > 1 ? mergeSWChartData(SW_DATA.sams, gSelSW) : SW_DATA.sams;
    autoBaeChartData = autoChartData; // SW_DATA no tiene auto_bae_chart, fallback
    autoTblData      = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_auto,     gSelSW) : SW_DATA.tbl_auto;
    baeTblData       = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_bae,      gSelSW) : SW_DATA.tbl_bae;
    autoBaeTblData   = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_auto_bae, gSelSW) : SW_DATA.tbl_auto_bae;
    samsTblData      = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_sams,     gSelSW) : SW_DATA.tbl_sams;
    chartPeriod = tblPeriod = swPer;
  } else {
    autoChartData    = DATA_AUTO_CHART;     baeChartData = DATA_BAE_CHART; samsChartData = DATA_SAMS_CHART;
    autoBaeChartData = DATA_AUTO_BAE_CHART;
    autoTblData      = DATA_AUTO;           baeTblData   = DATA_BAE;       samsTblData   = DATA_SAMS;
    autoBaeTblData   = DATA_AUTO_BAE;
    chartPeriod = tblPeriod = period;
  }

  // ── Elegir canal activo → una gráfica, una tabla ──────────────────────
  var chFilter = getChannelFilter();
  var activeChartData, activeTblData, activeObjData, activeRegions, activeTitle;
  if (chFilter === "sams") {
    activeChartData = samsChartData;
    activeTblData   = samsTblData;
    activeObjData   = OBJ_SAMS;
    activeRegions   = REGIONS_SAMS;
    activeTitle     = "\U0001F4CA SAM\u2019S Club \u2014 LOS por Proveedor";
  } else if (chFilter === "bae") {
    activeChartData = baeChartData;
    activeTblData   = baeTblData;
    activeObjData   = OBJ_AUTO;
    activeRegions   = REGIONS_BAE;
    activeTitle     = "\U0001F4CA BAE \u2014 LOS por Proveedor";
  } else if (chFilter === "auto") {
    activeChartData = autoChartData;
    activeTblData   = autoTblData;
    activeObjData   = OBJ_AUTO;
    activeRegions   = REGIONS_AUTO;
    activeTitle     = "\U0001F4CA Autoservicios \u2014 LOS por Proveedor";
  } else {
    // "all" → Total (AUTO + BAE combinados) – gráfica y tabla usan el mismo dataset
    activeChartData = autoBaeChartData;
    activeTblData   = autoBaeTblData;
    activeObjData   = OBJ_AUTO;
    activeRegions   = REGIONS_AUTO;
    activeTitle     = "\U0001F4CA Total \u2014 LOS por Proveedor";
  }

  document.getElementById("chartTitle").textContent = activeTitle;

  // ── Una sola gráfica ──────────────────────────────────────────────────
  gChartMain = paintChart("chartMain", activeChartData, activeObjData, DISPLAY_ORDER, loc, chartPeriod, gChartMain);

  // ── Una sola tabla ────────────────────────────────────────────────────
  var tableWrap = document.getElementById("tableMain");
  tableWrap.innerHTML = "";
  tableWrap.appendChild(buildTable(
    activeTitle.replace(/^\U0001F4CA /, ""),
    activeTblData, activeObjData, DISPLAY_ORDER,
    activeRegions, loc, tblPeriod, true, [tblPeriod]
  ));
}

function paintChart(canvasId, chartData, objData, order, loc, period, existing, filtLocs) {
  var canvas = document.getElementById(canvasId);
  if (!canvas) return null;
  var isFiltAll = (filtLocs && filtLocs.length > 0 && loc === "__all__");
  var locKey = isFiltAll ? null : (loc === "__all__") ? "2026" : loc;

  var vendors=[], lA=[], rA=[], sA=[], tA=[], oA=[];
  for (var i = 0; i < order.length; i++) {
    var v = order[i];
    var vd = chartData[v]; if (!vd) continue;
    if (isFiltAll) {
      var lS=0,rS=0,sS=0,tS=0,cn=0;
      for (var fi=0; fi<filtLocs.length; fi++) {
        var fld=vd[filtLocs[fi]]; if(!fld) continue;
        var fpd=fld[period]; if(!fpd) continue;
        lS+=fpd.l||0; rS+=fpd.r||0; sS+=fpd.s||0; tS+=fpd.t||0; cn++;
      }
      if (!cn) continue;
      vendors.push(v);
      lA.push(lS/cn); rA.push(rS/cn); sA.push(sS/cn); tA.push(tS/cn);
      var ov=objData?objData[v]:null, oSm=0, oc=0;
      if (ov) { for(var fi=0;fi<filtLocs.length;fi++){if(ov[filtLocs[fi]]!==undefined){oSm+=ov[filtLocs[fi]];oc++;}} }
      oA.push(oc?oSm/oc:(ov&&ov["2026"]!==undefined?ov["2026"]:null));
    } else {
      var ld = vd[locKey]; if (!ld) continue;
      var pd = ld[period]; if (!pd) continue;
      vendors.push(v);
      lA.push(pd.l || 0); rA.push(pd.r || 0); sA.push(pd.s || 0); tA.push(pd.t || 0);
      var ov = objData ? objData[v] : null;
      oA.push(ov ? (ov[locKey] !== undefined ? ov[locKey] : null) : null);
    }
  }

  var idx = [];
  for (var i = 0; i < vendors.length; i++) idx.push(i);
  idx.sort(function(a, b) { return tA[b] - tA[a]; });

  var labels = idx.map(function(i) { return SHORT_NAMES[vendors[i]] || vendors[i]; });
  var lS = idx.map(function(i){return lA[i];}), rS = idx.map(function(i){return rA[i];});
  var sS = idx.map(function(i){return sA[i];}), tS = idx.map(function(i){return tA[i];});
  var oS = idx.map(function(i){return oA[i];});

  if (existing) { try { existing.destroy(); } catch(e) {} }

  // Compute unified Y max so stacked-bar axis and absolute-line axis share the same range
  var barMax = 0;
  for (var mi = 0; mi < lS.length; mi++) {
    barMax = Math.max(barMax, (lS[mi]||0) + (rS[mi]||0) + (sS[mi]||0));
  }
  var lineVals = tS.concat(oS).filter(function(v){return v!=null;});
  var lineMax  = lineVals.length ? Math.max.apply(null, lineVals) : 0;
  var yMax     = Math.ceil(Math.max(barMax, lineMax)) + 3;

  return new Chart(canvas.getContext("2d"), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {type:"bar",  label:"Llegada",  data:lS, backgroundColor:"#9ca3af", stack:"stk", order:2,
         yAxisID:"y",
         datalabels:{display:true, color:"#374151", font:{size:8},
           formatter:function(v){return v>0.05?parseFloat(v).toFixed(1):"";}, anchor:"center", align:"center"}},
        {type:"bar",  label:"Recibo",   data:rS, backgroundColor:"#f97316", stack:"stk", order:2,
         yAxisID:"y",
         datalabels:{display:true, color:"#fff", font:{size:9,weight:"bold"},
           formatter:function(v){return v>0.05?parseFloat(v).toFixed(1):"";}, anchor:"center", align:"center"}},
        {type:"bar",  label:"Salida CD",data:sS, backgroundColor:"#fbbf24", stack:"stk", order:2,
         yAxisID:"y",
         datalabels:{display:true, color:"#78350f", font:{size:8,weight:"bold"},
           formatter:function(v){return v>0.05?parseFloat(v).toFixed(1):"";}, anchor:"center", align:"center"}},
        {type:"line", label:"Total",    data:tS, borderColor:"#111827", backgroundColor:"#111827",
         pointRadius:0, borderWidth:1, tension:0, fill:false, order:1,
         yAxisID:"y2",
         datalabels:{display:true, color:"#fff",
           backgroundColor:"#111827", borderRadius:3,
           padding:{top:2,bottom:2,left:4,right:4},
           font:{size:8,weight:"bold"},
           formatter:function(v){return v!=null?parseFloat(v).toFixed(1):"";},
           anchor:"end", align:"top", offset:4}},
        {type:"line", label:"Objetivo", data:oS, borderColor:"#2563eb", backgroundColor:"transparent",
         pointRadius:0, borderWidth:1, tension:0, fill:false, order:1,
         yAxisID:"y2",
         datalabels:{display:true, color:"#2563eb", font:{size:9,weight:"bold"},
           formatter:function(v){return v!=null?parseFloat(v).toFixed(1):"";},
           anchor:"end",
           align:function(ctx){
             var tot=ctx.chart.data.datasets[3].data[ctx.dataIndex];
             var obj=ctx.dataset.data[ctx.dataIndex];
             return (tot!=null&&obj!=null&&Math.abs(tot-obj)<2)?"right":"top";
           },
           offset:4}}
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      layout: {padding:{top:28}},
      plugins: {
        legend:  {position:"bottom", labels:{font:{size:10}, boxWidth:12}},
        tooltip: {mode:"index", intersect:false}
      },
      scales: {
        x:  {stacked:true, ticks:{font:{size:9}}},
        y:  {stacked:true, beginAtZero:true, min:0, max:yMax,
             afterBuildTicks:function(axis){
               var t=[];
               for(var i=0;i<=axis.max;i++) t.push({value:i});
               axis.ticks=t;
             },
             ticks:{font:{size:9}, autoSkip:false, maxTicksLimit:yMax+1},
             title:{display:true, text:"Horas", font:{size:9}}},
        y2: {stacked:false, beginAtZero:true, min:0, max:yMax,
             display:false}
      }
    }
  });
}

function buildTable(title, data, objData, order, regions, loc, period, hasObj, overridePeriods) {
  var allLocs = flatLocs(regions);
  var isAll   = (loc === "__all__");
  var inTable = (allLocs.indexOf(loc) >= 0) || (CUAU_SUBS.indexOf(loc) >= 0 && title === "AUTO Secos");
  var PERIODS = overridePeriods || ALL_PER;

  var wrap = document.createElement("div");
  wrap.className = "bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden";

  var hdr = document.createElement("div");
  hdr.className = "bg-slate-800 text-white px-4 py-2 font-bold text-sm";
  var titleSpan = document.createElement("span"); titleSpan.textContent = title; hdr.appendChild(titleSpan);
  wrap.appendChild(hdr);

  if (!isAll && !inTable) {
    var msg = document.createElement("p");
    msg.className = "p-4 text-sm text-slate-400 italic";
    msg.textContent = "Sin datos para la locaci\u00f3n seleccionada en esta tabla.";
    wrap.appendChild(msg);
    return wrap;
  }

  var scrl = document.createElement("div"); scrl.className = "overflow-x-auto";
  wrap.appendChild(scrl);
  var tbl = document.createElement("table"); tbl.className = "tbl";
  scrl.appendChild(tbl);
  var thead = document.createElement("thead");
  var tbody = document.createElement("tbody");
  tbl.appendChild(thead); tbl.appendChild(tbody);

  if (isAll) {
    // ── ALL LOCS: Objetivos | Ene | Feb | ... | Jul | YTD ──────────────────
    var numLoc = allLocs.length;
    var secCols = numLoc + 1; // locs + 2026

    // Row 1: section labels
    var tr1 = document.createElement("tr");
    var thV = th("Proveedores TOP", 3, 1, "v-name text-left", "#1e3a8a");
    tr1.appendChild(thV);
    // Objetivos block
    tr1.appendChild(th("Objetivos LOS 2026 por Cedis", 1, secCols, "", "#1e3a8a"));
    // Month blocks
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var bg = PER_BG[PERIODS[pi]] || "#1d4ed8";
      tr1.appendChild(th(PER_LBL[PERIODS[pi]], 1, secCols, "", bg));
    }
    thead.appendChild(tr1);

    // Row 2: region sub-headers (repeated for obj + each period)
    var tr2 = document.createElement("tr");
    var blocks = PERIODS.length + 1; // obj + periods
    for (var bi = 0; bi < blocks; bi++) {
      for (var ri = 0; ri < regions.length; ri++) {
        var rbg = REG_BG[regions[ri].n] || "#334155";
        tr2.appendChild(th(regions[ri].n, 1, regions[ri].l.length, "reg-hdr", rbg));
      }
      tr2.appendChild(th("2026", 1, 1, "", "#b45309"));
    }
    thead.appendChild(tr2);

    // Row 3: loc codes
    var tr3 = document.createElement("tr");
    for (var bi = 0; bi < blocks; bi++) {
      for (var li = 0; li < allLocs.length; li++) {
        tr3.appendChild(th(allLocs[li], 1, 1, "", "#3b82f6"));
      }
      tr3.appendChild(th("Prom", 1, 1, "", "#d97706"));
    }
    thead.appendChild(tr3);

    // Vendor rows
    for (var vi = 0; vi < order.length; vi++) {
      var v = order[vi];
      var vd = data[v]; if (!vd) continue;
      var ov = (objData && objData[v]) ? objData[v] : null;
      var row = document.createElement("tr");
      var tdN = td(v, "v-name"); tdN.title = v; row.appendChild(tdN);

      // Objectives block
      for (var li = 0; li < allLocs.length; li++) {
        row.appendChild(td(fmt(ov ? ov[allLocs[li]] : null), "cell-obj"));
      }
      row.appendChild(td(fmt(ov ? ov["2026"] : null), "cell-obj"));

      // Period blocks
      for (var pi = 0; pi < PERIODS.length; pi++) {
        var pp = PERIODS[pi];
        var pd = vd[pp] || {};
        for (var li = 0; li < allLocs.length; li++) {
          var lc = allLocs[li];
          var val = (pd[lc] !== undefined) ? pd[lc] : null;
          var objL = ov ? (ov[lc] !== undefined ? ov[lc] : null) : null;
          row.appendChild(td(fmt(val), hasObj ? cellCls(val, objL) : (val===null?"cell-null":"")));
        }
        var val26 = (pd["2026"] !== undefined) ? pd["2026"] : null;
        var obj26 = ov ? (ov["2026"] !== undefined ? ov["2026"] : null) : null;
        row.appendChild(td(fmt(val26), hasObj ? cellCls(val26, obj26) : (val26===null?"cell-null":"")));
      }
      tbody.appendChild(row);
    }

    // Total row
    var totRow = document.createElement("tr");
    totRow.appendChild(td("Total", "v-name cell-tot font-bold"));
    var objTot = (objData && objData["__total__"]) ? objData["__total__"] : null;
    // Total objectives
    for (var li = 0; li < allLocs.length; li++) {
      totRow.appendChild(td(fmt(objTot ? objTot[allLocs[li]] : null), "cell-tot font-bold"));
    }
    totRow.appendChild(td(fmt(objTot ? objTot["2026"] : null), "cell-tot font-bold"));
    // Total per period
    var totData = data["__total__"] || {};
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var pp = PERIODS[pi];
      var pd = totData[pp] || {};
      var objLT = objTot;
      for (var li = 0; li < allLocs.length; li++) {
        var lc = allLocs[li];
        var val = (pd[lc] !== undefined) ? pd[lc] : null;
        var objL = objLT ? (objLT[lc] !== undefined ? objLT[lc] : null) : null;
        totRow.appendChild(td(fmt(val), hasObj ? cellCls(val, objL) : "cell-tot"));
      }
      var val26 = (pd["2026"] !== undefined) ? pd["2026"] : null;
      var obj26 = objLT ? (objLT["2026"] !== undefined ? objLT["2026"] : null) : null;
      totRow.appendChild(td(fmt(val26), hasObj ? cellCls(val26, obj26) : "cell-tot"));
    }
    tbody.appendChild(totRow);

  } else {
    // ── SINGLE LOC: months as columns ──────────────────────────────────────
    var trH = document.createElement("tr");
    trH.appendChild(th("Proveedores TOP", 1, 1, "v-name text-left", "#1e3a8a"));
    trH.appendChild(th("Objetivo", 1, 1, "", "#1e3a8a"));
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var bg = PER_BG[PERIODS[pi]] || "#1d4ed8";
      trH.appendChild(th(PER_LBL[PERIODS[pi]], 1, 1, "", bg));
    }
    thead.appendChild(trH);

    for (var vi = 0; vi < order.length; vi++) {
      var v = order[vi];
      var vd = data[v]; if (!vd) continue;
      var ov = (objData && objData[v]) ? objData[v] : null;
      var objL = ov ? (ov[loc] !== undefined ? ov[loc] : null) : null;
      var anyData = false;
      for (var pi = 0; pi < PERIODS.length; pi++) {
        if (vd[PERIODS[pi]] && vd[PERIODS[pi]][loc] != null) { anyData=true; break; }
      }
      if (!anyData) continue;
      var row = document.createElement("tr");
      var tdN = td(v, "v-name"); tdN.title=v; row.appendChild(tdN);
      row.appendChild(td(fmt(objL), "cell-obj"));
      for (var pi = 0; pi < PERIODS.length; pi++) {
        var val = (vd[PERIODS[pi]] && vd[PERIODS[pi]][loc] != null) ? vd[PERIODS[pi]][loc] : null;
        row.appendChild(td(fmt(val), hasObj ? cellCls(val, objL) : (val===null?"cell-null":"")));
      }
      tbody.appendChild(row);
    }

    var totRow = document.createElement("tr");
    var objLT = (objData && objData["__total__"]) ? (objData["__total__"][loc] !== undefined ? objData["__total__"][loc] : null) : null;
    totRow.appendChild(td("Total", "v-name cell-tot"));
    totRow.appendChild(td(fmt(objLT), "cell-tot"));
    var totData = data["__total__"] || {};
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var val = (totData[PERIODS[pi]] && totData[PERIODS[pi]][loc] != null) ? totData[PERIODS[pi]][loc] : null;
      totRow.appendChild(td(fmt(val), hasObj ? cellCls(val, objLT) : "cell-tot"));
    }
    tbody.appendChild(totRow);
  }

  return wrap;
}

function th(text, rs, cs, cls, bg) {
  var el = document.createElement("th");
  el.textContent = text;
  if (rs > 1) el.rowSpan = rs;
  if (cs > 1) el.colSpan = cs;
  if (cls)    el.className = cls;
  if (bg)     el.style.background = bg;
  return el;
}
function td(text, cls) {
  var el = document.createElement("td");
  el.textContent = (text === null || text === undefined) ? "\u2014" : text;
  if (cls) el.className = cls;
  return el;
}

// ── Panel Actualizar ─────────────────────────────────────────────────────
var _pollTimer = null;
function lanzarActualizacion() {
  var fi = document.getElementById('fIni').value;
  var ff = document.getElementById('fFin').value;
  if (!fi || !ff) { alert('Selecciona fecha inicio y fin'); return; }
  if (fi > ff)    { alert('La fecha inicio debe ser menor o igual a la fecha fin'); return; }
  document.getElementById('btnUpdate').disabled = true;
  document.getElementById('spinnerBox').style.display = 'inline-flex';
  document.getElementById('updateBar-progress').style.display = 'block';
  document.getElementById('statusMsg').textContent = 'Iniciando...';
  fetch('/api/actualizar', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({fecha_inicio: fi, fecha_fin: ff})
  }).then(function(r){ return r.json(); }).then(function(){
    _pollTimer = setInterval(pollEstado, 1500);
  }).catch(function(e){
    mostrarEstado({ok:false, msg:'Error: '+e, pct:0, running:false});
  });
}
function pollEstado() {
  fetch('/api/estado').then(function(r){ return r.json(); }).then(function(e){
    mostrarEstado(e);
    if (!e.running) {
      clearInterval(_pollTimer);
      if (e.ok) {
        setTimeout(function(){
          location.reload();
        }, 1200);
      }
    }
  });
}
function mostrarEstado(e) {
  document.getElementById('statusMsg').textContent = e.msg;
  document.getElementById('progressFill').style.width = (e.pct||0) + '%';
  if (!e.running) {
    document.getElementById('btnUpdate').disabled = false;
    document.getElementById('spinnerBox').style.display = 'none';
    document.getElementById('progressFill').style.background = e.ok ? '#16a34a' : '#dc2626';
  }
}
// Fecha fin default = hoy
document.getElementById('fFin').value = new Date().toISOString().slice(0,10);

try { Chart.register(ChartDataLabels); } catch(e) { /* datalabels optional */ }
fetch('sw_data.json').then(function(r){return r.json();}).then(function(d){
  SW_DATA = d;
  var panel = document.getElementById('swPanel');
  var clear = document.createElement('div');
  clear.className='ms-item'; clear.style.cssText='color:#64748b;font-style:italic;border-bottom:1px solid #e2e8f0;';
  clear.innerHTML='&#10005; Limpiar';
  clear.onclick=function(e){e.stopPropagation();gSelSW=[];syncSwPanel();renderAll();};
  panel.appendChild(clear);
  d.sw_list.forEach(function(sw){
    var key='SW'+sw; var mes=d.sw_mes_map[key]||'';
    var item=document.createElement('div'); item.className='ms-item'; item.dataset.val=key;
    var cb=document.createElement('input'); cb.type='checkbox'; cb.checked=false;
    var sp=document.createElement('span'); sp.textContent='SW '+sw+' \u2014 '+mes;
    item.appendChild(cb); item.appendChild(sp);
    item.addEventListener('click',function(e){e.stopPropagation();toggleSW(key);});
    panel.appendChild(item);
  });
}).catch(function(){console.warn('sw_data.json no encontrado');});
onLocChange();  // poblar panel CD al cargar
renderAll();
</script>
</body>
</html>
"""


def build_html(final, perec):
    parts = [HEAD]

    # ── Embedded data ──
    parts.append("<script>\n")
    parts.append(f"var DATA_AUTO       = {j(final.get('auto', {}))};\n")
    parts.append(f"var DATA_BAE        = {j(final.get('bae', {}))};\n")
    parts.append(f"var DATA_AUTO_BAE   = {j(final.get('auto_bae', {}))};\n")
    parts.append(f"var DATA_SAMS       = {j(final.get('sams', {}))};\n")
    parts.append(f"var DATA_AUTO_CHART     = {j(final.get('auto_chart', {}))};\n")
    parts.append(f"var DATA_BAE_CHART      = {j(final.get('bae_chart', {}))};\n")
    parts.append(f"var DATA_AUTO_BAE_CHART = {j(final.get('auto_bae_chart', {}))};\n")
    parts.append(f"var DATA_SAMS_CHART     = {j(final.get('sams_chart', {}))};\n")
    parts.append(f"var DATA_PEREC_AUTO = {j(perec.get('auto', {}))};\n")
    parts.append(f"var DATA_PEREC_SAMS = {j(perec.get('sams', {}))};\n")
    parts.append(f"var DISPLAY_ORDER   = {j(final.get('display_order', []))};\n")
    parts.append(f"var PEREC_ORDER     = {j(perec.get('display_order', []))};\n")
    parts.append(OBJECTIVES_JS)
    parts.append(LOGIC_JS)

    return "".join(parts)


def main():
    print("Loading JSON data...")
    final, perec = load_data()
    print(f"  AUTO vendors:  {len(final.get('auto', {}))}")
    print(f"  SAMS vendors:  {len(final.get('sams', {}))}")
    print(f"  PEREC vendors: {len(perec.get('auto', {}))}")

    print("Building HTML...")
    html = build_html(final, perec)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUT) // 1024
    print(f"Done! -> {OUT}")
    print(f"Size: {size_kb} KB ({len(html):,} chars)")


if __name__ == "__main__":
    main()
