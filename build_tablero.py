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
    # Fecha hasta la que están actualizados los datos (última SW en sw_data.json)
    last_date = ""
    sw_path = os.path.join(BASE, "sw_data.json")
    if os.path.exists(sw_path):
        with open(sw_path, encoding="utf-8-sig") as f:
            sw = json.load(f)
        dates = sw.get("sw_dates", {})
        if dates:
            last_fin = sorted(v["fin"] for v in dates.values() if "fin" in v)[-1]
            # Formatear como DD/MM/YYYY
            y, m, d = last_fin.split("-")
            meses = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
            last_date = f"{d} {meses[int(m)-1]} {y}"
    return final, perec, last_date


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
  text-align: center; border-bottom: 1px solid #edf0f3;
}
.tbl tr:hover td { background: #f1f5f9; }
.v-name { text-align:left !important; max-width:200px; white-space:nowrap;
          overflow:hidden; text-overflow:ellipsis; }
.cell-ok   { color:#16a34a; }
.cell-bad  { color:#b91c1c; }
.col-active { background:rgba(245,158,11,0.07); }
.cell-null { color:#94a3b8; }
.cell-obj  { background:#dbeafe; color:#1e40af; font-weight:700; }
.cell-tot  { background:#f1f5f9; font-weight:700; font-size:0.78rem; border-top:2px solid #d1d5db; }
.reg-split { border-left:2px solid #e2e8f0 !important; }
.cell-prom { font-weight:700; font-size:0.78rem; border-left:2px solid #e2e8f0; }
.tbl td.font-bold { font-size:0.78rem; }
.delta-ok  { color:#16a34a; font-size:0.60rem; font-weight:700; display:block; line-height:1.2; }
.delta-bad { color:#dc2626; font-size:0.60rem; font-weight:700; display:block; line-height:1.2; }
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
    <label class="text-slate-600 text-xs">Actualizar hasta:</label>
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
  <span class="ml-auto text-xs text-slate-500">&#128197; Datos al: <b id="lastDateLbl" class="text-slate-700">DATA_LAST_DATE_PH</b></span>
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
    <label class="text-sm font-semibold text-slate-600">Formato:</label>
    <select id="fmtSel" onchange="onFmtChange()" class="text-sm border border-slate-300 rounded-md px-2 py-1">
      <option value="all">Todos</option>
      <option value="auto">Solo Auto</option>
      <option value="bae">Solo BAE</option>
      <option value="sams">Solo SAM'S</option>
      <option value="auto_bae">Auto + BAE</option>
      <option value="auto_sams">Auto + SAM'S</option>
      <option value="bae_sams">BAE + SAM'S</option>
    </select>
  </div>
  <div class="flex items-center gap-2">
    <label class="text-sm font-semibold text-slate-600">Mes:</label>
    <select id="mesSel" onchange="renderAll()" class="text-sm border border-slate-300 rounded-md px-2 py-1">
      <option value="Enero">Enero</option>
      <option value="Febrero">Febrero</option>
      <option value="Marzo">Marzo</option>
      <option value="Abril">Abril</option>
      <option value="Mayo">Mayo</option>
      <option value="Junio">Junio</option>
      <option value="Julio">Julio</option>
      <option value="ytd" selected>YTD</option>
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
  <h2 class="font-bold text-slate-700 text-sm mb-3" id="chartTitle">&#128202; LOS por Proveedor</h2>
  <div id="chartsContainer"><div class="chart-wrap"><canvas id="chartMain"></canvas></div></div>
</div>

<div id="tables" class="space-y-8"></div>

<div class="mt-2 text-center text-xs text-slate-400">
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
// ── Objetivos Solo Autoservicios (sin BAE) ─────────────────────────────────
var OBJ_SOLO_AUTO = {
  "KIMBERLY CLARK DE MEX SA B CV":  {CUU:7.9,CLN:5.0,MXL:9.8,MTY:10.2,CUAU:10.2,CUAU7494:10.2,CUAU7464:7.5,CUAU7492:10.2,"7464":7.5,"5907":7.5,STB:10.2,CHL:10.2,GDL:10.2,MER:4.8,VHSA:8.5,"2026":7.0},
  "ENBOTELLAD NIAGARA D MX":        null,
  "JUGOS DEL VALLE":                 {CUU:6.6,CLN:4.4,MXL:8.4,MTY:8.6,CUAU:8.6,CUAU7494:8.6,CUAU7464:5.5,CUAU7492:8.6,"7464":5.5,"5907":5.3,STB:9.9,CHL:8.8,GDL:8.8,MER:4.1,VHSA:8.5,"2026":4.0},
  "SANTA CLARA MERC PACHU S RL CV": null,
  "PROCTER AND GAMBLE MEXICO INC":  {CUU:9.6,CLN:6.9,MXL:9.8,MTY:10.0,CUAU:10.0,CUAU7494:10.0,CUAU7464:9.0,CUAU7492:10.0,"7464":9.0,"5907":10.7,STB:10.0,CHL:10.0,GDL:10.0,MER:5.3,VHSA:9.2,"2026":7.8},
  "MARCAS NESTLE":                   {CUU:8.3,CLN:5.0,MXL:10.0,MTY:10.0,CUAU:10.0,CUAU7494:10.0,CUAU7464:7.2,CUAU7492:10.0,"7464":7.2,"5907":8.7,STB:10.0,CHL:10.0,GDL:10.0,MER:4.6,VHSA:8.5,"2026":5.0},
  "COLGATE PALMOLIVE SA CV":         {CUU:11.1,CLN:8.6,MXL:10.4,MTY:11.5,CUAU:11.5,CUAU7494:11.5,CUAU7464:8.5,CUAU7492:11.5,"7464":8.5,"5907":8.0,STB:11.5,CHL:11.5,GDL:11.5,MER:6.3,VHSA:10.0,"2026":5.0},
  "COMERC PEPSICO MEXICO S RL CV":   {CUU:10.8,CLN:7.1,MXL:10.0,MTY:9.3,CUAU:9.3,CUAU7494:9.3,CUAU7464:6.0,CUAU7492:9.3,"7464":6.0,"5907":7.4,STB:10.5,CHL:10.8,GDL:10.8,MER:5.3,VHSA:8.8,"2026":6.0},
  "BONAFONT + ENVASASORA":           {CUU:6.8,CLN:4.5,MXL:10.1,MTY:8.0,CUAU:8.0,CUAU7494:8.0,CUAU7464:5.0,CUAU7492:8.0,"7464":5.0,"5907":5.5,STB:8.0,CHL:8.0,GDL:8.4,MER:5.7,VHSA:7.1,"2026":6.0},
  "UNILEVER DE MEXICO S RL CV":      {CUU:13.5,CLN:10.1,MXL:13.1,MTY:9.1,CUAU:9.1,CUAU7494:9.1,CUAU7464:6.5,CUAU7492:9.1,"7464":6.5,"5907":11.1,STB:13.5,CHL:13.5,GDL:13.5,MER:8.4,VHSA:11.0,"2026":9.5},
  "HERDEZ SA DE CV":                 {CUU:10.1,CLN:6.6,MXL:10.5,MTY:10.1,CUAU:10.1,CUAU7494:10.1,CUAU7464:7.3,CUAU7492:10.1,"7464":7.3,"5907":8.0,STB:10.9,CHL:11.0,GDL:11.0,MER:5.0,VHSA:9.1,"2026":6.5},
  "CERVEZA CANAL MO S D":            {CUU:7.4,CLN:5.0,MXL:9.2,MTY:8.0,CUAU:8.0,CUAU7494:8.0,CUAU7464:5.5,CUAU7492:8.0,"7464":5.5,"5907":5.5,STB:8.0,CHL:8.0,GDL:8.0,MER:4.8,VHSA:7.0,"2026":5.0},
  "FRABEL SA DE CV":                 {CUU:12.0,CLN:11.4,MXL:13.9,MTY:12.0,CUAU:12.0,CUAU7494:12.0,CUAU7464:6.0,CUAU7492:12.0,"7464":6.0,"5907":9.9,STB:12.0,CHL:12.0,GDL:12.0,MER:10.2,VHSA:11.2,"2026":10.5},
  "MONDELEZ MEXICO S DE RL DE CV":   null,
  "KELLOGG COMPANY MEXICO SRL CV":   {CUU:8.9,CLN:6.8,MXL:12.3,MTY:12.5,CUAU:12.5,CUAU7494:12.5,CUAU7464:7.0,CUAU7492:12.5,"7464":7.0,"5907":10.2,STB:12.5,CHL:12.5,GDL:12.5,MER:5.8,VHSA:10.3,"2026":7.0},
  "__total__":                        {CUU:9.4,CLN:6.8,MXL:10.6,MTY:10.0,CUAU:10.0,CUAU7494:10.0,CUAU7464:6.8,CUAU7492:10.0,"7464":6.8,"5907":8.2,STB:10.6,CHL:10.5,GDL:10.6,MER:5.9,VHSA:9.3,"2026":6.6}
};
// ── Objetivos Solo BAE ───────────────────────────────────────────────────────
// Clave "8806" = objetivo especifico para CD BAE SUR (MTY), usado en buildCDObjMap
var OBJ_SOLO_BAE = {
  "KIMBERLY CLARK DE MEX SA B CV":  {CLN:6.0,MTY:2.5,"8806":3.1,STB:2.2,CHL:0.9,GDL:3.6,MER:4.8,VHSA:3.6,"2026":4.5},
  "ENBOTELLAD NIAGARA D MX":        null,
  "JUGOS DEL VALLE":                 {CLN:6.0,MTY:2.5,"8806":3.0,STB:2.5,CHL:1.2,GDL:3.2,MER:4.7,VHSA:3.5,"2026":4.2},
  "SANTA CLARA MERC PACHU S RL CV": null,
  "PROCTER AND GAMBLE MEXICO INC":  {CLN:4.5,MTY:2.9,"8806":3.0,STB:3.4,CHL:1.6,GDL:3.9,MER:4.4,VHSA:2.8,"2026":4.1},
  "MARCAS NESTLE":                   {CLN:3.5,MTY:2.9,"8806":3.2,STB:2.6,CHL:1.1,GDL:3.0,MER:4.7,VHSA:4.1,"2026":4.5},
  "COLGATE PALMOLIVE SA CV":         {CLN:5.5,MTY:3.7,"8806":3.1,STB:4.0,CHL:2.2,GDL:3.9,MER:4.2,VHSA:3.1,"2026":4.3},
  "COMERC PEPSICO MEXICO S RL CV":   {CLN:5.0,MTY:3.5,"8806":4.0,STB:4.1,CHL:2.0,GDL:4.1,MER:5.3,VHSA:3.8,"2026":4.3},
  "BONAFONT + ENVASASORA":           {CLN:6.0,MTY:2.0,"8806":8.0,STB:2.4,CHL:null,GDL:4.9,MER:5.0,VHSA:1.8,"2026":3.8},
  "UNILEVER DE MEXICO S RL CV":      {CLN:6.5,MTY:4.2,"8806":6.2,STB:4.3,CHL:2.3,GDL:5.5,MER:5.7,VHSA:3.8,"2026":5.6},
  "HERDEZ SA DE CV":                 {CLN:4.0,MTY:3.6,"8806":5.5,STB:2.7,CHL:1.0,GDL:3.9,MER:6.0,VHSA:4.2,"2026":5.2},
  "CERVEZA CANAL MO S D":            {CLN:3.5,MTY:2.0,"8806":2.4,STB:2.0,CHL:1.0,GDL:2.7,MER:5.8,VHSA:3.9,"2026":4.7},
  "FRABEL SA DE CV":                 {CLN:5.5,MTY:2.5,"8806":5.0,STB:2.3,CHL:1.3,GDL:4.5,MER:3.5,VHSA:2.8,"2026":3.6},
  "MONDELEZ MEXICO S DE RL DE CV":   null,
  "KELLOGG COMPANY MEXICO SRL CV":   {CLN:6.0,MTY:3.0,"8806":3.3,STB:2.9,CHL:1.1,GDL:3.9,MER:4.6,VHSA:3.1,"2026":4.2},
  "__total__":                        {CLN:5.2,MTY:2.9,"8806":4.2,STB:3.0,CHL:1.4,GDL:3.9,MER:4.7,VHSA:3.6,"2026":4.4}
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
// ── Locacion → CDs mapping completo ────────────────────────────────────
var LOC_LABELS = {
  CUU:"Chihuahua", CLN:"Culiac\u00e1n", MXL:"Mexicali", MTY:"Monterrey",
  CUAU:"Cuautitl\u00e1n", STB:"Sta. B\u00e1rbara", SMO:"S. Mart\u00edn Obispo",
  CHL:"Chalco", GDL:"Guadalajara", MER:"M\u00e9rida", VHSA:"Villahermosa"
};
var LOC_ORDER = ["CUU","CLN","MXL","MTY","CUAU","STB","SMO","CHL","GDL","MER","VHSA"];
var LOC_CDS = {
  CUU:  [{n:"4640 AUTOSERVICIOS", k:"4640"},
          {n:"5780 SAM'S",         k:"5780"}],
  CLN:  [{n:"7455 BAE",           k:"7455"},
          {n:"7487 AUTOSERVICIOS", k:"7487"},
          {n:"4971 SAM'S",         k:"4971"}],
  MXL:  [{n:"4924 AUTOSERVICIOS", k:"4924"},
          {n:"6140 SAM'S",         k:"6140"}],
  MTY:  [{n:"7461 BAE",                k:"7461"},
          {n:"7490 AUTOSERVICIOS",      k:"7490"},
          {n:"8806 BAE SUR",            k:"8806"},
          {n:"4995 SAM'S",              k:"4995"},
          {n:"7498 PERE AUTOSERVICIOS", k:"7498"},
          {n:"7502 PERE SAM'S",         k:"7502"}],
  CUAU: [{n:"7494 NAVE 1",          k:"7494"},
          {n:"7464 NAVE 2",        k:"7464"},
          {n:"7492 NAVE 3",        k:"7492"}],
  STB:  [{n:"7457 BAE",           k:"7457"},
          {n:"7482 AUTOSERVICIOS", k:"7482"}],
  SMO:  [{n:"7466 PERE AUTOSERVICIOS", k:"7466"},
          {n:"4996 PERE SAM'S",         k:"4996"},
          {n:"6388 SAM'S",              k:"6388"}],
  CHL:  [{n:"7454 PERE BAE",           k:"7454"},
          {n:"7459 BAE",               k:"7459"},
          {n:"7471 AUTOSERVICIOS",     k:"7471"},
          {n:"7505 SAM'S",             k:"7505"},
          {n:"8801 PERE AUTOSERVICIOS",k:"8801"}],
  GDL:  [{n:"5907 MI BODEGA",           k:"5907"},
          {n:"7460 BAE",               k:"7460"},
          {n:"7493 AUTOSERVICIOS",     k:"7493"},
          {n:"6238 SAM'S",             k:"6238"},
          {n:"6239 PERE SAM'S",        k:"6239"},
          {n:"7495 PERE AUTOSERVICIOS",k:"7495"}],
  MER:  [{n:"4188 AUTOSERVICIOS", k:"4188"},
          {n:"7103 BAE",           k:"7103"},
          {n:"7506 SAM'S",         k:"7506"}],
  VHSA: [{n:"7453 BAE",               k:"7453"},
          {n:"7468 AUTOSERVICIOS",     k:"7468"},
          {n:"6550 SAM'S",             k:"6550"},
          {n:"4659 PERE AUTOSERVICIOS",k:"4659"},
          {n:"6151 PERE SAM'S",        k:"6151"}]
};
// CD metadata ──────────────────────────────────────────────────
var CD_TO_CITY = {
  "4640":"CUU","5780":"CUU",
  "7455":"CLN","7487":"CLN","4971":"CLN",
  "4924":"MXL","6140":"MXL",
  "7461":"MTY","7490":"MTY","8806":"MTY","4995":"MTY","7498":"MTY","7502":"MTY",
  "7494":"CUAU","7464":"CUAU","7492":"CUAU",
  "7457":"STB","7482":"STB",
  "7466":"SMO","4996":"SMO","6388":"SMO",
  "7459":"CHL","7471":"CHL","7505":"CHL","8801":"CHL","7454":"CHL",
  "5907":"GDL","7460":"GDL","7493":"GDL","6238":"GDL","6239":"GDL","7495":"GDL",
  "4188":"MER","7103":"MER","7506":"MER",
  "7453":"VHSA","7468":"VHSA","6550":"VHSA","4659":"VHSA","6151":"VHSA"
};
// auto = AUTOSERVICIOS+BAE (secos), sams = SAM'S (secos), perec_auto/perec_sams = perecederos
var CD_CHANNEL = {
  "4640":"auto","7487":"auto","4924":"auto","7490":"auto",
  "7494":"auto","7464":"auto","7492":"auto",
  "7482":"auto","7471":"auto","7493":"auto","4188":"auto","7468":"auto",
  "5907":"auto",
  "7455":"bae","7461":"bae","8806":"bae","7457":"bae","7459":"bae",
  "7460":"bae","7103":"bae","7453":"bae",
  "5780":"sams","4971":"sams","6140":"sams","4995":"sams",
  "6388":"sams","7505":"sams","6238":"sams","7506":"sams","6550":"sams",
  "8801":"perec_auto","7495":"perec_auto","4659":"perec_auto",
  "7498":"perec_auto","7466":"perec_auto",
  "6239":"perec_sams","6151":"perec_sams",
  "7502":"perec_sams","4996":"perec_sams",
  "7454":"perec_bae"
};
// PEREC CDs -> ciudad key (los datos PEREC usan claves de ciudad)
var CD_CITY_PEREC = {
  "8801":"CHL","7495":"GDL","4659":"VHSA",
  "7498":"MTY","7466":"SMO",
  "6239":"GDL","6151":"VHSA",
  "7502":"MTY","4996":"SMO",
  "7454":"CHL"
};
function buildCDObjMap(cityObjMap) {
  if (!cityObjMap) return null;
  var out = {};
  Object.keys(cityObjMap).forEach(function(v) {
    var co = cityObjMap[v];
    if (!co) { out[v] = null; return; }
    var cdMap = {"2026": co["2026"]};
    Object.keys(CD_TO_CITY).forEach(function(cd) {
      // CD-specific override (e.g. "8806", "7464", "5907") takes priority over city key
      if (co[cd] !== undefined) { cdMap[cd] = co[cd]; return; }
      var city = CD_TO_CITY[cd];
      if (co[city] !== undefined) cdMap[cd] = co[city];
    });
    out[v] = cdMap;
  });
  return out;
}
var OBJ_AUTO_CD      = buildCDObjMap(OBJ_AUTO);
var OBJ_SAMS_CD      = buildCDObjMap(OBJ_SAMS);
var OBJ_SOLO_AUTO_CD = buildCDObjMap(OBJ_SOLO_AUTO);
var OBJ_SOLO_BAE_CD  = buildCDObjMap(OBJ_SOLO_BAE);
// Helpers formato-conscientes para seleccionar el mapa correcto de objetivos
function fmtObjCity(){var f=gFormato;if(f==='sams')return OBJ_SAMS;if(f==='bae')return OBJ_SOLO_BAE;if(f==='auto')return OBJ_SOLO_AUTO;return OBJ_AUTO;}
function fmtObjCD(){var f=gFormato;if(f==='sams')return OBJ_SAMS_CD;if(f==='bae')return OBJ_SOLO_BAE_CD;if(f==='auto')return OBJ_SOLO_AUTO_CD;return OBJ_AUTO_CD;}
// Agrupa CDs por ciudad para usarlos como "regiones" en la tabla
function buildCDRegions(cdKeys) {
  var groups = {}, order = [];
  cdKeys.forEach(function(cd) {
    var city = CD_TO_CITY[cd];
    var lbl  = (city && LOC_LABELS[city]) ? LOC_LABELS[city] : cd;
    if (!groups[lbl]) { groups[lbl] = []; order.push(lbl); }
    groups[lbl].push(cd);
  });
  return order.map(function(lbl) { return {n:lbl, l:groups[lbl]}; });
}
// Filtra las regiones PEREC existentes por ciudades seleccionadas
function buildCityRegions(cityKeys, baseRegions) {
  if (!cityKeys.length) return [];
  return baseRegions.map(function(r) {
    return {n:r.n, l:r.l.filter(function(l){ return cityKeys.indexOf(l)>=0; })};
  }).filter(function(r){ return r.l.length>0; });
}

// Devuelve CDs de un canal para una ciudad: canal 'auto_bae' incluye auto+bae
function cdsForCanalCity(canal, cityKey) {
  var cds = [];
  Object.keys(CD_TO_CITY).forEach(function(cd) {
    if (CD_TO_CITY[cd] !== cityKey) return;
    var ch = CD_CHANNEL[cd];
    if (canal === 'auto_bae' ? (ch === 'auto' || ch === 'bae') : ch === canal) cds.push(cd);
  });
  return cds;
}
// Todos los CDs del canal agrupados por ciudad: {city: [cds]}
function cdsByCanalNacional(canal) {
  var out = {};
  Object.keys(CD_TO_CITY).forEach(function(cd) {
    var ch = CD_CHANNEL[cd];
    var ok = canal === 'auto_bae' ? (ch === 'auto' || ch === 'bae') : ch === canal;
    if (!ok) return;
    var city = CD_TO_CITY[cd];
    if (!out[city]) out[city] = [];
    out[city].push(cd);
  });
  return out;
}

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

// ── FILTRO FORMATO ─────────────────────────────────────────────────────────────────
var gFormato = 'all'; // clave activa: all | auto | bae | sams | auto_bae | auto_sams | bae_sams


function onFmtChange() {
  gFormato = document.getElementById('fmtSel').value;
  gSelCD = [];
  syncCdPanel();
  renderAll();
}

function onLocChange() {
  var loc = document.getElementById('locSel').value;
  gSelCD = [];
  var panel = document.getElementById('cdPanel');
  panel.innerHTML = '';
  // Boton limpiar
  var clear = document.createElement('div');
  clear.className = 'ms-item'; clear.style.cssText = 'color:#64748b;font-style:italic;border-bottom:1px solid #e2e8f0;';
  clear.innerHTML = '&#10005; Todos';
  clear.onclick = function(e){ e.stopPropagation(); gSelCD=[]; syncCdPanel(); renderAll(); };
  panel.appendChild(clear);

  function addCdItem(d) {
    // CDs perecederos se excluyen del filtro — tienen sus propias tablas abajo
    if ((CD_CHANNEL[d.k] || '').indexOf('perec') >= 0) return;
    var item = document.createElement('div'); item.className='ms-item'; item.dataset.val=d.k; item.dataset.lbl=d.n;
    var cb = document.createElement('input'); cb.type='checkbox'; cb.checked=false;
    var sp = document.createElement('span'); sp.textContent=d.n;
    item.appendChild(cb); item.appendChild(sp);
    item.addEventListener('click', function(e){ e.stopPropagation(); toggleCD(d.k, d.n); });
    panel.appendChild(item);
  }

  if (loc === '__all__') {
    LOC_ORDER.forEach(function(city) {
      if (!LOC_CDS[city] || !LOC_CDS[city].length) return;
      var sep = document.createElement('div');
      sep.style.cssText = 'padding:4px 12px 2px;font-size:0.7rem;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.06em;background:#f8fafc;border-top:1px solid #e2e8f0;pointer-events:none;';
      sep.textContent = LOC_LABELS[city] || city;
      panel.appendChild(sep);
      LOC_CDS[city].forEach(addCdItem);
    });
  } else if (LOC_CDS[loc]) {
    LOC_CDS[loc].forEach(addCdItem);
  }
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

function getFiltLocs() {
  if (!gSelCD.length) return null;
  var keys = [];
  gSelCD.forEach(function(x){ if (keys.indexOf(x.k) < 0) keys.push(x.k); });
  return keys;
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
var REGIONS_ALL   = [{n:"NORTE",l:["CUU","CLN","MXL","MTY"]},{n:"CENTRO",l:["CUAU","STB","SMO"]},{n:"SUR",l:["CHL","GDL","MER","VHSA"]}];
var CUAU_SUBS     = ["CUAU7494","CUAU7464","CUAU7492"];
var REGIONS_SAMS  = [{n:"NORTE",l:["CUU","CLN","MTY"]},{n:"CENTRO",l:["SMO"]},{n:"SUR",l:["CHL","GDL","MER","VHSA"]}];
var REGIONS_PAUTO = [{n:"NORTE",l:["MTY"]},{n:"CENTRO",l:["SMO"]},{n:"SUR",l:["CHL","GDL","VHSA"]}];
var REGIONS_PSAMS = [{n:"NORTE",l:["MTY"]},{n:"CENTRO",l:["SMO"]},{n:"SUR",l:["GDL","VHSA"]}];
var SW_DATA = null;
var gSelSW  = [];

var ALL_PER = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","ytd"];
var PER_LBL = {Enero:"Enero",Febrero:"Febrero",Marzo:"Marzo",Abril:"Abril",Mayo:"Mayo",Junio:"Junio",Julio:"Julio",ytd:"YTD"};
var PER_BG  = {Enero:"#1d4ed8",Febrero:"#1d4ed8",Marzo:"#1d4ed8",Abril:"#1d4ed8",Mayo:"#1d4ed8",Junio:"#1d4ed8",Julio:"#d97706",ytd:"#065f46"};
var REG_BG  = {NORTE:"#1d4ed8",CENTRO:"#6d28d9",SUR:"#065f46"};

var gChartInstances = [];  // array: un chart por CD seleccionado (o 1 si no hay filtro)
var gCanal = 'auto_bae';
var gCanalTitle = 'LOS Total';

var CANAL_CFG = {
  auto_bae: { label: 'Auto + BAE' },
  auto:     { label: 'Autoservicios' },
  bae:      { label: 'BAE' },
  sams:     { label: "SAM'S" }
};
var CANAL_TITLE = {
  auto_bae: 'Autoservicios m\u00e1s BAE',
  auto:     'Autoservicios',
  bae:      'BAE',
  sams:     "SAM'S"
};

function updateChartTitle(override) {
  gCanalTitle = override || CANAL_TITLE[gCanal] || 'LOS por Proveedor';
  var titleEl = document.getElementById('chartTitle');
  if (titleEl) titleEl.textContent = gCanalTitle;
}
function fmt(v) {
  if (v === null || v === undefined) return "\u2014";
  return parseFloat(v).toFixed(1);
}

function cellCls(val, obj) {
  if (val === null || val === undefined) return "cell-null";
  if (obj === null || obj === undefined) return "";
  var diff = Math.round((val - obj) * 10) / 10;  // precision fix: 12.0001 vs 12.0 => 0.0
  return diff <= 0 ? "cell-ok" : "cell-bad";
}

// Celda Total: solo color, sin marcadores ni delta
function tdDelta(val, obj, extraCls) {
  var el = document.createElement("td");
  if (val === null || val === undefined) {
    el.textContent = "\u2014";
    el.className = "cell-null" + (extraCls ? " " + extraCls : "");
    return el;
  }
  el.textContent = fmt(val);
  if (obj === null || obj === undefined) {
    if (extraCls) el.className = extraCls;
    return el;
  }
  var diff = Math.round((val - obj) * 10) / 10;
  var colorCls = diff <= 0 ? "cell-ok" : "cell-bad";
  el.className = colorCls + (extraCls ? " " + extraCls : "");
  return el;
}

// Devuelve true si la locacion en el indice 'li' es la primera de una región (salvo la región 0)
function isRegStart(li, allLocs, regions) {
  if (li === 0) return false; // la primera columna no necesita separador izquierdo
  var cum = 0;
  for (var ri = 1; ri < regions.length; ri++) {
    for (var rj = 0; rj < ri; rj++) cum += regions[rj].l.length;
    if (li === cum) return true;
    cum = 0;
  }
  return false;
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

// ── HELPERS DE CHARTS ───────────────────────────────────────────────────────────────────
// Destruye todas las instancias Chart.js previas
function destroyAllCharts() {
  gChartInstances.forEach(function(ch) { try { ch.destroy(); } catch(e) {} });
  gChartInstances = [];
}
// Construye el HTML del contenedor de charts:
// defs = [{id, title}] — uno por CD o uno solo
function renderChartsContainer(defs) {
  var cont = document.getElementById('chartsContainer');
  cont.innerHTML = '';
  var isMulti = defs.length > 1;
  if (isMulti) {
    cont.className = 'grid gap-4';
    cont.style.gridTemplateColumns = 'repeat(' + Math.min(defs.length, 3) + ', 1fr)';
  } else {
    cont.className = '';
    cont.style.gridTemplateColumns = '';
  }
  defs.forEach(function(def) {
    var outer = document.createElement('div');
    if (def.title) {
      var lbl = document.createElement('div');
      lbl.className = 'text-xs font-bold text-slate-600 mb-1 text-center';
      lbl.textContent = def.title;
      outer.appendChild(lbl);
    }
    var wrap = document.createElement('div');
    wrap.className = 'chart-wrap';
    if (isMulti) wrap.style.height = '320px';
    var canvas = document.createElement('canvas');
    canvas.id = def.id;
    wrap.appendChild(canvas);
    outer.appendChild(wrap);
    cont.appendChild(outer);
  });
}
// Devuelve el label de un CD: busca en gSelCD, si no usa el id
function getCDLabel(cd) {
  for (var i = 0; i < gSelCD.length; i++) {
    if (String(gSelCD[i].k) === String(cd)) return gSelCD[i].n;
  }
  return cd;
}

function renderAll() {
  var loc    = getLocKey();
  var period = document.getElementById("mesSel").value;
  var locEl  = document.getElementById("locSel");
  var locLbl = locEl.value === "__all__" ? "Todas" : locEl.options[locEl.selectedIndex].text;
  var cdLbl  = gSelCD.length ? (" \u00b7 " + (gSelCD.length===1 ? gSelCD[0].n : gSelCD.length+" CDs")) : "";
  var perLbl = document.getElementById("mesSel").options[document.getElementById("mesSel").selectedIndex].text;
  var swLbl  = gSelSW.length ? (" \u00b7 " + (gSelSW.length===1 ? gSelSW[0] : gSelSW.length+" semanas")) : "";
  document.getElementById("badge").textContent = "Mostrando: " + locLbl + cdLbl + " \u00b7 " + perLbl + swLbl;

  var useSW  = gSelSW.length > 0 && SW_DATA;
  var swPer  = gSelSW.length === 1 ? gSelSW[0] : '__sel__';
  var activePeriod = useSW ? swPer : period;
  // Multi-SW: tabla muestra cada semana como columna separada; grAfica sigue con el promedio
  var swOver = useSW ? (gSelSW.length > 1 ? gSelSW.slice().sort(function(a, b) {
    var list = (SW_DATA && SW_DATA.sw_list) ? SW_DATA.sw_list : [];
    var ia = list.indexOf(parseInt(a.slice(2)));
    var ib = list.indexOf(parseInt(b.slice(2)));
    if (ia < 0) ia = parseInt(a.slice(2));
    if (ib < 0) ib = parseInt(b.slice(2));
    return ia - ib;
  }) : [activePeriod]) : null;

  // ── BUNDLE UNICO: misma fuente para grafica Y tabla ─────────────────────
  var b = resolveViewBundle(loc, activePeriod, useSW, swPer);

  // Actualiza titulo en grafica y tabla (mismo string)
  gCanalTitle = b.title;
  var titleEl = document.getElementById('chartTitle');
  if (titleEl) titleEl.textContent = gCanalTitle;

  var wrap = document.getElementById('tables');
  wrap.innerHTML = '';

  if (b.empty) {
    destroyAllCharts();
    renderChartsContainer([{ id:'chartMain', title:null }]);
    var msgEl = document.createElement('p');
    msgEl.className = 'p-4 text-sm text-slate-400 italic';
    msgEl.textContent = b.emptyMsg;
    wrap.appendChild(msgEl);
  } else {
    // ── SIEMPRE UNA SOLA GRAFICA ──
    renderChartsContainer([{ id:'chartMain', title:null }]);
    destroyAllCharts();
    var chartLoc = (b.filtLocs && b.filtLocs.length > 0) ? '__all__' : b.loc;
    var inst = paintChart('chartMain', b.chartData, b.objData, DISPLAY_ORDER,
      chartLoc, b.period, null, b.filtLocs);
    if (inst) gChartInstances.push(inst);

    // Tabla — usa b.tblData (misma fuente, diferente estructura)
    wrap.appendChild(buildTable(
      b.title, b.tblData, b.objData, DISPLAY_ORDER,
      b.regions, b.loc, b.period, true, swOver, b.filtLocs
    ));
  }

  // Perecederos — siempre visibles, independiente del filtro formato
  var perLoc = document.getElementById('locSel').value;
  if (perLoc === '__all__') perLoc = '__all__';
  wrap.appendChild(buildTable(
    'Perecederos \u2014 AUTO',
    DATA_PEREC_AUTO, OBJ_PEREC_AUTO, PEREC_ORDER,
    REGIONS_PAUTO, perLoc, period, true, null, null
  ));
  wrap.appendChild(buildTable(
    "Perecederos \u2014 SAM'S",
    DATA_PEREC_SAMS, null, PEREC_ORDER,
    REGIONS_PSAMS, perLoc, period, false, null, null
  ));
}

// ── RESOLUTOR UNICO DE DATOS ────────────────────────────────────────────────
// Una sola funcion determina que datos ver. Grafica y tabla la llaman juntas.
// Si los datos cambian, se cambia aqui y se propaga a ambos automaticamente.
function resolveViewBundle(loc, activePeriod, useSW, swPer) {
  var filtLocs = getFiltLocs();
  var cdMode   = filtLocs && filtLocs.length > 0;

  // Determina canal segun seleccion de CDs
  if (cdMode) {
    var channels = {};
    filtLocs.forEach(function(cd){
      var ch = CD_CHANNEL[cd];
      if (ch) channels[ch] = true;
    });
    var hasAuto = channels['auto'], hasBae = channels['bae'], hasSams = channels['sams'];
    if      (hasSams && !hasAuto && !hasBae) gCanal = 'sams';
    else if (hasAuto && !hasBae  && !hasSams) gCanal = 'auto';
    else if (hasBae  && !hasAuto && !hasSams) gCanal = 'bae';
    else gCanal = 'auto_bae';
  }

  // ── Datos SW base (elegir dataset correcto por formato)
  var _swChart, _swTbl;
  if (useSW) {
    if (gFormato === 'sams') {
      _swChart = gSelSW.length > 1 ? mergeSWChartData(SW_DATA.sams, gSelSW) : SW_DATA.sams;
      _swTbl   = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_sams, gSelSW) : SW_DATA.tbl_sams;
    } else if (gFormato === 'bae') {
      _swChart = gSelSW.length > 1 ? mergeSWChartData(SW_DATA.bae, gSelSW) : SW_DATA.bae;
      _swTbl   = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_bae, gSelSW) : SW_DATA.tbl_bae;
    } else if (gFormato === 'auto') {
      _swChart = gSelSW.length > 1 ? mergeSWChartData(SW_DATA.auto, gSelSW) : SW_DATA.auto;
      _swTbl   = gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_auto, gSelSW) : SW_DATA.tbl_auto;
    } else {
      // all, auto_bae, auto_sams, bae_sams -> auto_bae si existe, sino auto
      var _srcC = SW_DATA.auto_bae || SW_DATA.auto;
      var _srcT = SW_DATA.tbl_auto_bae || SW_DATA.tbl_auto;
      _swChart = gSelSW.length > 1 ? mergeSWChartData(_srcC, gSelSW) : _srcC;
      _swTbl   = gSelSW.length > 1 ? mergeSWTblData(_srcT, gSelSW) : _srcT;
    }
  }
  var swAutoChart = _swChart || null;
  var swAutoTbl   = _swTbl   || null;
  // Alias sams para cdMode
  var swSamsChart = useSW ? (gSelSW.length > 1 ? mergeSWChartData(SW_DATA.sams, gSelSW) : SW_DATA.sams) : null;
  var swSamsTbl   = useSW ? (gSelSW.length > 1 ? mergeSWTblData(SW_DATA.tbl_sams, gSelSW) : SW_DATA.tbl_sams) : null;

  // ── MODO CD: datos por cedis especifico ──
  if (cdMode) {
    var cdsCanal = filtLocs.filter(function(cd){
      var ch = CD_CHANNEL[cd];
      return gCanal === 'auto_bae' ? (ch==='auto'||ch==='bae') : ch===gCanal;
    });
    if (cdsCanal.length === 0) {
      return { empty:true, emptyMsg:'Sin CDs del canal seleccionado.' };
    }
    var title = CANAL_TITLE[gCanal] || gCanal;

    // Con SW activo: mapear CDs -> city codes y usar datos SW
    if (useSW) {
      var isSamsCd = gCanal === 'sams';
      var swChCd = isSamsCd ? swSamsChart : swAutoChart;
      var swTbCd = isSamsCd ? swSamsTbl   : swAutoTbl;
      var citySet = {};
      cdsCanal.forEach(function(cd){ var c=CD_TO_CITY[cd]; if(c) citySet[c]=true; });
      var cities = Object.keys(citySet);
      var swLoc  = cities.length === 1 ? cities[0] : '__all__';
      var swFilt = cities.length > 1   ? cities    : null;
      var cityRegions = [];
      cities.forEach(function(c){ var lbl=LOC_LABELS[c]||c; cityRegions.push({n:lbl,l:[c]}); });
      return {
        empty:false, title:title,
        chartData: swChCd, tblData: swTbCd,
        objData:   fmtObjCity(),
        regions:   cityRegions,
        loc:       swLoc, filtLocs: swFilt,
        period:    activePeriod
      };
    }

    // Sin SW: usar datos mensuales cd_chart.json
    if (!DATA_CHART_CD) {
      fetch('/bigquery_results/cd_chart.json')
        .then(function(r){ return r.json(); })
        .then(function(data){ DATA_CHART_CD = data; renderAll(); });
      return { empty:true, emptyMsg:'Cargando datos de CD...' };
    }
    return {
      empty:false, title:title,
      chartData: DATA_CHART_CD,
      tblData:   DATA_CD,
      objData:   fmtObjCD(),
      regions:   buildCDRegions(cdsCanal),
      loc:       cdsCanal.length===1 ? cdsCanal[0] : '__all__',
      filtLocs:  cdsCanal.length>1  ? cdsCanal    : null,
      period:    activePeriod
    };
  }

  // ── MODO NORMAL: gFormato determina el dataset directamente ─────────────────
  var FMT_DATA = {
    'all':       {tbl:DATA_ALL,       chart:DATA_ALL_CHART,       reg:REGIONS_ALL,  title:'LOS Total',     obj:OBJ_AUTO},
    'auto':      {tbl:DATA_AUTO,      chart:DATA_AUTO_CHART,      reg:REGIONS_AUTO, title:'Autoservicios', obj:OBJ_SOLO_AUTO},
    'bae':       {tbl:DATA_BAE,       chart:DATA_BAE_CHART,       reg:REGIONS_AUTO, title:'BAE',           obj:OBJ_SOLO_BAE},
    'sams':      {tbl:DATA_SAMS,      chart:DATA_SAMS_CHART,      reg:REGIONS_SAMS, title:"SAM'S",         obj:OBJ_SAMS},
    'auto_bae':  {tbl:DATA_AUTO_BAE,  chart:DATA_AUTO_BAE_CHART,  reg:REGIONS_AUTO, title:'Auto + BAE',    obj:OBJ_AUTO},
    'auto_sams': {tbl:DATA_AUTO_SAMS, chart:DATA_AUTO_SAMS_CHART, reg:REGIONS_ALL,  title:"Auto + SAM'S",  obj:OBJ_AUTO},
    'bae_sams':  {tbl:DATA_BAE_SAMS,  chart:DATA_BAE_SAMS_CHART,  reg:REGIONS_ALL,  title:"BAE + SAM'S",   obj:OBJ_AUTO},
  };
  var fd = FMT_DATA[gFormato] || FMT_DATA['all'];
  if (useSW) {
    var swCh = (gFormato === 'sams') ? swSamsChart : swAutoChart;
    var swTb = (gFormato === 'sams') ? swSamsTbl   : swAutoTbl;
    return {
      empty:false, title:fd.title,
      chartData:swCh, tblData:swTb,
      objData:fd.obj, regions:fd.reg,
      loc:loc, filtLocs:null, period:activePeriod
    };
  }
  return {
    empty:false, title:fd.title,
    chartData:fd.chart, tblData:fd.tbl,
    objData:fd.obj, regions:fd.reg,
    loc:loc, filtLocs:null, period:activePeriod
  };
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
        var fpd=fld[period]; if(!fpd || fpd.t==null) continue;
        lS+=fpd.l||0; rS+=fpd.r||0; sS+=fpd.s||0; tS+=fpd.t; cn++;
      }
      if (!cn) continue;
      vendors.push(v);
      lA.push(lS/cn); rA.push(rS/cn); sA.push(sS/cn); tA.push(tS/cn);
      var ov=objData?objData[v]:null, oSm=0, oc=0;
      if (ov) { for(var fi=0;fi<filtLocs.length;fi++){if(ov[filtLocs[fi]]!=null){oSm+=ov[filtLocs[fi]];oc++;}} }
      oA.push(oc?oSm/oc:(ov&&ov["2026"]!=null?ov["2026"]:null));
    } else {
      var ld = vd[locKey]; if (!ld) continue;
      var pd = ld[period]; if (!pd || pd.t==null) continue;
      vendors.push(v);
      lA.push(pd.l || 0); rA.push(pd.r || 0); sA.push(pd.s || 0); tA.push(pd.t);
      var ov = objData ? objData[v] : null;
      oA.push(ov ? (ov[locKey] != null ? ov[locKey] : (ov["2026"] != null ? ov["2026"] : null)) : null);
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
         datalabels:{display:true, color:"#374151", font:{size:13,weight:"bold"},
           formatter:function(v){return v>0.05?parseFloat(v).toFixed(1):"";}, anchor:"center", align:"center"}},
        {type:"bar",  label:"Recibo",   data:rS, backgroundColor:"#f97316", stack:"stk", order:2,
         yAxisID:"y",
         datalabels:{display:true, color:"#fff", font:{size:13,weight:"bold"},
           formatter:function(v){return v>0.05?parseFloat(v).toFixed(1):"";}, anchor:"center", align:"center"}},
        {type:"bar",  label:"Salida CD",data:sS, backgroundColor:"#fbbf24", stack:"stk", order:2,
         yAxisID:"y",
         datalabels:{display:true, color:"#78350f", font:{size:13,weight:"bold"},
           formatter:function(v){return v>0.05?parseFloat(v).toFixed(1):"";}, anchor:"center", align:"center"}},
        {type:"line", label:"Total",    data:tS, borderColor:"#111827", backgroundColor:"#111827",
         pointRadius:0, borderWidth:1, tension:0, fill:false, order:1,
         yAxisID:"y2",
         datalabels:{display:true, color:"#fff",
           backgroundColor:"#111827", borderRadius:3,
           padding:{top:2,bottom:2,left:4,right:4},
           font:{size:13,weight:"bold"},
           formatter:function(v){return v!=null?parseFloat(v).toFixed(1):"";},
           anchor:"end", align:"top", offset:4}},
        {type:"line", label:"Objetivo", data:oS, borderColor:"#2563eb", backgroundColor:"transparent",
         pointRadius:0, borderWidth:1, tension:0, fill:false, order:1,
         yAxisID:"y2",
         datalabels:{display:true, color:"#2563eb", font:{size:13,weight:"bold"},
           formatter:function(v){return v!=null?parseFloat(v).toFixed(1):"";},
           anchor:"end",
           align:function(ctx){
             var tot=ctx.chart.data.datasets[3].data[ctx.dataIndex];
             var obj=ctx.dataset.data[ctx.dataIndex];
             return (tot!=null&&obj!=null&&Math.abs(tot-obj)<3)?"right":"top";
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
        x:  {stacked:true, ticks:{font:{size:9}}, grid:{display:false}},
        y:  {stacked:true, beginAtZero:true, min:0, max:yMax, grid:{display:false},
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

function buildTable(title, data, objData, order, regions, loc, period, hasObj, overridePeriods, filtLocs) {
  var rawAllLocs = flatLocs(regions);
  var PERIODS = overridePeriods || ALL_PER;

  // Resuelve el valor de una celda (vendor, period, loc)
  // Si filtLocs tiene varios CDs: promedio simple entre ellos.
  // Si filtLocs tiene 1 CD o es null: valor directo.
  function getVal(data, v, per, singleLoc, fLocs) {
    var vd = data[v]; if (!vd) return null;
    var pd = vd[per]; if (!pd) return null;
    if (fLocs && fLocs.length > 1) {
      var sum = 0, cnt = 0;
      fLocs.forEach(function(l) { if (pd[l] != null) { sum += pd[l]; cnt++; } });
      return cnt ? Math.round(sum / cnt * 10) / 10 : null;
    }
    var key = (fLocs && fLocs.length === 1) ? fLocs[0] : singleLoc;
    // Fallback: si key no existe (ej '__all__'), usar '2026' (nacional)
    if (pd[key] !== undefined) return pd[key];
    if (key === '__all__' && pd['2026'] !== undefined) return pd['2026'];
    return null;
  }
  function getObjVal(objData, v, fLocs, singleLoc) {
    var ov = objData && objData[v]; if (!ov) return null;
    if (fLocs && fLocs.length > 1) {
      var sum = 0, cnt = 0;
      fLocs.forEach(function(l) { if (ov[l] != null) { sum += ov[l]; cnt++; } });
      return cnt ? Math.round(sum / cnt * 10) / 10 : null;
    }
    var key = (fLocs && fLocs.length === 1) ? fLocs[0] : singleLoc;
    if (ov[key] !== undefined) return ov[key];
    if (key === '__all__' && ov['2026'] !== undefined) return ov['2026'];
    return null;
  }

  // Determina si hay datos para mostrar
  var hasData = false;
  for (var vi = 0; vi < order.length; vi++) {
    var v = order[vi]; var vd = data[v]; if (!vd) continue;
    for (var pi = 0; pi < PERIODS.length; pi++) {
      if (getVal(data, v, PERIODS[pi], loc, filtLocs) != null) { hasData = true; break; }
    }
    if (hasData) break;
  }

  var wrap = document.createElement("div");
  wrap.className = "bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden";
  var hdr = document.createElement("div");
  hdr.className = "bg-slate-800 text-white px-4 py-2 font-bold text-sm";
  hdr.appendChild(Object.assign(document.createElement("span"), {textContent: title}));
  wrap.appendChild(hdr);

  if (!hasData) {
    wrap.appendChild(Object.assign(document.createElement("p"), {
      className: "p-4 text-sm text-slate-400 italic",
      textContent: "Sin datos para la selecci\u00f3n actual."
    }));
    return wrap;
  }

  var scrl = document.createElement("div"); scrl.className = "overflow-x-auto";
  wrap.appendChild(scrl);
  var tbl = document.createElement("table"); tbl.className = "tbl";
  scrl.appendChild(tbl);
  var thead = document.createElement("thead");
  var tbody = document.createElement("tbody");
  tbl.appendChild(thead); tbl.appendChild(tbody);

  var isAll   = (loc === "__all__") && (!filtLocs || filtLocs.length === 0);
  var allLocs  = flatLocs(regions);
  var activePeriodIdx = PERIODS.indexOf(period);

  if (isAll && !overridePeriods) {
    // ── VISTA COMPLETA (modo mes): objetivos + meses expandidos/colapsados ──
    // Detectar el último mes con datos (= mes en curso)
    var ALL_MONTHS_FULL = ALL_PER.filter(function(p){ return p !== 'ytd'; });
    var lastMIdx = -1;
    for (var mi2 = ALL_MONTHS_FULL.length - 1; mi2 >= 0; mi2--) {
      var hasMD = false;
      for (var vi2 = 0; vi2 < order.length && !hasMD; vi2++) {
        var vd2 = data[order[vi2]];
        if (!vd2 || !vd2[ALL_MONTHS_FULL[mi2]]) continue;
        var pd2 = vd2[ALL_MONTHS_FULL[mi2]];
        for (var kk in pd2) { if (pd2[kk] != null) { hasMD = true; break; } }
      }
      if (hasMD) { lastMIdx = mi2; break; }
    }
    // Periodos siempre expandidos (bloque completo por CEDIS)
    var alwaysExp = [];
    if (lastMIdx >= 0) alwaysExp.push(ALL_MONTHS_FULL[lastMIdx]); // solo mes en curso
    alwaysExp.push('ytd');
    // Si el usuario seleccionó un mes colapsado, ese también se expande
    var selExp = (period && alwaysExp.indexOf(period) < 0) ? period : null;

    // Lista de meses a mostrar: todos los que tienen datos + ytd
    var visMonths = [];
    for (var mi2 = 0; mi2 < ALL_MONTHS_FULL.length; mi2++) {
      var pm = ALL_MONTHS_FULL[mi2];
      var hasMD2 = false;
      for (var vi2 = 0; vi2 < order.length && !hasMD2; vi2++) {
        var vd2 = data[order[vi2]];
        if (!vd2 || !vd2[pm]) continue;
        var pd2 = vd2[pm];
        for (var kk in pd2) { if (pd2[kk] != null) { hasMD2 = true; break; } }
      }
      if (hasMD2) visMonths.push(pm);
    }
    visMonths.push('ytd');

    var secCols = allLocs.length + 1;

    // ── Fila 1: secciones ──
    var tr1 = document.createElement("tr");
    tr1.appendChild(th("Proveedores TOP", 3, 1, "v-name text-left", "#1e3a8a"));
    tr1.appendChild(th("Objetivos LOS 2026 por Cedis", 1, secCols, "", "#1e3a8a"));
    for (var pi = 0; pi < visMonths.length; pi++) {
      var pp = visMonths[pi];
      var bg = PER_BG[pp] || "#1d4ed8";
      var lbl = PER_LBL[pp] || pp;
      var isAct = (pp === period);
      var isExp = (alwaysExp.indexOf(pp) >= 0 || pp === selExp);
      if (isExp) {
        var thEl = th(lbl, 1, secCols, "", bg);
        if (isAct) { thEl.style.outline="3px solid #f59e0b"; thEl.style.outlineOffset="-3px"; }
        tr1.appendChild(thEl);
      } else {
        // Colapsado: una sola celda, rowspan=3 para cubrir las filas de región y CD
        var thEl = th(lbl, 3, 1, "text-xs", bg);
        thEl.style.writingMode="vertical-rl";
        thEl.style.textOrientation="mixed";
        thEl.style.padding="4px 2px";
        thEl.title="Selecciona este mes en el filtro para ver el detalle por CD";
        if (isAct) { thEl.style.outline="3px solid #f59e0b"; thEl.style.outlineOffset="-3px"; }
        tr1.appendChild(thEl);
      }
    }
    thead.appendChild(tr1);

    // ── Fila 2: regiones (solo bloques expandidos) ──
    var tr2 = document.createElement("tr");
    // Bloque objetivos
    for (var ri = 0; ri < regions.length; ri++) {
      var rbg = REG_BG[regions[ri].n] || "#334155";
      tr2.appendChild(th(regions[ri].n, 1, regions[ri].l.length, "reg-hdr", rbg));
    }
    tr2.appendChild(th("2026", 1, 1, "", "#b45309"));
    // Bloques expandidos
    for (var pi = 0; pi < visMonths.length; pi++) {
      var pp = visMonths[pi];
      if (alwaysExp.indexOf(pp) < 0 && pp !== selExp) continue; // colapsado → skip
      for (var ri = 0; ri < regions.length; ri++) {
        var rbg = REG_BG[regions[ri].n] || "#334155";
        tr2.appendChild(th(regions[ri].n, 1, regions[ri].l.length, "reg-hdr", rbg));
      }
      tr2.appendChild(th("2026", 1, 1, "", "#b45309"));
    }
    thead.appendChild(tr2);

    // ── Fila 3: códigos de CD (solo bloques expandidos) ──
    var tr3 = document.createElement("tr");
    for (var li = 0; li < allLocs.length; li++) {
      tr3.appendChild(th(allLocs[li], 1, 1, "", "#3b82f6"));
    }
    tr3.appendChild(th("Prom", 1, 1, "", "#d97706"));
    for (var pi = 0; pi < visMonths.length; pi++) {
      var pp = visMonths[pi];
      if (alwaysExp.indexOf(pp) < 0 && pp !== selExp) continue;
      for (var li = 0; li < allLocs.length; li++) {
        tr3.appendChild(th(allLocs[li], 1, 1, "", "#3b82f6"));
      }
      tr3.appendChild(th("Prom", 1, 1, "", "#d97706"));
    }
    thead.appendChild(tr3);

    // ── Filas de proveedores ──
    for (var vi = 0; vi < order.length; vi++) {
      var v = order[vi];
      var vd = data[v]; if (!vd) continue;
      var ov = (objData && objData[v]) ? objData[v] : null;
      var obj26 = ov ? (ov["2026"] !== undefined ? ov["2026"] : null) : null;
      var row = document.createElement("tr");
      var tdN = td(v, "v-name"); tdN.title = v; row.appendChild(tdN);
      // Bloque objetivos
      for (var li = 0; li < allLocs.length; li++) {
        var rsObj = isRegStart(li, allLocs, regions) ? " reg-split" : "";
        row.appendChild(td(fmt(ov ? ov[allLocs[li]] : null), "cell-obj" + rsObj));
      }
      row.appendChild(td(fmt(obj26), "cell-obj"));
      // Bloques por periodo
      for (var pi = 0; pi < visMonths.length; pi++) {
        var pp = visMonths[pi];
        var pd = vd[pp] || {};
        var actCls = (pp === period) ? " col-active" : "";
        var isExp = (alwaysExp.indexOf(pp) >= 0 || pp === selExp);
        if (isExp) {
          for (var li = 0; li < allLocs.length; li++) {
            var lc = allLocs[li];
            var val = (pd[lc] !== undefined) ? pd[lc] : null;
            var objL = ov ? (ov[lc] !== undefined ? ov[lc] : null) : null;
            var rsP = (li === 0 || isRegStart(li, allLocs, regions)) ? " reg-split" : "";
            row.appendChild(td(fmt(val), ((hasObj?cellCls(val,objL):(val===null?"cell-null":""))+actCls+rsP).trim()));
          }
          var val26 = (pd["2026"] !== undefined) ? pd["2026"] : null;
          row.appendChild(td(fmt(val26), ((hasObj?cellCls(val26,obj26):(val26===null?"cell-null":""))+" cell-prom"+actCls).trim()));
        } else {
          // Colapsado: solo el total nacional
          var val26 = (pd["2026"] !== undefined) ? pd["2026"] : null;
          row.appendChild(td(fmt(val26), ((hasObj?cellCls(val26,obj26):(val26===null?"cell-null":""))+" cell-prom text-center"+actCls).trim()));
        }
      }
      tbody.appendChild(row);
    }

    // ── Fila Total ──
    var totRow = document.createElement("tr");
    totRow.appendChild(td("Total", "v-name cell-tot font-bold"));
    var objTot = (objData && objData["__total__"]) ? objData["__total__"] : null;
    var objTot26 = objTot ? (objTot["2026"] !== undefined ? objTot["2026"] : null) : null;
    for (var li = 0; li < allLocs.length; li++) {
      var rsTotObj = isRegStart(li, allLocs, regions) ? " reg-split" : "";
      totRow.appendChild(td(fmt(objTot ? objTot[allLocs[li]] : null), "cell-tot font-bold" + rsTotObj));
    }
    totRow.appendChild(td(fmt(objTot26), "cell-tot font-bold"));
    var totData2 = data["__total__"] || {};
    for (var pi = 0; pi < visMonths.length; pi++) {
      var pp = visMonths[pi];
      var pd = totData2[pp] || {};
      var actCls = (pp === period) ? " col-active" : "";
      var isExp = (alwaysExp.indexOf(pp) >= 0 || pp === selExp);
      if (isExp) {
        for (var li = 0; li < allLocs.length; li++) {
          var lc = allLocs[li];
          var val = (pd[lc] !== undefined) ? pd[lc] : null;
          var objL = objTot ? (objTot[lc] !== undefined ? objTot[lc] : null) : null;
          var rsTot = (li === 0 || isRegStart(li, allLocs, regions)) ? " reg-split" : "";
          var tcls = (hasObj?cellCls(val,objL):(val===null?"cell-null":""))+" font-bold"+actCls+rsTot;
          totRow.appendChild(td(fmt(val), tcls.trim()));
        }
        var val26 = (pd["2026"] !== undefined) ? pd["2026"] : null;
        var tcls26 = (hasObj?cellCls(val26,objTot26):(val26===null?"cell-null":""))+" font-bold cell-prom"+actCls;
        totRow.appendChild(td(fmt(val26), tcls26.trim()));
      } else {
        var val26 = (pd["2026"] !== undefined) ? pd["2026"] : null;
        var tcls26 = (hasObj?cellCls(val26,objTot26):(val26===null?"cell-null":""))+" font-bold cell-prom text-center"+actCls;
        totRow.appendChild(td(fmt(val26), tcls26.trim()));
      }
    }
    tbody.appendChild(totRow);

  } else if (isAll && overridePeriods) {
    // ── VISTA COMPLETA modo SW: bloque por SW con CDs ──
    var secCols = allLocs.length + 1;
    var tr1 = document.createElement("tr");
    tr1.appendChild(th("Proveedores TOP", 3, 1, "v-name text-left", "#1e3a8a"));
    tr1.appendChild(th("Objetivos LOS 2026 por Cedis", 1, secCols, "", "#1e3a8a"));
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var bg = PER_BG[PERIODS[pi]] || "#6d28d9";
      tr1.appendChild(th(PER_LBL[PERIODS[pi]] || PERIODS[pi], 1, secCols, "", bg));
    }
    thead.appendChild(tr1);
    var tr2 = document.createElement("tr");
    for (var bi = 0; bi <= PERIODS.length; bi++) {
      for (var ri = 0; ri < regions.length; ri++) {
        var rbg = REG_BG[regions[ri].n] || "#334155";
        tr2.appendChild(th(regions[ri].n, 1, regions[ri].l.length, "reg-hdr", rbg));
      }
      tr2.appendChild(th("2026", 1, 1, "", "#b45309"));
    }
    thead.appendChild(tr2);
    var tr3 = document.createElement("tr");
    for (var bi = 0; bi <= PERIODS.length; bi++) {
      for (var li = 0; li < allLocs.length; li++) tr3.appendChild(th(allLocs[li], 1, 1, "", "#3b82f6"));
      tr3.appendChild(th("Prom", 1, 1, "", "#d97706"));
    }
    thead.appendChild(tr3);
    for (var vi = 0; vi < order.length; vi++) {
      var v = order[vi];
      var vd = data[v]; if (!vd) continue;
      var ov = (objData && objData[v]) ? objData[v] : null;
      var obj26 = ov ? (ov["2026"] !== undefined ? ov["2026"] : null) : null;
      var row = document.createElement("tr");
      var tdN = td(v, "v-name"); tdN.title = v; row.appendChild(tdN);
      for (var li = 0; li < allLocs.length; li++) {
        var rsObjB2 = isRegStart(li, allLocs, regions) ? " reg-split" : "";
        row.appendChild(td(fmt(ov?ov[allLocs[li]]:null), "cell-obj" + rsObjB2));
      }
      row.appendChild(td(fmt(obj26), "cell-obj"));
      for (var pi = 0; pi < PERIODS.length; pi++) {
        var pd = vd[PERIODS[pi]] || {};
        var actCls = (pi === activePeriodIdx) ? " col-active" : "";
        for (var li = 0; li < allLocs.length; li++) {
          var lc = allLocs[li]; var val = pd[lc] !== undefined ? pd[lc] : null;
          var objL = ov ? (ov[lc] !== undefined ? ov[lc] : null) : null;
          var rsPB2 = (li === 0 || isRegStart(li, allLocs, regions)) ? " reg-split" : "";
          row.appendChild(td(fmt(val), ((hasObj?cellCls(val,objL):(val===null?"cell-null":""))+actCls+rsPB2).trim()));
        }
        var val26 = pd["2026"] !== undefined ? pd["2026"] : null;
        row.appendChild(td(fmt(val26), (hasObj?cellCls(val26,obj26):(val26===null?"cell-null":""))+actCls));
      }
      tbody.appendChild(row);
    }
    var totRow = document.createElement("tr");
    totRow.appendChild(td("Total", "v-name cell-tot font-bold"));
    var objTot = (objData&&objData["__total__"])?objData["__total__"]:null;
    var objTot26 = objTot?(objTot["2026"]!==undefined?objTot["2026"]:null):null;
    for (var li = 0; li < allLocs.length; li++) {
      var rsTotB2 = isRegStart(li, allLocs, regions) ? " reg-split" : "";
      totRow.appendChild(td(fmt(objTot?objTot[allLocs[li]]:null), "cell-tot font-bold" + rsTotB2));
    }
    totRow.appendChild(td(fmt(objTot26), "cell-tot font-bold"));
    var totData2 = data["__total__"] || {};
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var pd = totData2[PERIODS[pi]] || {};
      var actCls = (pi === activePeriodIdx) ? " col-active" : "";
      for (var li = 0; li < allLocs.length; li++) {
        var lc = allLocs[li]; var val = pd[lc] !== undefined ? pd[lc] : null;
        var objL = objTot?(objTot[lc]!==undefined?objTot[lc]:null):null;
        var rsTotPB2 = (li === 0 || isRegStart(li, allLocs, regions)) ? " reg-split" : "";
        var tcls = (hasObj?cellCls(val,objL):(val===null?"cell-null":""))+" font-bold"+actCls+rsTotPB2;
        totRow.appendChild(td(fmt(val), tcls.trim()));
      }
      var val26 = pd["2026"]!==undefined?pd["2026"]:null;
      var tcls26 = (hasObj?cellCls(val26,objTot26):(val26===null?"cell-null":""))+" font-bold"+actCls;
      totRow.appendChild(td(fmt(val26), tcls26.trim()));
    }
    tbody.appendChild(totRow);

  } else {
    // ── VISTA SIMPLE (loc específica o filtro CD): muestra todos los periodos ──
    var trH = document.createElement("tr");
    trH.appendChild(th("Proveedores TOP", 1, 1, "v-name text-left", "#1e3a8a"));
    trH.appendChild(th("Objetivo", 1, 1, "", "#1e3a8a"));
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var bg = PER_BG[PERIODS[pi]] || "#1d4ed8";
      trH.appendChild(th(PER_LBL[PERIODS[pi]] || PERIODS[pi], 1, 1, "", bg));
    }
    thead.appendChild(trH);
    for (var vi = 0; vi < order.length; vi++) {
      var v = order[vi];
      var vd = data[v]; if (!vd) continue;
      var objL = getObjVal(objData, v, filtLocs, loc);
      var anyData = false;
      for (var pi = 0; pi < PERIODS.length; pi++) {
        if (getVal(data, v, PERIODS[pi], loc, filtLocs) != null) { anyData = true; break; }
      }
      if (!anyData) continue;
      var row = document.createElement("tr");
      var tdN = td(v, "v-name"); tdN.title = v; row.appendChild(tdN);
      row.appendChild(td(fmt(objL), "cell-obj"));
      for (var pi = 0; pi < PERIODS.length; pi++) {
        var val = getVal(data, v, PERIODS[pi], loc, filtLocs);
        var cls = (hasObj ? cellCls(val, objL) : (val === null ? "cell-null" : ""));
        if (pi === activePeriodIdx) cls += " col-active";
        row.appendChild(td(fmt(val), cls.trim()));
      }
      tbody.appendChild(row);
    }
    var totRow = document.createElement("tr");
    var objLT = objData && objData["__total__"]
      ? getObjVal(objData, "__total__", filtLocs, loc) : null;
    totRow.appendChild(td("Total", "v-name cell-tot font-bold"));
    totRow.appendChild(td(fmt(objLT), "cell-tot font-bold"));
    for (var pi = 0; pi < PERIODS.length; pi++) {
      var val    = getVal(data, "__total__", PERIODS[pi], loc, filtLocs);
      var clsTot = (hasObj?cellCls(val,objLT):(val===null?"cell-null":""))+" font-bold"+(pi===activePeriodIdx?" col-active":"");
      totRow.appendChild(td(fmt(val), clsTot.trim()));
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
  var ff = document.getElementById('fFin').value;
  if (!ff) { alert('Selecciona la fecha hasta la que quieres actualizar'); return; }
  var fi = '2026-01-01';
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
// Máximo = hoy (no permitir fechas futuras)
document.getElementById('fFin').max = new Date().toISOString().slice(0,10);

try { Chart.register(ChartDataLabels); } catch(e) { /* datalabels optional */ }
fetch('sw_data.json?v='+Date.now()).then(function(r){return r.json();}).then(function(d){
  SW_DATA = d;
  var panel = document.getElementById('swPanel');
  var clear = document.createElement('div');
  clear.className='ms-item'; clear.style.cssText='color:#64748b;font-style:italic;border-bottom:1px solid #e2e8f0;';
  clear.innerHTML='&#10005; Limpiar';
  clear.onclick=function(e){e.stopPropagation();gSelSW=[];syncSwPanel();renderAll();};
  panel.appendChild(clear);
  d.sw_list.forEach(function(sw){
    var key='SW'+sw;
    var mes=d.sw_mes_map[key]||'';
    var dateInfo=(d.sw_dates&&d.sw_dates[key])||null;
    var sublbl=dateInfo?dateInfo.label:mes;
    var fullLbl='SW'+sw+' \u00b7 '+sublbl;
    // Registra label y color dinamicamente para buildTable
    PER_LBL[key] = fullLbl;
    PER_BG[key]  = '#6d28d9';
    var item=document.createElement('div'); item.className='ms-item'; item.dataset.val=key;
    var cb=document.createElement('input'); cb.type='checkbox'; cb.checked=false;
    var sp=document.createElement('span');
    sp.innerHTML='<b>SW'+sw+'</b> <span style="color:#64748b;font-size:.85em">· '+sublbl+'</span>';
    item.appendChild(cb); item.appendChild(sp);
    item.addEventListener('click',function(e){e.stopPropagation();toggleSW(key);});
    panel.appendChild(item);
  });
}).catch(function(){console.warn('sw_data.json no encontrado');});
onLocChange();
updateChartTitle();
renderAll();

</script>
</body>
</html>
"""


def build_html(final, perec, last_date=""):
    # Inyectar fecha al HTML estático antes de incrustar scripts
    head_with_date = HEAD.replace("DATA_LAST_DATE_PH", last_date)
    parts = [head_with_date]

    # ── Embedded data ──
    parts.append("<script>\n")
    parts.append(f"var DATA_AUTO            = {j(final.get('auto', {}))};\n")
    parts.append(f"var DATA_BAE             = {j(final.get('bae', {}))};\n")
    parts.append(f"var DATA_SAMS            = {j(final.get('sams', {}))};\n")
    parts.append(f"var DATA_AUTO_BAE        = {j(final.get('auto_bae', {}))};\n")
    parts.append(f"var DATA_AUTO_SAMS       = {j(final.get('auto_sams', {}))};\n")
    parts.append(f"var DATA_BAE_SAMS        = {j(final.get('bae_sams', {}))};\n")
    parts.append(f"var DATA_ALL             = {j(final.get('all', {}))};\n")
    parts.append(f"var DATA_AUTO_CHART      = {j(final.get('auto_chart', {}))};\n")
    parts.append(f"var DATA_BAE_CHART       = {j(final.get('bae_chart', {}))};\n")
    parts.append(f"var DATA_AUTO_BAE_CHART  = {j(final.get('auto_bae_chart', {}))};\n")
    parts.append(f"var DATA_AUTO_SAMS_CHART = {j(final.get('auto_sams_chart', {}))};\n")
    parts.append(f"var DATA_BAE_SAMS_CHART  = {j(final.get('bae_sams_chart', {}))};\n")
    parts.append(f"var DATA_SAMS_CHART      = {j(final.get('sams_chart', {}))};\n")
    parts.append(f"var DATA_ALL_CHART       = {j(final.get('all_chart', {}))};\n")
    parts.append(f"var DATA_CD         = {j(final.get('cd_matrix', {}))};\n")
    parts.append(f"var DATA_CHART_CD   = null; // cargado lazy via fetch\n")
    parts.append(f"var DATA_LAST_DATE  = {j(last_date)};\n")
    parts.append(f"var DATA_PEREC_AUTO = {j(perec.get('auto', {}))};\n")
    parts.append(f"var DATA_PEREC_SAMS = {j(perec.get('sams', {}))};\n")
    parts.append(f"var DISPLAY_ORDER   = {j(final.get('display_order', []))};\n")
    parts.append(f"var PEREC_ORDER     = {j(perec.get('display_order', []))};\n")
    parts.append(OBJECTIVES_JS)
    parts.append(LOGIC_JS)

    return "".join(parts)


def main():
    print("Loading JSON data...")
    final, perec, last_date = load_data()
    print(f"  AUTO vendors:  {len(final.get('auto', {}))}")
    print(f"  SAMS vendors:  {len(final.get('sams', {}))}")
    print(f"  PEREC vendors: {len(perec.get('auto', {}))}")
    print(f"  Datos al:      {last_date}")

    # Guardar cd_chart separado para carga lazy
    cd_chart_path = os.path.join(BQ, "cd_chart.json")
    with open(cd_chart_path, "w", encoding="utf-8") as f:
        json.dump(final.get('cd_chart', {}), f, ensure_ascii=False, separators=(',', ':'))
    print(f"  cd_chart.json: {os.path.getsize(cd_chart_path)//1024} KB")

    print("Building HTML...")
    html = build_html(final, perec, last_date)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUT) // 1024
    print(f"Done! -> {OUT}")
    print(f"Size: {size_kb} KB ({len(html):,} chars)")


if __name__ == "__main__":
    main()
