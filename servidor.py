"""
servidor.py
FastAPI que sirve el tablero LOS Proveedores TOP 2026
y expone /api/actualizar para refrescar datos desde BigQuery.
"""
import json, os, re, unicodedata, subprocess, sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import pandas as pd
import openpyxl

BASE = Path(__file__).parent
BQ_TABLE = "wmt-intl-supplychain-gcp-prod.MR101_WM_AD_HOC.SCH_YMS_SEMANAL"

app = FastAPI(title="Tablero LOS Proveedores TOP 2026")

# ── Estado global del proceso de actualización ──────────────────────────────
_estado = {"running": False, "msg": "Listo", "ok": True, "pct": 0, "puppy_url": None}

# ── Helpers de normalización (idénticos a build_sw_data.py) ─────────────────
def _ascii(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower().strip()

_CEDIS_MAP = [
    ("chihu", "CUU"), ("culia", "CLN"), ("mexic", "MXL"),
    ("monte", "MTY"), ("cuaut", "CUAU"), ("santa bar", "STB"),
    ("chalco", "CHL"), ("guada", "GDL"), ("merid", "MER"),
    ("villaher", "VHSA"), ("san mart", "SMO"),
]
def get_cedis(loc: str) -> str:
    c = _ascii(loc)
    for pre, code in _CEDIS_MAP:
        if c.startswith(pre): return code
    return loc[:4].upper()

def get_cat(nombre: str) -> str:
    n = _ascii(str(nombre or ""))
    if "perecedero" in n: return "EXCLUIR"
    if "sam" in n:        return "SAM'S Club"
    if any(k in n for k in ("secos", "sstk", "bae", "nave")): return "Autoservicios"
    return "EXCLUIR"

# ── Carga vendors del Excel ──────────────────────────────────────────────────
def cargar_vendors():
    wb = openpyxl.load_workbook(BASE / "YMS TOP 15 2026.xlsx", data_only=True)
    CATS = {"Top 15 Proveedores": "Autoservicios", "Top 15 Sams ": "SAM'S Club"}
    kw_by_cat = {}
    seen = set()
    for sname, cat in CATS.items():
        if cat in seen: continue
        seen.add(cat)
        ws = wb[sname]
        vendors = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            rank, _, vname = row[0], row[1], row[2]
            if not isinstance(rank, int) or not isinstance(vname, str): continue
            vendors.append(vname)
        kw = {}
        for v in vendors:
            toks = re.findall(r"[A-Z]{4,}", v.upper())
            for t in toks[:3]:
                if t not in ("MEXI", "COMP", "GRUP", "CORP", "COME", "COMER", "MERC"):
                    kw[t] = v; break
        kw_by_cat[cat] = kw
    return kw_by_cat

def match_vendor(csv_v: str, kw_by_cat: dict, cat: str) -> Optional[str]:
    if not isinstance(csv_v, str): return None
    n = _ascii(csv_v.upper().strip())
    for kw, excel_v in kw_by_cat.get(cat, {}).items():
        if _ascii(kw) in n: return excel_v
    return None

# ── Query BigQuery ───────────────────────────────────────────────────────────
def run_bq_query(fecha_inicio: str, fecha_fin: str) -> pd.DataFrame:
    from google.cloud import bigquery
    client = bigquery.Client()
    query = f"""
    SELECT
      ANIO, APPOINTMENT_NBR, ARRIVAL_DATE, CEDIS,
      ARRIVAL_TS, DRIVER_ARRIVAL_TS, TRAILER_OPEN_TS,
      DOCK_DOOR_CLOSE, POD, DEPARTURE_TS,
      LOCACION, SW, TIPO_CITA, NOMBRE_CEDIS, VENDOR, NODO,
      DURACION_DE_SERVICIO, LLEGADA_A_TRAFICO, SALIDA_DE_CD,
      MES, ABRIR_CORTINA, CERRAR_CORTINA, PAPER_W,
      CITAS_CORRECTAS
    FROM `{BQ_TABLE}`
    WHERE ARRIVAL_DATE BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
      AND TRIM(UPPER(TIPO_CITA)) IN ('PROVEEDOR', 'CITA NUEVA')
      AND CITAS_CORRECTAS = 1
    ORDER BY ARRIVAL_DATE
    """
    return client.query(query).to_dataframe()

# ── Calcular métricas desde timestamps (para filas con NULLs) ───────────────
def calcular_metricas(df: pd.DataFrame) -> pd.DataFrame:
    ts_cols = ["ARRIVAL_TS", "DRIVER_ARRIVAL_TS", "TRAILER_OPEN_TS",
               "DOCK_DOOR_CLOSE", "POD", "DEPARTURE_TS"]
    for c in ts_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", utc=True)

    def diff_min(a, b):
        return (a - b).dt.total_seconds() / 60

    # Guardar valores BQ pre-calculados antes de sobreescribir
    bq_llegada  = pd.to_numeric(df.get("LLEGADA_A_TRAFICO"),   errors="coerce")
    bq_duracion = pd.to_numeric(df.get("DURACION_DE_SERVICIO"), errors="coerce")
    bq_salida   = pd.to_numeric(df.get("SALIDA_DE_CD"),         errors="coerce")

    df["LLEGADA_A_TRAFICO"]    = diff_min(df["DRIVER_ARRIVAL_TS"], df["ARRIVAL_TS"])
    df["ABRIR"]                = diff_min(df["TRAILER_OPEN_TS"],   df["DRIVER_ARRIVAL_TS"])
    df["CERRAR"]               = diff_min(df["DOCK_DOOR_CLOSE"],   df["TRAILER_OPEN_TS"])
    df["PAPER"]                = diff_min(df["POD"],               df["DOCK_DOOR_CLOSE"])
    df["SALIDA_DE_CD"]         = diff_min(df["DEPARTURE_TS"],      df["POD"])
    df["DURACION_DE_SERVICIO"] = df[["ABRIR", "CERRAR", "PAPER"]].sum(axis=1, min_count=1)

    # formula_2: preferir valores BQ (idénticos a lo que el usuario ve en BQ)
    # Solo usar timestamps recalculados cuando BQ no tiene los tres campos
    bq_los = (bq_llegada.fillna(0) + bq_duracion.fillna(0) + bq_salida.fillna(0)) / 60
    bq_has_values = (bq_llegada.notna() | bq_duracion.notna() | bq_salida.notna())
    ts_los = (
        df["LLEGADA_A_TRAFICO"].fillna(0) +
        df["DURACION_DE_SERVICIO"].fillna(0) +
        df["SALIDA_DE_CD"].fillna(0)
    ) / 60
    df["formula_2"] = bq_los.where(bq_has_values & (bq_los > 0), ts_los)
    df.loc[df["formula_2"] <= 0, "formula_2"] = float("nan")
    return df

# ── Agregar datos por SW ─────────────────────────────────────────────────────
from sw_calendar import SW_MES_MAP, SW_DATES, sw_range_label  # fuente única de verdad

DISPLAY_ORDER = [
    "KIMBERLY CLARK DE MEX SA B CV","ENBOTELLAD NIAGARA D MX","JUGOS DEL VALLE",
    "SANTA CLARA MERC PACHU S RL CV","PROCTER AND GAMBLE MEXICO INC","MARCAS NESTLE",
    "COLGATE PALMOLIVE SA CV","COMERC PEPSICO MEXICO S RL CV","BONAFONT + ENVASASORA",
    "UNILEVER DE MEXICO S RL CV","HERDEZ SA DE CV","CERVEZA CANAL MO S D",
    "FRABEL SA DE CV","MONDELEZ MEXICO S DE RL DE CV","KELLOGG COMPANY MEXICO SRL CV",
]

def agregar_sw(df: pd.DataFrame, kw_by_cat: dict, fecha_fin: str = None) -> dict:
    """Genera estructura sw_data.json compatible con el tablero."""
    from collections import defaultdict

    # Convertir SW a int limpio
    df["SW"] = pd.to_numeric(df["SW"], errors="coerce").dropna().astype(int)

    # Convertir horas → horas (ya están en minutos, dividir por 60)
    for col in ["LLEGADA_A_TRAFICO","DURACION_DE_SERVICIO","SALIDA_DE_CD","formula_2"]:
        if col not in df.columns:
            df[col] = float("nan")

    raw = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {
        "c":0,"wl":0.0,"wr":0.0,"ws":0.0,"wt":0.0
    })))

    for _, row in df.iterrows():
        cat = get_cat(str(row.get("NOMBRE_CEDIS","")))
        if cat == "EXCLUIR": continue
        v = match_vendor(str(row.get("VENDOR","")), kw_by_cat, cat)
        if not v: continue
        cedis_lbl = get_cedis(str(row.get("LOCACION","")))
        sw_num = row.get("SW")
        if pd.isna(sw_num): continue
        sw_key = f"SW{int(sw_num)}"

        l = float(row.get("LLEGADA_A_TRAFICO") or 0) / 60
        r = float(row.get("DURACION_DE_SERVICIO") or 0) / 60
        s = float(row.get("SALIDA_DE_CD") or 0) / 60
        t = float(row.get("formula_2") or 0)
        if t <= 0: t = l + r + s

        rec = raw[v][cedis_lbl][sw_key]
        rec["c"]  += 1
        rec["wl"] += l
        rec["wr"] += r
        rec["ws"] += s
        rec["wt"] += t

    def avg(rec): return {k: round(rec[k]/rec["c"],2) if rec["c"] else None for k in ["wl","wr","ws","wt"]}

    sw_list_nums = sorted(SW_MES_MAP.keys())
    sw_mes_map   = {f"SW{k}": v for k,v in SW_MES_MAP.items()}

    # Construir auto / sams (estructura para gráficas)
    AUTO_CEDIS = ["CUU","CLN","MXL","MTY","CUAU","STB","SMO","CHL","GDL","MER","VHSA"]
    SAMS_CEDIS = ["CUU","CLN","MXL","MTY","SMO","CHL","GDL","MER","VHSA"]

    def build_chart(cedis_list):
        out = {}
        for v in DISPLAY_ORDER:
            out[v] = {}
            # Promedio por locacion
            for cedis in cedis_list:
                out[v][cedis] = {}
                for sw_key in [f"SW{n}" for n in sw_list_nums]:
                    rec = raw[v][cedis].get(sw_key)
                    if rec and rec["c"]:
                        a = avg(rec)
                        out[v][cedis][sw_key] = {"l":a["wl"],"r":a["wr"],"s":a["ws"],"t":a["wt"]}
            # Promedio nacional '2026' — requerido por paintChart cuando loc='__all__'
            nat = {}
            for sw_key in [f"SW{n}" for n in sw_list_nums]:
                wl, wr, ws, wt, cnt = 0.0, 0.0, 0.0, 0.0, 0
                for cedis in cedis_list:
                    rec = raw[v][cedis].get(sw_key)
                    if rec and rec["c"]:
                        wl += rec["wl"]; wr += rec["wr"]
                        ws += rec["ws"]; wt += rec["wt"]
                        cnt += rec["c"]
                if cnt:
                    nat[sw_key] = {"l":round(wl/cnt,2),"r":round(wr/cnt,2),
                                   "s":round(ws/cnt,2),"t":round(wt/cnt,2)}
            out[v]["2026"] = nat
        return out

    def build_tbl(cedis_list):
        out = {}
        for v in DISPLAY_ORDER:
            out[v] = {}
            for sw_key in [f"SW{n}" for n in sw_list_nums]:
                out[v][sw_key] = {}
                for cedis in cedis_list:
                    rec = raw[v][cedis].get(sw_key)
                    if rec and rec["c"]:
                        out[v][sw_key][cedis] = round(raw[v][cedis][sw_key]["wt"]/rec["c"],2)
        return out

    # Filtrar sw_list a solo SWs con datos reales
    auto_data = build_chart(AUTO_CEDIS)
    sams_data = build_chart(SAMS_CEDIS)
    sw_has_data = set()
    for section in (auto_data, sams_data):
        for vendor, cedis_map in section.items():
            nat = cedis_map.get("2026", {})
            for sw_key, vals in nat.items():
                if vals and vals.get("t"):
                    sw_has_data.add(int(sw_key[2:]))
    # Excluir SWs cuyo inicio sea posterior a fecha_fin (evita que BQ
    # meta SWs futuros por discrepancias en el campo SW de la tabla)
    if fecha_fin:
        sw_has_data = {
            k for k in sw_has_data
            if k in SW_DATES and SW_DATES[k]["inicio"] <= fecha_fin
        }
    # Purgar SWs excluidos de auto_data y sams_data para que no contaminen
    # promedios ni graficas (no basta con quitarlos de sw_dates)
    bad_keys = {
        f"SW{k}" for k in SW_DATES
        if fecha_fin and SW_DATES[k]["inicio"] > fecha_fin
    }
    if bad_keys:
        for section in (auto_data, sams_data):
            for vendor, cedis_map in section.items():
                for cedis, sw_map in cedis_map.items():
                    for bk in bad_keys:
                        sw_map.pop(bk, None)

    sw_list_final    = [n for n in sw_list_nums if n in sw_has_data]
    sw_mes_map_final = {f"SW{k}": v for k, v in SW_MES_MAP.items() if k in sw_has_data}
    sw_dates_final   = {
        f"SW{k}": {"inicio": SW_DATES[k]["inicio"], "fin": SW_DATES[k]["fin"],
                   "label": sw_range_label(k)}
        for k in sw_has_data if k in SW_DATES
    }

    tbl_auto = build_tbl(AUTO_CEDIS)
    return {
        "sw_list":    sw_list_final,
        "sw_mes_map": sw_mes_map_final,
        "sw_dates":   sw_dates_final,
        "auto":       auto_data,
        "sams":       sams_data,
        "auto_bae":   auto_data,      # alias: mismos vendors, BAE es subconjunto de Auto
        "tbl_auto":   tbl_auto,
        "tbl_sams":   build_tbl(SAMS_CEDIS),
        "tbl_auto_bae": tbl_auto,     # alias
    }

# ── Generar vendor_cedis_mes_FINAL.csv desde datos BQ frescos ───────────────
MES_NUM_MAP = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}

def generar_csv_mensual(df: pd.DataFrame, kw_by_cat: dict):
    """Genera vendor_cedis_mes_FINAL.csv desde el DataFrame fresco de BigQuery."""
    d = df.copy()

    # Derivar MES desde columna BQ si existe, si no desde ARRIVAL_DATE
    if "MES" in d.columns and d["MES"].notna().any():
        d["_mes"] = d["MES"].astype(str).str.strip()
    else:
        d["ARRIVAL_DATE"] = pd.to_datetime(d["ARRIVAL_DATE"], errors="coerce")
        d["_mes"] = d["ARRIVAL_DATE"].dt.month.map(MES_NUM_MAP)

    # Filtrar categorias validas
    d["_cat"] = d["NOMBRE_CEDIS"].apply(get_cat)
    d = d[d["_cat"] != "EXCLUIR"].copy()

    # Solo tipos de cita validos (candado: nunca incluir Backhaul ni otros)
    TIPOS_OK = {"PROVEEDOR", "CITA NUEVA"}
    if "TIPO_CITA" in d.columns:
        d = d[d["TIPO_CITA"].str.upper().str.strip().isin(TIPOS_OK)].copy()

    # Solo uso correcto
    if "CITAS_CORRECTAS" in d.columns:
        d = d[pd.to_numeric(d["CITAS_CORRECTAS"], errors="coerce") == 1].copy()

    # Solo filas con LOS valido
    d = d[d["formula_2"] > 0].copy()

    # Columnas de tiempo en minutos — usar las calculadas por calcular_metricas
    # (ABRIR, CERRAR, PAPER ya estan en el df si calcular_metricas corrio antes)
    for col_orig, col_bq in [("ABRIR", "ABRIR_CORTINA"), ("CERRAR", "CERRAR_CORTINA"), ("PAPER", "PAPER_W")]:
        if col_orig not in d.columns and col_bq in d.columns:
            d[col_orig] = pd.to_numeric(d[col_bq], errors="coerce").fillna(0)
        elif col_orig not in d.columns:
            d[col_orig] = 0.0

    # Convertir minutos -> horas
    for col in ["LLEGADA_A_TRAFICO", "ABRIR", "CERRAR", "PAPER", "SALIDA_DE_CD"]:
        d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0) / 60

    # Agrupar: VENDOR x CEDIS x MES x ANIO
    grp = d.groupby(["VENDOR", "CEDIS", "_mes", "ANIO"], as_index=False).agg(
        TOTAL_CITAS=("APPOINTMENT_NBR", "count"),
        LLEGADA=("LLEGADA_A_TRAFICO", "mean"),
        ABRIR=("ABRIR", "mean"),
        CERRAR=("CERRAR", "mean"),
        PAPER=("PAPER", "mean"),
        SALIDA=("SALIDA_DE_CD", "mean"),
        TOTAL_HRS=("formula_2", "mean"),
        LOS_SUM=("formula_2", "sum"),   # suma exacta — evita error de redondeo en YTD
    )
    grp = grp.rename(columns={"_mes": "MES"})
    for col in ["LLEGADA", "ABRIR", "CERRAR", "PAPER", "SALIDA", "TOTAL_HRS"]:
        grp[col] = grp[col].round(4)
    # LOS_SUM sin redondear para preservar precisión completa

    # Guardar en ambas rutas
    for path in [
        BASE / "vendor_cedis_mes_FINAL.csv",
        BASE / "bigquery_results" / "vendor_cedis_mes_FINAL.csv",
    ]:
        grp.to_csv(path, index=False, encoding="utf-8-sig")

    return len(grp), int(grp["TOTAL_CITAS"].sum())

# ── Vendors de Perecederos (top 5) y sus variantes de nombre en BQ ──────────
PEREC_VENDOR_PREFIX = {
    "DRISCOLL":                    "DRISCOLL S OPERACIONES SA C",
    "PILGRIMS PRIDE":              "PILGRIMS PRIDE S DE RL DE C",
    "LANDEROS PALAZUELOS":         "LANDEROS PALAZUELOS EDUARDO",
    "MJ INTERNATIONAL":            "MJ INTERNATIONAL MARKETIN S",
    "FRUTAS Y LEGUMBRES ALPHA":    "FRUTAS Y LEGUMBRES ALPHA SA CV",
}

def _perec_display(vendor_raw: str):
    v = str(vendor_raw or "").upper().strip()
    for pre, disp in PEREC_VENDOR_PREFIX.items():
        if v.startswith(pre):
            return disp
    return None

def generar_csv_perec(df: pd.DataFrame):
    """Genera vendor_cedis_mes_PEREC.csv (Top 5 proveedores de Perecederos)
    desde el MISMO DataFrame fresco de BigQuery -- estas filas normalmente se
    descartan en generar_csv_mensual() porque NOMBRE_CEDIS contiene
    'perecedero' (ver get_cat). Aqui las rescatamos aparte."""
    d = df.copy()

    if "MES" in d.columns and d["MES"].notna().any():
        d["_mes"] = d["MES"].astype(str).str.strip()
    else:
        d["ARRIVAL_DATE"] = pd.to_datetime(d["ARRIVAL_DATE"], errors="coerce")
        d["_mes"] = d["ARRIVAL_DATE"].dt.month.map(MES_NUM_MAP)

    d["_disp"] = d["VENDOR"].apply(_perec_display)
    d = d[d["_disp"].notna()].copy()
    d = d[d["NOMBRE_CEDIS"].apply(lambda n: "perecedero" in _ascii(n))].copy()

    TIPOS_OK = {"PROVEEDOR", "CITA NUEVA"}
    if "TIPO_CITA" in d.columns:
        d = d[d["TIPO_CITA"].str.upper().str.strip().isin(TIPOS_OK)].copy()
    if "CITAS_CORRECTAS" in d.columns:
        d = d[pd.to_numeric(d["CITAS_CORRECTAS"], errors="coerce") == 1].copy()
    d = d[d["formula_2"] > 0].copy()

    for col_orig, col_bq in [("ABRIR", "ABRIR_CORTINA"), ("CERRAR", "CERRAR_CORTINA"), ("PAPER", "PAPER_W")]:
        if col_orig not in d.columns and col_bq in d.columns:
            d[col_orig] = pd.to_numeric(d[col_bq], errors="coerce").fillna(0)
        elif col_orig not in d.columns:
            d[col_orig] = 0.0
    for col in ["LLEGADA_A_TRAFICO", "ABRIR", "CERRAR", "PAPER", "SALIDA_DE_CD"]:
        d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0) / 60

    if d.empty:
        grp = pd.DataFrame(columns=["VENDOR", "CEDIS", "MES", "ANIO", "TOTAL_CITAS",
                                     "LLEGADA", "CERRAR", "PAPER", "SALIDA",
                                     "TOTAL_HRS", "LOS_SUM"])
    else:
        d["VENDOR"] = d["_disp"]
        grp = d.groupby(["VENDOR", "CEDIS", "_mes", "ANIO"], as_index=False).agg(
            TOTAL_CITAS=("APPOINTMENT_NBR", "count"),
            LLEGADA=("LLEGADA_A_TRAFICO", "mean"),
            ABRIR=("ABRIR", "mean"),
            CERRAR=("CERRAR", "mean"),
            PAPER=("PAPER", "mean"),
            SALIDA=("SALIDA_DE_CD", "mean"),
            TOTAL_HRS=("formula_2", "mean"),
            LOS_SUM=("formula_2", "sum"),
        )
        grp = grp.rename(columns={"_mes": "MES"})
        for col in ["LLEGADA", "ABRIR", "CERRAR", "PAPER", "SALIDA", "TOTAL_HRS"]:
            grp[col] = grp[col].round(4)

    for path in [
        BASE / "bigquery_results" / "vendor_cedis_mes_PEREC.csv",
    ]:
        grp.to_csv(path, index=False, encoding="utf-8-sig")

    return len(grp), int(grp["TOTAL_CITAS"].sum())

# ── Pipeline completo de actualización ──────────────────────────────────────
def pipeline_actualizar(fecha_inicio: str, fecha_fin: str):
    global _estado
    try:
        _estado = {"running": True, "msg": "Conectando a BigQuery...", "ok": True, "pct": 5}
        df_bq = run_bq_query(fecha_inicio, fecha_fin)
        _estado["msg"] = f"Descargadas {len(df_bq):,} filas. Calculando métricas..."
        _estado["pct"] = 35

        df_bq = calcular_metricas(df_bq)

        # Descartar registros cuyo campo SW en BQ corresponde a una semana
        # que aún no ha comenzado según fecha_fin (BQ a veces asigna SW25
        # a citas que llegaron el último día de SW24, contaminando el conteo)
        if "SW" in df_bq.columns:
            sw_num_col = pd.to_numeric(df_bq["SW"], errors="coerce")
            valid_sws  = {
                k for k, v in SW_DATES.items()
                if v["inicio"] <= fecha_fin
            }
            mask = sw_num_col.isna() | sw_num_col.isin(valid_sws)
            dropped = int((~mask).sum())
            df_bq = df_bq[mask].copy()
            if dropped:
                _estado["msg"] = f"Filtrados {dropped} registros con SW fuera de rango. Cargando vendors..."

        _estado["msg"] = "Métricas calculadas. Cargando vendors del Excel..."
        _estado["pct"] = 55

        kw_by_cat = cargar_vendors()
        _estado["msg"] = "Generando vendor_cedis_mes_FINAL.csv..."
        _estado["pct"] = 65

        filas_csv, citas_csv = generar_csv_mensual(df_bq, kw_by_cat)
        filas_perec, citas_perec = generar_csv_perec(df_bq)
        _estado["msg"] = (f"CSV mensual: {filas_csv} filas / {citas_csv:,} citas. "
                           f"Perecederos: {filas_perec} filas / {citas_perec:,} citas. Agregando por SW...")
        _estado["pct"] = 70

        sw_data = agregar_sw(df_bq, kw_by_cat, fecha_fin=fecha_fin)

        # Guardar sw_data.json (reemplaza NaN/Infinity con null para JSON valido)
        sw_path = BASE / "sw_data.json"
        import re as _re
        sw_raw = json.dumps(sw_data, ensure_ascii=False, separators=(",", ":"))
        sw_raw = _re.sub(r'\bNaN\b', 'null', sw_raw)
        sw_raw = _re.sub(r'-?Infinity\b', 'null', sw_raw)
        sw_path.write_text(sw_raw, encoding="utf-8")
        _estado["msg"] = f"sw_data.json guardado ({sw_path.stat().st_size//1024} KB). Regenerando tablero..."
        _estado["pct"] = 85

        # Regenerar también matrix_FINAL.json y el HTML si los datos mensuales cambiaron
        # (se hace vía subprocess para reusar la lógica existente)
        result = subprocess.run(
            [sys.executable, str(BASE / "gen_matrix_FINAL.py")],
            capture_output=True, text=True, cwd=str(BASE)
        )
        if result.returncode != 0:
            _estado["msg"] = f"Advertencia gen_matrix: {result.stderr[:200]}"
        _estado["pct"] = 90

        result_perec = subprocess.run(
            [sys.executable, str(BASE / "gen_matrix_perec.py")],
            capture_output=True, text=True, cwd=str(BASE)
        )
        if result_perec.returncode != 0:
            _estado["msg"] = f"Advertencia gen_matrix_perec: {result_perec.stderr[:200]}"
        _estado["pct"] = 92

        result2 = subprocess.run(
            [sys.executable, str(BASE / "build_tablero.py")],
            capture_output=True, text=True, cwd=str(BASE)
        )
        if result2.returncode != 0:
            raise RuntimeError(f"Error build_tablero: {result2.stderr[:300]}")

        # ── Generar standalone + publicar a Puppy Pages ─────────────────────
        _estado["msg"] = "Generando versión standalone para Puppy Pages..."
        _estado["pct"] = 95
        result3 = subprocess.run(
            [sys.executable, str(BASE / "make_standalone.py")],
            capture_output=True, text=True, cwd=str(BASE)
        )
        if result3.returncode != 0:
            # No es fatal — el tablero local ya está actualizado
            _estado = {
                "running": False,
                "msg": f" Actualización completa pero standalone falló: {result3.stderr[:200]}",
                "ok": True,
                "pct": 100,
                "puppy_url": None,
            }
            return

        _estado["msg"] = "Publicando en Puppy Pages..."
        try:
            puppy_url = publicar_a_puppy_pages()
        except Exception as pub_err:
            puppy_url = None
            _estado["msg"] = f" Actualización completa pero no se pudo publicar: {str(pub_err)[:200]}"

        _estado = {
            "running": False,
            "msg": (
                f" Actualización completa — {len(df_bq):,} registros procesados ({fecha_inicio} → {fecha_fin}). "
                + (f"Publicado en Puppy Pages " if puppy_url else "Publicación falló ")
            ),
            "ok": True,
            "pct": 100,
            "puppy_url": puppy_url,
        }
    except Exception as e:
        _estado = {"running": False, "msg": f" Error: {str(e)[:300]}", "ok": False, "pct": 0, "puppy_url": None}

# ── Publicar a Puppy Pages ──────────────────────────────────────────────────
PUPPY_API  = "https://puppy.walmart.com/api/sharing"
PUPPY_PAGE = "tablero-yms-top-2026"

def _leer_token() -> str:
    import configparser
    cfg = configparser.ConfigParser()
    cfg.read(Path.home() / ".code_puppy" / "puppy.cfg")
    return cfg["puppy"]["puppy_token"]

def publicar_a_puppy_pages() -> str:
    """Sube tablero_standalone.html a Puppy Pages. Devuelve la URL pública."""
    import requests
    standalone = BASE / "tablero_standalone.html"
    if not standalone.exists():
        raise FileNotFoundError("tablero_standalone.html no encontrado — corre make_standalone.py primero")
    html = standalone.read_text(encoding="utf-8")
    token = _leer_token()
    resp = requests.post(
        f"{PUPPY_API}/pages",
        json={
            "name":        PUPPY_PAGE,
            "business":    "general",
            "html":        html,
            "description": "Tablero LOS Proveedores TOP 2026",
        },
        headers={"Authorization": f"Bearer {token}"},
        timeout=120,
    )
    resp.raise_for_status()
    return resp.json().get("url", f"https://puppy.walmart.com/sharing/mmvhern/{PUPPY_PAGE}")

# ── Rutas ────────────────────────────────────────────────────────────────────
class ActualizarRequest(BaseModel):
    fecha_inicio: str   # YYYY-MM-DD
    fecha_fin: str      # YYYY-MM-DD

@app.post("/api/actualizar")
async def actualizar(req: ActualizarRequest):
    global _estado
    if _estado["running"]:
        raise HTTPException(status_code=409, detail="Ya hay una actualización en curso")

    # Validar fechas
    try:
        fi = datetime.strptime(req.fecha_inicio, "%Y-%m-%d")
        ff = datetime.strptime(req.fecha_fin,    "%Y-%m-%d")
        if fi > ff:
            raise ValueError("fecha_inicio > fecha_fin")
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    # Correr en hilo background para no bloquear
    import threading
    threading.Thread(
        target=pipeline_actualizar,
        args=(req.fecha_inicio, req.fecha_fin),
        daemon=True
    ).start()

    return {"status": "iniciado"}

@app.get("/api/estado")
async def estado():
    return JSONResponse(_estado)

@app.get("/")
@app.get("/tablero_los_proveedores.html")
async def dashboard():
    from fastapi.responses import Response
    content = (BASE / "tablero_los_proveedores.html").read_bytes()
    return Response(content=content, media_type="text/html",
                    headers={"Cache-Control": "no-store, no-cache, must-revalidate",
                             "Pragma": "no-cache"})

@app.get("/cd_chart.json")
async def cd_chart_json():
    path = BASE / "bigquery_results" / "cd_chart.json"
    if path.exists():
        return FileResponse(path, media_type="application/json")
    raise HTTPException(status_code=404, detail="cd_chart.json no encontrado")

@app.get("/{path:path}")
async def static(path: str):
    file_path = BASE / path
    if file_path.exists() and file_path.is_file():
        # JSON y CSV siempre frescos — evita que el browser cachee sw_data.json viejo
        if file_path.suffix in (".json", ".csv"):
            from fastapi.responses import Response
            content = file_path.read_bytes()
            media = "application/json" if file_path.suffix == ".json" else "text/csv"
            return Response(
                content=content, media_type=media,
                headers={"Cache-Control": "no-store, no-cache, must-revalidate",
                         "Pragma": "no-cache"}
            )
        return FileResponse(file_path)
    raise HTTPException(status_code=404, detail="Archivo no encontrado")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=False)
