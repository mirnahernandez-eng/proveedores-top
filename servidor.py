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
_estado = {"running": False, "msg": "Listo", "ok": True, "pct": 0}

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
        # Sin exclusiones: se conservan negativos y tiempos altos
        return (a - b).dt.total_seconds() / 60

    df["LLEGADA_A_TRAFICO"]    = diff_min(df["DRIVER_ARRIVAL_TS"], df["ARRIVAL_TS"])
    df["ABRIR"]                = diff_min(df["TRAILER_OPEN_TS"],   df["DRIVER_ARRIVAL_TS"])
    df["CERRAR"]               = diff_min(df["DOCK_DOOR_CLOSE"],   df["TRAILER_OPEN_TS"])
    df["PAPER"]                = diff_min(df["POD"],               df["DOCK_DOOR_CLOSE"])
    df["SALIDA_DE_CD"]         = diff_min(df["DEPARTURE_TS"],      df["POD"])
    df["DURACION_DE_SERVICIO"] = df[["ABRIR", "CERRAR", "PAPER"]].sum(axis=1, min_count=1)
    df["formula_2"] = (
        df["LLEGADA_A_TRAFICO"].fillna(0) +
        df["DURACION_DE_SERVICIO"].fillna(0) +
        df["SALIDA_DE_CD"].fillna(0)
    ) / 60
    df.loc[df["formula_2"] <= 0, "formula_2"] = float("nan")
    return df

# ── Agregar datos por SW ─────────────────────────────────────────────────────
SW_MES_MAP = {
    48:"Enero",49:"Enero",50:"Enero",51:"Enero",52:"Enero",
    1:"Febrero",2:"Febrero",3:"Febrero",4:"Febrero",
    5:"Marzo",6:"Marzo",7:"Marzo",8:"Marzo",9:"Marzo",
    10:"Abril",11:"Abril",12:"Abril",13:"Abril",
    14:"Mayo",15:"Mayo",16:"Mayo",17:"Mayo",
    18:"Junio",19:"Junio",20:"Junio",21:"Junio",22:"Junio",
}

DISPLAY_ORDER = [
    "KIMBERLY CLARK DE MEX SA B CV","ENBOTELLAD NIAGARA D MX","JUGOS DEL VALLE",
    "SANTA CLARA MERC PACHU S RL CV","PROCTER AND GAMBLE MEXICO INC","MARCAS NESTLE",
    "COLGATE PALMOLIVE SA CV","COMERC PEPSICO MEXICO S RL CV","BONAFONT + ENVASASORA",
    "UNILEVER DE MEXICO S RL CV","HERDEZ SA DE CV","CERVEZA CANAL MO S D",
    "FRABEL SA DE CV","MONDELEZ MEXICO S DE RL DE CV","KELLOGG COMPANY MEXICO SRL CV",
]

def agregar_sw(df: pd.DataFrame, kw_by_cat: dict) -> dict:
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
            for cedis in cedis_list:
                out[v][cedis] = {}
                for sw_key in [f"SW{n}" for n in sw_list_nums]:
                    rec = raw[v][cedis].get(sw_key)
                    if rec and rec["c"]:
                        a = avg(rec)
                        out[v][cedis][sw_key] = {"l":a["wl"],"r":a["wr"],"s":a["ws"],"t":a["wt"]}
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

    return {
        "sw_list":    sw_list_nums,
        "sw_mes_map": sw_mes_map,
        "auto":       build_chart(AUTO_CEDIS),
        "sams":       build_chart(SAMS_CEDIS),
        "tbl_auto":   build_tbl(AUTO_CEDIS),
        "tbl_sams":   build_tbl(SAMS_CEDIS),
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
    )
    grp = grp.rename(columns={"_mes": "MES"})
    for col in ["LLEGADA", "ABRIR", "CERRAR", "PAPER", "SALIDA", "TOTAL_HRS"]:
        grp[col] = grp[col].round(4)

    # Guardar en ambas rutas
    for path in [
        BASE / "vendor_cedis_mes_FINAL.csv",
        BASE / "bigquery_results" / "vendor_cedis_mes_FINAL.csv",
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
        _estado["msg"] = "Métricas calculadas. Cargando vendors del Excel..."
        _estado["pct"] = 55

        kw_by_cat = cargar_vendors()
        _estado["msg"] = "Generando vendor_cedis_mes_FINAL.csv..."
        _estado["pct"] = 65

        filas_csv, citas_csv = generar_csv_mensual(df_bq, kw_by_cat)
        _estado["msg"] = f"CSV mensual: {filas_csv} filas / {citas_csv:,} citas. Agregando por SW..."
        _estado["pct"] = 70

        sw_data = agregar_sw(df_bq, kw_by_cat)

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
        _estado["pct"] = 92

        result2 = subprocess.run(
            [sys.executable, str(BASE / "build_tablero.py")],
            capture_output=True, text=True, cwd=str(BASE)
        )
        if result2.returncode != 0:
            raise RuntimeError(f"Error build_tablero: {result2.stderr[:300]}")

        _estado = {
            "running": False,
            "msg": f" Actualización completa — {len(df_bq):,} registros procesados ({fecha_inicio} → {fecha_fin})",
            "ok": True,
            "pct": 100
        }
    except Exception as e:
        _estado = {"running": False, "msg": f" Error: {str(e)[:300]}", "ok": False, "pct": 0}

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
        return FileResponse(file_path)
    raise HTTPException(status_code=404, detail="Archivo no encontrado")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=False)
