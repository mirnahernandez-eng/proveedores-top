"""
rebuild_csv_mensual.py
======================
Reconstruye vendor_cedis_mes_FINAL.csv desde la fuente de verdad BQ.

FUENTE UNICA: bigquery_results/yms_2026_completo.csv
  - Descargado directamente de BQ con los filtros aplicados en query:
      UPPER(TIPO_CITA) IN ('PROVEEDOR', 'CITA NUEVA')
      CITAS_CORRECTAS = 1
      ARRIVAL_DATE BETWEEN '2026-01-01' AND '2026-07-03'
  - Columnas de tiempo ya en minutos: LLEGADA_A_TRAFICO, ABRIR_CORTINA,
    CERRAR_CORTINA, PAPER_W, SALIDA_DE_CD
  - Sin exclusion de tiempos altos ni negativos

Reglas de calculo:
  - formula_2 (hrs) = (LLEGADA + ABRIR + CERRAR + PAPER + SALIDA) / 60
  - formula_2 > 0   = registro valido (descarta citas sin timestamps)
  - Razon social    = .upper().strip() para unificar variantes
"""
import os, sys, pandas as pd
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE     = os.path.dirname(os.path.abspath(__file__))
BQ_CSV   = os.path.join(BASE, 'bigquery_results', 'yms_2026_completo.csv')

MES_MAP = {
    1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
    7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre',
}

# ─────────────────────────────────────────────────────────────────────────────
# 1. Lee yms_2026_completo.csv  (fuente de verdad BQ)
# ─────────────────────────────────────────────────────────────────────────────
print(f"Leyendo {BQ_CSV}...", flush=True)
if not os.path.exists(BQ_CSV):
    print(f"  ERROR: archivo no encontrado. Descarga primero desde BQ.")
    sys.exit(1)

df = pd.read_csv(BQ_CSV, encoding='utf-8-sig', low_memory=False)
print(f"  Raw: {len(df):,}")

# Normalizar razones sociales y NOMBRE_CEDIS
df['VENDOR']      = df['VENDOR'].astype(str).str.upper().str.strip()
df['NOMBRE_CEDIS'] = df['NOMBRE_CEDIS'].astype(str).str.upper().str.strip()

# ─────────────────────────────────────────────────────────────────────────────
# 2. Calcular LOS con la formula oficial:
#    LOS (hrs) = (AVG(LLEGADA_A_TRAFICO) + AVG(DURACION_DE_SERVICIO) + AVG(SALIDA_DE_CD)) / 60
#    Sin exclusion de tiempos altos ni negativos
# ─────────────────────────────────────────────────────────────────────────────
for col in ['LLEGADA_A_TRAFICO', 'DURACION_DE_SERVICIO', 'SALIDA_DE_CD']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['formula_2'] = (
    df['LLEGADA_A_TRAFICO']    +
    df['DURACION_DE_SERVICIO'] +
    df['SALIDA_DE_CD']
) / 60

# Solo registros con LOS valido (> 0)
antes = len(df)
df = df[df['formula_2'] > 0].copy()
print(f"  Tras formula_2>0: {antes:,} -> {len(df):,}")

# ─────────────────────────────────────────────────────────────────────────────
# 3. MES desde columna BQ o desde ARRIVAL_DATE
# ─────────────────────────────────────────────────────────────────────────────
if 'MES' in df.columns and df['MES'].notna().any():
    df['MES'] = df['MES'].astype(str).str.strip()
else:
    df['ARRIVAL_DATE'] = pd.to_datetime(df['ARRIVAL_DATE'], errors='coerce')
    df['MES'] = df['ARRIVAL_DATE'].dt.month.map(MES_MAP)

# ─────────────────────────────────────────────────────────────────────────────
# 4. Agrega por VENDOR + CEDIS + NOMBRE_CEDIS + MES
# ─────────────────────────────────────────────────────────────────────────────
print("Agregando...", flush=True)
grp_cols = [c for c in ['VENDOR','CEDIS','NOMBRE_CEDIS','MES'] if c in df.columns]
rows = []
for keys, sub in df.groupby(grp_cols, dropna=False):
    kd = dict(zip(grp_cols, keys if isinstance(keys, tuple) else [keys]))
    rows.append({**kd,
        'TOTAL_CITAS': len(sub),
        'LLEGADA':  round(sub['LLEGADA_A_TRAFICO'].mean()    / 60, 4),
        'RECIBO':   round(sub['DURACION_DE_SERVICIO'].mean() / 60, 4),
        'SALIDA':   round(sub['SALIDA_DE_CD'].mean()         / 60, 4),
        'TOTAL_HRS':round(sub['formula_2'].mean(),           4),
    })

out = pd.DataFrame(rows)
print(f"Filas resultado: {len(out):,}")

# ─────────────────────────────────────────────────────────────────────────────
# 5. Guarda en ambas rutas
# ─────────────────────────────────────────────────────────────────────────────
for dest in [
    os.path.join(BASE, 'vendor_cedis_mes_FINAL.csv'),
    os.path.join(BASE, 'bigquery_results', 'vendor_cedis_mes_FINAL.csv'),
]:
    out.to_csv(dest, index=False, encoding='utf-8-sig')
    print(f"Guardado: {dest}")

# ─────────────────────────────────────────────────────────────────────────────
# 6. Validacion Frabel
# ─────────────────────────────────────────────────────────────────────────────
print("\n=== Meses con datos ===")
mes_order = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio']
mes_counts = df.groupby('MES').size()
for m in mes_order:
    print(f"  {m:10s}: {mes_counts.get(m,0):,}")

print("\n=== Frabel por mes ===")
f = out[out['VENDOR'].str.contains('FRABEL', na=False, case=False)]
total_f = f['TOTAL_CITAS'].sum()
for m in mes_order:
    n = f[f['MES']==m]['TOTAL_CITAS'].sum()
    if n: print(f"  {m:10s}: {n}")
print(f"  {'TOTAL':10s}: {total_f}")
