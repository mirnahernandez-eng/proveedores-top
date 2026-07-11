"""
append_sw_data.py
Lee sw21_22_new.csv, calcula métricas desde timestamps,
hace append al YMS_24.CSV y re-genera sw_data.json.
"""
import pandas as pd
import os

BASE    = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP'
NEW_CSV = os.path.join(BASE, 'sw21_22_new.csv')
MAIN_CSV= os.path.join(BASE, 'YMS_24.CSV')

# ── 1. Leer nuevo CSV ─────────────────────────────────────────────────────────
print('Leyendo sw21_22_new.csv...')
df = pd.read_csv(NEW_CSV, encoding='utf-8', encoding_errors='replace', low_memory=False)
print(f'  Filas raw: {len(df):,}  | SW: {sorted(df["SW"].dropna().unique().tolist())}')

# Candado: solo tipos de cita validos
TIPOS_OK = {'Proveedor', 'Cita Nueva'}
if 'TIPO_CITA' in df.columns:
    antes = len(df)
    df = df[df['TIPO_CITA'].isin(TIPOS_OK)].copy()
    print(f'  Filtro TIPO_CITA ({TIPOS_OK}): {antes:,} -> {len(df):,} filas')

# ── 2. Parsear timestamps ─────────────────────────────────────────────────────
ts_cols = ['ARRIVAL_TS','DRIVER_ARRIVAL_TS','TRAILER_OPEN_TS',
           'DOCK_DOOR_CLOSE','POD','DEPARTURE_TS']
for c in ts_cols:
    df[c] = pd.to_datetime(df[c], errors='coerce', utc=True)

def diff_min(a, b):
    """Diferencia (a - b) en minutos; negativo o gigante -> NaN."""
    d = (a - b).dt.total_seconds() / 60
    d = d.where((d >= 0) & (d < 1440))   # 0..24h válido
    return d

# Calcular segmentos en minutos
df['LLEGADA_A_TRAFICO']    = diff_min(df['DRIVER_ARRIVAL_TS'], df['ARRIVAL_TS'])
df['ABRIR_CORTINA']        = diff_min(df['TRAILER_OPEN_TS'],   df['DRIVER_ARRIVAL_TS'])
df['CERRAR_CORTINA']       = diff_min(df['DOCK_DOOR_CLOSE'],   df['TRAILER_OPEN_TS'])
df['PAPER_W']              = diff_min(df['POD'],               df['DOCK_DOOR_CLOSE'])
df['SALIDA_DE_CD']         = diff_min(df['DEPARTURE_TS'],      df['POD'])
df['DURACION_DE_SERVICIO'] = df[['ABRIR_CORTINA','CERRAR_CORTINA','PAPER_W']].sum(axis=1, min_count=1)

# formula_2 = LOS total en HORAS (como en el CSV original)
df['formula_2'] = (df['LLEGADA_A_TRAFICO'].fillna(0) +
                   df['DURACION_DE_SERVICIO'].fillna(0) +
                   df['SALIDA_DE_CD'].fillna(0)) / 60
df.loc[df['formula_2'] <= 0, 'formula_2'] = float('nan')

print(f'  formula_2 no-null: {df["formula_2"].notna().sum():,}')
print(f'  formula_2 mean: {df["formula_2"].mean():.2f}h')

# ── 3. Alinear columnas con el CSV principal ──────────────────────────────────
print('Leyendo cabecera del CSV principal...')
main_cols = pd.read_csv(MAIN_CSV, nrows=0).columns.tolist()
print(f'  Columnas originales: {len(main_cols)}')

# Convertir timestamps de vuelta a string (para compatibilidad con CSV)
for c in ts_cols:
    df[c] = df[c].dt.strftime('%Y-%m-%d %H:%M:%S UTC').where(df[c].notna(), other='')

# Añadir columnas faltantes como vacías
for col in main_cols:
    if col not in df.columns:
        df[col] = ''

df_out = df[main_cols]

# ── 4. Escribir CSV temporal y luego append via OS ─────────────────────────────
TEMP_CSV = r'C:\Users\mmvhern\sw21_22_append.csv'
print(f'Escribiendo CSV temporal en {TEMP_CSV}...')
df_out.to_csv(TEMP_CSV, index=False, header=False, encoding='utf-8', errors='replace')
print(f'  OK! {len(df_out):,} filas escritas')

# Append via Windows type command
print('Haciendo append al CSV principal...')
import subprocess
result = subprocess.run(
    f'type "{TEMP_CSV}" >> "{MAIN_CSV}"',
    shell=True, capture_output=True, text=True
)
if result.returncode != 0:
    print('Error:', result.stderr)
else:
    print('  Append OK!')

# ── 5. Verificar ─────────────────────────────────────────────────────────────
df_verify = pd.read_csv(MAIN_CSV, usecols=['SW','MES'],
                        encoding='utf-8', encoding_errors='replace', low_memory=False)
df_verify['SW'] = pd.to_numeric(df_verify['SW'], errors='coerce')
df_verify = df_verify.dropna(subset=['SW'])
df_verify['SW'] = df_verify['SW'].astype(int)
all_sw = sorted(df_verify[df_verify['SW'] < 40]['SW'].unique().tolist())
print(f'\n=== SW disponibles en CSV ahora: {all_sw}')
for sw in [19, 20, 21, 22]:
    cnt = (df_verify['SW'] == sw).sum()
    print(f'  SW {sw}: {cnt:,} filas')
print('\nDone. Ahora ejecuta: .sw_venv\\Scripts\\python.exe build_sw_data.py')
