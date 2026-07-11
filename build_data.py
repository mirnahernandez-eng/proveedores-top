"""
Build LOS dashboard data — YMS_24.CSV + YMS TOP 15 2026.xlsx
Requirements:
  - SECOS only (exclude PERECEDEROS)
  - No outlier cap on LOS (include high exit times)
  - Weighted averages by citas
  - Matrix: all months side-by-side + YTD column
  - 2026 Prom extracted from Excel
"""
import json
import os
import re
import unicodedata

import openpyxl
import pandas as pd

BASE = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP'

# ─── Helpers ─────────────────────────────────────────────────────────────────
def _ascii(s) -> str:
    return unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower().strip()

_CEDIS_PREFIX_MAP = [
    ('chihu',     ('CUU',  'Chihuahua')),
    ('culia',     ('CLN',  'Culiac\u00e1n')),
    ('mexic',     ('MXL',  'Mexicali')),
    ('monte',     ('MTY',  'Monterrey')),
    ('cuaut',     ('CUAU', 'Cuautitl\u00e1n')),
    ('santa bar', ('STB',  'Santa Barbara')),
    ('chalco',    ('CHL',  'Chalco')),
    ('guada',     ('GDL',  'Guadalajara')),
    ('merid',     ('MER',  'M\u00e9rida')),
    ('villaher',  ('VHSA', 'Villahermosa')),
    ('san mart',  ('SMO',  'San Mart\u00edn Obispo')),
]

def get_cedis(loc: str):
    clean = _ascii(loc)
    for prefix, val in _CEDIS_PREFIX_MAP:
        if clean.startswith(prefix):
            return val
    return (loc[:4].upper(), loc)

def get_categoria(nombre) -> str:
    """Classify as Autoservicios or SAM'S Club (SECOS only — skip PERECEDEROS)."""
    n = _ascii(str(nombre or ''))
    if 'perecedero' in n:
        return 'EXCLUIR'
    if 'sam' in n:
        return "SAM'S Club"
    if any(k in n for k in ('secos', 'sstk', 'bae', 'nave')):
        return 'Autoservicios'
    return 'EXCLUIR'

MES_ORDER = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
             'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

SHORT_NAMES = [
    ('KIMBERLY',     'KIMBERLY'), ('ENBOTELLAD', 'NIAGARA'),
    ('NIAGARA',      'NIAGARA'),  ('JUGOS',      'JUGOS'),
    ('SANTA CLARA',  'STA CLARA'),('PROCTER',    'PROCTER'),
    ('MARCAS NESTLE','NESTLE'),   ('NESTLE',     'NESTLE'),
    ('COLGATE',      'COLGATE'),  ('COMERC PEPSICO', 'PEPSICO'),
    ('PEPSICO',      'PEPSICO'),  ('BONAFONT',   'BONAFONT'),
    ('BEBIDAS PURIF','BONAFONT'), ('UNILEVER',   'UNILEVER'),
    ('HERDEZ',       'HERDEZ'),   ('CERVEZA CANAL', 'CANAL MO'),
    ('CANAL MO',     'CANAL MO'), ('FRABEL',     'FRABEL'),
    ('MONDELEZ',     'MONDELEZ'), ('KELLOGG',    'KELLOGG'),
]

def vendor_short(name) -> str:
    if not isinstance(name, str):
        return 'N/A'
    n = name.upper()
    for k, v in SHORT_NAMES:
        if k in n:
            return v
    return name.split()[0][:12]

# ─── 1. Read Excel — vendors, objectives (with 2026 Prom) ────────────────────
print('Reading Excel...')
wb = openpyxl.load_workbook(
    os.path.join(BASE, 'YMS TOP 15 2026.xlsx'),
    read_only=True, data_only=True
)

SHEET_INFO = {
    'Top 15 Proveedores': {
        'categoria': 'Autoservicios',
        'cedis_order': ['CUU', 'CLN', 'MXL', 'MTY', 'CUAU', 'STB', 'CHL', 'GDL', 'MER', 'VHSA'],
    },
    'Top 15 Sams ': {
        'categoria': "SAM'S Club",
        'cedis_order': ['CUU', 'CLN', 'MTY', 'SMO', 'CHL', 'GDL', 'MER', 'VHSA'],
    },
}

vendors_by_cat: dict  = {}
objectives_raw: dict  = {}   # cat -> vendor -> cedis -> float
obj_prom_2026: dict   = {}   # cat -> vendor -> float  (national prom)

for sheet_name, info in SHEET_INFO.items():
    cat         = info['categoria']
    cedis_order = info['cedis_order']
    ws          = wb[sheet_name]
    all_rows    = list(ws.iter_rows(values_only=True))

    # Find first header row with 'Proveedores TOP'
    header_row_idx = None
    for ridx, row in enumerate(all_rows[:15]):
        if row[2] == 'Proveedores TOP':
            header_row_idx = ridx
            break
    if header_row_idx is None:
        print(f'  WARNING: header not found in {sheet_name}')
        continue

    # Find objective section start col
    obj_start_col = None
    for ridx, row in enumerate(all_rows[:header_row_idx]):
        for cidx, val in enumerate(row):
            if isinstance(val, str) and 'Objetivo' in val:
                obj_start_col = cidx
                break
        if obj_start_col is not None:
            break

    # Build col_map: only first occurrence of each CEDIS in objectives section
    header    = all_rows[header_row_idx]
    col_map: dict[int, str] = {}
    prom_col  = None   # column right after last CEDIS = national 2026 prom
    if obj_start_col is not None:
        for cidx, val in enumerate(header):
            if cidx < obj_start_col:
                continue
            if val in cedis_order and val not in col_map.values():
                col_map[cidx] = val
        # The column AFTER the last objective CEDIS col is the 2026 Prom
        if col_map:
            last_obj_col = max(col_map.keys())
            # scan forward for next numeric / year value
            for cidx in range(last_obj_col + 1, last_obj_col + 5):
                v = header[cidx] if cidx < len(header) else None
                if v is None or isinstance(v, (int, float)):
                    prom_col = cidx
                    break

    print(f'  {sheet_name}: obj_col={obj_start_col}, cedis_cols={col_map}, prom_col={prom_col}')

    vendors: list[dict] = []
    obj_data: dict      = {}
    prom_data: dict     = {}

    _seen_rank1 = False
    for row in all_rows[header_row_idx + 1:]:
        rank        = row[0]
        vendor_name = row[2]
        if not isinstance(rank, int) or not isinstance(vendor_name, str):
            continue
        if rank == 1 and _seen_rank1:
            break
        if rank == 1:
            _seen_rank1 = True
        if rank > 16:
            break

        short = vendor_short(vendor_name)
        vendors.append({'excel_name': vendor_name, 'short': short, 'rank': rank})

        # Objectives per CEDIS
        obj_data[vendor_name] = {}
        for cidx, cedis_code in col_map.items():
            val = row[cidx] if cidx < len(row) else None
            if isinstance(val, (int, float)):
                obj_data[vendor_name][cedis_code] = round(float(val), 1)

        # 2026 national prom
        if prom_col is not None and prom_col < len(row):
            val = row[prom_col]
            if isinstance(val, (int, float)):
                prom_data[vendor_name] = round(float(val), 1)

    vendors_by_cat[cat] = vendors
    objectives_raw[cat] = obj_data
    obj_prom_2026[cat]  = prom_data
    print(f'    Vendors ({len(vendors)}): {[v["short"] for v in vendors]}')
    print(f'    Proms loaded: {len(prom_data)}')

# ─── 2. Build keyword map for CSV matching ────────────────────────────────────
KEYWORDS_BY_CAT: dict = {}
for cat, vendors in vendors_by_cat.items():
    kw_map: dict[str, str] = {}
    for v in vendors:
        n = v['excel_name'].upper()
        tokens = re.findall(r'[A-Z]{4,}', n)
        for tok in tokens[:3]:
            if tok not in ('MEXI', 'COMP', 'GRUP', 'CORP', 'COME', 'COMER', 'MERC'):
                kw_map[tok] = v['excel_name']
                break
    KEYWORDS_BY_CAT[cat] = kw_map

def match_vendor(csv_vendor: str, kw_map: dict):
    if not isinstance(csv_vendor, str):
        return None
    n = _ascii(csv_vendor)
    for kw, excel_name in kw_map.items():
        if _ascii(kw) in n:
            return excel_name
    return None

# ─── 3. Read & prepare CSV ───────────────────────────────────────────────────
print('\nReading CSV (120MB)...')
df = pd.read_csv(
    os.path.join(BASE, 'YMS_24.CSV'),
    encoding='utf-8', encoding_errors='replace', low_memory=False,
)
print(f'  Loaded {len(df):,} rows')

# Numeric
for col in ['LLEGADA_A_TRAFICO', 'DURACION_DE_SERVICIO', 'SALIDA_DE_CD', 'formula_2']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['llegada_h'] = df['LLEGADA_A_TRAFICO']   / 60
df['recibo_h']  = df['DURACION_DE_SERVICIO'] / 60
df['salida_h']  = df['SALIDA_DE_CD']         / 60
df['total_h']   = df['formula_2']

# CEDIS
cedis_pairs     = df['LOCACION'].apply(get_cedis)
df['cedis_code'] = cedis_pairs.apply(lambda x: x[0])
df['cedis_name'] = cedis_pairs.apply(lambda x: x[1])

# Category — SECOS only, exclude PERECEDEROS
df['categoria'] = df['NOMBRE_CEDIS'].apply(get_categoria)
df = df[df['categoria'].isin(['Autoservicios', "SAM'S Club"])].copy()

# No hard cap — only remove physically impossible negatives
df = df[df['total_h'] > 0].copy()
print(f'  After SECOS filter: {len(df):,} rows')
print(f'  NOMBRE_CEDIS unique: {df["NOMBRE_CEDIS"].unique().tolist()}')

# Match vendors
def match_row(row):
    return match_vendor(row['VENDOR'], KEYWORDS_BY_CAT.get(row['categoria'], {}))

print('  Matching vendors...')
df['excel_vendor'] = df.apply(match_row, axis=1)
df_top = df[df['excel_vendor'].notna()].copy()

short_map = {}
rank_map  = {}
for cat, vendors in vendors_by_cat.items():
    for v in vendors:
        short_map[(cat, v['excel_name'])] = v['short']
        rank_map[ (cat, v['excel_name'])] = v['rank']

df_top['vendor_short'] = df_top.apply(
    lambda r: short_map.get((r['categoria'], r['excel_vendor']), r['excel_vendor'][:10]), axis=1
)
df_top['vendor_rank'] = df_top.apply(
    lambda r: rank_map.get((r['categoria'], r['excel_vendor']), 99), axis=1
)
print(f'  Matched: {len(df_top):,} rows')

# ─── 4. Weighted aggregation helpers ─────────────────────────────────────────
def w_avg(series_vals, series_citas):
    """Weighted average: sum(val*citas)/sum(citas)."""
    total_citas = series_citas.sum()
    if total_citas == 0:
        return None
    return round((series_vals * series_citas).sum() / total_citas, 1)

# ─── 5. Per-chart data (CEDIS+mes view, for the bar charts) ──────────────────
grp = df_top.groupby(
    ['cedis_code', 'cedis_name', 'MES', 'categoria', 'excel_vendor', 'vendor_short', 'vendor_rank'],
    as_index=False
).agg(
    llegada_h=('llegada_h', 'mean'),
    recibo_h =('recibo_h',  'mean'),
    salida_h =('salida_h',  'mean'),
    total_h  =('total_h',   'mean'),
    citas    =('APPOINTMENT_NBR', 'count'),
)
for col in ['llegada_h', 'recibo_h', 'salida_h', 'total_h']:
    grp[col] = grp[col].round(1)

hist = df_top.groupby(
    ['cedis_code', 'categoria', 'excel_vendor', 'vendor_short'], as_index=False
).agg(hist_total=('total_h', 'mean'), hist_citas=('APPOINTMENT_NBR', 'count'))
hist['hist_total'] = hist['hist_total'].round(1)

cedis_list   = (
    df_top[['cedis_code', 'cedis_name']].drop_duplicates().sort_values('cedis_code').to_dict('records')
)
mes_available = [m for m in MES_ORDER if m in df_top['MES'].unique()]

charts: dict = {}
for cedis_code in df_top['cedis_code'].unique():
    charts[cedis_code] = {}
    for mes in mes_available:
        charts[cedis_code][mes] = {}
        for cat in ['Autoservicios', "SAM'S Club"]:
            subset   = grp[(grp['cedis_code'] == cedis_code) & (grp['MES'] == mes) & (grp['categoria'] == cat)].copy()
            hist_sub = hist[(hist['cedis_code'] == cedis_code) & (hist['categoria'] == cat)].set_index('excel_vendor')['hist_total'].to_dict()
            obj_data = objectives_raw.get(cat, {})
            subset   = subset.sort_values('total_h', ascending=False)
            out = []
            for _, row in subset.iterrows():
                ev  = row['excel_vendor']
                out.append({
                    'vendor': row['vendor_short'], 'vendor_full': ev,
                    'llegada': row['llegada_h'], 'recibo': row['recibo_h'],
                    'salida': row['salida_h'],  'total': row['total_h'],
                    'citas': int(row['citas']),
                    'objetivo': obj_data.get(ev, {}).get(cedis_code),
                    'hist_prom': hist_sub.get(ev),
                })
            charts[cedis_code][mes][cat] = out

# ─── 6. Monthly trend ────────────────────────────────────────────────────────
monthly_trend: dict = {}
for cat in ['Autoservicios', "SAM'S Club"]:
    monthly_trend[cat] = []
    for mes in mes_available:
        sub = df_top[(df_top['MES'] == mes) & (df_top['categoria'] == cat)]
        avg = round(sub['total_h'].mean(), 1) if not sub.empty else None
        monthly_trend[cat].append({'mes': mes, 'avg_total': avg, 'citas': len(sub)})

# ─── 7. MATRIX — all months + YTD, weighted prom ─────────────────────────────
REGIONS = {
    'Autoservicios': [
        {'name': 'NORTE',  'cedis': ['CUU', 'CLN', 'MXL', 'MTY']},
        {'name': 'CENTRO', 'cedis': ['CUAU', 'STB']},
        {'name': 'SUR',    'cedis': ['CHL', 'GDL', 'MER', 'VHSA']},
    ],
    "SAM'S Club": [
        {'name': 'NORTE',  'cedis': ['CUU', 'CLN', 'MTY']},
        {'name': 'CENTRO', 'cedis': ['SMO']},
        {'name': 'SUR',    'cedis': ['CHL', 'GDL', 'MER', 'VHSA']},
    ],
}

# Aggregate: categoria + excel_vendor + cedis_code + mes -> weighted total_h, citas
mat_grp = df_top.groupby(
    ['categoria', 'MES', 'excel_vendor', 'cedis_code'],
    as_index=False
).agg(
    total_sum=('total_h', 'sum'),  # sum of all LOS hours
    citas=('APPOINTMENT_NBR', 'count'),
)
# Weighted average = total_sum / citas (since each row = 1 appointment with its LOS)
mat_grp['total_w'] = (mat_grp['total_sum'] / mat_grp['citas']).round(1)

matrix: dict = {}
for cat in ['Autoservicios', "SAM'S Club"]:
    cat_vendors = vendors_by_cat.get(cat, [])
    all_cedis   = [c for r in REGIONS[cat] for c in r['cedis']]

    matrix[cat] = {
        'regions':    REGIONS[cat],
        'vendors':    [{'name': v['excel_name'], 'short': v['short'], 'rank': v['rank']}
                       for v in cat_vendors],
        'objectives': objectives_raw.get(cat, {}),
        'obj_prom':   obj_prom_2026.get(cat, {}),   # national 2026 prom per vendor
        'meses':      {},
    }

    sub_cat = mat_grp[mat_grp['categoria'] == cat]

    # Per month
    for mes in mes_available:
        sub_mes = sub_cat[sub_cat['MES'] == mes]
        mes_data: dict = {}

        for v in cat_vendors:
            vname    = v['excel_name']
            sub_v    = sub_mes[sub_mes['excel_vendor'] == vname]
            row_vals = {r['cedis_code']: r['total_w'] for _, r in sub_v.iterrows()}
            row_citas= {r['cedis_code']: r['citas']   for _, r in sub_v.iterrows()}

            # Weighted prom across CEDIS for this month
            total_los   = sum(row_vals[c] * row_citas[c] for c in all_cedis if c in row_vals)
            total_citas = sum(row_citas[c] for c in all_cedis if c in row_citas)
            prom = round(total_los / total_citas, 1) if total_citas > 0 else None

            mes_data[vname] = {'vals': row_vals, 'citas': row_citas, 'prom': prom}

        # Total row (weighted across all vendors)
        tot_vals: dict = {}
        tot_citas: dict = {}
        for c in all_cedis:
            sub_c = sub_mes[sub_mes['cedis_code'] == c]
            if not sub_c.empty:
                tot_citas[c] = int(sub_c['citas'].sum())
                tot_vals[c]  = round(sub_c['total_sum'].sum() / tot_citas[c], 1)
        tot_total = sum(tot_vals[c]*tot_citas[c] for c in all_cedis if c in tot_vals)
        tot_c_sum = sum(tot_citas[c] for c in all_cedis if c in tot_citas)
        mes_data['__total__'] = {
            'vals': tot_vals, 'citas': tot_citas,
            'prom': round(tot_total / tot_c_sum, 1) if tot_c_sum else None
        }

        matrix[cat]['meses'][mes] = mes_data

    # YTD (all months combined)
    sub_ytd = sub_cat
    ytd_data: dict = {}
    for v in cat_vendors:
        vname   = v['excel_name']
        sub_v   = sub_ytd[sub_ytd['excel_vendor'] == vname]
        row_vals = {}
        row_citas= {}
        for c in all_cedis:
            sub_c = sub_v[sub_v['cedis_code'] == c]
            if not sub_c.empty:
                tc = sub_c['citas'].sum()
                ts = sub_c['total_sum'].sum()
                row_vals[c]  = round(ts / tc, 1)
                row_citas[c] = int(tc)
        total_los   = sum(row_vals[c] * row_citas[c] for c in all_cedis if c in row_vals)
        total_citas = sum(row_citas[c] for c in all_cedis if c in row_citas)
        prom = round(total_los / total_citas, 1) if total_citas > 0 else None
        ytd_data[vname] = {'vals': row_vals, 'citas': row_citas, 'prom': prom}

    # YTD total row
    tot_vals = {}; tot_citas = {}
    for c in all_cedis:
        sub_c = sub_ytd[sub_ytd['cedis_code'] == c]
        if not sub_c.empty:
            tot_citas[c] = int(sub_c['citas'].sum())
            tot_vals[c]  = round(sub_c['total_sum'].sum() / tot_citas[c], 1)
    tot_total = sum(tot_vals[c]*tot_citas[c] for c in all_cedis if c in tot_vals)
    tot_c_sum = sum(tot_citas[c] for c in all_cedis if c in tot_citas)
    ytd_data['__total__'] = {
        'vals': tot_vals, 'citas': tot_citas,
        'prom': round(tot_total / tot_c_sum, 1) if tot_c_sum else None
    }
    matrix[cat]['ytd'] = ytd_data

print('Matrix built.')

# ─── 8. Save ──────────────────────────────────────────────────────────────────
output = {
    'cedis_list':     cedis_list,
    'meses':          mes_available,
    'charts':         charts,
    'monthly_trend':  monthly_trend,
    'matrix':         matrix,
}

out_path = os.path.join(BASE, 'dashboard_data.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False)

size_kb = os.path.getsize(out_path) // 1024
print(f'\nSaved: {out_path}  ({size_kb} KB)')
