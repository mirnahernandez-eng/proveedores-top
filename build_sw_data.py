"""
build_sw_data.py
Genera sw_data.json con datos agregados por Semana Walmart (SW)
misma estructura que dashboard_data.json pero por SW.
"""
import json, os, re, unicodedata, sys
import openpyxl, pandas as pd
from sw_calendar import SW_MES_MAP, SW_DATES, sw_range_label  # fuente única de verdad

BASE = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP'

# Semana maxima a mostrar (inclusive). SWs > MAX_SW y < 48 se excluyen.
# Actualizar este valor al inicio de cada nueva semana.
MAX_SW = 23

# ── helpers ───────────────────────────────────────────────────────────────────
def _ascii(s):
    return unicodedata.normalize('NFKD', str(s)).encode('ascii','ignore').decode().lower().strip()

_CEDIS_MAP = [
    ('chihu',     'CUU'), ('culia',     'CLN'), ('mexic',     'MXL'),
    ('monte',     'MTY'), ('cuaut',     'CUAU'),('santa bar', 'STB'),
    ('chalco',    'CHL'), ('guada',     'GDL'), ('merid',     'MER'),
    ('villaher',  'VHSA'),('san mart',  'SMO'),
]
def get_cedis(loc):
    c = _ascii(loc)
    for pre, code in _CEDIS_MAP:
        if c.startswith(pre): return code
    return loc[:4].upper()

def get_cat(nombre):
    n = _ascii(str(nombre or ''))
    if 'perecedero' in n: return 'EXCLUIR'
    if 'sam' in n:        return "SAM'S Club"
    if 'bae' in n:        return 'BAE'
    if any(k in n for k in ('secos','sstk','nave')): return 'Autoservicios'
    return 'EXCLUIR'

# ── vendors del Excel ─────────────────────────────────────────────────────────
print('Leyendo Excel...')
wb   = openpyxl.load_workbook(os.path.join(BASE,'YMS TOP 15 2026.xlsx'), data_only=True)
CATS = {'Top 15 Proveedores': 'Autoservicios', 'Top 15 Sams ': "SAM'S Club"}
# BAE comparte vendors con Autoservicios

vendors_by_cat = {}
kw_by_cat      = {}
seen_cats      = set()

for sname, cat in CATS.items():
    if cat in seen_cats: continue
    seen_cats.add(cat)
    ws = wb[sname]
    vendors = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rank, _, vname = row[0], row[1], row[2]
        if not isinstance(rank, int) or not isinstance(vname, str): continue
        vendors.append(vname)
    vendors_by_cat[cat] = vendors
    kw = {}
    for v in vendors:
        toks = re.findall(r'[A-Z]{4,}', v.upper())
        for t in toks[:3]:
            if t not in ('MEXI','COMP','GRUP','CORP','COME','COMER','MERC'):
                kw[t] = v; break
    kw_by_cat[cat] = kw
    print(f'  {cat}: {len(vendors)} vendors, {len(kw)} keywords')

def match_vendor(csv_v, cat):
    if not isinstance(csv_v, str): return None
    n = _ascii(csv_v)
    # BAE usa los mismos vendors que Autoservicios
    lookup_cat = 'Autoservicios' if cat == 'BAE' else cat
    for kw, excel_v in kw_by_cat.get(lookup_cat, {}).items():
        if _ascii(kw) in n: return excel_v
    return None

# ── CSV principal + SW 21-22 extra ─────────────────────────────────────────────────────────────────
print('Leyendo CSV...')
COLS = ['SW','MES','LOCACION','NOMBRE_CEDIS','VENDOR','APPOINTMENT_NBR',
        'LLEGADA_A_TRAFICO','DURACION_DE_SERVICIO','SALIDA_DE_CD','formula_2']
df = pd.read_csv(os.path.join(BASE,'YMS_24.CSV'), usecols=COLS,
                 encoding='utf-8', encoding_errors='replace', low_memory=False)
print(f'  CSV principal: {len(df):,} rows')

# Cargar datos SW 21-22 del archivo extra de BigQuery
EXTRA_CSV = os.path.join(BASE, 'sw21_22_new.csv')
if os.path.exists(EXTRA_CSV):
    print('  Procesando sw21_22_new.csv...')
    ts_cols = ['ARRIVAL_TS','DRIVER_ARRIVAL_TS','TRAILER_OPEN_TS',
               'DOCK_DOOR_CLOSE','POD','DEPARTURE_TS']
    df_extra = pd.read_csv(EXTRA_CSV, encoding='utf-8', encoding_errors='replace',
                           low_memory=False)
    # Candado: solo tipos de cita validos
    TIPOS_OK = {'Proveedor', 'Cita Nueva'}
    if 'TIPO_CITA' in df_extra.columns:
        antes = len(df_extra)
        df_extra = df_extra[df_extra['TIPO_CITA'].isin(TIPOS_OK)].copy()
        print(f'  Filtro TIPO_CITA: {antes:,} -> {len(df_extra):,} filas')
    # Calcular métricas desde timestamps
    for c in ts_cols:
        df_extra[c] = pd.to_datetime(df_extra[c], errors='coerce', utc=True)
    def diff_min(a, b):
        d = (a - b).dt.total_seconds() / 60
        return d.where((d >= 0) & (d < 1440))
    df_extra['LLEGADA_A_TRAFICO']    = diff_min(df_extra['DRIVER_ARRIVAL_TS'], df_extra['ARRIVAL_TS'])
    df_extra['ABRIR_CORTINA']        = diff_min(df_extra['TRAILER_OPEN_TS'],   df_extra['DRIVER_ARRIVAL_TS'])
    df_extra['CERRAR_CORTINA']       = diff_min(df_extra['DOCK_DOOR_CLOSE'],   df_extra['TRAILER_OPEN_TS'])
    df_extra['PAPER_W']              = diff_min(df_extra['POD'],               df_extra['DOCK_DOOR_CLOSE'])
    df_extra['SALIDA_DE_CD']         = diff_min(df_extra['DEPARTURE_TS'],      df_extra['POD'])
    df_extra['DURACION_DE_SERVICIO'] = df_extra[['ABRIR_CORTINA','CERRAR_CORTINA','PAPER_W']].sum(axis=1, min_count=1)
    df_extra['formula_2'] = (df_extra['LLEGADA_A_TRAFICO'].fillna(0) +
                             df_extra['DURACION_DE_SERVICIO'].fillna(0) +
                             df_extra['SALIDA_DE_CD'].fillna(0)) / 60
    df_extra.loc[df_extra['formula_2'] <= 0, 'formula_2'] = float('nan')
    # Unir columnas necesarias
    df_extra = df_extra[[c for c in COLS if c in df_extra.columns]]
    for c in COLS:
        if c not in df_extra.columns:
            df_extra[c] = float('nan')
    df_extra = df_extra[COLS]
    print(f'  SW extra: {len(df_extra):,} rows | SW={sorted(df_extra["SW"].dropna().astype(int).unique().tolist())}')
    df = pd.concat([df, df_extra], ignore_index=True)
    print(f'  Total combinado: {len(df):,} rows')
else:
    print(f'  (sw21_22_new.csv no encontrado, solo datos del CSV principal)')

# Cargar SWs faltantes desde yms_2026_completo.csv (cubre hasta SW 23/24 de Julio)
COMPLETO_CSV = os.path.join(BASE, 'bigquery_results', 'yms_2026_completo.csv')
if os.path.exists(COMPLETO_CSV):
    # Solo traer SWs que NO esten ya en df (evitar duplicados)
    sw_ya_en_df = set(pd.to_numeric(df['SW'], errors='coerce').dropna().astype(int).unique())
    df_comp = pd.read_csv(COMPLETO_CSV, encoding='utf-8', encoding_errors='replace', low_memory=False)
    df_comp['SW'] = pd.to_numeric(df_comp['SW'], errors='coerce')
    # Filtrar: solo SWs nuevas y dentro del tope MAX_SW (o >= 48 para semanas de enero)
    df_comp = df_comp[
        df_comp['SW'].notna() &
        ((df_comp['SW'] >= 48) | (df_comp['SW'] <= MAX_SW)) &
        ~df_comp['SW'].isin(sw_ya_en_df)
    ].copy()
    if len(df_comp) > 0:
        TIPOS_OK = {'Proveedor', 'Cita Nueva', 'PROVEEDOR', 'CITA NUEVA'}
        if 'TIPO_CITA' in df_comp.columns:
            df_comp = df_comp[df_comp['TIPO_CITA'].str.strip().isin(TIPOS_OK)].copy()
        if 'CITAS_CORRECTAS' in df_comp.columns:
            df_comp = df_comp[pd.to_numeric(df_comp['CITAS_CORRECTAS'], errors='coerce') == 1].copy()
        # Los valores ya vienen en minutos desde BQ — calcular formula_2
        for col in ['LLEGADA_A_TRAFICO', 'DURACION_DE_SERVICIO', 'SALIDA_DE_CD']:
            if col in df_comp.columns:
                df_comp[col] = pd.to_numeric(df_comp[col], errors='coerce').fillna(0)
        df_comp['formula_2'] = (
            df_comp['LLEGADA_A_TRAFICO'].fillna(0) +
            df_comp['DURACION_DE_SERVICIO'].fillna(0) +
            df_comp['SALIDA_DE_CD'].fillna(0)
        ) / 60
        df_comp.loc[df_comp['formula_2'] <= 0, 'formula_2'] = float('nan')
        df_comp = df_comp[[c for c in COLS if c in df_comp.columns]]
        for c in COLS:
            if c not in df_comp.columns:
                df_comp[c] = float('nan')
        df_comp = df_comp[COLS]
        nuevas_sw = sorted(df_comp['SW'].dropna().astype(int).unique().tolist())
        print(f'  yms_2026_completo: {len(df_comp):,} rows nuevas | SW={nuevas_sw}')
        df = pd.concat([df, df_comp], ignore_index=True)
        print(f'  Total final: {len(df):,} rows')
    else:
        print('  yms_2026_completo: sin SWs nuevas que agregar')

for col in ['LLEGADA_A_TRAFICO','DURACION_DE_SERVICIO','SALIDA_DE_CD','formula_2']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['llegada_h'] = df['LLEGADA_A_TRAFICO']    / 60
df['recibo_h']  = df['DURACION_DE_SERVICIO'] / 60
df['salida_h']  = df['SALIDA_DE_CD']         / 60
df['total_h']   = df['formula_2']
df['cedis_code']= df['LOCACION'].apply(get_cedis)
df['categoria'] = df['NOMBRE_CEDIS'].apply(get_cat)
df['SW']        = pd.to_numeric(df['SW'], errors='coerce')
df = df.dropna(subset=['SW'])
df['SW']        = df['SW'].astype(int)

df = df[df['categoria'].isin(['Autoservicios', 'BAE', "SAM'S Club"]) & (df['total_h'] > 0)].copy()
print(f'  Tras filtro SECOS: {len(df):,} (auto={len(df[df.categoria=="Autoservicios"]):,}, bae={len(df[df.categoria=="BAE"]):,}, sams={len(df[df.categoria=="SAM\'S Club"]):,})')

df['excel_vendor'] = df.apply(lambda r: match_vendor(r['VENDOR'], r['categoria']), axis=1)
df = df[df['excel_vendor'].notna()].copy()
print(f'  Matched vendors: {len(df):,}')

# ── SW x mes mapping ──────────────────────────────────────────────────────────
sw_mes = df.groupby('SW')['MES'].agg(lambda x: x.mode()[0]).to_dict()
# Excluir SWs futuras: quitar cualquier SW > MAX_SW que no sea semana de cierre de año (48+)
df = df[((df['SW'] >= 48) | (df['SW'] <= MAX_SW))].copy()

sw_list_raw = sorted(df['SW'].unique().tolist())

# Ordenar: poner 48-52 antes que 1
sw_ordered = [sw for sw in sw_list_raw if sw >= 48] + [sw for sw in sw_list_raw if sw < 48]
print(f'  SW disponibles: {sw_ordered}')

# ── Agregacion por SW + cedis + vendor ────────────────────────────────────────
print('Agregando...')
grp = df.groupby(['SW','cedis_code','categoria','excel_vendor'], as_index=False).agg(
    llegada_h=('llegada_h','mean'),
    recibo_h =('recibo_h', 'mean'),
    salida_h =('salida_h', 'mean'),
    total_h  =('total_h',  'mean'),
    citas    =('APPOINTMENT_NBR','count'),
)
for col in ['llegada_h','recibo_h','salida_h','total_h']:
    grp[col] = grp[col].round(1)

# Estructura: cat -> vendor -> cedis -> SW -> {l,r,s,t}
CEDIS_AUTO = ['CUU','CLN','MXL','MTY','CUAU','STB','CHL','GDL','MER','VHSA']
CEDIS_BAE  = ['CLN','MTY','STB','CHL','GDL','MER','VHSA']
CEDIS_SAMS = ['CUU','CLN','MTY','SMO','CHL','GDL','MER','VHSA']

def build_chart_data(cat, cedis_list, vendors):
    data = {}
    for v in vendors:
        data[v] = {}
        for loc in cedis_list:
            sub = grp[(grp['categoria']==cat) & (grp['excel_vendor']==v) & (grp['cedis_code']==loc)]
            loc_d = {}
            for _, row in sub.iterrows():
                sw_key = f"SW{int(row['SW'])}"
                loc_d[sw_key] = {'l':row['llegada_h'],'r':row['recibo_h'],
                                 's':row['salida_h'], 't':row['total_h'],
                                 'c':int(row['citas'])}
            if loc_d:
                data[v][loc] = loc_d
        # national avg per SW
        sub_v = grp[(grp['categoria']==cat) & (grp['excel_vendor']==v)]
        nat = sub_v.groupby('SW', as_index=False).agg(
            l=('llegada_h','mean'),r=('recibo_h','mean'),
            s=('salida_h','mean'), t=('total_h','mean'))
        nat_d = {}
        for _, row in nat.iterrows():
            nat_d[f"SW{int(row['SW'])}"] = {
                'l':round(row['l'],1),'r':round(row['r'],1),
                's':round(row['s'],1),'t':round(row['t'],1)}
        data[v]['2026'] = nat_d

    # __total__
    tot_sub = grp[grp['categoria']==cat]
    data['__total__'] = {}
    for loc in cedis_list:
        sub = tot_sub[tot_sub['cedis_code']==loc]
        loc_d = {}
        for sw, g in sub.groupby('SW'):
            loc_d[f"SW{int(sw)}"] = {
                'l':round(g['llegada_h'].mean(),1),'r':round(g['recibo_h'].mean(),1),
                's':round(g['salida_h'].mean(),1), 't':round(g['total_h'].mean(),1),
                'c':int(g['citas'].sum())}
        if loc_d: data['__total__'][loc] = loc_d
    nat = tot_sub.groupby('SW', as_index=False).agg(
        l=('llegada_h','mean'),r=('recibo_h','mean'),
        s=('salida_h','mean'),t=('total_h','mean'))
    data['__total__']['2026'] = {f"SW{int(r['SW'])}":{'l':round(r['l'],1),'r':round(r['r'],1),
        's':round(r['s'],1),'t':round(r['t'],1)} for _,r in nat.iterrows()}
    return data


def build_table_data(cat, cedis_list, vendors):
    """Format: data[vendor][SW_key][cedis] = total_h (matches buildTable expected structure)"""
    data = {}
    for v in vendors:
        data[v] = {}
        sub_v = grp[(grp['categoria']==cat) & (grp['excel_vendor']==v)]
        for sw, g in sub_v.groupby('SW'):
            sw_key = f"SW{int(sw)}"
            data[v][sw_key] = {}
            for loc in cedis_list:
                row = g[g['cedis_code']==loc]
                if not row.empty:
                    data[v][sw_key][loc] = round(row['total_h'].mean(), 1)
            # national avg
            data[v][sw_key]['2026'] = round(g['total_h'].mean(), 1)

    # __total__
    tot_sub = grp[grp['categoria']==cat]
    data['__total__'] = {}
    for sw, g in tot_sub.groupby('SW'):
        sw_key = f"SW{int(sw)}"
        data['__total__'][sw_key] = {}
        for loc in cedis_list:
            row = g[g['cedis_code']==loc]
            if not row.empty:
                data['__total__'][sw_key][loc] = round(row['total_h'].mean(), 1)
        data['__total__'][sw_key]['2026'] = round(g['total_h'].mean(), 1)
    return data

auto_data      = build_chart_data('Autoservicios', CEDIS_AUTO, vendors_by_cat['Autoservicios'])
bae_data       = build_chart_data('BAE',            CEDIS_BAE,  vendors_by_cat['Autoservicios'])
sams_data      = build_chart_data("SAM'S Club",     CEDIS_SAMS, vendors_by_cat["SAM'S Club"])
auto_tbl_data  = build_table_data('Autoservicios',  CEDIS_AUTO, vendors_by_cat['Autoservicios'])
bae_tbl_data   = build_table_data('BAE',            CEDIS_BAE,  vendors_by_cat['Autoservicios'])
sams_tbl_data  = build_table_data("SAM'S Club",     CEDIS_SAMS, vendors_by_cat["SAM'S Club"])

# Combined auto+bae usando categoria 'Autoservicios' OR 'BAE'
def build_chart_data_multi(cats, cedis_list, vendors):
    """Agrega multiples categorias como si fueran una sola."""
    grp_m = grp[grp['categoria'].isin(cats)]
    data = {}
    for v in vendors:
        data[v] = {}
        for loc in cedis_list:
            sub = grp_m[(grp_m['excel_vendor']==v) & (grp_m['cedis_code']==loc)]
            loc_d = {}
            for sw, g in sub.groupby('SW'):
                sw_key = f"SW{int(sw)}"
                loc_d[sw_key] = {'l':round(g['llegada_h'].mean(),1),'r':round(g['recibo_h'].mean(),1),
                                 's':round(g['salida_h'].mean(),1),'t':round(g['total_h'].mean(),1),
                                 'c':int(g['citas'].sum())}
            if loc_d: data[v][loc] = loc_d
        sub_v = grp_m[grp_m['excel_vendor']==v]
        nat_d = {}
        for sw, g in sub_v.groupby('SW'):
            nat_d[f"SW{int(sw)}"] = {'l':round(g['llegada_h'].mean(),1),'r':round(g['recibo_h'].mean(),1),
                                      's':round(g['salida_h'].mean(),1),'t':round(g['total_h'].mean(),1)}
        data[v]['2026'] = nat_d
    tot_sub = grp_m
    data['__total__'] = {}
    for loc in cedis_list:
        sub = tot_sub[tot_sub['cedis_code']==loc]
        loc_d = {}
        for sw, g in sub.groupby('SW'):
            loc_d[f"SW{int(sw)}"] = {'l':round(g['llegada_h'].mean(),1),'r':round(g['recibo_h'].mean(),1),
                                      's':round(g['salida_h'].mean(),1),'t':round(g['total_h'].mean(),1),'c':int(g['citas'].sum())}
        if loc_d: data['__total__'][loc] = loc_d
    nat = tot_sub.groupby('SW', as_index=False).agg(l=('llegada_h','mean'),r=('recibo_h','mean'),
        s=('salida_h','mean'),t=('total_h','mean'))
    data['__total__']['2026'] = {f"SW{int(r['SW'])}":{'l':round(r['l'],1),'r':round(r['r'],1),'s':round(r['s'],1),'t':round(r['t'],1)} for _,r in nat.iterrows()}
    return data

def build_table_data_multi(cats, cedis_list, vendors):
    grp_m = grp[grp['categoria'].isin(cats)]
    data = {}
    for v in vendors:
        data[v] = {}
        sub_v = grp_m[grp_m['excel_vendor']==v]
        for sw, g in sub_v.groupby('SW'):
            sw_key = f"SW{int(sw)}"
            data[v][sw_key] = {}
            for loc in cedis_list:
                row = g[g['cedis_code']==loc]
                if not row.empty: data[v][sw_key][loc] = round(row['total_h'].mean(), 1)
            data[v][sw_key]['2026'] = round(g['total_h'].mean(), 1)
    tot_sub = grp_m
    data['__total__'] = {}
    for sw, g in tot_sub.groupby('SW'):
        sw_key = f"SW{int(sw)}"
        data['__total__'][sw_key] = {}
        for loc in cedis_list:
            row = g[g['cedis_code']==loc]
            if not row.empty: data['__total__'][sw_key][loc] = round(row['total_h'].mean(), 1)
        data['__total__'][sw_key]['2026'] = round(g['total_h'].mean(), 1)
    return data

auto_bae_data     = build_chart_data_multi(['Autoservicios','BAE'], CEDIS_AUTO, vendors_by_cat['Autoservicios'])
auto_bae_tbl_data = build_table_data_multi(['Autoservicios','BAE'], CEDIS_AUTO, vendors_by_cat['Autoservicios'])

# ── Output ────────────────────────────────────────────────────────────────────
# sw_mes_map y sw_dates vienen de sw_calendar (fuente única de verdad)
out = {
    'sw_list':    sw_ordered,
    'sw_mes_map': {f"SW{sw}": SW_MES_MAP.get(sw, sw_mes.get(sw, '?')) for sw in sw_ordered},
    'sw_dates':   {
        f"SW{sw}": {"inicio": SW_DATES[sw]["inicio"],
                    "fin":    SW_DATES[sw]["fin"],
                    "label":  sw_range_label(sw)}
        for sw in sw_ordered if sw in SW_DATES
    },
    'auto':       auto_data,       # solo Autoservicios (secos)
    'bae':        bae_data,        # solo BAE
    'auto_bae':   auto_bae_data,   # auto + bae combinados
    'sams':       sams_data,       # SAM'S Club
    'tbl_auto':   auto_tbl_data,
    'tbl_bae':    bae_tbl_data,
    'tbl_auto_bae': auto_bae_tbl_data,
    'tbl_sams':   sams_tbl_data,
}
out_path = os.path.join(BASE, 'sw_data.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',',':'))

size_kb = os.path.getsize(out_path) // 1024
print(f'Listo! sw_data.json -> {size_kb}KB')
print('SW list:', sw_ordered)
