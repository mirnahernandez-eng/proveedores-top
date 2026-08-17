"""
gen_matrix_perec.py
Genera matrix_PEREC.json (Top 5 proveedores de Perecederos) desde
vendor_cedis_mes_PEREC.csv. Companero de gen_matrix_FINAL.py -- misma
logica de promedio ponderado por citas, mismo enfoque de meses dinamicos
(solo incluye meses con datos reales, para no romperse cada mes nuevo).
"""
import csv, json, os
from collections import defaultdict

BASE = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP\bigquery_results'
CSV  = os.path.join(BASE, 'vendor_cedis_mes_PEREC.csv')
OUT  = os.path.join(BASE, 'matrix_PEREC.json')

_MES_ORDER_FULL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
                    'Septiembre','Octubre','Noviembre','Diciembre']
MESES_PROM = ['Enero','Febrero','Marzo','Abril','Mayo','Junio']

DISPLAY_ORDER = [
    'DRISCOLL S OPERACIONES SA C',
    'PILGRIMS PRIDE S DE RL DE C',
    'LANDEROS PALAZUELOS EDUARDO',
    'MJ INTERNATIONAL MARKETIN S',
    'FRUTAS Y LEGUMBRES ALPHA SA CV',
]

# Codigos de CEDIS de Perecederos por region (ver build_tablero.py CD_CHANNEL)
AUTO_LOCS = {'MTY': [7498], 'SMO': [7466], 'CHL': [8801], 'GDL': [7495], 'VHSA': [4659]}
SAMS_LOCS = {'MTY': [7502], 'SMO': [4996], 'GDL': [6239], 'VHSA': [6151]}
AUTO_ALL = [c for cl in AUTO_LOCS.values() for c in cl]
SAMS_ALL = [c for cl in SAMS_LOCS.values() for c in cl]

# raw[vendor][cedis][mes] = {'c':n, 'ws':suma_los}
raw = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {'c': 0, 'ws': 0.0})))

if os.path.exists(CSV):
    with open(CSV, encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            vendor = (r.get('VENDOR') or '').strip()
            if vendor not in DISPLAY_ORDER:
                continue
            cedis = int(r['CEDIS'])
            mes   = r['MES']
            c     = int(r['TOTAL_CITAS'] or 0)
            if r.get('LOS_SUM'):
                los_ws = float(r['LOS_SUM'])
            else:
                los_ws = float(r['TOTAL_HRS'] or 0) * c
            raw[vendor][cedis][mes]['c']  += c
            raw[vendor][cedis][mes]['ws'] += los_ws

# Meses dinamicos: solo los que tienen al menos una cita real
_meses_con_datos = {
    mes for vendor_d in raw.values()
    for cedis_d in vendor_d.values()
    for mes, vals in cedis_d.items()
    if vals['c'] > 0
}
MESES_ALL = [m for m in _MES_ORDER_FULL if m in _meses_con_datos]


def wavg(vendors, cedis_list, mes_list):
    sw, sc = 0.0, 0
    for v in vendors:
        for c in cedis_list:
            for m in mes_list:
                r = raw[v][c].get(m)
                if r and r['c']:
                    sw += r['ws']
                    sc += r['c']
    return round(sw / sc, 2) if sc else None


def build_matrix(locs_dict, all_cedis):
    m = {}
    for v in DISPLAY_ORDER:
        m[v] = {'prom': {}, 'ytd': {}}
        for loc, cl in locs_dict.items():
            m[v]['prom'][loc] = wavg([v], cl, MESES_PROM) if cl else None
            m[v]['ytd'][loc]  = wavg([v], cl, MESES_ALL)  if cl else None
        m[v]['prom']['2026'] = wavg([v], all_cedis, MESES_PROM)
        m[v]['ytd']['2026']  = wavg([v], all_cedis, MESES_ALL)
        for mes in MESES_ALL:
            m[v][mes] = {}
            for loc, cl in locs_dict.items():
                m[v][mes][loc] = wavg([v], cl, [mes]) if cl else None
            m[v][mes]['2026'] = wavg([v], all_cedis, [mes])
    m['__total__'] = {'prom': {}, 'ytd': {}}
    for loc, cl in locs_dict.items():
        m['__total__']['prom'][loc] = wavg(DISPLAY_ORDER, cl, MESES_PROM) if cl else None
        m['__total__']['ytd'][loc]  = wavg(DISPLAY_ORDER, cl, MESES_ALL)  if cl else None
    m['__total__']['prom']['2026'] = wavg(DISPLAY_ORDER, all_cedis, MESES_PROM)
    m['__total__']['ytd']['2026']  = wavg(DISPLAY_ORDER, all_cedis, MESES_ALL)
    for mes in MESES_ALL:
        m['__total__'][mes] = {}
        for loc, cl in locs_dict.items():
            m['__total__'][mes][loc] = wavg(DISPLAY_ORDER, cl, [mes]) if cl else None
        m['__total__'][mes]['2026'] = wavg(DISPLAY_ORDER, all_cedis, [mes])
    return m


out = {
    'display_order': DISPLAY_ORDER,
    'auto': build_matrix(AUTO_LOCS, AUTO_ALL),
    'sams': build_matrix(SAMS_LOCS, SAMS_ALL),
}
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

print('OK - matrix_PEREC.json')
print('  Meses incluidos:', MESES_ALL)
for v in DISPLAY_ORDER:
    print(f"  {v[:30]:30s} AUTO_YTD:{wavg([v], AUTO_ALL, MESES_ALL)}  SAMS_YTD:{wavg([v], SAMS_ALL, MESES_ALL)}")
