"""
Genera matriz mensual vendor x locacion — objetivos + mes a mes.
Incluye desglose de componentes (LLEGADA, RECIBO, SALIDA) para la grafica.
"""
import csv, json, os
from collections import defaultdict

BASE = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP\bigquery_results'
CSV  = r'C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP\vendor_cedis_mes_FINAL.csv'
OUT  = os.path.join(BASE, 'matrix_FINAL.json')

MESES_PROM = ['Enero','Febrero','Marzo','Abril','Mayo','Junio']
MESES_ALL  = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio']
MESES_IND  = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio']

VENDOR_DISPLAY = {
    'KIMBERLY CLARK DE MEX SA DE CV':                 'KIMBERLY CLARK DE MEX SA B CV',
    'EMBOTELLADORA NIAGARA SA DE CV':                 'ENBOTELLAD NIAGARA D MX',
    'JUGOS DEL VALLE SAPI DE CV':                     'JUGOS DEL VALLE',
    'SANTA CLARA MERCANTIL DE PACHUCA S DE RL DE CV': 'SANTA CLARA MERC PACHU S RL CV',
    'SANTA CLARA MERCANTI':                           'SANTA CLARA MERC PACHU S RL CV',
    'SANTA CLARA MERCANTIL DE PA':                    'SANTA CLARA MERC PACHU S RL CV',
    'PROCTER & GAMBLE MÉXICO S DE RL DE CV':          'PROCTER AND GAMBLE MEXICO INC',
    'MARCAS NESTLE SA DE CV':                         'MARCAS NESTLE',
    'COLGATE PALMOLIVE SA DE CV':                     'COLGATE PALMOLIVE SA CV',
    'COMERC PEPSICO MEXICO S RL CV':                  'COMERC PEPSICO MEXICO S RL CV',
    'BONAFONT SA DE CV':                              'BONAFONT + ENVASASORA',
    'ENVASADORA LA SUPREM':                           'BONAFONT + ENVASASORA',
    'UNILEVER DE MEXICO S RL CV':                     'UNILEVER DE MEXICO S RL CV',
    'CIA COMERCIAL HERDEZ SA DE CV':                  'HERDEZ SA DE CV',
    'CERVEZA CANAL MO S DE RL DE CV':                 'CERVEZA CANAL MO S D',
    'FRABEL SA DE CV':                                'FRABEL SA DE CV',
    'MONDELEZ MEXICO S DE RL DE CV':                  'MONDELEZ MEXICO S DE RL DE CV',
    'MONDELEZ MEXICO S DE':                           'MONDELEZ MEXICO S DE RL DE CV',
    'KELLOGG COMPANY MEXICO SRL DE CV':               'KELLOGG COMPANY MEXICO SRL CV',
}
DISPLAY_ORDER = [
    'KIMBERLY CLARK DE MEX SA B CV','ENBOTELLAD NIAGARA D MX','JUGOS DEL VALLE',
    'SANTA CLARA MERC PACHU S RL CV','PROCTER AND GAMBLE MEXICO INC','MARCAS NESTLE',
    'COLGATE PALMOLIVE SA CV','COMERC PEPSICO MEXICO S RL CV','BONAFONT + ENVASASORA',
    'UNILEVER DE MEXICO S RL CV','HERDEZ SA DE CV','CERVEZA CANAL MO S D',
    'FRABEL SA DE CV','MONDELEZ MEXICO S DE RL DE CV','KELLOGG COMPANY MEXICO SRL CV',
]

REGIONS     = {'NORTE':['CUU','CLN','MXL','MTY'],'CENTRO':['CUAU','STB','SMO'],'SUR':['CHL','GDL','MER','VHSA']}
REGIONS_BAE = {'NORTE':['CLN','MTY'],'CENTRO':['STB'],'SUR':['CHL','GDL','MER','VHSA']}
ALL_LOCS_LIST = [l for g in REGIONS.values() for l in g]

# Autoservicios puro
AUTO_LOCS = {
    'CUU':[4640],
    'CLN':[7487],
    'MXL':[4924],
    'MTY':[7490],          # 7498 es PERECEDEROS AUTO
    'CUAU':[7464,7492,7494],'CUAU7494':[7494],'CUAU7464':[7464],'CUAU7492':[7492],
    'STB':[7482],
    'SMO':[],              # 7466 es PERECEDEROS AUTO
    'CHL':[7471],
    'GDL':[7493, 5907],    # 5907 Mi Bodega = Autoservicios
    'MER':[4188],
    'VHSA':[7468],
}
# BAE (Bodega Aurrera Express)
BAE_LOCS = {
    'CUU':[],
    'CLN':[7455],
    'MXL':[],
    'MTY':[7461,8806],
    'CUAU':[],'STB':[7457],'SMO':[],
    'CHL':[7459],
    'GDL':[7460],          # 5907 ya se movio a AUTO
    'MER':[7103],
    'VHSA':[7453],
}
# SAM'S Club
SAMS_LOCS = {
    'CUU':[5780],
    'CLN':[4971],
    'MXL':[6140],
    'MTY':[4995],          # 7502 es PERECEDEROS SAM'S
    'CUAU':[],'STB':[],'SMO':[6388],  # 4996 es PERECEDEROS SAM'S
    'CHL':[7505],
    'GDL':[6238],
    'MER':[7506],
    'VHSA':[6550],
}

# Lookup normalizado: upper + strip para capturar variantes de mayusculas/espacios
VENDOR_DISPLAY_NORM = {k.upper().strip(): v for k, v in VENDOR_DISPLAY.items()}

# raw[disp][cedis][mes] = {c, ws_total, ws_llegada, ws_recibo, ws_salida}
raw = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {
    'c':0,'ws':0.0,'wl':0.0,'wr':0.0,'wsal':0.0
})))

with open(CSV, encoding='utf-8-sig') as f:
    for r in csv.DictReader(f):
        disp = VENDOR_DISPLAY_NORM.get(r['VENDOR'].upper().strip())
        if not disp: continue
        cedis  = int(r['CEDIS'])
        mes    = r['MES']
        c      = int(r['TOTAL_CITAS'] or 0)
        total  = float(r['TOTAL_HRS']  or 0)
        llegada= float(r['LLEGADA']    or 0)
        recibo = float(r['RECIBO']     or 0)
        salida = float(r['SALIDA']     or 0)
        raw[disp][cedis][mes]['c']    += c
        raw[disp][cedis][mes]['ws']   += total   * c
        raw[disp][cedis][mes]['wl']   += llegada * c
        raw[disp][cedis][mes]['wr']   += recibo  * c
        raw[disp][cedis][mes]['wsal'] += salida  * c

def _wavg(vendors, cedis_list, mes_list, field='ws'):
    """Promedio simple de todas las citas individuales.
    sw = sum(LOS_medio_grupo * n_citas) = suma de todos los LOS individuales
    sc = sum(n_citas) = total de citas
    resultado = media de todas las citas una por una.
    """
    sw, sc = 0.0, 0
    for v in vendors:
        for c in cedis_list:
            for m in mes_list:
                r = raw[v][c].get(m)
                if r and r['c']:
                    sw += r[field]   # TOTAL_HRS * citas = suma de LOS individuales
                    sc += r['c']     # numero de citas
    return round(sw / sc, 2) if sc else None

def avg_by_vendor(vendors, cedis_list, mes_list, field='ws'):
    """Promedio simple de LOS por proveedor: cada TOP vendor pesa igual."""
    vals = []
    for v in vendors:
        v_val = _wavg([v], cedis_list, mes_list, field)
        if v_val is not None:
            vals.append(v_val)
    if not vals:
        return None
    return round(sum(vals) / len(vals), 2)

def wavg(vendors, cedis_list, mes_list):
    return _wavg(vendors, cedis_list, mes_list, 'ws')

def wavg_comp(vendors, cedis_list, mes_list):
    """Devuelve dict con llegada, recibo, salida."""
    l = _wavg(vendors, cedis_list, mes_list, 'wl')
    r = _wavg(vendors, cedis_list, mes_list, 'wr')
    s = _wavg(vendors, cedis_list, mes_list, 'wsal')
    return {'l': l, 'r': r, 's': s}

def build_matrix(locs_dict, all_cedis):
    m = {}
    for v in DISPLAY_ORDER:
        m[v] = {'prom': {}, 'ytd': {}}
        for loc, cl in locs_dict.items():
            m[v]['prom'][loc] = wavg([v], cl, MESES_PROM) if cl else None
            m[v]['ytd'][loc]  = wavg([v], cl, MESES_ALL)  if cl else None
        m[v]['prom']['2026'] = wavg([v], all_cedis, MESES_PROM)
        m[v]['ytd']['2026']  = wavg([v], all_cedis, MESES_ALL)
        for mes in MESES_IND:
            m[v][mes] = {}
            for loc, cl in locs_dict.items():
                m[v][mes][loc] = wavg([v], cl, [mes]) if cl else None
            m[v][mes]['2026'] = wavg([v], all_cedis, [mes])
    # total row — promedio ponderado por citas (todos los TOP vendors juntos)
    m['__total__'] = {'prom': {}, 'ytd': {}}
    for loc, cl in locs_dict.items():
        m['__total__']['prom'][loc] = wavg(DISPLAY_ORDER, cl, MESES_PROM) if cl else None
        m['__total__']['ytd'][loc]  = wavg(DISPLAY_ORDER, cl, MESES_ALL)  if cl else None
    m['__total__']['prom']['2026'] = wavg(DISPLAY_ORDER, all_cedis, MESES_PROM)
    m['__total__']['ytd']['2026']  = wavg(DISPLAY_ORDER, all_cedis, MESES_ALL)
    for mes in MESES_IND:
        m['__total__'][mes] = {}
        for loc, cl in locs_dict.items():
            m['__total__'][mes][loc] = wavg(DISPLAY_ORDER, cl, [mes]) if cl else None
        m['__total__'][mes]['2026'] = wavg(DISPLAY_ORDER, all_cedis, [mes])
    return m

def build_chart(locs_dict, all_cedis):
    """Componentes por vendor por loc por periodo (para la grafica)."""
    c = {}
    locs = {'2026': all_cedis, **locs_dict}
    periods = {'prom': MESES_PROM, 'ytd': MESES_ALL, **{m:[m] for m in MESES_IND}}
    for v in DISPLAY_ORDER:
        c[v] = {}
        for loc, cl in locs.items():
            if cl is None: cl = []
            c[v][loc] = {}
            for period, mes_list in periods.items():
                comp = wavg_comp([v], cl if cl else [], mes_list)
                tot  = wavg([v], cl if cl else [], mes_list)
                c[v][loc][period] = {'l': comp['l'], 'r': comp['r'], 's': comp['s'], 't': tot}
    return c

# Listas de CDs sin duplicados (las sub-keys CUAU7464 etc. repiten CDs)
AUTO_ALL     = list(dict.fromkeys(c for cl in AUTO_LOCS.values() for c in cl))
BAE_ALL      = list(dict.fromkeys(c for cl in BAE_LOCS.values()  for c in cl))
SAMS_ALL     = list(dict.fromkeys(c for cl in SAMS_LOCS.values() for c in cl))

# Combinaciones de canales
AUTO_BAE_LOCS  = {loc: AUTO_LOCS.get(loc,[]) + BAE_LOCS.get(loc,[])  for loc in AUTO_LOCS}
AUTO_SAMS_LOCS = {loc: AUTO_LOCS.get(loc,[]) + SAMS_LOCS.get(loc,[]) for loc in AUTO_LOCS}
BAE_SAMS_LOCS  = {loc: BAE_LOCS.get(loc,[])  + SAMS_LOCS.get(loc,[]) for loc in AUTO_LOCS}
ALL_LOCS       = {loc: AUTO_LOCS.get(loc,[]) + BAE_LOCS.get(loc,[]) + SAMS_LOCS.get(loc,[]) for loc in AUTO_LOCS}

AUTO_BAE_ALL  = list(dict.fromkeys(AUTO_ALL + BAE_ALL))
AUTO_SAMS_ALL = list(dict.fromkeys(AUTO_ALL + SAMS_ALL))
BAE_SAMS_ALL  = list(dict.fromkeys(BAE_ALL  + SAMS_ALL))
ALL_CDS       = list(dict.fromkeys(AUTO_ALL + BAE_ALL + SAMS_ALL))

def build_cd_matrix():
    """Genera matriz indexada por numero de cedis individual (string)."""
    all_cedis = sorted({c for v in raw for c in raw[v]})
    m = {}
    for v in DISPLAY_ORDER + ['__total__']:
        vendors = [v] if v != '__total__' else DISPLAY_ORDER
        m[v] = {'prom': {}, 'ytd': {}}
        for cedis in all_cedis:
            ck = str(cedis)
            m[v]['prom'][ck] = wavg(vendors, [cedis], MESES_PROM)
            m[v]['ytd'][ck]  = wavg(vendors, [cedis], MESES_ALL)
        m[v]['prom']['2026'] = wavg(vendors, all_cedis, MESES_PROM)
        m[v]['ytd']['2026']  = wavg(vendors, all_cedis, MESES_ALL)
        for mes in MESES_IND:
            m[v][mes] = {}
            for cedis in all_cedis:
                ck = str(cedis)
                m[v][mes][ck] = wavg(vendors, [cedis], [mes])
            m[v][mes]['2026'] = wavg(vendors, all_cedis, [mes])
    return m

def build_cd_chart():
    """Chart components indexados por cedis individual."""
    all_cedis = sorted({c for v in raw for c in raw[v]})
    periods = {'prom': MESES_PROM, 'ytd': MESES_ALL, **{m: [m] for m in MESES_IND}}
    c = {}
    for v in DISPLAY_ORDER:
        c[v] = {}
        for cedis in all_cedis:
            ck = str(cedis)
            c[v][ck] = {}
            for period, mes_list in periods.items():
                comp = wavg_comp([v], [cedis], mes_list)
                tot  = wavg([v], [cedis], mes_list)
                c[v][ck][period] = {'l': comp['l'], 'r': comp['r'], 's': comp['s'], 't': tot}
    return c

out = {
    'display_order': DISPLAY_ORDER,
    'regions':       REGIONS,
    'regions_bae':   REGIONS_BAE,
    'meses':         MESES_IND,
    'auto':           build_matrix(AUTO_LOCS,      AUTO_ALL),
    'bae':            build_matrix(BAE_LOCS,       BAE_ALL),
    'sams':           build_matrix(SAMS_LOCS,      SAMS_ALL),
    'auto_bae':       build_matrix(AUTO_BAE_LOCS,  AUTO_BAE_ALL),
    'auto_sams':      build_matrix(AUTO_SAMS_LOCS, AUTO_SAMS_ALL),
    'bae_sams':       build_matrix(BAE_SAMS_LOCS,  BAE_SAMS_ALL),
    'all':            build_matrix(ALL_LOCS,        ALL_CDS),
    'auto_chart':      build_chart(AUTO_LOCS,       AUTO_ALL),
    'bae_chart':       build_chart(BAE_LOCS,        BAE_ALL),
    'auto_bae_chart':  build_chart(AUTO_BAE_LOCS,   AUTO_BAE_ALL),
    'auto_sams_chart': build_chart(AUTO_SAMS_LOCS,  AUTO_SAMS_ALL),
    'bae_sams_chart':  build_chart(BAE_SAMS_LOCS,   BAE_SAMS_ALL),
    'sams_chart':      build_chart(SAMS_LOCS,       SAMS_ALL),
    'all_chart':       build_chart(ALL_LOCS,         ALL_CDS),
    'cd_matrix':     build_cd_matrix(),
    'cd_chart':      build_cd_chart(),
}
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',',':'))

print('OK')
for v in DISPLAY_ORDER:
    comp = wavg_comp([v], AUTO_ALL, MESES_PROM)
    tot  = wavg([v], AUTO_ALL, MESES_PROM)
    print(f"  {v[:30]:30s} L:{comp['l']}  R:{comp['r']}  S:{comp['s']}  T:{tot}")
