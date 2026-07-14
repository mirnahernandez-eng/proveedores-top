# -*- coding: utf-8 -*-
"""
patch_sw_dates.py
Inyecta sw_dates en el sw_data.json existente usando sw_calendar.py.
Corre una sola vez — no toca los datos de proveedores, solo añade/actualiza la clave sw_dates.
"""
import json, re, sys
sys.path.insert(0, '.')
from sw_calendar import SW_DATES, SW_MES_MAP, sw_range_label

SW_JSON = 'sw_data.json'

print(f'Leyendo {SW_JSON}...')
raw = open(SW_JSON, encoding='utf-8').read()
# Quitar NaN/Infinity que JSON no soporta
raw = re.sub(r'\bNaN\b', 'null', raw)
raw = re.sub(r'-?Infinity\b', 'null', raw)
d = json.loads(raw)

sw_list = d.get('sw_list', [])
print(f'SWs actuales en sw_list: {sw_list}')

# Reconstruir sw_mes_map desde fuente oficial
d['sw_mes_map'] = {f"SW{sw}": SW_MES_MAP.get(sw, d['sw_mes_map'].get(f'SW{sw}', '?'))
                   for sw in sw_list}

# Añadir/actualizar sw_dates
d['sw_dates'] = {
    f"SW{sw}": {
        "inicio": SW_DATES[sw]["inicio"],
        "fin":    SW_DATES[sw]["fin"],
        "label":  sw_range_label(sw),
    }
    for sw in sw_list if sw in SW_DATES
}

# Guardar
out_raw = json.dumps(d, ensure_ascii=False, separators=(',', ':'))
open(SW_JSON, 'w', encoding='utf-8').write(out_raw)
size_kb = len(out_raw) // 1024
print(f'Listo! {SW_JSON} actualizado ({size_kb} KB)')
print('sw_dates generados:')
for k, v in d['sw_dates'].items():
    print(f'  {k}: {v["inicio"]} -> {v["fin"]}  ({v["label"]})')
