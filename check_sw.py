# -*- coding: utf-8 -*-
import json, sys
sys.stdout.reconfigure(encoding='utf-8')

d = json.load(open('sw_data.json', encoding='utf-8'))

print("=== sw_list actual ===")
print(d['sw_list'])

print("\n=== sw_mes_map ===")
for k,v in d['sw_mes_map'].items():
    print(f"  {k}: {v}")

# Revisar cuantos vendors tienen datos en SW23 y SW24
auto = d['auto']
sams = d['sams']

for sw_key in ['SW23', 'SW24']:
    print(f"\n=== {sw_key} ===")
    for cat_name, cat_data in [('AUTO', auto), ('SAMS', sams)]:
        con_datos = 0
        for vendor, cedis_map in cat_data.items():
            if vendor.startswith('__'): continue
            nat = cedis_map.get('2026', {})
            if sw_key in nat and nat[sw_key] and nat[sw_key].get('t'):
                con_datos += 1
        print(f"  {cat_name}: {con_datos} vendors con datos en {sw_key}")
    # Mostrar ejemplo de valores
    kimberly = auto.get('KIMBERLY CLARK DE MEX SA B CV', {})
    nat_kim = kimberly.get('2026', {})
    if sw_key in nat_kim:
        print(f"  Ejemplo (Kimberly Clark nacional): {nat_kim[sw_key]}")
    else:
        print(f"  Ejemplo (Kimberly Clark nacional): SIN DATOS")

# Checar fecha de modificacion del json
import os
stat = os.stat('sw_data.json')
from datetime import datetime
print(f"\nsw_data.json modificado: {datetime.fromtimestamp(stat.st_mtime)}")
