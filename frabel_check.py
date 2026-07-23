import json, sys

sw = json.load(open('sw_data.json', encoding='utf-8-sig'))
nat = sw['auto']['FRABEL SA DE CV']['2026']

# Datos por SW
print("=== FRABEL - Datos nacionales por SW ===")
print(f"{'SW':<6} {'Total':>7} {'Llegada':>8} {'Recibo':>8} {'Salida':>8}")
print("-"*45)
for k in sorted(nat.keys(), key=lambda x: int(x[2:])):
    v = nat[k]
    if v:
        print(f"{k:<6} {v.get('t',0):>7.2f} {v.get('l',0):>8.2f} {v.get('r',0):>8.2f} {v.get('s',0):>8.2f}")

vals = [v['t'] for v in nat.values() if v and v.get('t')]
n = len(vals)
avg = sum(vals) / n if n else 0

print("-"*45)
print(f"Total SWs con datos: {n}")
print(f"SW23 presente: {'SW23' in nat and bool(nat.get('SW23'))}")
print(f"SW25 presente: {'SW25' in nat and bool(nat.get('SW25'))}")
print(f"Promedio YTD (prom. de SWs): {round(avg,2)}")

# Citas en vendor_cedis_mes_FINAL.csv
import csv
total_citas = 0
with open('vendor_cedis_mes_FINAL.csv', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    for row in reader:
        if 'FRABEL' in str(row.get('VENDOR','')).upper():
            total_citas += int(float(row.get('TOTAL_CITAS', 0)))
print(f"\nTotal citas FRABEL en CSV mensual: {total_citas}")
