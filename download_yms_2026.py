"""
Descarga el dataset YMS 2026 completo desde BigQuery con TRIM fix en TIPO_CITA.
Guarda en bigquery_results/yms_2026_completo.csv
"""
import os
import csv
from google.cloud import bigquery

OUTPUT = r"C:\Users\mmvhern\OneDrive - Walmart Inc\Escritorio\puppy\YMS_TOP\bigquery_results\yms_2026_completo.csv"

QUERY = """
SELECT
  APPOINTMENT_NBR, ARRIVAL_DATE, CEDIS, NOMBRE_CEDIS, LOCACION,
  VENDOR, TIPO_CITA, CITAS_CORRECTAS, MES, SW, ANIO,
  LLEGADA_A_TRAFICO, ABRIR_CORTINA, CERRAR_CORTINA, PAPER_W, SALIDA_DE_CD,
  DURACION_DE_SERVICIO
FROM `wmt-intl-supplychain-gcp-prod.MR101_WM_AD_HOC.SCH_YMS_SEMANAL`
WHERE ARRIVAL_DATE BETWEEN '2026-01-01' AND '2026-07-03'
  AND TRIM(UPPER(TIPO_CITA)) IN ('PROVEEDOR', 'CITA NUEVA')
  AND CITAS_CORRECTAS = 1
ORDER BY ARRIVAL_DATE, CEDIS
"""

print("Conectando a BigQuery...")
client = bigquery.Client()

print("Ejecutando query (sin pandas - CSV directo)...")
job = client.query(QUERY)
result = job.result()  # espera a que termine

os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

total = 0
frabel = 0

with open(OUTPUT, 'w', newline='', encoding='utf-8-sig') as f:
    writer = None
    for row in result:
        row_dict = dict(row)
        if writer is None:
            fieldnames = list(row_dict.keys())
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
        writer.writerow(row_dict)
        total += 1
        vendor = str(row_dict.get('VENDOR', '')).upper().strip()
        if 'FRABEL' in vendor:
            frabel += 1
        if total % 50000 == 0:
            print(f"  ...{total:,} filas escritas")

print(f"Filas descargadas: {total:,}")
print(f"Filas de FRABEL  : {frabel:,}")
print(f"\nGuardado en: {OUTPUT}")
print(f"Tamano archivo: {os.path.getsize(OUTPUT) / 1_048_576:.1f} MB")
