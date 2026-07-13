import pandas as pd

paths = [
    ('root', 'vendor_cedis_mes_FINAL.csv'),
    ('bigquery_results', 'bigquery_results/vendor_cedis_mes_FINAL.csv'),
]

for label, path in paths:
    try:
        df = pd.read_csv(path, encoding='utf-8-sig')
        frabel = df[df['VENDOR'].str.upper().str.contains('FRABEL', na=False)]
        total = frabel['TOTAL_CITAS'].sum()
        print(f"[{label}] -> {path}")
        print(f"  Filas totales CSV : {len(df):,}")
        print(f"  Citas Frabel      : {total}")
        print(f"  Por CEDIS:")
        print(frabel.groupby('CEDIS')['TOTAL_CITAS'].sum().to_string())
        print()
    except Exception as e:
        print(f"[{label}] ERROR: {e}\n")
