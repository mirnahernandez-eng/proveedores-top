import pandas as pd
df = pd.read_csv('sw21_22_new.csv', nrows=100)
print('=== Todas las columnas con su % de nulls ===')
for col in df.columns:
    pct_null = df[col].isna().sum()
    sample = df[col].dropna().head(2).tolist()
    print(f'  {col}: {pct_null}/100 nulls | sample={sample}')
