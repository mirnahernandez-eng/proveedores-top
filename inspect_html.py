import re

with open('tablero_standalone.html', encoding='utf-8') as f:
    html = f.read()

print('Tamanio HTML:', len(html), 'chars')

# scripts src externos
scripts = re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', html)
print('\n=== Scripts externos ===')
for s in scripts:
    print(' ', s)

# fetch calls
fetches = re.findall(r'fetch\(["\']([^"\']+)["\']', html)
print('\n=== fetch() calls ===')
for f in fetches:
    print(' ', f)

# DATA_ variables definidas
data_vars = re.findall(r'var\s+(DATA_\w+)\s*=', html)
print('\n=== Variables DATA_ definidas ===')
for d in data_vars:
    print(' ', d)

# link/img src externos (CDN, etc.)
extern = re.findall(r'(?:src|href)=["\']https?://[^"\']+["\']', html)
print('\n=== URLs externas (CDN) ===')
for e in set(extern):
    print(' ', e[:100])
