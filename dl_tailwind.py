"""Descarga Tailwind a traves del proxy Walmart."""
import urllib.request, sys

URLS = [
    "https://cdn.tailwindcss.com/3.4.17/tailwind.min.css",
    "https://unpkg.com/tailwindcss@3.4.7/dist/tailwind.min.css",
]
PROXY = "http://sysproxy.wal-mart.com:8080"

proxy_handler = urllib.request.ProxyHandler({'http': PROXY, 'https': PROXY})
opener = urllib.request.build_opener(proxy_handler)

for url in URLS:
    try:
        sys.stdout.write(f"Intentando: {url}\n")
        sys.stdout.flush()
        r = opener.open(url, timeout=15)
        data = r.read()
        with open("tailwind.min.css", "wb") as f:
            f.write(data)
        sys.stdout.write(f"OK: {len(data)} bytes descargados\n")
        sys.stdout.flush()
        break
    except Exception as e:
        sys.stdout.write(f"FALLO: {e}\n")
        sys.stdout.flush()
else:
    sys.stdout.write("No se pudo descargar Tailwind por ningun proxy\n")
    sys.stdout.flush()
