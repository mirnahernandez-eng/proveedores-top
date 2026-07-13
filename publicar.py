"""
publicar.py — Pipeline completo para publicar el tablero en Puppy Pages
  Paso 1: build_tablero.py   → regenera tablero_los_proveedores.html
  Paso 2: make_standalone.py → incrusta todos los assets para Puppy Pages
  Paso 3: Kory llama a share-puppy para subir tablero_standalone.html

Uso:
  .venv\\Scripts\\python.exe publicar.py
  — o —
  Dile a Kory: "publica el tablero"
"""
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).parent
PY   = sys.executable

STEPS = [
    ("build_tablero.py",   "Regenerando tablero HTML desde datos actuales..."),
    ("make_standalone.py", "Generando version standalone (incrustando assets)..."),
]

def run_step(script: str, label: str) -> bool:
    print(f"\n{'-' * 52}")
    print(f"  {label}")
    print(f"{'-' * 52}")
    t0 = time.time()
    result = subprocess.run(
        [PY, str(BASE / script)],
        cwd=str(BASE),
        capture_output=False,  # deja que imprima en consola
        text=True,
    )
    elapsed = time.time() - t0
    if result.returncode != 0:
        print(f"\n  ERROR en {script} (codigo {result.returncode})")
        return False
    print(f"  OK - {script} en {elapsed:.1f}s")
    return True

def main() -> int:
    print("\n" + "=" * 52)
    print("  PUBLICAR TABLERO - inicio")
    print("=" * 52)

    for script, label in STEPS:
        if not run_step(script, label):
            print("\n  Abortando pipeline.")
            return 1

    standalone = BASE / "tablero_standalone.html"
    size_kb = standalone.stat().st_size // 1024 if standalone.exists() else 0

    print("\n" + "=" * 52)
    print(f"  tablero_standalone.html listo ({size_kb} KB) - OK")
    print("  Paso 3: Kory sube a Puppy Pages automaticamente [LISTO]")
    print("=" * 52)
    return 0

if __name__ == "__main__":
    sys.exit(main())
