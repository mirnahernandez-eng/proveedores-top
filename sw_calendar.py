# -*- coding: utf-8 -*-
"""
sw_calendar.py
Calendario oficial de Supply Weeks de Walmart – fuente única de verdad.

Estructura: SW_num → (inicio_YYYY-MM-DD, fin_YYYY-MM-DD, mes_fiscal)
El 'mes_fiscal' refleja la designación del periodo en el tablero YMS,
no necesariamente el mes calendario con más días en la semana.
"""

SW_CALENDAR = {
    # ── Enero (Q4 año anterior, SWs 48-52) ──────────────────────────────────
    48: ("2025-12-27", "2026-01-02", "Enero"),
    49: ("2026-01-03", "2026-01-09", "Enero"),
    50: ("2026-01-10", "2026-01-16", "Enero"),
    51: ("2026-01-17", "2026-01-23", "Enero"),
    52: ("2026-01-24", "2026-01-30", "Enero"),
    # ── Febrero ──────────────────────────────────────────────────────────────
     1: ("2026-01-31", "2026-02-06", "Febrero"),
     2: ("2026-02-07", "2026-02-13", "Febrero"),
     3: ("2026-02-14", "2026-02-20", "Febrero"),
     4: ("2026-02-21", "2026-02-27", "Febrero"),
    # ── Marzo ────────────────────────────────────────────────────────────────
     5: ("2026-02-28", "2026-03-06", "Marzo"),
     6: ("2026-03-07", "2026-03-13", "Marzo"),
     7: ("2026-03-14", "2026-03-20", "Marzo"),
     8: ("2026-03-21", "2026-03-27", "Marzo"),
     9: ("2026-03-28", "2026-04-03", "Marzo"),   # 4 días mar, 3 días abr
    # ── Abril ────────────────────────────────────────────────────────────────
    10: ("2026-04-04", "2026-04-10", "Abril"),
    11: ("2026-04-11", "2026-04-17", "Abril"),
    12: ("2026-04-18", "2026-04-24", "Abril"),
    13: ("2026-04-25", "2026-05-01", "Abril"),   # 6 días abr, 1 día may
    # ── Mayo ─────────────────────────────────────────────────────────────────
    14: ("2026-05-02", "2026-05-08", "Mayo"),
    15: ("2026-05-09", "2026-05-15", "Mayo"),
    16: ("2026-05-16", "2026-05-22", "Mayo"),
    17: ("2026-05-23", "2026-05-29", "Mayo"),
    # ── Junio ────────────────────────────────────────────────────────────────
    18: ("2026-05-30", "2026-06-05", "Junio"),   # 5 días jun, 2 días may
    19: ("2026-06-06", "2026-06-12", "Junio"),
    20: ("2026-06-13", "2026-06-19", "Junio"),
    21: ("2026-06-20", "2026-06-26", "Junio"),
    22: ("2026-06-27", "2026-07-03", "Junio"),   # 4 días jun, 3 días jul
    # ── Julio ────────────────────────────────────────────────────────────────
    23: ("2026-07-04", "2026-07-10", "Julio"),
    24: ("2026-07-11", "2026-07-17", "Julio"),
    25: ("2026-07-18", "2026-07-24", "Julio"),
    26: ("2026-07-25", "2026-07-31", "Julio"),
    # ── Agosto ───────────────────────────────────────────────────────────────
    27: ("2026-08-01", "2026-08-07", "Agosto"),
    28: ("2026-08-08", "2026-08-14", "Agosto"),
    29: ("2026-08-15", "2026-08-21", "Agosto"),
    30: ("2026-08-22", "2026-08-28", "Agosto"),
    # ── Septiembre ───────────────────────────────────────────────────────────
    31: ("2026-08-29", "2026-09-04", "Septiembre"),  # 3 días ago, 4 días sep
    32: ("2026-09-05", "2026-09-11", "Septiembre"),
    33: ("2026-09-12", "2026-09-18", "Septiembre"),
    34: ("2026-09-19", "2026-09-25", "Septiembre"),
    35: ("2026-09-26", "2026-10-02", "Septiembre"),  # 5 días sep, 2 días oct
    # ── Octubre ──────────────────────────────────────────────────────────────
    36: ("2026-10-03", "2026-10-09", "Octubre"),
    37: ("2026-10-10", "2026-10-16", "Octubre"),
    38: ("2026-10-17", "2026-10-23", "Octubre"),
    39: ("2026-10-24", "2026-10-30", "Octubre"),
    # ── Noviembre ────────────────────────────────────────────────────────────
    40: ("2026-10-31", "2026-11-06", "Noviembre"),   # 1 día oct, 6 días nov
    41: ("2026-11-07", "2026-11-13", "Noviembre"),
    42: ("2026-11-14", "2026-11-20", "Noviembre"),
    43: ("2026-11-21", "2026-11-27", "Noviembre"),
    # ── Diciembre ────────────────────────────────────────────────────────────
    44: ("2026-11-28", "2026-12-04", "Diciembre"),   # 3 días nov, 4 días dic
    45: ("2026-12-05", "2026-12-11", "Diciembre"),
    46: ("2026-12-12", "2026-12-18", "Diciembre"),
    47: ("2026-12-19", "2026-12-25", "Diciembre"),
}

# Derivados (para importar directamente sin reconstruir)
SW_MES_MAP = {sw: data[2] for sw, data in SW_CALENDAR.items()}
SW_DATES   = {sw: {"inicio": data[0], "fin": data[1]} for sw, data in SW_CALENDAR.items()}

# Orden fiscal: 48-52 primero, luego 1-47
SW_ORDER = [sw for sw in range(48, 53)] + [sw for sw in range(1, 48)]

# Helper: dado un SW num, devuelve su rango como string legible (ej. "Jul 4-10")
_MES_ABREV = {
    1:"Ene", 2:"Feb", 3:"Mar", 4:"Abr", 5:"May", 6:"Jun",
    7:"Jul", 8:"Ago", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dic",
}

def sw_range_label(sw_num: int) -> str:
    """Devuelve '4 Jul - 10 Jul' o '27 Jun - 3 Jul' para SWs que cruzan mes."""
    if sw_num not in SW_CALENDAR:
        return ""
    ini_str, fin_str, _ = SW_CALENDAR[sw_num]
    from datetime import date
    ini = date.fromisoformat(ini_str)
    fin = date.fromisoformat(fin_str)
    m1 = _MES_ABREV[ini.month]
    m2 = _MES_ABREV[fin.month]
    if ini.month == fin.month:
        return f"{m1} {ini.day}-{fin.day}"
    return f"{m1} {ini.day} - {m2} {fin.day}"
