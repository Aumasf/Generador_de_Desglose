"""costos_partes.py

Este módulo contiene SOLO la lógica de cálculo para las partes A/B/D/E/F a partir del CDT.

Objetivos
- Mantener pdf_utils.py más legible.
- Evitar tocar la lógica que ya funciona (resumen CDT..CU+IVA, match, etc.).
- Entregar números ENTEROS, sin decimales y nunca negativos.

Convención (según el template)
- CDT = D + E + F
- C = 1 (siempre)
- D = (A + B) / C  => con C=1, A+B = D

Reglas (según lo último acordado)
1) F (Transporte) NUNCA es 0: siempre es 10% del CDT (redondeo tipo Excel).
2) El reparto D/E depende del tipo de ítem:
   - ambiguo:
       D : E mantiene la proporción 45 : 40 sobre el remanente (CDT - F)
   - mano_obra:
       E = 0
       D = CDT - F   (≈ 90%)
   - materiales:
       D : E mantiene la proporción 5 : 85 sobre el remanente (CDT - F)
3) A y B se derivan de D:
   - Si el ítem tiene mano de obra (mano_obra o ambiguo):
       A = 10% de D
       B = D - A
   - Si el ítem es de materiales:
       B = 0  (mano de obra inexistente)
       A = D  (porque D = A + B)

Notas
- No imprimimos decimales.
- No imprimimos negativos.
- Garantizamos que D + E + F == CDT.
"""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List


def _round_half_up(x: float) -> int:
    """Redondeo 0 decimales tipo Excel (ROUND_HALF_UP)."""
    try:
        return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except Exception:
        return 0


def _clamp_nonneg_int(x) -> int:
    """Convierte a int y fuerza >= 0."""
    try:
        v = int(x)
    except Exception:
        v = 0
    return v if v >= 0 else 0


def _alloc_porcentajes(total: int, porcentajes: List[float]) -> List[int]:
    """Asigna 'total' en partes enteras según porcentajes, cerrando suma exacta.

    Estrategia:
    1) floor de cada parte
    2) distribuir el residuo (+1) según mayor parte fraccional
    """
    total = _clamp_nonneg_int(total)
    if total == 0:
        return [0] * len(porcentajes)

    raws = [total * p for p in porcentajes]
    floors = [int(r) if r >= 0 else 0 for r in raws]
    fracs = [r - f for r, f in zip(raws, floors)]

    s = sum(floors)
    residuo = total - s

    # índices por fracción desc
    orden = sorted(range(len(fracs)), key=lambda i: fracs[i], reverse=True)

    out = floors[:]
    k = 0
    while residuo > 0 and orden:
        i = orden[k % len(orden)]
        out[i] += 1
        residuo -= 1
        k += 1

    # seguridad: si por algún motivo nos pasamos (no debería), recortamos
    exceso = sum(out) - total
    k = 0
    while exceso > 0 and orden:
        i = orden[::-1][k % len(orden)]
        if out[i] > 0:
            out[i] -= 1
            exceso -= 1
        k += 1

    return out


def _norm_tipo(tipo_item: str) -> str:
    t = (tipo_item or "").strip().lower()

    # tolerancia de entradas
    if t in ("mano de obra", "mano_obra", "manoobra", "mo", "mano-obra"):
        return "mano_obra"
    if t in ("materiales", "material", "mat"):
        return "materiales"
    if t in ("ambiguo", "ambigua", "amb"):
        return "ambiguo"

    # default conservador
    return "ambiguo"


def _split_ratio(total: int, w1: float, w2: float) -> List[int]:
    """Divide 'total' en 2 enteros manteniendo proporción w1:w2 y cerrando suma."""
    total = _clamp_nonneg_int(total)
    if total == 0:
        return [0, 0]
    s = float(w1) + float(w2)
    if s <= 0:
        return [total, 0]
    return _alloc_porcentajes(total, [float(w1) / s, float(w2) / s])


def calcular_partes_desde_cdt(cdt: int, tipo_item: str) -> Dict[str, int]:
    """Devuelve las partes desde CDT según reglas.

    Retorna un dict con:
      - D, E, F
      - A, B
      - AB (A+B), C (=1)
      - CDT

    Todos enteros y >= 0, y con D + E + F == CDT.
    """
    cdt = _clamp_nonneg_int(cdt)
    t = _norm_tipo(tipo_item)

    # F siempre 10% (redondeo tipo Excel)
    f = _round_half_up(cdt * 0.10)
    f = min(_clamp_nonneg_int(f), cdt)

    rem = cdt - f
    rem = _clamp_nonneg_int(rem)

    if t == "mano_obra":
        d, e = rem, 0
    elif t == "materiales":
        d, e = _split_ratio(rem, 5, 85)  # D:E ≈ 5:85 sobre el remanente
    else:
        # ambiguo
        d, e = _split_ratio(rem, 45, 40)  # D:E ≈ 45:40 sobre el remanente

    # A y B a partir de D (C siempre 1)
    if d <= 0:
        a, b = 0, 0
    else:
        if t == "materiales":
            # Si no hay mano de obra, B=0 y A=D (porque D = A + B)
            a, b = d, 0
        else:
            # Mano de obra (o ambiguo): A 10% de D, B el resto
            a = _round_half_up(d * 0.10)
            a = min(_clamp_nonneg_int(a), d)
            b = d - a
            b = _clamp_nonneg_int(b)

    # seguridad final (no negativos)
    d = _clamp_nonneg_int(d)
    e = _clamp_nonneg_int(e)
    f = _clamp_nonneg_int(f)
    a = _clamp_nonneg_int(a)
    b = _clamp_nonneg_int(b)

    # cierre exacto por si algo raro pasó por clamps/redondeos
    suma = d + e + f
    if suma != cdt:
        # ajustamos E (materiales) para cerrar, sin permitir negativos
        ajuste = cdt - suma
        e = _clamp_nonneg_int(e + ajuste)
        # si aún no cierra (por clamps), ajustamos D
        suma2 = d + e + f
        if suma2 != cdt:
            d = _clamp_nonneg_int(d + (cdt - suma2))
            # y si se desfasó por arriba, recortamos
            if d + e + f > cdt:
                exceso = (d + e + f) - cdt
                if e >= exceso:
                    e -= exceso
                elif d >= exceso:
                    d -= exceso

    return {
        "CDT": cdt,
        "A": a,
        "B": b,
        "C": 1,
        "AB": _clamp_nonneg_int(a + b),  # == D (con C=1)
        "D": d,
        "E": e,
        "F": f,
    }
