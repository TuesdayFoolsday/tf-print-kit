"""Texas Class B contract bond premium, as carried on the FB proposal.

Sliding scale from FB's bond calculator, applied to the BASE BID only, then
marked up by bond_markup_pct — a per-estimate field, default 21%, editable
(David's ruling 2026-10-07; supersedes "defaults to project markup" in the spec).
Independent of project OH&P. Completion-time surcharge is deliberately not
applied. Result is whole dollars.
"""

TIERS = [            # (band width, rate) — cumulative bands
    (100_000,      0.02500),   # first $100,000
    (400_000,      0.01500),   # $100,001 – $500,000
    (2_000_000,    0.01000),   # $500,001 – $2,500,000
    (2_500_000,    0.00750),   # $2,500,001 – $5,000,000
    (2_500_000,    0.00700),   # $5,000,001 – $7,500,000
    (float("inf"), 0.00650),   # over $7,500,000
]

OHP = 0.21


def premium(base_bid: float) -> float:
    """Raw Class B premium on the contract amount (no markup)."""
    left, prem = float(base_bid), 0.0
    for band, rate in TIERS:
        x = min(left, band)
        prem += x * rate
        left -= x
        if left <= 0:
            break
    return prem


def bond_add(base_bid: float, ohp: float = OHP) -> int:
    """Dollar add shown on page 1: premium × (1 + OH&P), rounded to whole dollars."""
    return int(premium(base_bid) * (1 + ohp) + 0.5)  # half-up, not banker's rounding


if __name__ == "__main__":
    for b in (243_643, 500_000, 1_000_000, 2_000_000, 3_000_000):
        print(f"{b:>12,}  premium {premium(b):>10,.2f}  add {bond_add(b):>8,}")
