"""Compare one export's unit rates to history and print a short plain-text check.

    python3 metrics/bid_check.py <export.XML> <cost_map.csv> <history.csv> [--project NAME] [--append]

Flags any unit type whose labor $/SF, material $/SF, sell $/SF or SF/man-day sits outside the
historical range by more than 20%. --append adds this bid's row to history.csv (replacing a row
with the same project name).
"""
import argparse, csv, statistics, sys, os
sys.path.insert(0, os.path.dirname(__file__))
import bid_metrics as m

METRICS = [("labor_per_sf", "labor $/SF"), ("material_per_sf", "material $/SF"),
           ("sell_per_sf", "sell $/SF"), ("sf_per_manday", "SF/man-day")]
TOL = 0.20


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("export"); ap.add_argument("cost_map"); ap.add_argument("history")
    ap.add_argument("--project", default=""); ap.add_argument("--append", action="store_true")
    a = ap.parse_args()
    r = m.compute(a.export, a.cost_map, a.project)
    hist = []
    if os.path.exists(a.history):
        with open(a.history, newline="", encoding="utf-8") as f:
            hist = [h for h in csv.DictReader(f) if h.get("project") != r["project"]]
    out = [f"Bid Check — {r['project']} — ${r['bid_price']:,} — {r['wall_sf_total']:,.0f} wall SF — "
           f"${r['sell_per_wall_sf'] or 0:,.2f}/SF overall — OH {r['oh_pct']:.0%} / profit {r['profit_pct']:.0%}"]
    if r["reconciles"] != "y":
        out.append(f"⚠ export does not reconcile: buckets ${r['bucketed_total']:,} vs bid ${r['bid_price']:,}")
    if r["unmapped_items"]:
        out.append(f"⚠ items not in Cost Map (counted as other): {r['unmapped_items']}")
    flags = 0
    for u in m.UNIT_TYPES:
        if not r.get(f"{u}_sf"):
            continue
        parts = []
        for key, label in METRICS:
            v = r.get(f"{u}_{key}")
            past = [float(h[f"{u}_{key}"]) for h in hist if h.get(f"{u}_{key}") not in (None, "", "None")]
            if v is None or not past:
                parts.append(f"{label} {v}"); continue
            lo, hi, med = min(past), max(past), statistics.median(past)
            mark = ""
            if v < lo * (1 - TOL) or v > hi * (1 + TOL):
                mark = " ⚠"; flags += 1
            parts.append(f"{label} {v:,.2f} (hist {lo:,.2f}–{hi:,.2f}, med {med:,.2f}, n={len(past)}){mark}")
        out.append(f"• {u.replace('_', ' ')} {r[f'{u}_sf']:,.0f} SF: " + "; ".join(parts))
    out.append(f"{flags} flag(s) outside history ±{TOL:.0%}." if flags else "All unit rates inside historical range.")
    print("\n".join(out))
    if a.append:
        rows = hist + [r]
        keys = []
        [keys.append(k) for x in rows for k in x if k not in keys]
        with open(a.history, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys); w.writeheader(); w.writerows(rows)


if __name__ == "__main__":
    main()
