"""Compare one export's labor $/SF by unit type to history and print a short plain-text check.

    python3 metrics/bid_check.py <export.XML | compact.json> <cost_map.csv> <history.csv> [--project NAME] [--note TEXT] [--append]

Labor $/SF (base wage actually priced, before burden and markup) is the judged metric: it is the only one
flagged. Production, material and sell $/SF are printed for reference. --append adds or replaces this
project's row in history.csv (only when the export reconciles); --note sets the row's notes (e.g. "radius wall").
"""
import argparse, csv, statistics, sys, os
sys.path.insert(0, os.path.dirname(__file__))
import bid_metrics as m

TOL = 0.20           # flag labor $/SF more than 20% outside the historical range
CREW_GAP = 0.05      # flag when Crew Page labor differs from item labor by more than 5%


def fnum(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("export"); ap.add_argument("cost_map"); ap.add_argument("history")
    ap.add_argument("--project", default=""); ap.add_argument("--note", default="")
    ap.add_argument("--append", action="store_true")
    a = ap.parse_args()
    r = m.compute(a.export, a.cost_map, a.project)
    hist = []
    if os.path.exists(a.history):
        with open(a.history, newline="", encoding="utf-8") as f:
            hist = list(csv.DictReader(f))
    prev = next((h for h in hist if h.get("project") == r["project"]), None)
    r["notes"] = a.note or (prev or {}).get("notes", "")
    hist = [h for h in hist if h.get("project") != r["project"]]

    out = [f"Bid Check — {r['project']} — ${r['bid_price']:,} — {r['wall_sf_total']:,.0f} wall SF — "
           f"OH {r['oh_pct']:.0%} / profit {r['profit_pct']:.0%}"]
    if r["reconciles"] != "y":
        out.append(f"⚠ export does not reconcile — {r['reconcile_notes']}. Not added to history.")
    wv = r.get("wages_vs_standard")
    if wv is not None and abs(wv) > 0.01:
        out.append(f"⚠ Wage rates {abs(wv):.0%} {'above' if wv > 0 else 'below'} standard ({r['wages']}; standard "
                   + " / ".join(f"{k} {v:g}" for k, v in m.STD_WAGES.items()) + "). Labor $/SF below includes that.")
    gap = r["crew_vs_takeoff_labor"] - 1
    if abs(gap) > CREW_GAP:
        out.append(f"⚠ Crew Page labor is {abs(gap):.0%} {'below' if gap < 0 else 'above'} the line-item labor — check the crew mix. "
                   "Rates below use the Crew Page (what's in the price).")
    if r["unmapped_items"]:
        out.append(f"⚠ items not in Cost Map (counted as other): {r['unmapped_items']}")

    flags = 0
    for u in m.UNIT_TYPES:
        if not r.get(f"{u}_sf"):
            continue
        name = u.replace("_", " ")
        v = r.get(f"{u}_labor_per_sf")
        past = [(fnum(h.get(f"{u}_labor_per_sf")), h) for h in hist]
        past = [(x, h) for x, h in past if x is not None]
        line = f"• {name} {r[f'{u}_sf']:,.0f} SF — labor ${v:,.2f}/SF"
        if past:
            lo, hi = min(past, key=lambda t: t[0]), max(past, key=lambda t: t[0])
            med = statistics.median(x for x, _ in past)
            tag = lambda t: t[1]["project"] + (f" ({t[1]['notes']})" if t[1].get("notes") else "")
            mark = ""
            if v < lo[0] * (1 - TOL) or v > hi[0] * (1 + TOL):
                mark = " ⚠ " + ("HIGH" if v > hi[0] else "LOW"); flags += 1
            line += (f"{mark} | median ${med:,.2f} · low ${lo[0]:,.2f} {tag(lo)} · high ${hi[0]:,.2f} {tag(hi)} · n={len(past)}")
        else:
            line += " | no history yet"
        ref = [f"{r.get(f'{u}_sf_per_manday') or '–'} SF/man-day",
               f"material ${r.get(f'{u}_material_per_sf', 0):,.2f}/SF", f"sell ${r.get(f'{u}_sell_per_sf', 0):,.2f}/SF"]
        out.append(line + "\n   ref: " + " · ".join(ref))
    out.append(f"{flags} labor flag(s) outside history ±{TOL:.0%}." if flags else "Labor $/SF inside historical range on every unit type.")
    print("\n".join(out))

    if a.append and r["reconciles"] == "y":
        rows = hist + [r]
        keys = []
        [keys.append(k) for x in rows for k in x if k not in keys]
        with open(a.history, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys, lineterminator="\n"); w.writeheader(); w.writerows(rows)


if __name__ == "__main__":
    main()
