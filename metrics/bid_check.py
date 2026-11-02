"""Compare one export's labor $/SF by unit type to history and print a short plain-text check.

    python3 metrics/bid_check.py <export.XML | compact.json> <cost_map.csv> <merged Bid Metrics.csv> [--project NAME] [--note TEXT]
        [--status final|iteration|draft] [--pending-out DIR]     # scheduled runs (B7 / B16)
        [--append]                                               # chat only: rewrites the CSV given

Labor $/SF (base wage actually priced, before burden and markup) is the judged metric: it is the only one
flagged. Production, material and sell $/SF are printed for reference. --append adds or replaces this
project's row in history.csv (only when the export reconciles); --note sets the row's notes (e.g. "radius wall").
--pending-out writes the row as one pending JSON for '_Knowledge/Pending/Bid Metrics' instead (System Rules B16).
The range compares against estimate_status = final rows only (iterations and drafts stay for reference).
"""
import argparse, csv, statistics, sys, os
sys.path.insert(0, os.path.dirname(__file__))
import bid_metrics as m
import metrics_log as ml
import units as U

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
    ap.add_argument("--status", default="final", choices=["final", "iteration", "draft"])
    ap.add_argument("--pending-out", default="")
    ap.add_argument("--by", default="")
    ap.add_argument("--units", default="", help="merged Bid Metrics Units.csv (history by material, size, grout)")
    ap.add_argument("--units-pending-out", default="")
    ap.add_argument("--grout", default="", help="David's grout per CMU size, e.g. '12:full=0.63,16oc=0.37; 8:32oc'")
    ap.add_argument("--grout-walls", default="", help="by-wall export: 'Wall name=32oc; Other wall=full' (David)")
    a = ap.parse_args()
    wg = dict(x.split("=", 1) for x in (p.strip() for p in a.grout_walls.split(";")) if "=" in x)
    r = m.compute(a.export, a.cost_map, a.project, wall_grout={k.strip(): v.strip().lower() for k, v in wg.items()})
    raw, (lf, mf) = r.pop("_units_raw"), r.pop("_factors")
    nwalls = r.pop("walls", 0)
    r["metrics_id"] = ml.make_id("tradesmen", r.get("job_number", ""), r["project"])
    r["estimate_status"] = a.status
    hist = []
    if os.path.exists(a.history):
        with open(a.history, newline="", encoding="utf-8") as f:
            hist = list(csv.DictReader(f))
    prev = next((h for h in hist if h.get("project") == r["project"]), None)
    r["notes"] = a.note or (prev or {}).get("notes", "")
    hist = [h for h in hist if h.get("project") != r["project"] and h.get("metrics_id") != r["metrics_id"]]
    if prev and prev.get("bid_id") and not r.get("bid_id"):
        r["bid_id"] = prev["bid_id"]
    rng = [h for h in hist if h.get("estimate_status", "final") == "final"]
    urows = U.finish(raw, r["metrics_id"], r["project"], "tradesmen", a.status, lf, mf, grout_spec=a.grout,
                     basis="Tradesmen by wall (grout per wall, David)" if nwalls else "Tradesmen line items by size",
                     notes=r["notes"])
    uhist = []
    if a.units and os.path.exists(a.units):
        with open(a.units, newline="", encoding="utf-8") as f:
            uhist = [h for h in csv.DictReader(f) if h.get("metrics_id") != r["metrics_id"]
                     and h.get("estimate_status", "final") == "final" and U.metric(h) is not None]

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
    proj = {h.get("metrics_id"): h for h in hist}
    tag = lambda h: h["project"] + (f" ({proj[h['metrics_id']]['notes']})" if proj.get(h["metrics_id"], {}).get("notes") else "")
    for u in urows:
        v = U.metric(u)
        if v is None:
            continue
        per = "SF" if u["material"] in ("stone", "cast_stone") else ("block" if u["material"].startswith("cmu") else "brick")
        label = {"cmu_gray": "gray CMU", "cmu_arch": "arch CMU"}.get(u["material"], u["material"].replace("_", " "))
        size = (f'{u["size"]}" ' if u["material"].startswith("cmu") else (f'{u["size"]} ' if u["material"] == "brick" else ""))
        grout = f' · grout {u["grout"]}' if u["material"].startswith("cmu") else ""
        same = lambda h, keys: all(h.get(k) == u[k] for k in keys)
        basis, past = "", []
        if u["material"].startswith("cmu") and u["grout"] not in ("unknown", "mixed"):
            past = [h for h in uhist if same(h, ("material", "size", "grout"))]
            basis = "same size + grout"
        if len(past) < 3:
            past = [h for h in uhist if same(h, ("material", "size"))]
            basis = "same size, any grout" if u["material"].startswith("cmu") else "same size"
        line = f"• {size}{label}{grout} — {u['units']:,} {per}s ({u['sf']:,.0f} SF) — labor ${v:,.2f}/{per}"
        pv = [(U.metric(h), h) for h in past]
        if pv:
            lo, hi = min(pv, key=lambda t: t[0]), max(pv, key=lambda t: t[0])
            med = statistics.median(x for x, _ in pv)
            mark = ""
            if v < lo[0] * (1 - TOL) or v > hi[0] * (1 + TOL):
                mark = " ⚠ " + ("HIGH" if v > hi[0] else "LOW"); flags += 1
            line += f"{mark} | {basis}: median ${med:,.2f} · low ${lo[0]:,.2f} {tag(lo[1])} · high ${hi[0]:,.2f} {tag(hi[1])} · n={len(pv)}"
        else:
            line += " | no history for this size yet"
        out.append(line + f"\n   ref: {u['units_per_manday'] or '–'} {per}s/man-day · material ${u['material_per_unit'] or 0}/{per} · sell ${u['sell_per_unit'] or 0}/{per}")
    if not uhist:
        out.append("(no Bid Metrics Units history given — run with --units)")
    out.append(f"{flags} labor flag(s) outside history ±{TOL:.0%}." if flags else "Labor inside the historical range on every material and size.")
    print("\n".join(out))

    if a.units_pending_out and r["reconciles"] == "y" and a.units:
        uheader, _ = ml.bl.load(a.units)
        for u in urows:
            ml.write_patch(ml.row_patch(u, uheader, a.by), uheader, set(), a.units_pending_out)
    if a.pending_out and r["reconciles"] == "y":
        header, _ = ml.bl.load(a.history)
        missing = [k for k in r if k not in header]
        if missing:
            print("⚠ not written — columns not in Bid Metrics: " + ", ".join(missing))
            return
        ml.write_patch(ml.row_patch(r, header, a.by), header, {h.get("metrics_id") for h in hist + ([prev] if prev else [])}, a.pending_out)
    if a.append and r["reconciles"] == "y":
        rows = hist + [r]
        keys = []
        [keys.append(k) for x in rows for k in x if k not in keys]
        with open(a.history, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys, lineterminator="\n"); w.writeheader(); w.writerows(rows)


if __name__ == "__main__":
    main()
