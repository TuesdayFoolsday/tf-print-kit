"""Fill the Bid Metrics v2 columns on estimates that already exist, without touching anything else in them.

    python3 metrics/backfill_v2.py "Merged Bid Metrics.csv" "Merged Units.csv" "Cost Map.csv" sources.json \
        --out <dir> [--write] [--by "chat <date>"]

sources.json: {"<metrics_id>": {"file": "<export.XML | compact.json | workbook.xlsx>",
                                "grout_walls": "Wall=32oc; ..." (by-wall Tradesmen only, optional)}, ...}
India rows read their tabs from the row's own source_detail ("workbook tabs: ...").

For each estimate it rebuilds the row from its source and refuses unless the rebuild matches what is stored:
bid price within $2, every unit type's labor_base within $2, and the same Units rows (same unit_key set, each
labor_base within $2). The CMU grout split comes from the stored Units rows (their grout labels and SF shares), so
keys line up with what David approved. Only then does it write pending changes that set the v2 columns only:
  Bid Metrics: labor_std_ratio (+ openings_per_100sf / wages_bumped when the stored row has them blank)
  Units:       grout_fill_pct, rebar_lf_per_sf, labor_lay/rebar/grout/clean, hours, hours_per_unit, labor_std_per_unit
Never a whole-row rewrite, never another column. Without --write it only reports (dry run).
Exit 0 = every listed estimate matched (or was reported as skipped); 1 = something refused.
"""
import argparse, csv, io, json, os, re, sys, warnings
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import bid_metrics as m
import metrics_log as ml
import units as U

TOL = 2.0
ROW_V2 = ["labor_std_ratio"]
ROW_FILL_IF_BLANK = ["openings_per_100sf", "wages_bumped"]


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def fmt(v):
    """A value as pending-file text, the way row_patch writes it."""
    return "" if v is None else (f"{v:g}" if isinstance(v, float) else str(v))


def grout_spec_from_units(urows):
    """Rebuild the --grout spec from stored Units rows: '8:full=0.6,16oc=0.4; 12:full' (shares by SF)."""
    by = {}
    for u in urows:
        if not u["material"].startswith("cmu") or u["grout"] in ("", "n/a"):
            continue
        by.setdefault(u["size"], {}).setdefault(u["grout"], 0.0)
        by[u["size"]][u["grout"]] += fnum(u["sf"]) or 0.0
    parts = []
    for size, g in by.items():
        tot = sum(g.values())
        if len(g) == 1:
            parts.append(f"{size}:{next(iter(g))}")
        elif tot:
            parts.append(f"{size}:" + ",".join(f"{k}={v / tot:.6f}" for k, v in g.items()))
    return "; ".join(parts)


def rebuild(row, urows, src, cost_map):
    """-> (new_row, new_units list) or raises ValueError(reason)."""
    path = src["file"]
    if not os.path.exists(path):
        raise ValueError("source file not found")
    if row["source"] == "india":
        import india_metrics as im
        mt = re.search(r"workbook tabs: ([^;]+)", row.get("source_detail") or "")
        if not mt:
            raise ValueError("no workbook tabs in source_detail")
        tabs = [t.strip() for t in mt.group(1).split(", ")]
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            r = im.compute(path, cost_map, tabs, row["project"], row["metrics_id"], row.get("bid_id", ""),
                           row.get("estimate_status", ""), row.get("notes", ""))
        raw, gf, v2 = r.pop("_units_raw"), r.pop("_grout_found"), r.pop("_v2", {})
        lf, mf = fnum(r["labor_load_factor"]) or 0, fnum(r["material_load_factor"]) or 0
        nu = U.finish(raw, row["metrics_id"], row["project"], "india", row.get("estimate_status", ""), lf, mf,
                      grout_found=gf, v2=v2)
    else:
        wg = dict(x.split("=", 1) for x in (p.strip() for p in (src.get("grout_walls") or "").split(";")) if "=" in x)
        r = m.compute(path, cost_map, row["project"], wall_grout={k.strip(): v.strip().lower() for k, v in wg.items()})
        raw, (lf, mf), v2 = r.pop("_units_raw"), r.pop("_factors"), r.pop("_v2", {})
        r.pop("walls", None); r.pop("_qty_note", None)
        spec = None if wg else grout_spec_from_units(urows)
        nu = U.finish(raw, row["metrics_id"], row["project"], "tradesmen", row.get("estimate_status", ""), lf, mf,
                      grout_spec=spec, v2=v2)
    return r, nu


def compare(row, urows, r, nu):
    probs = []
    if abs((fnum(r.get("bid_price")) or 0) - (fnum(row.get("bid_price")) or 0)) > TOL:
        probs.append(f"bid price {r.get('bid_price')} vs stored {row.get('bid_price')}")
    for b in m.UNIT_TYPES + m.OTHER:
        k = f"{b}_labor_base"
        if k in row and row[k] != "" and abs((fnum(r.get(k)) or 0) - (fnum(row[k]) or 0)) > TOL:
            probs.append(f"{k} {r.get(k)} vs stored {row[k]}")
    old = {u["unit_key"]: u for u in urows}
    new = {u["unit_key"]: u for u in nu}
    if set(old) != set(new):
        probs.append(f"units rows differ: missing {sorted(set(old) - set(new))} new {sorted(set(new) - set(old))}")
    for k in set(old) & set(new):
        if abs((fnum(new[k]["labor_base"]) or 0) - (fnum(old[k]["labor_base"]) or 0)) > TOL:
            probs.append(f"{k}: labor {new[k]['labor_base']} vs stored {old[k]['labor_base']}")
    return probs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("metrics"); ap.add_argument("units"); ap.add_argument("cost_map"); ap.add_argument("sources")
    ap.add_argument("--out", required=True); ap.add_argument("--write", action="store_true"); ap.add_argument("--by", default="")
    a = ap.parse_args()
    mheader, mrows = ml.bl.load(a.metrics)
    uheader, urows_all = ml.bl.load(a.units)
    mheader, uheader = ml.with_v2(mheader, mrows), ml.with_v2(uheader, urows_all)
    rows = {r[0]: dict(zip(mheader, r)) for r in mrows}
    units = {}
    for r in urows_all:
        d = dict(zip(uheader, r)); units.setdefault(d["metrics_id"], []).append(d)
    sources = json.load(open(a.sources, encoding="utf-8"))
    report, bad = [], 0
    for mid in sorted(rows):
        row = rows[mid]
        if mid not in sources:
            report.append((mid, row["project"], "no source", "source file not given — left as is"))
            continue
        try:
            r, nu = rebuild(row, units.get(mid, []), sources[mid], a.cost_map)
        except Exception as e:   # one estimate's problem never stops the others
            report.append((mid, row["project"], "refused", str(e)[:200])); bad += 1
            continue
        probs = compare(row, units.get(mid, []), r, nu)
        if probs:
            report.append((mid, row["project"], "refused", "; ".join(probs)[:400])); bad += 1
            continue
        setm = {k: fmt(r.get(k)) for k in ROW_V2}
        for k in ROW_FILL_IF_BLANK:
            if row.get(k, "") == "":
                setm[k] = fmt(r.get(k))
        setm = {k: v for k, v in setm.items() if v != ""}
        n_units = 0
        buf = io.StringIO()
        with redirect_stdout(buf):
            if a.write and setm:
                ml.write_patch({"metrics_id": mid, "by": a.by, "set": setm}, mheader, set(rows), os.path.join(a.out, "Bid Metrics"))
            for u in nu:
                full = ml.row_patch(u, uheader, a.by)["set"]
                s = {k: full[k] for k in U.V2_UNITS_COLS if full.get(k, "") != ""}
                if s:
                    n_units += 1
                    if a.write:
                        ml.write_patch({"unit_key": u["unit_key"], "by": a.by, "set": s}, uheader,
                                       {x["unit_key"] for x in units.get(mid, [])}, os.path.join(a.out, "Bid Metrics Units"))
        report.append((mid, row["project"], "ok" if a.write else "would write",
                       f"row: {', '.join(f'{k}={v}' for k, v in setm.items()) or '—'} · {n_units} units rows"))
    w = max(len(x[0]) for x in report) if report else 10
    for mid, proj, st, note in report:
        print(f"{mid:<{w}}  {st:<11}  {proj[:40]:<40}  {note}")
    n = lambda s: sum(1 for x in report if x[2] == s)
    print(f"\n{n('ok') + n('would write')} matched · {n('refused')} refused · {n('no source')} without a source"
          + ("" if a.write else " · dry run, nothing written"))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
