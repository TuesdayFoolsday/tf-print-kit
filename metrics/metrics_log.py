"""Bid Metrics: base CSV + small pending change files (System Rules v1.13, B16), the same pattern as the Bid Log.

Unattended runs never rewrite 'Bid Metrics 2026.csv'. Each change is one small JSON file in
'_Knowledge/Pending/Bid Metrics', keyed by metrics_id, and every reader merges first:

    python3 metrics/metrics_log.py merge "Bid Metrics 2026.csv" <pending_dir> --out "Merged Bid Metrics.csv"
    python3 metrics/metrics_log.py patch "Bid Metrics 2026.csv" <out_dir> --metrics-id tm-26056 --by "hourly 11:52" \\
        --set notes="radius wall" [--set col=value ...] [--json patch.json]
    python3 metrics/metrics_log.py show  "Bid Metrics 2026.csv" <pending_dir> <metrics_id>

metrics_id: 'tm-<job number>' for a Tradesmen estimate, 'in-<project slug>' for an India workbook estimate
(make_id below). A metrics_id not in the base is a new row (needs project, source, estimate_status).
estimate_status: final (the estimate behind the price that went out; sets the Bid Check range) · iteration
(an earlier or alternate pricing) · draft (unfinished). Values are text; columns must exist in the base header.
Applied oldest first by (ts, file name); a later patch to the same field wins.
"""
import argparse, csv, datetime as dt, glob, json, os, re, sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "bidlog"))
import bid_log as bl   # load / write / read_patches / now_ct

KEY = "metrics_id"
STATUSES = {"final", "iteration", "draft", ""}
SOURCES = {"tradesmen", "india", ""}
# Columns added by Bid Metrics v2 (System Rules v1.29). A base file written before v2 lacks them: merge adds them
# (old rows read blank) and patches may set them.
V2_COLS = {"metrics_id": ["max_wall_height_ft", "openings_per_100sf", "chase_walls", "access", "wages_bumped",
                          "conditions_file"],
           "unit_key": ["grout_fill_pct", "rebar_lf_per_sf", "labor_lay", "labor_rebar", "labor_grout", "labor_clean"]}


def with_v2(header, rows=None):
    """Header (and rows) extended with any missing v2 columns, blank."""
    extra = [c for c in V2_COLS.get(header[0] if header else "", []) if c not in header]
    if rows is not None:
        for r in rows:
            r.extend([""] * (len(header) + len(extra) - len(r)))
    return header + extra


def make_id(source, job_number="", project=""):
    if source == "tradesmen" and str(job_number).strip():
        return f"tm-{str(job_number).strip()}"
    slug = re.sub(r"[^a-z0-9]+", "-", project.lower()).strip("-")[:40]
    return f"{'tm' if source == 'tradesmen' else 'in'}-{slug}"


def validate(p, header, base_ids):
    KEY = header[0]          # metrics_id (Bid Metrics) or unit_key (Bid Metrics Units)
    errs = []
    if not p.get(KEY):
        errs.append(f"missing {KEY}")
    s = p.get("set")
    if not isinstance(s, dict) or not s:
        return errs + ["missing set{}"]
    for k, v in s.items():
        if k not in header:
            errs.append(f"unknown column {k!r}")
        if k == KEY:
            errs.append(f"{KEY} can't be changed")
        if not isinstance(v, str):
            errs.append(f"{k}: value must be text (got {type(v).__name__})")
    if s.get("estimate_status") not in (None, *STATUSES):
        errs.append(f"estimate_status {s['estimate_status']!r} not final/iteration/draft")
    if s.get("source") not in (None, *SOURCES):
        errs.append(f"source {s['source']!r} not tradesmen/india")
    if p.get(KEY) not in base_ids:
        for k in ("project", "source", "estimate_status"):
            if not s.get(k):
                errs.append(f"new row {p.get(KEY)}: needs {k}")
    return errs


def merge(base_path, pdir):
    header, rows = bl.load(base_path)
    if not header or header[0] not in (KEY, "unit_key"):
        raise SystemExit(f"⚠ {base_path}: first column must be {KEY} or unit_key")
    header = with_v2(header, rows)
    K = header[0]
    idx = {r[0]: i for i, r in enumerate(rows)}
    applied, warns = [], []
    for f, p, err in bl.read_patches(pdir):
        name = os.path.basename(f)
        if err:
            warns.append(f"⚠ {name}: {err} — skipped"); continue
        errs = validate(p, header, set(idx))
        if errs:
            warns.append(f"⚠ {name}: " + "; ".join(errs) + " — skipped"); continue
        k = p.get(K)
        if k not in idx:
            rows.append([k] + [""] * (len(header) - 1)); idx[k] = len(rows) - 1
        row = rows[idx[k]]
        for c, v in p["set"].items():
            row[header.index(c)] = v
        applied.append(name)
    return header, rows, sorted(set(applied)), warns


def write_patch(p, header, base_ids, out_dir):
    KEY = header[0]
    p.setdefault("ts", bl.now_ct())
    errs = validate(p, header, base_ids)
    if errs:
        for e in errs:
            print("⚠ not written — " + e)
        return None
    os.makedirs(out_dir, exist_ok=True)
    stamp = p["ts"][:16].replace("T", " ").replace(":", "")
    name = f"{stamp} {re.sub(r'[^A-Za-z0-9._-]+', '-', p[KEY])[:110]}.json"
    path = os.path.join(out_dir, name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(p, f, ensure_ascii=False, indent=1); f.write("\n")
    kind = "change to" if p[KEY] in base_ids else "NEW row"
    folder = "Bid Metrics Units" if KEY == "unit_key" else "Bid Metrics"
    print(f"Upload (new file in _Knowledge/Pending/{folder}): {name}  ({os.path.getsize(path)} bytes) — {kind} {p[KEY]}")
    return path


def row_patch(row, header, by=""):
    """A whole row (estimate or units row) -> patch setting every column."""
    K = header[0]
    s = {}
    for c in header:
        if c == K or c not in row:
            continue
        v = row[c]
        s[c] = "" if v is None else (f"{v:g}" if isinstance(v, float) else str(v))
    return {K: row[K], "by": by, "set": s}


def cmd_merge(a):
    header, rows, applied, warns = merge(a.base, a.pending)
    bl.write(header, rows, a.out)
    for w in warns:
        print(w)
    print(f"Bid Metrics merged: {len(rows)} rows ({len(applied)} pending files applied, {len(warns)} skipped) -> {a.out}")
    if a.list:
        for n in applied:
            print("  applied: " + n)
    return 1 if warns else 0


def cmd_patch(a):
    header, rows = bl.load(a.base)
    header = with_v2(header, rows)
    p = json.load(open(a.json, encoding="utf-8")) if a.json else {}
    p.setdefault("set", {})
    for kv in a.set or []:
        k, _, v = kv.partition("=")
        p["set"][k.strip()] = v
    if a.metrics_id:
        p[header[0]] = a.metrics_id
    p.setdefault("by", a.by or "")
    return 0 if write_patch(p, header, {r[0] for r in rows}, a.out_dir) else 1


def cmd_show(a):
    header, rows, *_ = merge(a.base, a.pending)
    for r in rows:
        if r[0] == a.metrics_id:          # metrics_id or unit_key
            for k, v in zip(header, r):
                if v:
                    print(f"{k}: {v}")
            return 0
    print(f"{a.metrics_id}: not found"); return 1


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("merge"); p.add_argument("base"); p.add_argument("pending"); p.add_argument("--out", required=True)
    p.add_argument("--list", action="store_true"); p.set_defaults(f=cmd_merge)
    p = sp.add_parser("patch"); p.add_argument("base"); p.add_argument("out_dir"); p.add_argument("--metrics-id")
    p.add_argument("--set", action="append"); p.add_argument("--json"); p.add_argument("--by"); p.set_defaults(f=cmd_patch)
    p = sp.add_parser("show"); p.add_argument("base"); p.add_argument("pending"); p.add_argument("metrics_id"); p.set_defaults(f=cmd_show)
    a = ap.parse_args()
    sys.exit(a.f(a) or 0)


if __name__ == "__main__":
    main()
