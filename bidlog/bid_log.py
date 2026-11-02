"""Bid Log: base CSV + small pending change files (System Rules v1.8, C2/A15/B10).

Unattended runs never rewrite 'Bid Log 2026.csv'. Every change is one small JSON file in
'_Knowledge/Pending/Bid Log', and every reader merges first:

    python3 bidlog/bid_log.py patch "Bid Log 2026.csv" <out_dir> --bid-id 2026-213-suffolk --by "hourly 11:52" \\
        --set status=submitted --set submitted_amount=7024631 [--set col=value ...] [--json patch.json]
    python3 bidlog/bid_log.py merge "Bid Log 2026.csv" <pending_dir> --out "Merged Bid Log.csv"
    python3 bidlog/bid_log.py show  "Bid Log 2026.csv" <pending_dir> <bid_id>     # one merged row

Patch file = {"bid_id": "...", "ts": "YYYY-MM-DDTHH:MM-06:00", "by": "<run or chat>", "set": {"<column>": "<value>"}}
  * bid_id not in the base -> a new row (needs at least project and status); else only the named fields change.
  * Values are text, written exactly as given ("" clears a field). Columns must exist in the base header.
  * Applied oldest first by (ts, file name); a later patch to the same field wins (reported as a note).
File name: '<YYYY-MM-DD HHMM> <bid_id>.json'. Writing the same change twice is harmless.

`merge` writes the full merged CSV for scripts (hit rate, Dashboard, self-heal). In chat, the merged CSV is
uploaded over the base (helper agent + byte-for-byte check) and the applied files move to 'Pending/Bid Log/Merged'.
"""
import argparse, csv, datetime as dt, glob, json, os, re, sys

STATUSES = {"invited", "passed", "bidding", "submitted", "awarded", "lost", "closed"}
TAXES = {"included", "exempt", "excluded", ""}


def load(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        r = csv.reader(f)
        header = next(r)
        return header, [row for row in r]


def now_ct():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=-6))).strftime("%Y-%m-%dT%H:%M-06:00")


def validate(p, header, base_ids, strict_new=True):
    errs = []
    if not p.get("bid_id"):
        errs.append("missing bid_id")
    if not isinstance(p.get("set"), dict) or not p["set"]:
        errs.append("missing set{}")
        return errs
    for k, v in p["set"].items():
        if k not in header:
            errs.append(f"unknown column {k!r}")
        if k == "bid_id":
            errs.append("bid_id can't be changed")
        if not isinstance(v, str):
            errs.append(f"{k}: value must be text (got {type(v).__name__})")
    st = p["set"].get("status")
    if st is not None and st not in STATUSES:
        errs.append(f"status {st!r} not one of {sorted(STATUSES)}")
    tx = p["set"].get("tax")
    if tx is not None and tx not in TAXES:
        errs.append(f"tax {tx!r} not included/exempt/excluded")
    for k in ("submitted_amount", "base_bid", "alternates_total", "winning_amount"):
        v = p["set"].get(k)
        if v and not re.fullmatch(r"-?\d+(\.\d+)?", v):
            errs.append(f"{k} {v!r} must be a plain number (no $ or commas)")
    for k in ("invite_date", "bid_date", "bid_date_original", "submitted_date", "result_date"):
        v = p["set"].get(k)
        if v:
            try:
                dt.date.fromisoformat(v)
            except ValueError:
                errs.append(f"{k} {v!r} must be YYYY-MM-DD")
    if strict_new and p.get("bid_id") not in base_ids:
        for k in ("project", "status"):
            if not p["set"].get(k):
                errs.append(f"new row {p.get('bid_id')}: needs {k}")
    return errs


def read_patches(pdir):
    out = []
    for f in sorted(glob.glob(os.path.join(pdir, "*.json"))):
        try:
            p = json.load(open(f, encoding="utf-8"))
        except ValueError as e:
            out.append((f, None, f"not valid JSON ({e})"))
            continue
        for q in (p if isinstance(p, list) else [p]):
            out.append((f, q, None))
    out.sort(key=lambda t: ((t[1] or {}).get("ts", ""), os.path.basename(t[0])))
    return out


def merge(base_path, pdir):
    header, rows = load(base_path)
    idx = {r[0]: i for i, r in enumerate(rows)}
    applied, warns, notes, seen = [], [], [], {}
    for f, p, err in read_patches(pdir):
        name = os.path.basename(f)
        if err:
            warns.append(f"⚠ {name}: {err} — skipped")
            continue
        errs = validate(p, header, set(idx))
        if errs:
            warns.append(f"⚠ {name}: " + "; ".join(errs) + " — skipped")
            continue
        bid = p["bid_id"]
        if bid not in idx:
            rows.append([bid] + [""] * (len(header) - 1))
            idx[bid] = len(rows) - 1
        row = rows[idx[bid]]
        for k, v in p["set"].items():
            j = header.index(k)
            prev = seen.get((bid, k))
            if prev and prev[1] != v:
                notes.append(f"· {bid} {k}: {prev[1]!r} ({prev[0]}) → {v!r} ({name})")
            seen[(bid, k)] = (name, v)
            row[j] = v
        applied.append(name)
    return header, rows, sorted(set(applied)), warns, notes


def write(header, rows, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        w.writerows(rows)


def cmd_patch(a):
    header, rows = load(a.base)
    p = json.load(open(a.json, encoding="utf-8")) if a.json else {}
    p.setdefault("set", {})
    for kv in a.set or []:
        k, _, v = kv.partition("=")
        p["set"][k.strip()] = v
    if a.bid_id:
        p["bid_id"] = a.bid_id
    p.setdefault("ts", now_ct())
    p.setdefault("by", a.by or "")
    errs = validate(p, header, {r[0] for r in rows})
    if errs:
        for e in errs:
            print("⚠ not written — " + e)
        return 1
    os.makedirs(a.out_dir, exist_ok=True)
    stamp = p["ts"][:16].replace("T", " ").replace(":", "")
    name = f"{stamp} {re.sub(r'[^A-Za-z0-9._-]+', '-', p['bid_id'])}.json"
    path = os.path.join(a.out_dir, name)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(p, f, ensure_ascii=False, indent=1)
        f.write("\n")
    kind = "change to" if p["bid_id"] in {r[0] for r in rows} else "NEW row"
    print(f"Upload (new file in _Knowledge/Pending/Bid Log): {name}  ({os.path.getsize(path)} bytes) — {kind} {p['bid_id']}")
    return 0


def cmd_merge(a):
    header, rows, applied, warns, notes = merge(a.base, a.pending)
    write(header, rows, a.out)
    for w in warns:
        print(w)
    for n in notes:
        print(n)
    print(f"Bid Log merged: {len(rows)} rows ({len(applied)} pending files applied, {len(warns)} skipped) -> {a.out}")
    if a.list:
        for n in applied:
            print("  applied: " + n)
    return 1 if warns else 0


def cmd_show(a):
    header, rows, *_ = merge(a.base, a.pending)
    for r in rows:
        if r[0] == a.bid_id:
            for k, v in zip(header, r):
                if v:
                    print(f"{k}: {v}")
            return 0
    print(f"{a.bid_id}: not found")
    return 1


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("patch"); p.add_argument("base"); p.add_argument("out_dir"); p.add_argument("--bid-id")
    p.add_argument("--set", action="append"); p.add_argument("--json"); p.add_argument("--by"); p.set_defaults(f=cmd_patch)
    p = sp.add_parser("merge"); p.add_argument("base"); p.add_argument("pending"); p.add_argument("--out", required=True)
    p.add_argument("--list", action="store_true"); p.set_defaults(f=cmd_merge)
    p = sp.add_parser("show"); p.add_argument("base"); p.add_argument("pending"); p.add_argument("bid_id"); p.set_defaults(f=cmd_show)
    a = ap.parse_args()
    sys.exit(a.f(a) or 0)


if __name__ == "__main__":
    main()
