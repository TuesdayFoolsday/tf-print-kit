"""Price Book: supplier quotes -> one CSV row per quoted line item.

    python3 pricebook/price_book.py append <quotes.json> <folder>               # monthly files (preferred)
    python3 pricebook/price_book.py append <quotes.json> <price_book.csv>      # single file
    python3 pricebook/price_book.py latest <price_book.csv> [--item TEXT] [--days N]
    python3 pricebook/price_book.py movers <price_book.csv> [--days 90] [--min-pct 3]
    python3 pricebook/price_book.py check  <quotes.json>                        # validate only, no write
    python3 pricebook/price_book.py export <folder> <out.csv> [--days 90]       # latest rows + change vs prior quote; also Price Movers.csv beside it

quotes.json = a list of quotes (or one quote object):
  {"source_id": "<email message id or file item id>", "source_url": "<webLink/webUrl>",
   "quote_date": "YYYY-MM-DD", "supplier": "...", "class": "brick|cmu|stone|cast_stone|mortar_grout|
   accessories|insulation|rebar|sub_stucco|sub_waterproofing|sub_caulking|other",
   "project": "...", "city": "...", "quote_ref": "quote no / rev", "hold_days": 30 | "hold_until": "YYYY-MM-DD",
   "freight": "...", "tax": "excluded|included|unknown",
   "lines": [{"item": "brick modular", "product": "as written", "manufacturer": "...", "size_color": "...",
              "qty": 90000, "unit": "M", "unit_price": 800, "price_per_sf": 5.52, "notes": "..."}]}

Folder mode writes 'Price Book YYYY-MM.csv' (by quote month) + 'Price Book Quotes YYYY-MM.csv', and price lists to
'Price Lists YYYY.csv' + 'Price Lists Quotes YYYY.csv'; it prints exactly which files changed (upload those).
Single-file mode writes two files: <price_book.csv> (one row per line item) and its companion 'Price Book Quotes <year>.csv'
(one row per quote: link, freight, quote ref) joined on quote_key.
Rows are keyed by source_id + line number, so re-running the same quote replaces its rows (never duplicates).
`latest` = y on the newest row per supplier + item + unit + size_color.
"""
import argparse, csv, datetime as dt, hashlib, json, re, sys
from collections import defaultdict

COLS = ["quote_id", "quote_key", "quote_date", "supplier", "class", "project", "city", "item", "product",
        "manufacturer", "size_color", "qty", "unit", "unit_price", "price_per_sf", "unit_type", "tax", "hold_until",
        "latest", "notes"]
QCOLS = ["quote_key", "quote_date", "supplier", "class", "project", "city", "quote_ref", "freight", "tax",
         "hold_until", "lines", "source_url"]
CLASSES = {"brick", "cmu", "stone", "cast_stone", "mortar_grout", "accessories", "insulation", "rebar",
           "sub_stucco", "sub_waterproofing", "sub_caulking", "other"}
UNITS = {"M": ["m", "/m", "per m", "thousand", "per thousand", "1000", "mbf"],
         "EA": ["ea", "each", "pc", "pcs", "piece", "unit", "block", "blk"],
         "SF": ["sf", "sqft", "sq ft", "square feet", "ft2"], "LF": ["lf", "lin ft", "linear feet", "ft"],
         "TON": ["ton", "tons", "tn"], "LB": ["lb", "lbs", "pound"], "BAG": ["bag", "bags", "sack", "sk"],
         "CY": ["cy", "yd", "cubic yard"], "PAL": ["pal", "pallet", "cube"], "ROLL": ["roll", "rl"],
         "GAL": ["gal", "gallon", "pail", "bucket"], "BOX": ["box", "bx", "carton", "ctn"],
         "SHEET": ["sheet", "sht", "board", "bd"], "LOAD": ["load", "truck", "trip"],
         "LS": ["ls", "lump sum", "lump", "lot"]}
UNIT_MAP = {a: k for k, al in UNITS.items() for a in al + [k.lower()]}
ARCH_CMU = re.compile(r"trenwyth|burnish|split|ground|glaz|scored|ivany|color|arch|mesa|trendstone|fluted|ribbed|polish", re.I)
TAX = {"excluded", "included", "unknown"}


def unit_of(u):
    k = (u or "").strip().lower().lstrip("$/ ").rstrip(".")
    return UNIT_MAP.get(k)


def unit_type(cls, text):
    if cls == "cmu":
        return "cmu_arch" if ARCH_CMU.search(text or "") else "cmu_gray"
    if cls in ("brick", "stone", "cast_stone", "insulation", "rebar"):
        return cls
    if cls in ("mortar_grout", "accessories"):
        return "support"
    return "sub" if cls.startswith("sub_") else "other"


def money(x):
    if x in (None, ""):
        return ""
    if isinstance(x, (int, float)):
        return round(float(x), 4)
    s = re.sub(r"[,$\s]", "", str(x))
    try:
        return round(float(s), 4)
    except ValueError:
        return None


def iso(d):
    if not d:
        return ""
    try:
        return dt.date.fromisoformat(str(d)[:10]).isoformat()
    except ValueError:
        return None


def qkey(source_id):
    return hashlib.sha1(str(source_id).encode()).hexdigest()[:10]


def qid(source_id, n):
    return qkey(source_id) + f"-{n:02d}"


def quotes_path(csv_path):
    """Companion file: 'Price Book 2026.csv' -> 'Price Book Quotes 2026.csv'."""
    d, _, b = csv_path.rpartition("/")
    nb = b.replace("Price Book", "Price Book Quotes", 1) if "Price Book" in b else b[:-4] + " Quotes.csv"
    return (d + "/" if d else "") + nb


def flatten(quotes):
    """Validate quotes, return (rows, qrows, problems). A quote with any hard problem contributes no rows."""
    rows, qrows, problems = [], [], []
    for qi, q in enumerate(quotes):
        tag = f"quote {qi + 1} ({q.get('supplier', '?')} / {q.get('project', '?')})"
        hard = []
        for k in ("source_id", "quote_date", "supplier", "class", "lines"):
            if not q.get(k):
                hard.append(f"missing {k}")
        qd = iso(q.get("quote_date"))
        if qd is None:
            hard.append(f"bad quote_date {q.get('quote_date')!r}")
        cls = (q.get("class") or "").strip()
        if cls and cls not in CLASSES:
            hard.append(f"unknown class {cls!r}")
        tax = (q.get("tax") or "unknown").strip().lower()
        if tax not in TAX:
            hard.append(f"tax must be one of {sorted(TAX)}")
        hold = iso(q.get("hold_until")) if q.get("hold_until") else ""
        if hold is None:
            hard.append(f"bad hold_until {q.get('hold_until')!r}")
        if not hold and q.get("hold_days") and qd:
            hold = (dt.date.fromisoformat(qd) + dt.timedelta(days=int(q["hold_days"]))).isoformat()
        out = []
        for n, ln in enumerate(q.get("lines") or [], 1):
            u = unit_of(ln.get("unit"))
            p, psf, qty = money(ln.get("unit_price")), money(ln.get("price_per_sf")), money(ln.get("qty"))
            if not ln.get("item") or not ln.get("product"):
                hard.append(f"line {n}: item and product required")
            if u is None:
                hard.append(f"line {n}: unknown unit {ln.get('unit')!r}")
            if p is None or psf is None or qty is None:
                hard.append(f"line {n}: non-numeric price/qty")
            if p == "" and psf == "":
                hard.append(f"line {n}: no unit_price or price_per_sf")
            notes = (ln.get("notes") or "").strip()
            out.append({
                "quote_id": qid(q.get("source_id"), n), "quote_key": qkey(q.get("source_id")), "quote_date": qd, "supplier": q.get("supplier", "").strip(),
                "class": cls, "project": (q.get("project") or "").strip(), "city": (q.get("city") or "").strip(),
                "item": (ln.get("item") or "").strip().lower(), "product": (ln.get("product") or "").strip(),
                "manufacturer": (ln.get("manufacturer") or "").strip(), "size_color": (ln.get("size_color") or "").strip(),
                "qty": "" if qty in ("", None) else (int(qty) if float(qty).is_integer() else qty),
                "unit": u or "", "unit_price": p if p is not None else "", "price_per_sf": psf if psf is not None else "",
                "unit_type": unit_type(cls, f"{ln.get('item', '')} {ln.get('product', '')} {ln.get('size_color', '')}"),
                "tax": tax, "hold_until": hold, "latest": "", "notes": notes})
        if hard:
            problems.append(f"{tag}: " + "; ".join(hard))
        else:
            rows.extend(out)
            qrows.append({"quote_key": qkey(q["source_id"]), "quote_date": qd, "supplier": q["supplier"].strip(),
                          "class": cls, "project": (q.get("project") or "").strip(), "city": (q.get("city") or "").strip(),
                          "quote_ref": str(q.get("quote_ref") or "").strip(), "freight": (q.get("freight") or "").strip(),
                          "tax": tax, "hold_until": hold, "lines": len(out), "source_url": (q.get("source_url") or "").strip()})
    return rows, qrows, problems


def load_csv(path):
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))
    except FileNotFoundError:
        return []


def mark_latest(rows):
    best = {}
    for r in rows:
        k = (r["supplier"].lower(), r["item"], r["unit"], r["size_color"].lower())
        if k not in best or (r["quote_date"], r["quote_id"]) > (best[k]["quote_date"], best[k]["quote_id"]):
            best[k] = r
    ids = {id(r) for r in best.values()}
    for r in rows:
        r["latest"] = "y" if id(r) in ids else ""


def write_csv(rows, path, cols=COLS, key="quote_id"):
    rows.sort(key=lambda r: (r["quote_date"], r["supplier"], r[key]), reverse=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore", lineterminator="\n")
        w.writeheader()
        w.writerows(rows)


def read_quotes(path):
    data = json.load(open(path, encoding="utf-8"))
    return data if isinstance(data, list) else [data]


def is_dir_target(path):
    import os
    return os.path.isdir(path) or not path.lower().endswith(".csv")


def bucket(r):
    """Folder layout: 'Price Book YYYY-MM.csv' per quote month; price lists in 'Price Lists YYYY.csv'."""
    if r["project"].lower().startswith("price list"):
        return f"Price Lists {r['quote_date'][:4]}.csv"
    return f"Price Book {r['quote_date'][:7]}.csv"


def load_folder(folder):
    import glob, os
    rows, qrows = [], []
    for f in sorted(glob.glob(os.path.join(folder, "Price *.csv"))):
        if " Quotes " in os.path.basename(f):
            qrows += load_csv(f)
        else:
            rows += load_csv(f)
    return rows, qrows


def cmd_append_folder(a):
    import os
    os.makedirs(a.csv, exist_ok=True)
    new, qnew, problems = flatten(read_quotes(a.quotes))
    for p in problems:
        print("⚠ not written — " + p)
    newkeys = {q["quote_key"] for q in qnew}
    rows, qrows = load_folder(a.csv)
    before = {}
    for r in rows:
        before.setdefault(r["quote_key"], bucket(r))
    rows = [r for r in rows if r["quote_key"] not in newkeys] + new
    qrows = [q for q in qrows if q["quote_key"] not in newkeys] + qnew
    mark_latest(rows)
    qb = {}
    for r in rows:
        qb.setdefault(r["quote_key"], bucket(r))
    touched = {qb[k] for k in newkeys} | {before[k] for k in newkeys if k in before}
    # 'latest' flags can flip in older months too: rewrite any file whose content changed
    files = {}
    for r in rows:
        files.setdefault(bucket(r), []).append(r)
    changed = []
    for name, rs in sorted(files.items()):
        path = os.path.join(a.csv, name)
        old = open(path, encoding="utf-8").read() if os.path.exists(path) else None
        write_csv(rs, path)
        qname = name.replace("Price Book", "Price Book Quotes").replace("Price Lists", "Price Lists Quotes")
        qpath = os.path.join(a.csv, qname)
        keys = {r["quote_key"] for r in rs}
        oldq = open(qpath, encoding="utf-8").read() if os.path.exists(qpath) else None
        write_csv([q for q in qrows if q["quote_key"] in keys], qpath, QCOLS, "quote_key")
        for pth, o in ((path, old), (qpath, oldq)):
            if o != open(pth, encoding="utf-8").read():
                changed.append((os.path.basename(pth), os.path.getsize(pth)))
    for name in touched - set(files):
        print(f"⚠ {name} now empty — delete it")
    print(f"Price Book: +{len(new)} lines from {len(newkeys)} quotes; {len(rows)} lines / {len(qrows)} quotes")
    print("Upload these files (whole, replace):")
    for n, sz in changed:
        print(f"  {n}  ({sz // 1024} KB)" + ("  ⚠ over 900 KB" if sz > 900_000 else ""))
    return 1 if problems else 0


def all_rows(path):
    return load_folder(path)[0] if is_dir_target(path) else load_csv(path)


def cmd_append(a):
    if is_dir_target(a.csv):
        return cmd_append_folder(a)
    new, qnew, problems = flatten(read_quotes(a.quotes))
    for p in problems:
        print("⚠ not written — " + p)
    newkeys = {q["quote_key"] for q in qnew}
    rows = [r for r in load_csv(a.csv) if r["quote_key"] not in newkeys] + new
    qpath = quotes_path(a.csv)
    qrows = [q for q in load_csv(qpath) if q["quote_key"] not in newkeys] + qnew
    mark_latest(rows)
    write_csv(rows, a.csv)
    write_csv(qrows, qpath, QCOLS, "quote_key")
    import os
    sizes = os.path.getsize(a.csv), os.path.getsize(qpath)
    print(f"Price Book: +{len(new)} lines from {len(newkeys)} quotes; {len(rows)} lines / {len(qrows)} quotes "
          f"({sizes[0] // 1024} KB / {sizes[1] // 1024} KB)")
    if max(sizes) > 900_000:
        print("⚠ a Price Book file is over 900 KB — start a new yearly file")
    return 1 if problems else 0


def cmd_check(a):
    rows, _, problems = flatten(read_quotes(a.quotes))
    for p in problems:
        print("⚠ " + p)
    print(f"{len(rows)} valid lines; {len(problems)} quotes with problems")
    return 1 if problems else 0


def cmd_latest(a):
    rows = all_rows(a.csv)
    mark_latest(rows)  # recompute across every file read (a run may have refreshed only one month)
    rows = [r for r in rows if r["latest"] == "y"]
    if a.item:
        rx = re.compile(a.item, re.I)
        rows = [r for r in rows if rx.search(" ".join([r["item"], r["product"], r["size_color"], r["manufacturer"]]))]
    if a.days:
        cut = (dt.date.today() - dt.timedelta(days=a.days)).isoformat()
        rows = [r for r in rows if r["quote_date"] >= cut]
    rows.sort(key=lambda r: (r["item"], float(r["unit_price"] or 0)))
    for r in rows:
        price = f"${float(r['unit_price']):,.2f}/{r['unit']}" if r["unit_price"] else ""
        psf = f" (${float(r['price_per_sf']):,.2f}/SF)" if r["price_per_sf"] else ""
        print(f"{r['item']:<28} {r['supplier'][:28]:<28} {price}{psf}  {r['product'][:40]}  {r['quote_date']}  {r['project'][:30]}")
    print(f"{len(rows)} items")


def cmd_movers(a):
    rows = all_rows(a.csv)
    groups = defaultdict(list)
    for r in rows:
        if r["unit_price"] and r["unit"] != "LS" and r["item"] not in ("other", "delivery", "fuel surcharge") and "allowance" not in r["notes"].lower():
            prod = re.sub(r"[^a-z0-9]+", " ", r["product"].lower()).strip()
            groups[(r["supplier"], r["item"], prod, r["unit"], r["size_color"].lower())].append(r)
    today = dt.date.today()
    out = []
    for (sup, item, prod, unit, size), rs in groups.items():
        rs.sort(key=lambda r: r["quote_date"])
        last = rs[-1]
        if (today - dt.date.fromisoformat(last["quote_date"])).days > a.days:
            continue
        prior = [r for r in rs if (dt.date.fromisoformat(last["quote_date"]) - dt.date.fromisoformat(r["quote_date"])).days >= 30]
        if not prior:
            continue
        p0, p1 = float(prior[-1]["unit_price"]), float(last["unit_price"])
        if p0 <= 0:
            continue
        pct = (p1 - p0) / p0 * 100
        if abs(pct) >= a.min_pct:
            out.append((pct, f"{rs[-1]['product'][:40]} · {sup}: ${p0:,.2f} → ${p1:,.2f}/{unit} "
                              f"({pct:+.1f}%, {prior[-1]['quote_date']} → {last['quote_date']})"))
    out.sort(key=lambda x: -abs(x[0]))
    print("Price movers (last %d days, ≥ %.0f%%):" % (a.days, a.min_pct) if out else "Price movers: none")
    for _, line in out[:15]:
        print("· " + line)


LATEST_COLS = ["item", "supplier", "product", "size_color", "unit", "unit_price", "price_per_sf", "quote_date",
               "project", "hold_until", "prior_price", "prior_date", "change_pct", "mover", "notes"]


def _prod(r):
    return re.sub(r"[^a-z0-9]+", " ", r["product"].lower()).strip()


def cmd_export(a):
    """One small file for the Dashboard: every latest row, with the change vs the same supplier's same product
    quoted at least 30 days earlier. mover = y when that change is >= --min-pct and the latest quote is within --days."""
    rows = all_rows(a.csv)
    mark_latest(rows)
    hist = defaultdict(list)
    for r in rows:
        if r["unit_price"]:
            hist[(r["supplier"], r["item"], _prod(r), r["unit"], r["size_color"].lower())].append(r)
    today = dt.date.today()
    out = []
    for r in rows:
        if r["latest"] != "y":
            continue
        o = {k: r.get(k, "") for k in LATEST_COLS}
        o["notes"] = o["notes"][:80]
        o.update(prior_price="", prior_date="", change_pct="", mover="")
        if r["unit_price"] and r["unit"] != "LS" and r["item"] not in ("other", "delivery", "fuel surcharge") \
                and "allowance" not in r["notes"].lower():
            d1 = dt.date.fromisoformat(r["quote_date"])
            prior = sorted((x for x in hist[(r["supplier"], r["item"], _prod(r), r["unit"], r["size_color"].lower())]
                            if (d1 - dt.date.fromisoformat(x["quote_date"])).days >= 30), key=lambda x: x["quote_date"])
            if prior and float(prior[-1]["unit_price"]) > 0:
                p0, p1 = float(prior[-1]["unit_price"]), float(r["unit_price"])
                pct = (p1 - p0) / p0
                o.update(prior_price=prior[-1]["unit_price"], prior_date=prior[-1]["quote_date"],
                         change_pct=f"{pct:.4f}",
                         mover="y" if abs(pct) * 100 >= a.min_pct and (today - d1).days <= a.days else "")
        out.append(o)
    out.sort(key=lambda r: (r["item"], r["supplier"].lower(), r["product"].lower(), r["size_color"].lower()))
    with open(a.out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LATEST_COLS, lineterminator="\n")
        w.writeheader(); w.writerows(out)
    import os
    mpath = os.path.join(os.path.dirname(os.path.abspath(a.out)), "Price Movers.csv")
    with open(mpath, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LATEST_COLS, lineterminator="\n")
        w.writeheader(); w.writerows(r for r in out if r["mover"] == "y")
    print(f"{len(out)} latest rows ({sum(r['mover'] == 'y' for r in out)} movers) -> {a.out} ({os.path.getsize(a.out) // 1024} KB)"
          f" + {mpath} ({os.path.getsize(mpath)} bytes, movers only — the Dashboard reads this one)")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("append"); p.add_argument("quotes"); p.add_argument("csv"); p.set_defaults(f=cmd_append)
    p = sp.add_parser("check"); p.add_argument("quotes"); p.set_defaults(f=cmd_check)
    p = sp.add_parser("latest"); p.add_argument("csv"); p.add_argument("--item"); p.add_argument("--days", type=int); p.set_defaults(f=cmd_latest)
    p = sp.add_parser("movers"); p.add_argument("csv"); p.add_argument("--days", type=int, default=90)
    p.add_argument("--min-pct", type=float, default=3); p.set_defaults(f=cmd_movers)
    p = sp.add_parser("export"); p.add_argument("csv"); p.add_argument("out"); p.add_argument("--days", type=int, default=90)
    p.add_argument("--min-pct", type=float, default=3); p.set_defaults(f=cmd_export)
    a = ap.parse_args()
    sys.exit(a.f(a) or 0)


if __name__ == "__main__":
    main()
