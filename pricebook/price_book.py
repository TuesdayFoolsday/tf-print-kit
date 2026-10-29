"""Price Book: supplier quotes -> one CSV row per quoted line item.

    python3 pricebook/price_book.py append <quotes.json> <price_book.csv>      # add/refresh rows, rewrite whole file
    python3 pricebook/price_book.py latest <price_book.csv> [--item TEXT] [--days N]
    python3 pricebook/price_book.py movers <price_book.csv> [--days 90] [--min-pct 3]
    python3 pricebook/price_book.py check  <quotes.json>                        # validate only, no write

quotes.json = a list of quotes (or one quote object):
  {"source_id": "<email message id or file item id>", "source_url": "<webLink/webUrl>",
   "quote_date": "YYYY-MM-DD", "supplier": "...", "class": "brick|cmu|stone|cast_stone|mortar_grout|
   accessories|insulation|rebar|sub_stucco|sub_waterproofing|sub_caulking|other",
   "project": "...", "city": "...", "quote_ref": "quote no / rev", "hold_days": 30 | "hold_until": "YYYY-MM-DD",
   "freight": "...", "tax": "excluded|included|unknown",
   "lines": [{"item": "brick modular", "product": "as written", "manufacturer": "...", "size_color": "...",
              "qty": 90000, "unit": "M", "unit_price": 800, "price_per_sf": 5.52, "notes": "..."}]}

Rows are keyed by source_id + line number, so re-running the same quote replaces its rows (never duplicates).
`latest` = y on the newest row per supplier + item + unit + size_color.
"""
import argparse, csv, datetime as dt, hashlib, json, re, sys
from collections import defaultdict

COLS = ["quote_id", "quote_date", "supplier", "class", "project", "city", "item", "product", "manufacturer",
        "size_color", "qty", "unit", "unit_price", "price_per_sf", "unit_type", "freight", "tax", "hold_until",
        "quote_ref", "latest", "source_url", "notes"]
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


def qid(source_id, n):
    return hashlib.sha1(str(source_id).encode()).hexdigest()[:10] + f"-{n:02d}"


def flatten(quotes):
    """Validate quotes, return (rows, problems). A quote with any hard problem contributes no rows."""
    rows, problems = [], []
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
                "quote_id": qid(q.get("source_id"), n), "quote_date": qd, "supplier": q.get("supplier", "").strip(),
                "class": cls, "project": (q.get("project") or "").strip(), "city": (q.get("city") or "").strip(),
                "item": (ln.get("item") or "").strip().lower(), "product": (ln.get("product") or "").strip(),
                "manufacturer": (ln.get("manufacturer") or "").strip(), "size_color": (ln.get("size_color") or "").strip(),
                "qty": "" if qty in ("", None) else (int(qty) if float(qty).is_integer() else qty),
                "unit": u or "", "unit_price": p if p is not None else "", "price_per_sf": psf if psf is not None else "",
                "unit_type": unit_type(cls, f"{ln.get('item', '')} {ln.get('product', '')} {ln.get('size_color', '')}"),
                "freight": (q.get("freight") or "").strip(), "tax": tax, "hold_until": hold,
                "quote_ref": str(q.get("quote_ref") or "").strip(), "latest": "",
                "source_url": (q.get("source_url") or "").strip(), "notes": notes})
        if hard:
            problems.append(f"{tag}: " + "; ".join(hard))
        else:
            rows.extend(out)
    return rows, problems


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


def write_csv(rows, path):
    rows.sort(key=lambda r: (r["quote_date"], r["supplier"], r["quote_id"]), reverse=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def read_quotes(path):
    data = json.load(open(path, encoding="utf-8"))
    return data if isinstance(data, list) else [data]


def cmd_append(a):
    new, problems = flatten(read_quotes(a.quotes))
    for p in problems:
        print("⚠ not written — " + p)
    old = load_csv(a.csv)
    newids = {r["quote_id"].rsplit("-", 1)[0] for r in new}
    kept = [r for r in old if r["quote_id"].rsplit("-", 1)[0] not in newids]
    rows = kept + new
    mark_latest(rows)
    write_csv(rows, a.csv)
    size = sum(len(",".join(str(r.get(c, "")) for c in COLS)) + 1 for r in rows)
    print(f"Price Book: +{len(new)} lines from {len(newids)} quotes; {len(rows)} rows total (~{size // 1024} KB)")
    if size > 900_000:
        print("⚠ Price Book is over 900 KB — start a new yearly file")
    return 1 if problems else 0


def cmd_check(a):
    rows, problems = flatten(read_quotes(a.quotes))
    for p in problems:
        print("⚠ " + p)
    print(f"{len(rows)} valid lines; {len(problems)} quotes with problems")
    return 1 if problems else 0


def cmd_latest(a):
    rows = [r for r in load_csv(a.csv) if r["latest"] == "y"]
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
    rows = load_csv(a.csv)
    groups = defaultdict(list)
    for r in rows:
        if r["unit_price"] and r["unit"] != "LS":
            groups[(r["supplier"], r["item"], r["unit"], r["size_color"].lower())].append(r)
    today = dt.date.today()
    out = []
    for (sup, item, unit, size), rs in groups.items():
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
            out.append((pct, f"{item} ({size or 'any'}) · {sup}: ${p0:,.2f} → ${p1:,.2f}/{unit} "
                              f"({pct:+.1f}%, {prior[-1]['quote_date']} → {last['quote_date']})"))
    out.sort(key=lambda x: -abs(x[0]))
    print("Price movers (last %d days, ≥ %.0f%%):" % (a.days, a.min_pct) if out else "Price movers: none")
    for _, line in out[:15]:
        print("· " + line)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("append"); p.add_argument("quotes"); p.add_argument("csv"); p.set_defaults(f=cmd_append)
    p = sp.add_parser("check"); p.add_argument("quotes"); p.set_defaults(f=cmd_check)
    p = sp.add_parser("latest"); p.add_argument("csv"); p.add_argument("--item"); p.add_argument("--days", type=int); p.set_defaults(f=cmd_latest)
    p = sp.add_parser("movers"); p.add_argument("csv"); p.add_argument("--days", type=int, default=90)
    p.add_argument("--min-pct", type=float, default=3); p.set_defaults(f=cmd_movers)
    a = ap.parse_args()
    sys.exit(a.f(a) or 0)


if __name__ == "__main__":
    main()
