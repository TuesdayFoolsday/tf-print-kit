"""Turn an estimating export (Excel 2003 XML: sheets Material + Grand Totals) into one bid-metrics row.

    python3 metrics/bid_metrics.py <export.XML | compact.json> <cost_map.csv> [--project NAME] [--bid-id ID] [--json]
    python3 metrics/bid_metrics.py <export.XML> - --compact      # export -> compact JSON

Buckets labor/material by unit type using the cost map, spreads support items (grout, rebar, mortar,
anchors, flashing...) onto the units they serve, loads labor and material to selling price using the
export's own burden, tax and OH&P, and reconciles the bucketed total to the export's bid price.
"""
import argparse, csv, json, re, sys
import xml.etree.ElementTree as ET

SS = "urn:schemas-microsoft-com:office:spreadsheet"
UNIT_TYPES = ["cmu_gray", "cmu_arch", "brick", "stone", "cast_stone"]
OTHER = ["insulation", "scaffold", "other_items"]


def sheets(path):
    root = ET.parse(path).getroot()
    out = {}
    for ws in root.findall(f"{{{SS}}}Worksheet"):
        rows = []
        for r in ws.iter(f"{{{SS}}}Row"):
            row, ci = [], 0
            for c in r.findall(f"{{{SS}}}Cell"):
                idx = c.get(f"{{{SS}}}Index")
                if idx:
                    ci = int(idx) - 1
                while len(row) < ci:
                    row.append("")
                d = c.find(f"{{{SS}}}Data")
                row.append(d.text if d is not None and d.text else "")
                ci += 1
            rows.append(row)
        out[ws.get(f"{{{SS}}}Name")] = rows
    return out


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def load_map(path):
    rules = []
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rules.append((r["class"].strip(), re.compile(r["pattern"], re.I), r["bucket"].strip()))
    return rules


def bucket_for(rules, cls, mid, desc):
    text = f"{mid} {desc}"
    for c, pat, b in rules:
        if c == cls and pat.search(text):
            return b
    return None


def material_rows(wb):
    rows = []
    for r in wb["Material"]:
        if len(r) > 17 and r[1] and r[1] != "Class" and (num(r[4]) or num(r[16]) or num(r[12])):
            rows.append(dict(cls=r[1].strip(), id=r[2].strip(), desc=r[3].strip(), amount=num(r[4]), units=r[5],
                             qty=num(r[6]), qty_units=r[7], labor=num(r[12]), material=num(r[16]), mandays=num(r[17])))
    return rows


def grand_totals(wb):
    g = {"misc_items": {}}
    for r in wb["Grand Totals"]:
        cells = [c.strip() for c in r]
        label = next((c for c in cells if c), "")
        if label == "Material Costs":
            g["material"] = num(cells[4] if len(cells) > 4 else 0) or num(next((c for c in cells[1:] if num(c)), 0))
        elif label == "Misc. Taxable Costs":
            g["misc_taxable"] = num(next((c for c in cells[1:] if num(c)), 0))
        elif label == "Tax":
            vals = [num(c) for c in cells if num(c)]
            g["tax_rate"] = vals[0] if vals else 0.0
        elif label.startswith("Material + Misc"):
            vals = [num(c) for c in cells if num(c)]
            g["mat_cost_tax"], g["mat_oh"], g["mat_profit"], g["mat_ohp"], g["mat_total"] = vals[:5]
        elif label == "Crew Labor Cost":
            vals = [num(c) for c in cells if num(c)]
            g["crew_cost"], g["crew_oh"], g["crew_profit"], g["crew_ohp"], g["crew_total"] = vals[:5]
        elif label == "Equipment":
            vals = [num(c) for c in cells if num(c)]
            g["equip_total"] = vals[4] if len(vals) > 4 else (vals[-2] if len(vals) > 1 else 0)
        elif label == "Misc. Sub-Total":
            vals = [num(c) for c in cells if num(c)]
            g["misc_total"] = vals[4] if len(vals) > 4 else (vals[-2] if len(vals) > 1 else 0)
        elif label == "Subbed Out":
            vals = [num(c) for c in cells if num(c)]
            g["sub_total"] = vals[4] if len(vals) > 4 else 0
        elif label == "Cleaning Cost":
            vals = [num(c) for c in cells if num(c)]
            g["clean_total"] = vals[-2] if len(vals) > 1 else 0
        elif label == "Percent of Cost":
            vals = [num(c) for c in cells if num(c)]
            g["oh_pct"], g["profit_pct"] = (vals + [0, 0])[:2]
        elif label == "Grand Totals":
            vals = [num(c) for c in cells if num(c)]
            g["bid_price"] = vals[-1] if vals else 0
        elif "Performance Bond" in " ".join(cells):
            g["bond_note"] = next(c for c in cells if "Performance Bond" in c)
    for r in wb["Material"]:
        if len(r) > 16 and r[11].strip() == "Totals:":
            g["labor_base_total"], g["material_sheet_total"] = num(r[12]), num(r[16])
    g["job_title"] = (wb["Grand Totals"][0] + [""] * 6)[4].strip()
    jn = [c for c in wb["Grand Totals"][0] if c.strip()]
    g["job_number"] = jn[-1] if jn else ""
    return g


COMPACT_ROW = ["cls", "id", "desc", "amount", "units", "qty", "qty_units", "labor", "material", "mandays"]
COMPACT_TOTALS = ["material", "misc_taxable", "tax_rate", "mat_total", "crew_total", "equip_total", "misc_total",
                  "sub_total", "clean_total", "oh_pct", "profit_pct", "bid_price", "labor_base_total"]


def to_compact(path):
    """Export -> small JSON a run can transcribe or check: {job_title, job_number, rows:[[...]], totals:{...}}."""
    wb = sheets(path)
    g = grand_totals(wb)
    return {"job_title": g["job_title"], "job_number": g["job_number"],
            "rows": [[r[k] for k in COMPACT_ROW] for r in material_rows(wb)],
            "totals": {k: g.get(k, 0) for k in COMPACT_TOTALS}}


def load(path):
    if path.lower().endswith(".json"):
        c = json.load(open(path, encoding="utf-8"))
        rows = [dict(zip(COMPACT_ROW, r)) for r in c["rows"]]
        for r in rows:
            for k in ("amount", "qty", "labor", "material", "mandays"):
                r[k] = num(r[k])
            r["qty_units"] = r.get("qty_units") or ""
        g = {k: num(c["totals"].get(k, 0)) for k in COMPACT_TOTALS}
        g["job_title"], g["job_number"] = c.get("job_title", ""), c.get("job_number", "")
        return rows, g
    wb = sheets(path)
    return material_rows(wb), grand_totals(wb)


def compute(path, map_path, project="", bid_id=""):
    rules = load_map(map_path)
    rows, g = load(path)
    B = {b: dict(sf=0.0, pieces=0.0, labor=0.0, material=0.0, mandays=0.0) for b in UNIT_TYPES + OTHER}
    support = {"cmu": [], "veneer": [], "colored_mortar": [], "all": []}
    for r in rows:
        b = bucket_for(rules, r["cls"], r["id"], r["desc"]) or "other_items"
        if b in B:
            x = B[b]
            if b in UNIT_TYPES:
                x["sf"] += r["amount"]
                x["pieces"] += r["qty"] if r["qty_units"].lower().startswith("piece") else 0
            elif b == "insulation":
                x["sf"] += r["amount"]
            x["labor"] += r["labor"]; x["material"] += r["material"]; x["mandays"] += r["mandays"]
        else:  # support:<target>
            support.setdefault(b.split(":", 1)[1], []).append(r)

    def spread(items, targets):
        weights = {t: B[t]["sf"] for t in targets if B[t]["sf"] > 0}
        tot = sum(weights.values())
        if not tot:
            weights = {"other_items": 1.0}; tot = 1.0
        for r in items:
            for t, w in weights.items():
                B[t]["labor"] += r["labor"] * w / tot
                B[t]["material"] += r["material"] * w / tot
                B[t]["mandays"] += r["mandays"] * w / tot

    spread(support.get("cmu", []), ["cmu_gray", "cmu_arch"])
    spread(support.get("veneer", []), ["brick", "stone", "cast_stone"])
    veneer_present = any(B[t]["sf"] > 0 for t in ("brick", "stone", "cast_stone"))
    spread(support.get("colored_mortar", []), ["brick", "stone", "cast_stone"] if veneer_present else ["cmu_arch", "cmu_gray"])
    spread(support.get("all", []), UNIT_TYPES)

    lab_sum = sum(r["labor"] for r in rows)
    mat_sum = sum(r["material"] for r in rows)
    labor_factor = g.get("crew_total", 0) / lab_sum if lab_sum else 0          # base labor -> burdened + OH&P
    mat_factor = (g.get("mat_total", 0) * (g.get("material", mat_sum) / (g.get("material", mat_sum) + g.get("misc_taxable", 0)))) / mat_sum if mat_sum else 0
    misc_taxable_sell = g.get("mat_total", 0) - mat_sum * mat_factor

    row = {"bid_id": bid_id, "project": project or g["job_title"], "job_number": g["job_number"], "source": "tradesmen",
           "bid_price": round(g.get("bid_price", 0)), "oh_pct": g.get("oh_pct", 0), "profit_pct": g.get("profit_pct", 0),
           "tax_rate": g.get("tax_rate", 0), "labor_load_factor": round(labor_factor, 4), "material_load_factor": round(mat_factor, 4),
           "equipment_sell": round(g.get("equip_total", 0)), "misc_sell": round(g.get("misc_total", 0) + misc_taxable_sell),
           "subs_sell": round(g.get("sub_total", 0)), "cleaning_sell": round(g.get("clean_total", 0))}
    sold = 0.0
    for b in UNIT_TYPES + OTHER:
        x = B[b]
        lab_sell, mat_sell = x["labor"] * labor_factor, x["material"] * mat_factor
        sold += lab_sell + mat_sell
        row[f"{b}_sf"] = round(x["sf"], 1)
        if b in UNIT_TYPES:
            row[f"{b}_pieces"] = round(x["pieces"])
        row[f"{b}_labor_base"] = round(x["labor"])
        row[f"{b}_material_base"] = round(x["material"])
        row[f"{b}_mandays"] = round(x["mandays"], 1)
        row[f"{b}_sell"] = round(lab_sell + mat_sell)
        if b in UNIT_TYPES and x["sf"] > 0:
            row[f"{b}_labor_per_sf"] = round(x["labor"] / x["sf"], 2)
            row[f"{b}_material_per_sf"] = round(x["material"] / x["sf"], 2)
            row[f"{b}_sell_per_sf"] = round((lab_sell + mat_sell) / x["sf"], 2)
            row[f"{b}_sf_per_manday"] = round(x["sf"] / x["mandays"], 1) if x["mandays"] else None
            if x["pieces"]:
                row[f"{b}_labor_per_piece"] = round(x["labor"] / x["pieces"], 2)
                row[f"{b}_pieces_per_manday"] = round(x["pieces"] / x["mandays"], 1) if x["mandays"] else None
    total = sold + row["equipment_sell"] + row["misc_sell"] + row["subs_sell"] + row["cleaning_sell"]
    row["bucketed_total"] = round(total)
    # Three checks catch a bad transcription or a misread export: rows sum to the sheet totals, and the
    # Grand Totals parts sum to the bid price. Buckets themselves always sum by construction.
    parts = sum(g.get(k, 0) for k in ("mat_total", "crew_total", "equip_total", "misc_total", "sub_total", "clean_total"))
    problems = []
    if abs(lab_sum - g.get("labor_base_total", lab_sum)) > 1:
        problems.append(f"labor rows {lab_sum:,.0f} vs sheet total {g['labor_base_total']:,.0f}")
    if abs(mat_sum - g.get("material", mat_sum)) > 1:
        problems.append(f"material rows {mat_sum:,.0f} vs Material Costs {g['material']:,.0f}")
    if abs(parts - g.get("bid_price", 0)) > 2:
        problems.append(f"Grand Totals parts {parts:,.0f} vs bid price {g.get('bid_price', 0):,.0f}")
    row["reconciles"] = "n" if problems else "y"
    row["reconcile_notes"] = "; ".join(problems)
    row["mandays_total"] = round(sum(r["mandays"] for r in rows), 1)
    row["wall_sf_total"] = round(sum(B[b]["sf"] for b in UNIT_TYPES), 1)
    row["sell_per_wall_sf"] = round(g.get("bid_price", 0) / row["wall_sf_total"], 2) if row["wall_sf_total"] else None
    unmapped = sorted({f'{r["cls"]}:{r["id"]}' for r in rows if bucket_for(rules, r["cls"], r["id"], r["desc"]) is None})
    row["unmapped_items"] = "; ".join(unmapped)
    return row


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("export"); ap.add_argument("cost_map")
    ap.add_argument("--project", default=""); ap.add_argument("--bid-id", default="")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--compact", action="store_true", help="print the export as compact JSON and exit")
    a = ap.parse_args()
    if a.compact:
        print(json.dumps(to_compact(a.export), separators=(",", ":"))); sys.exit()
    r = compute(a.export, a.cost_map, a.project, a.bid_id)
    if a.json:
        print(json.dumps(r, indent=1))
    else:
        w = csv.DictWriter(sys.stdout, fieldnames=list(r)); w.writeheader(); w.writerow(r)
