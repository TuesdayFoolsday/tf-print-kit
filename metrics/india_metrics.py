"""India estimating workbook (takeoff tab + 'Proposal' pricing tabs) -> one Bid Metrics row per project.

    python3 metrics/india_metrics.py <workbook.xlsx> <cost_map.csv> --tabs "Tab A" "Tab B" ... --project NAME
        [--metrics-id in-<slug>] [--bid-id ID] [--status final|iteration|draft] [--notes TEXT] [--json]

How it reads a pricing tab (one tab per building; a project's tabs are summed into one row):
  * columns: B name · E final qty · F unit · G/H material unit price/tax · I material subtotal ·
    J/K labor unit price/burden · L labor subtotal · M/N equipment unit price/tax · O equipment subtotal.
    Base (before tax/burden) = subtotal x unit / (unit + tax-or-burden), so base wage is read, not assumed.
  * The markup chain (general conditions -> overhead -> profit) is one multiplier per tab:
    estimate total / sub-total. Every bucket's sell = its cost x that multiplier.
  * Crew labor (LABOR COST section: foreman, bricklayers, laborers, operators) is spread onto unit types by
    the takeoff's lay-rate man-days: the tab's crew-days cell points at a total cell in the takeoff, and the
    lay-rate rows above it (down to the previous building's total) give net SF and days per wall type.
    That is India's own production split, the analogue of Tradesmen's line-item labor.
  * Cleaning lines go to their unit type's labor (support:<unit type>, no SF; David 2026-11-02: cleaning is
    part of the unit labor), falling back to the sister type when that type has no lay-rate SF.
  * Item labor (grout, rebar, flashing, insulation) and material go by item name through the Cost Map rows
    with class INDIA; support items are spread by SF exactly as bid_metrics.py does.
  * Unit-type SF = lay-rate net SF (CMU counted in EA converts at 1.125 EA/SF). Material on a unit type with
    no lay-rate SF (e.g. burnished CMU priced but laid under a plain "8'' CMU" lay rate) folds into the
    sister type (cmu_arch -> cmu_gray, cast_stone/stone -> veneer present) and is noted.
  * man-days = bricklayer crew days (production masons), like Tradesmen.
Reconciliation: line subtotals must sum to the tab's Column Totals and Sub-total, and the markup chain to
its Estimate Total; otherwise reconciles = n.
"""
import argparse, csv, json, re, sys, os
import openpyxl
from openpyxl.utils import column_index_from_string as col_ix

sys.path.insert(0, os.path.dirname(__file__))
from bid_metrics import STD_WAGES, UNIT_TYPES, OTHER, load_map

SECTIONS = {"CMU MATERIALS", "BRICK MATERIALS", "MATERIAL CLEANING", "LABOR COST"}
FAMILY = {"cmu_gray": ["cmu_gray", "cmu_arch"], "cmu_arch": ["cmu_arch", "cmu_gray"],
          "brick": ["brick", "stone", "cast_stone"], "stone": ["stone", "brick", "cast_stone"],
          "cast_stone": ["cast_stone", "stone", "brick"]}
ROLE = [("FOREMAN", "Forman"), ("BRICKLAYER", "Bricklayer"), ("LABORER", "Labor"), ("OPERATOR", "Operator")]


def n(x):
    return float(x) if isinstance(x, (int, float)) else 0.0


def base_of(sub, unit, extra):
    sub, unit, extra = n(sub), n(unit), n(extra)
    if not sub:
        return 0.0
    return sub * unit / (unit + extra) if unit + extra else sub


def unit_type(label):
    l = label.lower()
    if "cast stone" in l or "precast" in l:
        return "cast_stone"
    if "stone" in l:
        return "stone"
    if "brick" in l:
        return "brick"
    if "cmu" in l or "block" in l:
        return "cmu_arch" if re.search(r"split|burnish|decor|ground|scored|glaz|trenwyth|color", l) else "cmu_gray"
    return None


def india_bucket(rules, section, name):
    if section == "LABOR COST":
        return "crew"
    for c, pat, b in rules:
        if c == "INDIA" and pat.search(name):
            return b
    return None


def crew_cell(ws_f):
    for c in ws_f[1]:
        if isinstance(c.value, str) and "!" in c.value:
            return c.value, re.findall(r"'?([^'!+=]+)'?!([A-Z]+)(\d+)", c.value)
    return None, []


def lay_rows(wbv, refs, all_refs):
    """Lay-rate rows feeding one pricing tab: [(type, sf, days, label)]."""
    out = []
    for sh, col, row in refs:
        ws, c, R = wbv[sh], col_ix(col), int(row)
        prev = [x for x in all_refs.get((sh, col), []) if x < R]
        lo = max(prev) if prev else 0
        found, loose = [], []
        for r in range(lo + 1, R):
            v = ws.cell(r, c).value
            if not isinstance(v, (int, float)):
                continue
            for cc in range(c - 1, max(0, c - 14), -1):
                s = ws.cell(r, cc).value
                if not isinstance(s, str) or not s.strip():
                    continue
                if s.strip().upper().rstrip(".") in ("DAYS", "DAY", "SF", "EA", "LF", "CY", "TN", "TONS", "MEN", "%"):
                    continue          # unit cells between the label and the days column
                lab = s.strip()
                nums = [ws.cell(r, k).value for k in range(cc + 1, cc + 4)]
                qty = next((x for x in nums if isinstance(x, (int, float))), 0)
                unit = next((ws.cell(r, k).value for k in range(cc + 1, cc + 5)
                             if isinstance(ws.cell(r, k).value, str)), "")
                rec = (unit_type(lab), qty, unit, v, lab)
                if lab.lower().startswith("lay rate"):
                    found.append(rec)
                elif unit_type(lab) and not re.match(r"\s*(total|mortar|grout|rebar|wire|#\d)", lab, re.I) \
                        and not re.search(r"(mortar|grout|wire rein|rebar)\s*$", lab, re.I):
                    loose.append(rec)
                break
        for t, qty, unit, days, lab in (found or loose):
            sf = n(qty)
            if str(unit).strip().upper().startswith("EA") and t and t.startswith("cmu"):
                sf = sf / 1.125
            elif not str(unit).strip().upper().startswith("SF"):
                sf = 0.0          # LF / EA bands and caps: days count, SF doesn't
            out.append((t or "other_items", sf, n(days), lab))
    return out


def read_tab(wbv, wbf, tab, rules, all_refs):
    ws = wbv[tab]
    rows, sect = [], "TOP"
    totals = {}
    crew = {"Forman": [0.0, 0.0], "Bricklayer": [0.0, 0.0], "Labor": [0.0, 0.0], "Operator": [0.0, 0.0]}  # base $, wage
    tax = burden = None
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        r = list(r) + [None] * 20
        name = r[1].strip() if isinstance(r[1], str) else ""
        if i == 4:
            tax, burden = n(r[7]), n(r[10])
        if name in SECTIONS:
            sect = name
            continue
        if name == "Column Totals":
            totals["mat"], totals["lab"], totals["eq"] = n(r[8]), n(r[11]), n(r[14])
            continue
        if name == "Sub-Total":
            totals["sub"] = next((n(x) for x in r[2:] if isinstance(x, (int, float))), 0)
        elif name == "Estimate Total":
            totals["total"] = next((n(x) for x in r[2:] if isinstance(x, (int, float))), 0)
        elif name in ("General Conditions", "Overhead", "Profit"):
            nums = [x for x in r[2:] if isinstance(x, (int, float))]
            totals[name] = nums[:2]
        if i < 5 or not name or "total" in totals:
            continue
        if "sub" in totals:
            continue
        m, l, e = n(r[8]), n(r[11]), n(r[14])
        if not (m or l or e):
            continue
        b = india_bucket(rules, sect, name)
        rec = dict(section=sect, name=name, qty=n(r[4]), unit=r[5], bucket=b,
                   mat=m, lab=l, eq=e,
                   mat_base=base_of(r[8], r[6], r[7]), lab_base=base_of(r[11], r[9], r[10]),
                   eq_base=base_of(r[14], r[12], r[13]))
        if b == "crew":
            up = name.upper()
            role = next((rl for key, rl in ROLE if key in up), None)
            if role:
                crew[role][0] += rec["lab_base"]
                crew[role][1] = n(r[9]) / 8 if n(r[9]) else crew[role][1]
        rows.append(rec)
    f, refs = crew_cell(wbf[tab])
    lay = lay_rows(wbv, refs, all_refs)
    return dict(tab=tab, rows=rows, totals=totals, crew=crew, tax=tax, burden=burden, lay=lay, crew_formula=f)


def compute(path, map_path, tabs, project, metrics_id="", bid_id="", status="", notes=""):
    rules = load_map(map_path)
    wbv = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path, data_only=False)
    all_refs = {}
    for t in wbf.sheetnames:
        if "proposal" in t.lower() or t.startswith(("Argyle (", "HEB (", "Talty PS")):
            for sh, col, row in crew_cell(wbf[t])[1]:
                all_refs.setdefault((sh, col), []).append(int(row))
    B = {b: dict(sf=0.0, pieces=0.0, labor=0.0, labor_takeoff=0.0, material=0.0, mandays=0.0, sell=0.0)
         for b in UNIT_TYPES + OTHER}
    extra = dict(equipment=0.0, misc=0.0, cleaning=0.0, price=0.0, sub=0.0, crew_cost=0.0, equip_cost=0.0)
    hours = {k: 0.0 for k in STD_WAGES}
    wage_seen = {}
    problems, warn, unmapped = [], [], set()
    oh, prof, taxes = set(), set(), set()
    lab_base_all = lab_sell_all = mat_base_all = mat_sell_all = 0.0
    for tab in tabs:
        T = read_tab(wbv, wbf, tab, rules, all_refs)
        tt = T["totals"]
        mult = tt.get("total", 0) / tt["sub"] if tt.get("sub") else 0
        extra["price"] += tt.get("total", 0)
        extra["sub"] += tt.get("sub", 0)
        taxes.add(round(T["tax"], 4))
        if tt.get("Overhead"): oh.add(tt["Overhead"][0])
        if tt.get("Profit"): prof.add(tt["Profit"][0])
        # reconciliation
        sm, sl, se = (sum(r[k] for r in T["rows"]) for k in ("mat", "lab", "eq"))
        if abs(sm - tt.get("mat", 0)) > 1 or abs(sl - tt.get("lab", 0)) > 1 or abs(se - tt.get("eq", 0)) > 1:
            problems.append(f"{tab}: lines {sm + sl + se:,.0f} vs Column Totals {tt.get('mat', 0) + tt.get('lab', 0) + tt.get('eq', 0):,.0f}")
        if abs(sm + sl + se - tt.get("sub", 0)) > 1:
            problems.append(f"{tab}: lines vs Sub-Total {tt.get('sub', 0):,.0f}")
        g = tt.get("General Conditions", [0, 0]); o = tt.get("Overhead", [0, 0]); p = tt.get("Profit", [0, 0])
        chain = tt.get("sub", 0) * (1 + n(g[0])) * (1 + n(o[0])) * (1 + n(p[0]))
        if abs(chain - tt.get("total", 0)) > 2:
            problems.append(f"{tab}: markup chain {chain:,.0f} vs Estimate Total {tt.get('total', 0):,.0f}")
        # lay-rate weights
        lay = T["lay"]
        typed = [x for x in lay if x[0] in UNIT_TYPES]
        days_by = {}
        for t, sf, d, lab in typed:
            days_by[t] = days_by.get(t, 0) + d
            B[t]["sf"] += sf
        dtot = sum(days_by.values())
        if not dtot:
            warn.append(f"{tab}: no lay-rate rows — crew labor counted as other items")
            days_by, dtot = {"other_items": 1.0}, 1.0
        crew_base = sum(v[0] for v in T["crew"].values())
        crew_cost = sum(r["lab"] for r in T["rows"] if r["bucket"] == "crew")
        mason_days = T["crew"]["Bricklayer"][0] / (T["crew"]["Bricklayer"][1] * 8) if T["crew"]["Bricklayer"][1] else 0
        for role, (base, wage) in T["crew"].items():
            if wage:
                hours[role] += base / wage
                wage_seen.setdefault(role, set()).add(round(wage, 2))
        for t, d in days_by.items():
            w = d / dtot
            B[t]["labor"] += crew_base * w
            B[t]["labor_takeoff"] += crew_base * w
            B[t]["mandays"] += mason_days * w
            B[t]["sell"] += crew_cost * mult * w
        extra["crew_cost"] += crew_cost
        lab_base_all += crew_base; lab_sell_all += crew_cost * mult
        # items
        support, folded, nolay = {}, {}, {}
        tab_types = set(days_by)
        for r in T["rows"]:
            b = r["bucket"]
            if b == "crew":
                continue
            if b is None:
                unmapped.add(r["name"]); b = "other_items"
            cost = r["mat"] + r["lab"] + r["eq"]
            if b in ("equipment", "misc", "cleaning"):
                extra[b] += cost * mult
                if b == "equipment":
                    extra["equip_cost"] += cost
                continue
            if b in UNIT_TYPES and b not in tab_types:
                sister = {"cmu_arch": "cmu_gray", "cmu_gray": "cmu_arch"}.get(b)
                if sister in tab_types:
                    folded.setdefault((b, sister), []).append(r["name"])
                    b = sister
                else:
                    nolay.setdefault(b, []).append(r["name"])   # kept in its own bucket; no SF, so no $/SF
            if b.startswith("support:"):
                support.setdefault(b.split(":", 1)[1], []).append(r)
                continue
            x = B[b]
            x["labor"] += r["lab_base"]; x["labor_takeoff"] += r["lab_base"]
            x["material"] += r["mat_base"] + r["eq_base"]
            x["sell"] += cost * mult
            if b in UNIT_TYPES and str(r["unit"]).upper().startswith("EA") and b != "cast_stone":
                x["pieces"] += r["qty"]
            lab_base_all += r["lab_base"]; lab_sell_all += r["lab"] * mult
            mat_base_all += r["mat_base"] + r["eq_base"]; mat_sell_all += (r["mat"] + r["eq"]) * mult

        for (a, b2), names in folded.items():
            warn.append(f"{tab}: {a.replace('_', ' ')} priced ({len(names)} lines) but laid under {b2.replace('_', ' ')} lay rates — counted as {b2.replace('_', ' ')}")
        for b2, names in nolay.items():
            warn.append(f"{tab}: {b2.replace('_', ' ')} priced ({len(names)} lines) with no lay-rate SF — in its own bucket, no $/SF")

        def spread(items, targets):
            wts = {t: B[t]["sf"] for t in targets if t in tab_types and B[t]["sf"] > 0}
            tot = sum(wts.values())
            if not tot:
                wts, tot = {"other_items": 1.0}, 1.0
            for r in items:
                cost = r["mat"] + r["lab"] + r["eq"]
                for t, w in wts.items():
                    B[t]["labor"] += r["lab_base"] * w / tot
                    B[t]["labor_takeoff"] += r["lab_base"] * w / tot
                    B[t]["material"] += (r["mat_base"] + r["eq_base"]) * w / tot
                    B[t]["sell"] += cost * mult * w / tot
        for r in sum(support.values(), []):
            lab_base_all += r["lab_base"]; lab_sell_all += r["lab"] * mult
            mat_base_all += r["mat_base"] + r["eq_base"]; mat_sell_all += (r["mat"] + r["eq"]) * mult
        spread(support.get("cmu", []), ["cmu_gray", "cmu_arch"])
        spread(support.get("veneer", []), ["brick", "stone", "cast_stone"])
        ven = any(t in tab_types for t in ("brick", "stone", "cast_stone"))
        spread(support.get("colored_mortar", []), ["brick", "stone", "cast_stone"] if ven else ["cmu_arch", "cmu_gray"])
        spread(support.get("all", []), UNIT_TYPES)
        # support:<unit type> (cleaning crews): labor/material onto that unit type, no SF; sister type if absent
        for u in UNIT_TYPES:
            if support.get(u):
                tgt = next((t for t in FAMILY[u] if t in tab_types and B[t]["sf"] > 0), None)
                spread(support[u], [tgt] if tgt else [])

    price = extra["price"]
    row = {"metrics_id": metrics_id, "bid_id": bid_id, "project": project, "job_number": "", "source": "india",
           "estimate_status": status, "estimator": "India", "source_detail": "",
           "bid_price": round(price), "oh_pct": next(iter(oh)) if len(oh) == 1 else "",
           "profit_pct": next(iter(prof)) if len(prof) == 1 else "",
           "tax_rate": next(iter(taxes)) if len(taxes) == 1 else "",
           "labor_load_factor": round(lab_sell_all / lab_base_all, 4) if lab_base_all else "",
           "material_load_factor": round(mat_sell_all / mat_base_all, 4) if mat_base_all else "",
           "equipment_sell": round(extra["equipment"]), "misc_sell": round(extra["misc"]), "subs_sell": 0,
           "cleaning_sell": round(extra["cleaning"])}
    if len(prof) > 1:
        warn.append("profit differs by building: " + ", ".join(f"{p:.1%}" for p in sorted(prof)))
    sold = 0.0
    for b in UNIT_TYPES + OTHER:
        x = B[b]
        sold += x["sell"]
        row[f"{b}_sf"] = round(x["sf"], 1)
        if b in UNIT_TYPES:
            row[f"{b}_pieces"] = round(x["pieces"])
        row[f"{b}_labor_base"] = round(x["labor"])
        row[f"{b}_labor_takeoff"] = round(x["labor_takeoff"])
        row[f"{b}_material_base"] = round(x["material"])
        row[f"{b}_mandays"] = round(x["mandays"], 1)
        row[f"{b}_sell"] = round(x["sell"])
        if b in UNIT_TYPES and x["sf"] > 0:
            row[f"{b}_labor_per_sf"] = round(x["labor"] / x["sf"], 2)
            row[f"{b}_labor_takeoff_per_sf"] = round(x["labor_takeoff"] / x["sf"], 2)
            row[f"{b}_material_per_sf"] = round(x["material"] / x["sf"], 2)
            row[f"{b}_sell_per_sf"] = round(x["sell"] / x["sf"], 2)
            row[f"{b}_sf_per_manday"] = round(x["sf"] / x["mandays"], 1) if x["mandays"] else ""
            if x["pieces"]:
                row[f"{b}_labor_per_piece"] = round(x["labor"] / x["pieces"], 2)
                row[f"{b}_pieces_per_manday"] = round(x["pieces"] / x["mandays"], 1) if x["mandays"] else ""
    total = sold + extra["equipment"] + extra["misc"] + extra["cleaning"]
    row["bucketed_total"] = round(total)
    if abs(total - price) > 2:
        problems.append(f"buckets {total:,.0f} vs estimate total {price:,.0f}")
    row["crew_vs_takeoff_labor"] = 1.0
    wages = {k: max(v) for k, v in wage_seen.items()}
    if any(len(v) > 1 for v in wage_seen.values()):
        warn.append("wages differ by building: " + "; ".join(f"{k} {sorted(v)}" for k, v in wage_seen.items() if len(v) > 1))
    row["wages"] = " / ".join(f"{k} {wages[k]:g}" for k in STD_WAGES if k in wages)
    diffs = [wages[k] / STD_WAGES[k] - 1 for k in STD_WAGES if wages.get(k)]
    row["wages_vs_standard"] = round(sum(diffs) / len(diffs), 3) if diffs else ""
    row["reconciles"] = "n" if problems else "y"
    row["reconcile_notes"] = "; ".join(problems)
    row["mandays_total"] = round(sum(B[b]["mandays"] for b in B), 1)
    row["wall_sf_total"] = round(sum(B[b]["sf"] for b in UNIT_TYPES), 1)
    row["sell_per_wall_sf"] = round(price / row["wall_sf_total"], 2) if row["wall_sf_total"] else ""
    row["unmapped_items"] = "; ".join(sorted(unmapped))
    row["notes"] = notes
    row["source_detail"] = "; ".join([f"workbook tabs: {', '.join(tabs)}"] + [f"⚠ {w}" for w in warn])
    row["matches_bid_log"] = ""
    for k, col in (("Bricklayer", "bricklayer"), ("Forman", "foreman"), ("Labor", "laborer"), ("Operator", "operator")):
        row[f"crew_hours_{col}"] = round(hours[k], 1) if hours[k] else ""
    row["crew_hours"] = round(sum(hours.values()), 1)
    row["crew_cost"] = round(extra["crew_cost"])
    row["equipment_cost"] = round(extra["equip_cost"])
    return row


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook"); ap.add_argument("cost_map")
    ap.add_argument("--tabs", nargs="+", required=True); ap.add_argument("--project", required=True)
    ap.add_argument("--metrics-id", default=""); ap.add_argument("--bid-id", default="")
    ap.add_argument("--status", default="", choices=["", "final", "iteration", "draft"])
    ap.add_argument("--notes", default=""); ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    r = compute(a.workbook, a.cost_map, a.tabs, a.project, a.metrics_id, a.bid_id, a.status, a.notes)
    if a.json:
        print(json.dumps(r, indent=1, ensure_ascii=False))
    else:
        w = csv.DictWriter(sys.stdout, fieldnames=list(r)); w.writeheader(); w.writerow(r)
