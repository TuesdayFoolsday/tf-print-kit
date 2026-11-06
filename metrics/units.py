"""Bid Metrics Units: one row per estimate x material x size x grout (System Rules v1.14, B7/B17).

David judges labor per UNIT for CMU and brick and per SF for stone (Nov 2, 2026):
    CMU units = SF x 1.15 · modular brick = SF x 7 · king = SF x 5 · Norman = SF x 4.75 · utility = SF x 3.1
    (the factors on India's MATERIAL FACTORS tab; unlabeled brick counts as modular).
CMU is split by size (4", 6", 8", 10", 12", ...) and grout (full, 16oc, 24oc, 32oc, 40oc, 48oc, mixed, unknown):
grout materially changes labor, so the Bid Check compares like with like.
"""
import re

MATERIALS = ["cmu_gray", "cmu_arch", "brick", "stone", "cast_stone"]
BRICK_FACTORS = {"modular": 7.0, "king": 5.0, "norman": 4.75, "utility": 3.1}
CMU_FACTOR = 1.15
UNITS_COLS = ["unit_key", "metrics_id", "project", "source", "estimate_status", "material", "size", "grout",
              "sf", "units", "unit_factor", "labor_base", "material_base", "mandays", "sell",
              "labor_per_unit", "labor_per_sf", "units_per_manday", "sf_per_manday", "material_per_unit",
              "sell_per_unit", "split_basis", "notes",
              # v2 (System Rules v1.29): measured conditions and the labor split (sum of the four = labor_per_unit)
              "grout_fill_pct", "rebar_lf_per_sf", "labor_lay", "labor_rebar", "labor_grout", "labor_clean",
              # v2b (David, Nov 5 evening: compare on hours, so a wage bump can't skew it): whole-crew hours the row
              # carries (Crew Page hours by labor share) and labor per unit re-priced at the standard wages
              "hours", "hours_per_unit", "labor_std_per_unit"]
V2_UNITS_COLS = UNITS_COLS[-9:]
KINDS = ("lay", "rebar", "grout", "clean")

# Blocks one 3,000-lb SpecMix grout sack (about 1 CY) fills solid, by nominal CMU width (Bid Metrics v2 plan).
BLOCKS_PER_SACK = {"6": 134.0, "8": 101.0, "10": 81.0, "12": 66.0}
# Expected fill for a grout label (cells grouted / all cells, vertical bars only; bond beams and jambs add more).
LABEL_FILL = {"full": 100, "8oc": 100, "16oc": 50, "24oc": 33, "32oc": 25, "40oc": 20, "48oc": 17}
FILL_GAP = 15     # label vs measured fill more than 15 points apart -> takeoff check


def labor_kind(cls, text):
    """Which part of unit labor a line is: grout crew, rebar, cleaning, or laying (everything else)."""
    c, t = (cls or "").upper(), (text or "").lower()
    if "clean" in t:
        return "clean"
    if c == "GRC" or re.search(r"\bgrout", t):
        return "grout"
    if c in ("REB", "RES") or re.search(r"rebar|#\d\b|#\d\s|positioner|epoxy dowel", t):
        return "rebar"
    return "lay"


def capacity(size, units):
    """Sacks needed to fill these blocks solid ('' when the size has no yield on file)."""
    y = BLOCKS_PER_SACK.get(str(size))
    return units / y if y else None


def label_gap(grout, fill_pct):
    """Points between the grout label's expected fill and the measured fill, or None."""
    exp = LABEL_FILL.get(grout)
    try:
        return None if exp is None or fill_pct in ("", None) else float(fill_pct) - exp
    except (TypeError, ValueError):
        return None

_SIZE = re.compile(r"(?<!['\d])(\d{1,2})\s*(?:\"|''|”|in\b|-in\b|inch)", re.I)


def cmu_size(text):
    """First nominal block width in a label ('8" CMU', "12'' CMU", '4"x8"x16 CMU'); '' when absent."""
    for m in _SIZE.finditer(text or ""):
        n = int(m.group(1))
        if 2 <= n <= 16:
            return str(n)
    return ""


def brick_size(text):
    t = (text or "").lower()
    for k in ("king", "norman", "utility"):
        if k in t:
            return k
    return "modular"


def size_of(material, text):
    if material.startswith("cmu"):
        return cmu_size(text) or "?"
    if material == "brick":
        return brick_size(text)
    return ""


def unit_factor(material, size):
    if material.startswith("cmu"):
        return CMU_FACTOR
    if material == "brick":
        return BRICK_FACTORS.get(size, BRICK_FACTORS["modular"])
    return 1.0          # stone, cast stone: per SF


_OC = re.compile(r"(?<!['\d])(\d{2})\s*(?:\"|''|”)?\s*o\.?\s*c", re.I)


def grout_of(text):
    """'full' / '<n>oc' / '' from a takeoff or grout label."""
    t = (text or "").lower()
    if re.search(r"full(y)?\s*-?\s*grout|solid\s*grout|grouted\s*solid|grout\s*solid", t):
        return "full"
    m = _OC.search(t)
    if m and re.search(r"rebar|vert|grout|#\d", t):
        return f"{int(m.group(1))}oc"
    return ""


def parse_grout_spec(spec):
    """'12:full=0.63,16oc=0.37; 8:32oc' -> {'12': [('full', .63), ('16oc', .37)], '8': [('32oc', 1.0)]}"""
    out = {}
    for part in (spec or "").split(";"):
        part = part.strip()
        if not part:
            continue
        size, _, rest = part.partition(":")
        items = []
        for it in rest.split(","):
            g, _, sh = it.strip().partition("=")
            items.append((g.strip().lower(), float(sh) if sh else None))
        known = sum(s for _, s in items if s is not None)
        missing = [i for i, (_, s) in enumerate(items) if s is None]
        for i in missing:
            items[i] = (items[i][0], (1 - known) / len(missing))
        out[size.strip().rstrip('"')] = items
    return out


def finish(rows, metrics_id, project, source, status, labor_factor, mat_factor, grout_spec=None, grout_found=None,
           basis="", notes="", v2=None):
    """rows: {(material, size): dict(sf, labor, material, mandays, sell)} -> list of units-table dicts.
    grout_spec (David) splits a size into grout rows by share; grout_found (from takeoff labels) labels a whole size."""
    spec = parse_grout_spec(grout_spec) if isinstance(grout_spec, str) else (grout_spec or {})
    out = []
    for key, x in sorted(rows.items()):
        if not (x["sf"] or x["labor"] or x["material"]):
            continue
        mat, size = key[0], key[1]
        if len(key) == 3:                      # grout already known per wall (by-wall export)
            parts = [(key[2], 1.0)]
        elif mat.startswith("cmu"):
            parts = spec.get(size) or [((grout_found or {}).get(size) or "unknown", 1.0)]
        else:
            parts = [("n/a", 1.0)]
        for grout, share in parts:
            sf, lab, matl, md = x["sf"] * share, x["labor"] * share, x["material"] * share, x["mandays"] * share
            kinds = {k: x.get(f"labor_{k}") for k in KINDS}
            have_split = all(v is not None for v in kinds.values())
            sell = x.get("sell", lab * labor_factor + matl * mat_factor) * share
            f = unit_factor(mat, size)
            units = sf * f
            key = f"{metrics_id}|{mat}|{size}|{grout}"
            out.append({
                "unit_key": key, "metrics_id": metrics_id, "project": project, "source": source,
                "estimate_status": status, "material": mat, "size": size, "grout": grout,
                "sf": round(sf, 1), "units": round(units), "unit_factor": f, "labor_base": round(lab),
                "material_base": round(matl), "mandays": round(md, 1), "sell": round(sell),
                "labor_per_unit": round(lab / units, 2) if units else "",
                "labor_per_sf": round(lab / sf, 2) if sf else "",
                "units_per_manday": round(units / md, 1) if md else "",
                "sf_per_manday": round(sf / md, 1) if md else "",
                "material_per_unit": round(matl / units, 2) if units else "",
                "sell_per_unit": round(sell / units, 2) if units else "",
                "split_basis": basis + (f"; grout share {share:.0%} (David)" if share < 1 else ""),
                "notes": notes,
                "grout_fill_pct": (round(100 * x["grout_sacks"] / x["grout_cap"]) if x.get("grout_cap") else "")
                                  if mat.startswith("cmu") and "grout_sacks" in x else "",
                "rebar_lf_per_sf": round(x["rebar_lf"] / x["sf"], 2) if "rebar_lf" in x and x["sf"] and mat.startswith("cmu") else "",
                **{f"labor_{k}": (round(kinds[k] * share / units, 2) if mat in ("cmu_gray", "cmu_arch", "brick") else
                                  round(kinds[k] * share / sf, 2)) if have_split and (units if mat in ("cmu_gray", "cmu_arch", "brick") else sf) else ""
                   for k in KINDS}})
            o = out[-1]
            if have_split and o["labor_lay"] != "" and (o["labor_per_unit"] != "" or o["labor_per_sf"] != ""):
                tot = o["labor_per_unit"] if mat in ("cmu_gray", "cmu_arch", "brick") else o["labor_per_sf"]
                o["labor_lay"] = round(tot - o["labor_rebar"] - o["labor_grout"] - o["labor_clean"], 2)  # parts sum exactly
            crew = x["crew_labor"] * share if "crew_labor" in x else None
            o.update(hours_fields(mat, units, sf, lab, v2, crew))
    return out


def hours_fields(mat, units, sf, lab, v2, crew=None):
    """hours, hours_per_unit, labor_std_per_unit for one row. v2 = {hours_per_labor, std_ratio} per $ of CREW labor
    (bid_metrics.v2_factors); blank when the estimate carries no crew hours. crew = the part of the row's labor paid
    by the hour (India: the crew lump; item labor priced per unit carries no hours and isn't re-priced). None = all
    of it (Tradesmen: every labor dollar is Crew Page base wage)."""
    crew = lab if crew is None else crew
    per = units if mat in ("cmu_gray", "cmu_arch", "brick") else sf
    f = {"hours": "", "hours_per_unit": "", "labor_std_per_unit": ""}
    v2 = v2 or {}
    if per and v2.get("hours_per_labor"):
        h = crew * v2["hours_per_labor"]
        f["hours"], f["hours_per_unit"] = round(h, 1), round(h / per, 3)
    if per and v2.get("std_ratio"):
        f["labor_std_per_unit"] = round((lab - crew + crew * v2["std_ratio"]) / per, 2)
    return f


def metric_std(row):
    """The judged number at the standard wages: labor_std_per_unit, else the priced figure (rows written before the
    standard-wage column, almost all priced at the standard wages anyway)."""
    try:
        return float(row.get("labor_std_per_unit"))
    except (TypeError, ValueError):
        return metric(row)


def lay_metric(row, kind="lay"):
    """One part of the judged number (labor_lay / labor_grout / ...) as a float, or None."""
    try:
        return float(row.get(f"labor_{kind}"))
    except (TypeError, ValueError):
        return None


def fill(row):
    try:
        return float(row.get("grout_fill_pct"))
    except (TypeError, ValueError):
        return None


def metric(row):
    """The judged number for a units row: labor per unit (CMU, brick) or per SF (stone, cast stone)."""
    v = row.get("labor_per_unit") if row.get("material") in ("cmu_gray", "cmu_arch", "brick") else row.get("labor_per_sf")
    try:
        return float(v)
    except (TypeError, ValueError):
        return None
