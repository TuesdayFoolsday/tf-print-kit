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
              "sell_per_unit", "split_basis", "notes"]

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
           basis="", notes=""):
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
                "notes": notes})
    return out


def metric(row):
    """The judged number for a units row: labor per unit (CMU, brick) or per SF (stone, cast stone)."""
    v = row.get("labor_per_unit") if row.get("material") in ("cmu_gray", "cmu_arch", "brick") else row.get("labor_per_sf")
    try:
        return float(v)
    except (TypeError, ValueError):
        return None
