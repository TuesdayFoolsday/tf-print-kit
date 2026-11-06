"""Read a bid's Wall Conditions file (System Rules v1.29, B7) into the Bid Metrics conditions fields.

The file is `07 Estimate/<Project> - Estimate - <date> - Wall Conditions.md`, written in chat from David's notes:
grout (and which drawing governs), vertical bars, corners, jambs, bond beams, wall heights, wall types, access,
wage bumps, anything unusual. It is free text; this reader picks up what it can and leaves the rest blank.
An optional `## Fields` section with `key: value` lines wins over anything guessed from the text:

    ## Fields
    - max_wall_height_ft: 14
    - openings: 6            (count; becomes openings_per_100sf with the wall SF)
    - chase_walls: y
    - access: tight          (normal / tight / difficult)
    - wages_bumped: n
    - grout: full            (or 16oc, 32oc ...; used as the CMU grout label when no --grout is given)
    - secure_site: n         (badging / escorts)
    - travel: n
    - prevailing_wage: y
    - difficulty: 3          (1-5)
The /job skill writes this section from what David says (he just talks; Claude asks a follow-up only when a missing
answer would change the labor comparison). A heading "## Answers" is read the same way. Nothing absent is assumed:
a file that never mentions chase walls leaves chase_walls blank.

    python3 metrics/conditions.py "<Wall Conditions.md>"      # prints what it read
"""
import json, re, sys, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import units as U

ACCESS = ("normal", "tight", "difficult")


def read(path):
    text = open(path, encoding="utf-8-sig").read()
    low = text.lower()
    out = {"max_wall_height_ft": None, "openings": None, "chase_walls": None, "access": None,
           "wages_bumped": None, "grout": None, "secure_site": None, "travel": None, "prevailing_wage": None,
           "difficulty": None}
    # wall heights: any number of feet on a line that talks about height
    hs = []
    for line in low.splitlines():
        if re.search(r"height|\btall\b", line):
            hs += [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s*(?:'|ft\b|feet|foot)", line)]
    out["max_wall_height_ft"] = max(hs) if hs else None
    m = re.search(r"(\d+)\s+openings|openings?\s*[:=]\s*(\d+)", low)
    if m:
        out["openings"] = int(m.group(1) or m.group(2))
    if re.search(r"\bchase", low):
        out["chase_walls"] = "n" if re.search(r"\bno\s+chase", low) else "y"
    m = re.search(r"access\s*[:=-]\s*(\w+)", low)
    if m and m.group(1) in ACCESS:
        out["access"] = m.group(1)
    elif re.search(r"(tight|restricted|limited)\s+access|access\s+(is\s+)?(tight|restricted|limited)", low):
        out["access"] = "tight"
    elif re.search(r"difficult\s+access|access\s+(is\s+)?difficult", low):
        out["access"] = "difficult"
    if re.search(r"wages?\s+(bump|bumped|raised|up)|bump(ed)?\s+(the\s+)?wages?|wage bump", low):
        out["wages_bumped"] = "y"
    elif re.search(r"wages?\s+(are\s+)?standard", low):
        out["wages_bumped"] = "n"
    m = re.search(r"^\s*-?\s*\**grout\**\s*:\s*(.+)$", low, re.M)
    if m:
        v = m.group(1)
        out["grout"] = "full" if re.search(r"\bfull\b|solid", v) else (U.grout_of(v) or None)
    # ## Fields wins
    fm = re.search(r"^##\s*(?:fields|answers)\s*$(.*?)(?=^##\s|\Z)", text, re.M | re.I | re.S)
    if fm:
        for k, v in re.findall(r"^\s*-?\s*([a-z_]+)\s*:\s*(.+?)\s*$", fm.group(1), re.M):
            k, v = k.lower(), v.split("(")[0].strip().lower()
            try:
                if k == "openings":
                    out["openings"] = int(float(v))
                elif k == "max_wall_height_ft":
                    out[k] = float(v)
                elif k == "difficulty":
                    out[k] = str(int(float(v))) if 1 <= float(v) <= 5 else None
                elif k in ("chase_walls", "secure_site", "travel", "prevailing_wage", "wages_bumped"):
                    out[k] = {"y": "y", "yes": "y", "n": "n", "no": "n"}.get(v)
                elif k == "access":
                    out[k] = v if v in ACCESS else None
                elif k in out:
                    out[k] = v
            except ValueError:
                pass                      # a value that doesn't fit its field stays as read from the text (or blank)
    return out


def fields(cond, wall_sf, link=""):
    """Conditions -> Bid Metrics columns (None = leave the export's own value)."""
    return {"max_wall_height_ft": cond.get("max_wall_height_ft"),
            "openings_per_100sf": round(100 * cond["openings"] / wall_sf, 2) if cond.get("openings") and wall_sf else None,
            "chase_walls": cond.get("chase_walls"), "access": cond.get("access"),
            "wages_bumped": cond.get("wages_bumped"), "conditions_file": link or None,
            "secure_site": cond.get("secure_site"), "travel": cond.get("travel"),
            "prevailing_wage": cond.get("prevailing_wage"), "difficulty": cond.get("difficulty")}


if __name__ == "__main__":
    print(json.dumps(read(sys.argv[1]), indent=1))
