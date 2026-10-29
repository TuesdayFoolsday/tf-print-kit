"""Job Cost History: closeout actuals vs the estimate, one row per job (plus one per unit type when it can be split).

    python3 jobcost/job_cost.py add <actuals.csv> <bid_metrics.csv> <history.csv> [--splits splits.csv]
    python3 jobcost/job_cost.py template <dir>        # writes the two blank intake files with one example row each

actuals.csv — one row per job x cost line (any accounting export mapped to these columns):
    job_number      accounting job number (e.g. 1807-06)
    estimate        the project name exactly as in Bid Metrics (links the job to its estimate)
    through_date    YYYY-MM-DD, date the actuals run through
    final           y = job closed out; n = in progress (kept, marked partial, never used for calibration)
    cost_type       labor | material | equipment | sub | other
    labor_class     bricklayer | foreman | laborer | operator   (labor rows only)
    hours           labor hours (labor rows)
    amount          dollars as booked
    burdened        y if the labor amount includes burden/taxes/fringe, n if base wage only, blank = unknown
splits.csv (optional) — when a job has more than one unit type and the timeline gives a clean break:
    job_number, unit_type (cmu_gray|cmu_arch|brick|stone|cast_stone), labor_hours_share (0-1)

Rules:
* The judged comparison is labor: actual base-wage labor $/SF and man-days vs the estimate's Crew Page labor.
  Labor dollars are compared only when the export says base wage (burdened = n). Burdened or unknown → hours only,
  and the row says so. Base wage is never backed out of a burdened number by a guess.
* Per-unit-type actuals exist only when the estimate has one unit type (then it gets 100%) or a split row is given.
  Everything else stays job-level.
* Re-adding a job replaces its rows. Rows from non-final actuals are marked partial.
"""
import argparse, csv, os, sys
from collections import defaultdict

UNIT_TYPES = ["cmu_gray", "cmu_arch", "brick", "stone", "cast_stone"]
CLASSES = {"bricklayer", "foreman", "laborer", "operator"}
COST_TYPES = {"labor", "material", "equipment", "sub", "other"}
ACOLS = ["job_number", "estimate", "through_date", "final", "cost_type", "labor_class", "hours", "amount", "burdened"]
SCOLS = ["job_number", "unit_type", "labor_hours_share"]
HCOLS = ["job_number", "estimate", "scope", "through_date", "status", "est_sf", "est_mandays", "act_mandays",
         "mandays_var_pct", "est_labor_base", "act_labor_base", "labor_var_pct", "est_labor_per_sf", "act_labor_per_sf",
         "est_sf_per_manday", "act_sf_per_manday", "est_material", "act_material", "material_var_pct",
         "act_equipment", "act_sub", "act_other", "labor_basis", "hours_by_class", "notes"]


def num(v):
    try:
        return float(v) if str(v).strip() != "" else None
    except ValueError:
        return None


def read(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def pct(a, b):
    return f"{(a - b) / b:.4f}" if a is not None and b else ""


def r2(x):
    return "" if x is None else f"{x:.2f}"


def cmd_template(a):
    os.makedirs(a.dir, exist_ok=True)
    with open(os.path.join(a.dir, "Vista Actuals - template.csv"), "w", newline="") as f:
        w = csv.writer(f, lineterminator="\n"); w.writerow(ACOLS)
        w.writerow(["1807-06", "Rowlett Creek", "2026-12-15", "y", "labor", "bricklayer", "1240", "37200", "n"])
    with open(os.path.join(a.dir, "Job Splits - template.csv"), "w", newline="") as f:
        w = csv.writer(f, lineterminator="\n"); w.writerow(SCOLS)
        w.writerow(["1807-06", "cmu_gray", "0.6"])
    print(f"templates written to {a.dir}")


def cmd_add(a):
    acts = read(a.actuals)
    est = {r["project"]: r for r in read(a.bid_metrics)}
    splits = defaultdict(dict)
    if a.splits:
        for s in read(a.splits):
            splits[s["job_number"]][s["unit_type"]] = num(s["labor_hours_share"]) or 0
    problems, jobs = [], defaultdict(list)
    for i, r in enumerate(acts, start=2):
        ct, lc = r["cost_type"].strip().lower(), r.get("labor_class", "").strip().lower()
        if ct not in COST_TYPES:
            problems.append(f"line {i}: cost_type '{r['cost_type']}'")
        elif ct == "labor" and lc not in CLASSES:
            problems.append(f"line {i}: labor_class '{r.get('labor_class')}'")
        elif num(r["amount"]) is None:
            problems.append(f"line {i}: amount '{r['amount']}'")
        else:
            jobs[r["job_number"].strip()].append(r)
    out = []
    for job, rs in jobs.items():
        names = {r["estimate"].strip() for r in rs}
        e = est.get(next(iter(names))) if len(names) == 1 else None
        if not e:
            problems.append(f"job {job}: estimate {sorted(names)} not found in Bid Metrics (or more than one name)")
            continue
        final = all(r["final"].strip().lower() == "y" for r in rs)
        thru = max(r["through_date"] for r in rs)
        lab = [r for r in rs if r["cost_type"].strip().lower() == "labor"]
        hours = sum(num(r["hours"]) or 0 for r in lab)
        by_cls = defaultdict(float)
        for r in lab:
            by_cls[r["labor_class"].strip().lower()] += num(r["hours"]) or 0
        burd = {r.get("burdened", "").strip().lower() for r in lab}
        base_ok = burd == {"n"}
        act_lab = sum(num(r["amount"]) for r in lab) if base_ok else None
        tot = lambda t: sum(num(r["amount"]) for r in rs if r["cost_type"].strip().lower() == t)
        act_md = hours / 8 if hours else None
        basis = "base wage" if base_ok else ("burdened — hours only" if "y" in burd else "unknown — hours only")
        cls_txt = " / ".join(f"{k} {by_cls[k]:,.0f}h" for k in ("bricklayer", "foreman", "laborer", "operator") if by_cls[k])
        status = "final" if final else "partial"
        ut_present = [u for u in UNIT_TYPES if num(e.get(f"{u}_sf"))]
        est_lab = sum(num(e.get(f"{u}_labor_base")) or 0 for u in UNIT_TYPES) + sum(
            num(e.get(f"{x}_labor_base")) or 0 for x in ("insulation", "scaffold", "other_items"))
        est_mat = sum(num(e.get(f"{u}_material_base")) or 0 for u in UNIT_TYPES + ["insulation", "scaffold", "other_items"])
        est_md, est_sf = num(e.get("mandays_total")), num(e.get("wall_sf_total"))
        row = dict(job_number=job, estimate=e["project"], scope="job", through_date=thru, status=status,
                   est_sf=r2(est_sf), est_mandays=r2(est_md), act_mandays=r2(act_md), mandays_var_pct=pct(act_md, est_md),
                   est_labor_base=r2(est_lab), act_labor_base=r2(act_lab), labor_var_pct=pct(act_lab, est_lab),
                   est_labor_per_sf=r2(est_lab / est_sf) if est_sf else "",
                   act_labor_per_sf=r2(act_lab / est_sf) if act_lab is not None and est_sf else "",
                   est_sf_per_manday=r2(est_sf / est_md) if est_sf and est_md else "",
                   act_sf_per_manday=r2(est_sf / act_md) if est_sf and act_md else "",
                   est_material=r2(est_mat), act_material=r2(tot("material")), material_var_pct=pct(tot("material"), est_mat),
                   act_equipment=r2(tot("equipment")), act_sub=r2(tot("sub")), act_other=r2(tot("other")),
                   labor_basis=basis, hours_by_class=cls_txt, notes=e.get("notes", ""))
        out.append(row)
        # unit-type rows: single unit type → 100%; else only from a split
        shares = splits.get(job) or ({ut_present[0]: 1.0} if len(ut_present) == 1 else {})
        if shares and abs(sum(shares.values()) - 1) > 0.01:
            problems.append(f"job {job}: split shares sum to {sum(shares.values()):.2f}, not 1 — unit rows skipped")
            shares = {}
        if len(ut_present) > 1 and not shares:
            row["notes"] = (row["notes"] + "; " if row["notes"] else "") + "multi-unit job: job-level only (add a split to calibrate unit types)"
        for u, sh in shares.items():
            sf, eu, emd = num(e.get(f"{u}_sf")), num(e.get(f"{u}_labor_base")), num(e.get(f"{u}_mandays"))
            amd = act_md * sh if act_md else None
            al = act_lab * sh if act_lab is not None else None
            out.append(dict(job_number=job, estimate=e["project"], scope=u, through_date=thru, status=status,
                            est_sf=r2(sf), est_mandays=r2(emd), act_mandays=r2(amd), mandays_var_pct=pct(amd, emd),
                            est_labor_base=r2(eu), act_labor_base=r2(al), labor_var_pct=pct(al, eu),
                            est_labor_per_sf=e.get(f"{u}_labor_per_sf", ""), act_labor_per_sf=r2(al / sf) if al is not None and sf else "",
                            est_sf_per_manday=e.get(f"{u}_sf_per_manday", ""), act_sf_per_manday=r2(sf / amd) if sf and amd else "",
                            est_material=r2(num(e.get(f"{u}_material_base"))), act_material="", material_var_pct="",
                            act_equipment="", act_sub="", act_other="", labor_basis=basis,
                            hours_by_class=f"{sh:.0%} of job hours" + (" (split)" if job in splits else " (single unit type)"),
                            notes=""))
    for p in problems:
        print("⚠ " + p)
    hist = read(a.history) if os.path.exists(a.history) else []
    done = {r["job_number"] for r in out}
    hist = [h for h in hist if h["job_number"] not in done] + out
    hist.sort(key=lambda r: (r["through_date"], r["job_number"], r["scope"] != "job", r["scope"]))
    with open(a.history, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HCOLS, lineterminator="\n"); w.writeheader(); w.writerows(hist)
    for r in out:
        if r["scope"] == "job":
            print(f"{r['job_number']} {r['estimate']} ({r['status']}, {r['labor_basis']}): man-days {r['act_mandays']} vs est "
                  f"{r['est_mandays']} ({float(r['mandays_var_pct'] or 0):+.0%})"
                  + (f"; labor ${float(r['act_labor_base']):,.0f} vs ${float(r['est_labor_base']):,.0f} ({float(r['labor_var_pct']):+.0%})"
                     if r["act_labor_base"] else ""))
    print(f"Job Cost History: {len(done)} job(s) written, {len(hist)} rows total -> {a.history}")
    return 1 if problems else 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("add"); p.add_argument("actuals"); p.add_argument("bid_metrics"); p.add_argument("history")
    p.add_argument("--splits"); p.set_defaults(f=cmd_add)
    p = sp.add_parser("template"); p.add_argument("dir"); p.set_defaults(f=cmd_template)
    a = ap.parse_args()
    sys.exit(a.f(a) or 0)


if __name__ == "__main__":
    main()
