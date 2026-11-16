"""Open jobs: cost to date vs estimate, % complete and the projected over/under, one row per open job.

    python3 jobcost/open_jobs.py --actuals "Open Job Actuals.csv" --jobs "Jobs.csv" --bid-metrics "Bid Metrics 2026.csv" \
        [--billing "JC Cost and Revenue.xls|.csv"] [--contracts "JC Contracts.xlsx"] \
        [--laid "Job Progress - Jobs.csv"] [--pm "PM Percent Complete.csv"] [--pct-source billed|pm|laid] \
        --out "Open Jobs.csv" [--summary open_jobs.json]

Inputs
* actuals — jobcost/jc_detail_multi.py output (weekly Vista JC Detail for every open job, System Rules B14 buckets).
* jobs — the Operations Jobs list (VistaJobNo, Title, Status, EstimateID). Open = Status not "Closed out".
  EVERY open job gets a row, with or without costs or an estimate.
* bid-metrics — merged Bid Metrics (pending files folded in first, B16).
* Estimate match, in this order, never by a guessed name:
    1. the Jobs row's EstimateID = a Bid Metrics metrics_id;
    2. exactly one Bid Metrics row whose vista_job = the Vista number (B18 bid-to-job link);
  (Job Cost History is keyed by the Tradesmen job number, not the Vista number, so it is not a match source.)
  Anything else -> flag "no estimate entered" (or "two estimates claim this job").
* % complete (pct_used), the switch is --pct-source (settings: courier/dashboard_config.json "pct_source"):
    billed (default, option B; David 2026-11-16: every job) — Vista billed to date (JC Cost and Revenue, Billed
                               Amount) / Vista current contract (JC Contracts - Detail). Nothing from the Jobs list.
    pm (option A)              — the PM's weekly % only; billed % shown for reference.
    laid                       — foreman daily reports: laid / estimated quantities (Job Progress - Jobs.csv).
  A contract of 0 or missing in Vista -> no billed %, no projection, noted on the row.
  Laid % is always shown in its own column (masonry jobs).
* Cross-check (David 2026-11-16: JC Cost and Revenue replaces JC Cost By Job, Phase and Cost Type): each job's JC Detail
  total must equal its JC Cost and Revenue Total Cost within $1; otherwise the row is flagged and not compared.
* Projection = cost to date / pct_used, set against the estimate (labor man-days, base-wage labor $, material with tax,
  total). Not projected below 5% complete. "over" = projected (or already actual) more than THRESHOLD over the estimate
  once the job is at least MIN_PCT complete.

Estimate side = job_cost.py (B14): man-days = Crew Page hours / 8; labor = base wage by unit type + insulation, scaffold,
other items; material = material base x (1 + tax rate); total = burdened crew labor (crew_cost) + material with tax +
equipment_cost (before OH&P). Vista's total includes everything booked (trucking, subs, other).
"""
import argparse, csv, json, os, re, sys
from collections import defaultdict

UNIT_TYPES = ["cmu_gray", "cmu_arch", "brick", "stone", "cast_stone"]
CLOSED = "Closed out"
PCT_SOURCES = ("billed", "pm", "laid")
THRESHOLD, MIN_PCT, MIN_PROJECT = 0.10, 0.15, 0.05
COLS = ["job", "project", "status", "estimate_id", "match", "through_date", "recon_ok",
        "pct_used", "pct_basis", "pct_billed", "pct_pm", "pct_laid", "contract", "billed_to_date",
        "est_mandays", "act_mandays", "est_labor_base", "act_labor_base", "est_material", "act_material",
        "est_cost_total", "act_total", "act_burden", "act_equipment", "act_sub", "act_other",
        "proj_mandays", "mandays_over_pct", "proj_labor_base", "labor_over_pct", "proj_material", "material_over_pct",
        "proj_total", "total_over_pct", "flag", "notes"]
COMPARE = [("mandays", "est_mandays", "act_mandays"), ("labor", "est_labor_base", "act_labor_base"),
           ("material", "est_material", "act_material"), ("total", "est_cost_total", "act_total")]


def num(v):
    try:
        s = str(v).replace(",", "").replace("$", "").replace("%", "").strip()
        return float(s) if s not in ("", "None") else None
    except ValueError:
        return None


def read_csv(path):
    if not path or not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def read_billing(path):
    """Vista 'JC Cost and Revenue' (.xls as exported: Contract, Description, L, M, S, E, R, O, P, I, ALL OTHER CT,
    Total Cost, Billed Amount) or its CSV twin (vista_no, ..., total_cost, billed) -> {job: (total_cost, billed)}."""
    out = {}
    if not path or not os.path.exists(path):
        return out
    if path.lower().endswith(".csv"):
        for r in read_csv(path):
            n = (r.get("vista_no") or r.get("job") or "").strip()
            if n:
                out[n] = (num(r.get("total_cost")), num(r.get("billed")))
        return out
    import xlrd
    s = xlrd.open_workbook(path).sheet_by_index(0)
    head = [str(v).replace("\n", " ").strip().lower() for v in s.row_values(1)]
    try:
        ci, ti, bi = head.index("contract"), head.index("total cost"), head.index("billed amount")
    except ValueError:
        raise SystemExit(f"✗ {path}: not the JC Cost and Revenue layout (header row 2: {head})")
    for r in range(2, s.nrows):
        row = s.row_values(r)
        m = re.match(r"^\s*(\d{4})-", str(row[ci]))
        if m:
            out[m.group(1)] = (num(row[ti]), num(row[bi]))
    return out


def read_contracts(path):
    """Vista 'JC Contracts - Detail' (.xlsx/.xls) -> {job: current contract} for amounts above 0."""
    out = {}
    if not path or not os.path.exists(path):
        return out
    with open(path, "rb") as f:
        is_xlsx = f.read(2) == b"PK"        # by content: the courier saves either kind under one local name
    if is_xlsx:
        import openpyxl
        rows = openpyxl.load_workbook(path, read_only=True).worksheets[0].iter_rows(values_only=True)
    else:
        import xlrd
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        rows = (sh.row_values(r) for r in range(sh.nrows))
    cur = None
    for row in rows:
        v0 = row[0] if row else None
        m = re.match(r"^\s*(\d{4})-\s", v0) if isinstance(v0, str) else None
        if m:
            cur = m.group(1)
            continue
        cells = [v for v in row if v not in (None, "")]
        if cur and "Original:" in cells and "Current:" in cells:
            c = num(cells[cells.index("Current:") + 1]) if cells.index("Current:") + 1 < len(cells) else None
            if c and c > 0:
                out[cur] = c
    return out


def estimate_figures(e):
    tax = 1 + (num(e.get("tax_rate")) or 0)
    hrs = num(e.get("crew_hours"))
    lab = sum(num(e.get(f"{u}_labor_base")) or 0 for u in UNIT_TYPES + ["insulation", "scaffold", "other_items"])
    mat = tax * sum(num(e.get(f"{u}_material_base")) or 0 for u in UNIT_TYPES + ["insulation", "scaffold", "other_items"])
    crew, eq = num(e.get("crew_cost")), num(e.get("equipment_cost"))
    return {"est_mandays": round(hrs / 8, 2) if hrs else None, "est_labor_base": round(lab, 2) if lab else None,
            "est_material": round(mat, 2) if mat else None,
            "est_cost_total": round(crew + mat + (eq or 0), 2) if crew else None}


def match_estimate(job, row, metrics, by_vista):
    mid = (row.get("EstimateID") or "").strip()
    if mid:
        return (metrics.get(mid), "Jobs row EstimateID", "" if mid in metrics else f"EstimateID {mid} not in Bid Metrics")
    hits = by_vista.get(job, [])
    if len(hits) == 1:
        return hits[0], "Bid Metrics vista_job", ""
    if len(hits) > 1:
        finals = [h for h in hits if h.get("estimate_status") == "final"]
        if len(finals) == 1:
            return finals[0], "Bid Metrics vista_job (final row)", ""
        return None, "", "two estimates claim this job (" + ", ".join(h["metrics_id"] for h in hits) + ")"
    return None, "", ""


def build(a):
    acts = {r["job"]: r for r in read_csv(a.actuals)}
    jobs = [r for r in read_csv(a.jobs) if (r.get("VistaJobNo") or "").strip() and r.get("Status") != CLOSED]
    bm = read_csv(a.bid_metrics)
    metrics = {r.get("metrics_id", ""): r for r in bm if r.get("metrics_id")}
    by_vista = defaultdict(list)
    for r in bm:
        if (r.get("vista_job") or "").strip():
            by_vista[r["vista_job"].strip()].append(r)
    billing = read_billing(a.billing)
    contracts = read_contracts(a.contracts)
    laid = {r["job"].strip(): num(r.get("job_pct_complete")) for r in read_csv(a.laid) if r.get("job")}
    pm = {}
    for r in read_csv(a.pm):
        p = num(r.get("pct_complete"))
        if r.get("job") and p is not None:
            pm[r["job"].strip()] = p / 100 if p > 1 else p
    out, problems, over_lines, over_jobs = [], [], [], []
    open_numbers = set()
    for j in sorted(jobs, key=lambda r: r["VistaJobNo"].strip()):
        n = j["VistaJobNo"].strip()
        open_numbers.add(n)
        act = acts.get(n, {})
        e, how, why = match_estimate(n, j, metrics, by_vista)
        est = estimate_figures(e) if e else {}
        flags, notes = [], []
        if not e:
            flags.append(why or "no estimate entered")
        elif est.get("est_mandays") is None:
            notes.append("estimate has no Crew Page hours — re-run its Bid Check (man-days not compared)")
        if not act:
            notes.append("no Vista cost in this export")
        elif act.get("recon_ok") == "n":
            flags.append("Vista subtotals don't match — not compared")
        if act.get("notes"):
            notes.append(act["notes"])
        contract = contracts.get(n)
        cr_total, billed = billing.get(n) or (None, None)
        pct_billed = round(billed / contract, 4) if billed is not None and contract else None
        if not contract:
            notes.append("no contract amount in Vista (JC Contracts) — billed % not computed")
        elif billed is None:
            notes.append("not in JC Cost and Revenue — billed % not computed")
        cross_ok = True
        if act and billing and cr_total is None:
            notes.append("not in JC Cost and Revenue — cost cross-check not done")
        elif act and cr_total is not None and abs((num(act.get("act_total")) or 0) - cr_total) > 1:
            cross_ok = False
            flags.append(f"Vista reports disagree: JC Detail ${num(act.get('act_total')) or 0:,.2f} vs JC Cost and "
                         f"Revenue ${cr_total:,.2f} (different dates?) — not compared")
        pct_pm, pct_laid = pm.get(n), (laid.get(n) / 100 if laid.get(n) is not None else None)
        if a.pct_source == "pm":
            pct, basis = pct_pm, "PM %"
        elif a.pct_source == "laid":
            pct, basis = pct_laid, "laid %"
        else:
            pct, basis = pct_billed, "billed / contract"
        if pct is None:
            basis = ""
        row = {"job": n, "project": j.get("Title", ""), "status": j.get("Status", ""),
               "estimate_id": (e or {}).get("metrics_id", ""), "match": how,
               "through_date": act.get("through_date", ""), "recon_ok": act.get("recon_ok", ""),
               "pct_used": pct, "pct_basis": basis, "pct_billed": pct_billed, "pct_pm": pct_pm, "pct_laid": pct_laid,
               "contract": contract, "billed_to_date": billed, **est,
               **{k: num(act.get(k)) for k in ("act_mandays", "act_labor_base", "act_material", "act_total", "act_burden",
                                                "act_equipment", "act_sub", "act_other")}}
        no_hours = bool(act) and not num(act.get("labor_hours"))
        if no_hours and ((num(act.get("act_sub")) or 0) > 0 or "SL" in act.get("notes", "")):
            notes.append("no FB labor hours (labor subcontracted?) — man-days and labor $ not compared")
        worst = []
        comparable = e and act and act.get("recon_ok") == "y" and cross_ok
        for key, ek, ak in COMPARE:
            ev, av = row.get(ek), row.get(ak)
            if no_hours and key in ("mandays", "labor"):
                ev = None
            if not comparable or ev in (None, 0) or av is None:
                continue
            proj = av / pct if pct and pct >= MIN_PROJECT else None
            row[f"proj_{'labor_base' if key == 'labor' else key}"] = round(proj, 2) if proj is not None else None
            over = (proj - ev) / ev if proj is not None else None
            row[f"{key}_over_pct"] = round(over, 4) if over is not None else None
            already = (av - ev) / ev
            if already > THRESHOLD:
                worst.append((already, f"{key} already {already:+.0%} over the estimate"))
            elif over is not None and pct >= MIN_PCT and over > THRESHOLD:
                worst.append((over, f"{key} projected {over:+.0%}"))
        if worst:
            worst.sort(reverse=True)
            flags.insert(0, "over")
            title = (j.get("Title") or "").strip()
            over_lines.append(f"{n} {title}: " + "; ".join(w[1] for w in worst)
                              + (f" ({pct:.0%} complete by {basis})" if pct is not None else ""))
            over_jobs.append({"job": n, "project": title,
                              "parts": [{"what": w[1].split()[0], "pct": w[0],
                                         "kind": "already" if "already" in w[1] else "projected"} for w in worst],
                              "pct_complete": pct, "pct_basis": basis})
        row["flag"] = "; ".join(flags)
        row["notes"] = "; ".join(notes)
        out.append(row)
    extra = sorted(set(acts) - open_numbers)
    if extra:
        problems.append("Vista jobs with cost but not open on the Jobs list (left out): " + ", ".join(extra))
    return out, problems, over_lines, over_jobs


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--actuals", required=True); ap.add_argument("--jobs", required=True)
    ap.add_argument("--bid-metrics", required=True)
    ap.add_argument("--billing", default=""); ap.add_argument("--contracts", default="")
    ap.add_argument("--laid", default=""); ap.add_argument("--pm", default="")
    ap.add_argument("--pct-source", default="billed", choices=PCT_SOURCES)
    ap.add_argument("--out", required=True); ap.add_argument("--summary", default="")
    a = ap.parse_args(argv)
    rows, problems, over, over_jobs = build(a)
    with open(a.out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLS, lineterminator="\n", extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if r.get(k) is None else r.get(k)) for k in COLS})
    no_est = [r["job"] for r in rows if "no estimate" in r["flag"] or "claim" in r["flag"]]
    summary = {"open_jobs": len(rows), "with_estimate": len(rows) - len(no_est), "no_estimate": no_est,
               "with_pct": sum(1 for r in rows if r["pct_used"] is not None), "pct_source": a.pct_source,
               "over": over, "over_jobs": over_jobs, "problems": problems}
    if a.summary:
        json.dump(summary, open(a.summary, "w"), indent=1)
    print(f"Open Jobs: {len(rows)} open jobs · {summary['with_estimate']} matched to an estimate · "
          f"{len(no_est)} flagged no estimate · {summary['with_pct']} with a % complete ({a.pct_source}) · {len(over)} over")
    for o in over:
        print("↳ over: " + o)
    for p in problems:
        print("⚠ " + p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
