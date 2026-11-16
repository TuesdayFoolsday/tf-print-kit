"""Vista JC Detail for MANY jobs (.xls, JCDetail.rpt saved as Excel) -> one row of actuals per job, B14 mapping.

    python3 jobcost/jc_detail_multi.py "<JC Detail YYYY-MM-DD.xls>" "<out Open Job Actuals.csv>" [--through YYYY-MM-DD]
        [--jobs 2502,2622] [--actuals-dir <dir>]

The weekly export David drops in SharePoint `_Drop Box` (B2 files it to Leadership › Job Cost) covers every open job.
This reader follows the job / phase / cost-type headers on every sheet (Excel splits the report at 65k rows) with the
same layout rules as jobcost/jc_detail_ledger.py, and checks every "Total for Cost Type / For Phase / For Job" line
against the rows read under it. A job whose totals don't match is marked recon_ok = n and must not be compared.

Buckets = System Rules B14 (David 2026-11-01), exactly as jobcost/vista_jc_detail.py applies them:
* cost type L with hours > 0 = base-wage labor (burdened = n); L lines without hours (Payroll Tax, Workmen's Comp,
  401K, Liability/Health Insurance, SUTA, FUTA, ...) = burden, left out of labor and reported as a %.
* phase 01650 Trucking, any L line = other (shown, not in the masonry comparison).
* M and I = material (Vista AP includes tax); E and R = equipment; O = other; S = sub; anything else
  (SL, INV, ...) = other, listed in the job's notes so nothing is hidden.
* Man-days = WHOLE-CREW hours / 8 (all L hours outside 01650), compared with the estimate's Crew Page hours / 8.
* Classes (shown, not used in the man-day total): craft 1/... = bricklayer, except a name on the foreman roster
  = foreman; craft 2/... = laborer, except anyone averaging $25-28/hr on the job = operator.

--actuals-dir also writes, per job, the job_cost.py actuals.csv (same columns and rows vista_jc_detail.py writes),
with final = n (open jobs are never final) and estimate left blank for the caller to fill from the match.
"""
import argparse, csv, os, re, sys
from collections import defaultdict

JOB = re.compile(r"^\s*(\d{4})-\s{2}(.*?)\s*(- Continued)?\s*$")
PH = re.compile(r"^\s*(\d{5})-\s+-\s+(.*?)\s{2,}(\d+)\s+([A-Z]+)\s*$")
TOT = re.compile(r"^Total (for Cost Type|For Phase|For Job):\s*(.*)$", re.I)
PRD = re.compile(r"^(?:(\d+)/(.+?)\s+)?([\d.]+)\s+/\s+(\d+)\s+/\s+(.*?)\s*(\d{1,3})\s+(\S.*)$")
DATED = re.compile(r"(\d{4}-\d{2}-\d{2})")

# Foreman roster (jc_detail_summary.py, David Nov 4 2026). Edit there and here together.
FOREMEN = {"Leon, Abel", "Borja II, Frank R.", "Barragan, Ignacio", "Angulo-Nagy, Wilman A.", "Banuelos, Oscar",
           "Rico-Robles Jr, Ruben", "Crissinger, Matthew E.", "Barker, Donald P", "Frilot, Antoine R", "Flores, Abel",
           "Porras, Tommy", "Guizar, Juan J", "Valencia, Ignacio", "Valladares, Jose", "Mills, Jeffrey", "Perez, Jose"}
BUCKET = {"M": "material", "I": "material", "E": "equipment", "R": "equipment", "O": "other", "S": "sub"}
COLS = ["job", "job_name", "through_date", "recon_ok", "rows", "first_month", "last_month", "last_actual_date",
        "labor_hours", "act_mandays", "hours_bricklayer", "hours_foreman", "hours_laborer", "hours_operator",
        "act_labor_base", "act_burden", "burden_pct", "act_trucking", "act_material", "act_equipment", "act_sub",
        "act_other", "act_total", "notes"]
ACOLS = ["job_number", "estimate", "through_date", "final", "cost_type", "labor_class", "hours", "amount", "burdened",
         "labor_sub", "tax_exempt"]


def _f(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def read(path):
    """-> {rows, recon, failures, jobs (number -> name, header order), problems}. Pure layout, no mapping."""
    import xlrd
    wb = xlrd.open_workbook(path, on_demand=True)
    day = lambda v: xlrd.xldate_as_datetime(v, wb.datemode).date().isoformat() if isinstance(v, (int, float)) and v > 1000 else ""
    rows, recon, problems, jobs = [], [], [], {}
    job = name = ph = None
    s_ct, s_ph, s_job = [0.0, 0.0], [0.0, 0.0], [0.0, 0.0]
    for sh in wb.sheets():
        for r in range(sh.nrows):
            row = sh.row(r)
            c0 = row[0].value
            c0s = str(c0).strip()
            if isinstance(c0, str) and c0s:
                m = JOB.match(c0)
                if m and not m.group(3):
                    job, name = m.group(1), m.group(2).strip()
                    jobs.setdefault(job, name)
                    s_job = [0.0, 0.0]
                    continue
                if m:
                    continue
                m = PH.match(c0)
                if m:
                    ph = (m.group(1), m.group(2).strip(), m.group(3), m.group(4))
                    s_ct = [0.0, 0.0]
                    continue
                if re.match(r"^\s*\d{5}-", c0):
                    problems.append(f"unparsed phase line {sh.name}!{r + 1}: {c0s}")
                continue
            t15 = str(row[15].value).strip() if len(row) > 15 else ""
            m = TOT.match(t15)
            if m:
                kind = m.group(1).lower()
                hrs, cost = _f(row[33].value) if len(row) > 33 else 0.0, _f(row[36].value) if len(row) > 36 else 0.0
                if kind.startswith("for cost"):
                    recon.append(("cost type", job, ph[0] if ph else "", s_ct[0], hrs, s_ct[1], cost))
                    s_ph = [s_ph[0] + s_ct[0], s_ph[1] + s_ct[1]]
                    s_ct = [0.0, 0.0]
                elif kind.startswith("for phase"):
                    recon.append(("phase", job, ph[0] if ph else "", s_ph[0], hrs, s_ph[1], cost))
                    s_job = [s_job[0] + s_ph[0], s_job[1] + s_ph[1]]
                    s_ph = [0.0, 0.0]
                else:
                    recon.append(("job", job, "", s_job[0], hrs, s_job[1], cost))
                continue
            if isinstance(c0, (int, float)) and c0s:
                if job is None or ph is None:
                    problems.append(f"detail line before a job/phase header {sh.name}!{r + 1}")
                    continue
                desc = str(row[14].value).strip()
                hrs, cost = _f(row[34].value), _f(row[36].value)
                rec = {"job": job, "phase": ph[0], "phase_desc": ph[1], "ct_num": ph[2], "ct": ph[3], "mth": day(c0)[:7],
                       "actual_date": day(row[4].value), "trans_type": str(row[7].value).strip(), "description": desc,
                       "hours": hrs, "cost": cost, "craft": "", "employee": "", "earn_desc": ""}
                if rec["trans_type"] == "PR":
                    pm = PRD.match(desc)
                    if pm:
                        rec.update(craft=pm.group(1) or "", employee=re.sub(r"\s+", " ", pm.group(5)).replace(" ,", ",").strip(),
                                   earn_desc=pm.group(7).strip())
                    elif hrs:
                        problems.append(f"payroll line with hours not understood {sh.name}!{r + 1}: {desc[:60]}")
                rows.append(rec)
                s_ct = [s_ct[0] + hrs, s_ct[1] + cost]
    failures = [x for x in recon if abs(x[3] - x[4]) > 0.011 or abs(x[5] - x[6]) > 0.011]
    return {"rows": rows, "recon": recon, "failures": failures, "jobs": jobs, "problems": problems}


def through_from_name(path):
    m = DATED.search(os.path.basename(path))
    return m.group(1) if m else ""


def summarize(data, through="", only=None):
    """B14 buckets per job -> list of dicts (COLS)."""
    by_job = defaultdict(list)
    for r in data["rows"]:
        by_job[r["job"]].append(r)
    bad = defaultdict(int)
    for x in data["failures"]:
        bad[x[1]] += 1
    jobs_with_total = {x[1] for x in data["recon"] if x[0] == "job"}
    out = []
    for job in sorted(by_job):
        if only and job not in only:
            continue
        rs = by_job[job]
        rate = defaultdict(lambda: [0.0, 0.0])
        for r in rs:
            if r["ct"] == "L" and r["hours"] > 0:
                p = r["employee"] or r["description"]
                rate[p][0] += r["hours"]
                rate[p][1] += r["cost"]

        def cls(r):
            p = r["employee"] or r["description"]
            if p in FOREMEN:
                return "foreman"
            if r["craft"] == "2" or r["description"].startswith("2/"):
                h, c = rate[p]
                return "operator" if h and 25 <= c / h <= 28.0001 else "laborer"
            return "bricklayer"

        t = defaultdict(float)
        hcls = defaultdict(float)
        odd = defaultdict(float)
        for r in rs:
            if r["ct"] == "L":
                if r["phase"].startswith("01650"):
                    t["trucking"] += r["cost"]
                elif r["hours"] > 0:
                    t["labor"] += r["cost"]
                    t["hours"] += r["hours"]
                    hcls[cls(r)] += r["hours"]
                else:
                    t["burden"] += r["cost"]
            else:
                b = BUCKET.get(r["ct"])
                if b is None:
                    b = "other"
                    odd[f"{r['ct_num']}-{r['ct']}"] += r["cost"]
                t[b] += r["cost"]
        total = sum(r["cost"] for r in rs)
        notes = []
        if odd:
            notes.append("in other: " + ", ".join(f"{k} ${v:,.0f}" for k, v in sorted(odd.items())))
        neg_hours = sum(r["hours"] for r in rs if r["ct"] == "L" and r["hours"] < 0)
        if neg_hours:
            notes.append(f"{-neg_hours:g} reversed hours counted as burden (as the B14 converter does)")
        if not t["hours"] and t["sub"]:
            notes.append("no FB labor hours, sub cost booked: labor subcontract? (B14 --labor-sub)")
        ok = job in jobs_with_total and not bad[job]
        if not ok:
            notes.append("Vista subtotals don't match the lines" if bad[job] else "no 'Total For Job' line read")
        months = sorted(r["mth"] for r in rs if r["mth"])
        dates = sorted(r["actual_date"] for r in rs if r["actual_date"])
        lab = t["labor"] + t["trucking"] * 0
        out.append({"job": job, "job_name": data["jobs"].get(job, ""), "through_date": through or (dates[-1] if dates else ""),
                    "recon_ok": "y" if ok else "n", "rows": len(rs),
                    "first_month": months[0] if months else "", "last_month": months[-1] if months else "",
                    "last_actual_date": dates[-1] if dates else "",
                    "labor_hours": round(t["hours"], 2), "act_mandays": round(t["hours"] / 8, 2),
                    **{f"hours_{c}": round(hcls[c], 2) for c in ("bricklayer", "foreman", "laborer", "operator")},
                    "act_labor_base": round(lab, 2), "act_burden": round(t["burden"], 2),
                    "burden_pct": round(t["burden"] / lab, 4) if lab else "",
                    "act_trucking": round(t["trucking"], 2), "act_material": round(t["material"], 2),
                    "act_equipment": round(t["equipment"], 2), "act_sub": round(t["sub"], 2),
                    "act_other": round(t["other"] + t["trucking"], 2), "act_total": round(total, 2),
                    "notes": "; ".join(notes)})
    return out


def actuals_rows(data, job, through, estimate=""):
    """job_cost.py actuals.csv rows for one job, line for line as vista_jc_detail.py writes them (final = n)."""
    rs = [r for r in data["rows"] if r["job"] == job]
    s = summarize({**data, "rows": rs}, through)[0] if rs else None
    rate = defaultdict(lambda: [0.0, 0.0])
    for r in rs:
        if r["ct"] == "L" and r["hours"] > 0:
            p = r["employee"] or r["description"]
            rate[p][0] += r["hours"]; rate[p][1] += r["cost"]
    out = []
    base = dict(job_number=job, estimate=estimate, through_date=through, final="n", labor_class="", hours="", burdened="",
                labor_sub="", tax_exempt="")
    for r in rs:
        if r["ct"] == "L":
            if r["phase"].startswith("01650"):
                out.append({**base, "cost_type": "other", "amount": f"{r['cost']:.2f}"})
            elif r["hours"] > 0:
                p = r["employee"] or r["description"]
                if p in FOREMEN:
                    c = "foreman"
                elif r["craft"] == "2" or r["description"].startswith("2/"):
                    h, cc = rate[p]; c = "operator" if 25 <= cc / h <= 28.0001 else "laborer"
                else:
                    c = "bricklayer"
                out.append({**base, "cost_type": "labor", "labor_class": c, "hours": r["hours"], "amount": f"{r['cost']:.2f}", "burdened": "n"})
        else:
            out.append({**base, "cost_type": {"M": "material", "I": "material", "E": "equipment", "R": "equipment",
                                              "O": "other", "S": "sub"}.get(r["ct"], "other"), "amount": f"{r['cost']:.2f}"})
    return out, s


def write_csv(path, rows, cols):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, lineterminator="\n", extrasaction="ignore")
        w.writeheader(); w.writerows(rows)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xls"); ap.add_argument("out")
    ap.add_argument("--through", default="", help="YYYY-MM-DD; default: the date in the file name, else the last line's date")
    ap.add_argument("--jobs", default="", help="comma list of Vista job numbers to keep (default: all)")
    ap.add_argument("--actuals-dir", default="")
    a = ap.parse_args(argv)
    data = read(a.xls)
    through = a.through or through_from_name(a.xls)
    only = {j.strip() for j in a.jobs.split(",") if j.strip()} or None
    rows = summarize(data, through, only)
    write_csv(a.out, rows, COLS)
    if a.actuals_dir:
        os.makedirs(a.actuals_dir, exist_ok=True)
        for r in rows:
            act, _ = actuals_rows(data, r["job"], r["through_date"])
            write_csv(os.path.join(a.actuals_dir, f"{r['job']} actuals.csv"), act, ACOLS)
    tot = sum(r["cost"] for r in data["rows"])
    print(f"{len(rows)} jobs · {len(data['rows'])} lines · ${tot:,.2f} · {len(data['recon'])} Vista subtotals checked, "
          f"{len(data['failures'])} mismatched · through {through or 'last line date per job'}")
    for p in data["problems"][:20]:
        print("⚠ " + p)
    for x in data["failures"][:20]:
        print(f"⚠ job {x[1]} {x[0]} {x[2]}: lines {x[3]:,.2f} h / ${x[5]:,.2f} vs Vista {x[4]:,.2f} h / ${x[6]:,.2f}")
    return 1 if data["failures"] or data["problems"] else 0


if __name__ == "__main__":
    sys.exit(main())
