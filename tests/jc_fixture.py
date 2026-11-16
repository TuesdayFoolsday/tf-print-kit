"""Write a small JC Detail .xls in the layout of the real Vista export (JCDetail.rpt saved as Excel, Nov 2026).

Column positions copied from a real multi-job export (Nov 2026): job / phase headers in column A,
detail lines with Mth (A), Trans# (B), posted date (C), actual date (E), trans type (H), source (J), description (O),
units (AF), hours (AI), cost (AK); subtotal lines with the label in P, hours in AH and cost in AK.
Synthetic people and amounts; no real payroll data is committed.
"""
import xlwt

MTH = {"2026-09": 46266.0, "2026-10": 46296.0, "2026-11": 46327.0}   # Excel serials for the 1st of each month


def pr(emp_no, name, craft, rate, earn_code, earn, hours, cost, mth="2026-10", day=46300.0):
    return dict(mth=mth, day=day, tt="PR", src="PR Entry", hours=hours, cost=cost,
                desc=f"{craft} {rate:.2f} / {emp_no} / {name}{earn_code} {earn}")


def ap(vendor_no, vendor, invoice, cost, mth="2026-10", day=46301.0, note="Material"):
    return dict(mth=mth, day=day, tt="AP", src="AP Entry", hours=0.0, cost=cost,
                desc=f"{vendor_no} {vendor} {invoice} / TR# 1234 / 1 / APCo: 1 {note}")


def write(path, jobs, break_total=None, sheets=1):
    """jobs: [(number, name, [(phase, phase_desc, ct_num, ct_code, [lines])])]. break_total: (job, kind) whose Vista
    subtotal is written wrong on purpose. sheets=2 puts the second half of the jobs on a second sheet."""
    wb = xlwt.Workbook()
    books = [wb.add_sheet(f"Sheet{i + 1}") for i in range(sheets)]
    state = {"r": 0, "s": books[0]}

    def put(cells):
        for c, v in cells.items():
            state["s"].write(state["r"], c, v)
        state["r"] += 2

    put({0: "JC Detail"}); put({1: "Company:", 13: 1.0})
    trans = 1000
    for i, (num, name, phases) in enumerate(jobs):
        if sheets > 1 and i == (len(jobs) + 1) // 2:
            state.update(r=0, s=books[1])
        put({0: f" {num}-  {name}   "})
        jh = jc = 0.0
        for ph, pdesc, ctn, ctc, lines in phases:
            put({0: f" {ph}-  -                     {pdesc}   {ctn}    {ctc}  "})
            h = c = 0.0
            for ln in lines:
                trans += 1
                put({0: MTH[ln["mth"]], 1: float(trans), 2: ln["day"] + 3, 4: ln["day"], 7: ln["tt"], 9: ln["src"],
                     14: ln["desc"], 31: 0.0, 34: float(ln["hours"]), 36: float(ln["cost"])})
                h += ln["hours"]; c += ln["cost"]
            bump = 1.0 if break_total == (num, "cost type") else 0.0
            put({15: "Total for Cost Type:", 33: h, 36: c + bump})
            put({15: "Total For Phase:", 33: h, 36: c})
            jh += h; jc += c
        put({15: "Total For Job:", 33: jh, 36: jc})
    wb.save(path)
