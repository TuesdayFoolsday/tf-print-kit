"""open_jobs.py: every open job gets a row, estimate match rules, the % complete switch and the over flag."""
import csv
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "jobcost"))
import open_jobs as oj  # noqa: E402


def contracts_xlsx(path, amounts):
    """Vista 'JC Contracts - Detail' layout: job header in column A, then a row with Original:/Current: (cols V, AC)."""
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active
    r = 3
    for job, amount in amounts.items():
        ws.cell(r, 1, f" {job}-   Job {job}")
        ws.cell(r + 2, 2, "Department:"); ws.cell(r + 2, 6, "1   Fenimore-Blythe Department")
        ws.cell(r + 4, 2, "Interface Taxes:"); ws.cell(r + 4, 8, "N")
        ws.cell(r + 4, 22, "Original:"); ws.cell(r + 4, 24, amount)
        ws.cell(r + 4, 29, "Current:"); ws.cell(r + 4, 32, amount)
        r += 8
    wb.save(path)
    return str(path)


def write(path, cols, rows):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(rows)
    return str(path)


class OpenJobs(unittest.TestCase):
    def setUp(self):
        self.t = tempfile.TemporaryDirectory(); d = Path(self.t.name); self.d = d
        acols = ["job", "through_date", "recon_ok", "labor_hours", "act_mandays", "act_labor_base", "act_material",
                 "act_total", "act_burden", "act_equipment", "act_sub", "act_other", "notes"]
        self.actuals = write(d / "a.csv", acols, [
            dict(job="2601", through_date="2026-11-13", recon_ok="y", labor_hours=800, act_mandays=100, act_labor_base=24000,
                 act_material=30000, act_total=70000, act_burden=6000, act_equipment=1000, act_sub=0, act_other=0, notes=""),
            dict(job="2602", through_date="2026-11-13", recon_ok="y", labor_hours=80, act_mandays=10, act_labor_base=2000,
                 act_material=1000, act_total=4000, act_burden=500, act_equipment=0, act_sub=0, act_other=0, notes=""),
            dict(job="2604", through_date="2026-11-13", recon_ok="n", labor_hours=8, act_mandays=1, act_labor_base=200,
                 act_material=0, act_total=300, act_burden=0, act_equipment=0, act_sub=0, act_other=0, notes=""),
            dict(job="2999", through_date="2026-11-13", recon_ok="y", labor_hours=8, act_mandays=1, act_labor_base=1,
                 act_material=0, act_total=1, act_burden=0, act_equipment=0, act_sub=0, act_other=0, notes="")])
        self.jobs = write(d / "j.csv", ["VistaJobNo", "Title", "Status", "EstimateID", "ContractValue"], [
            dict(VistaJobNo="2601", Title="Big School", Status="Active", EstimateID="", ContractValue="200000"),
            dict(VistaJobNo="2602", Title="Small Shop", Status="Active", EstimateID="tm-1", ContractValue=""),
            dict(VistaJobNo="2603", Title="No Estimate Yet", Status="Setup", EstimateID="", ContractValue=""),
            dict(VistaJobNo="2604", Title="Bad Totals", Status="Active", EstimateID="tm-1", ContractValue="5000"),
            dict(VistaJobNo="2605", Title="Done", Status="Closed out", EstimateID="", ContractValue="")])
        bcols = ["metrics_id", "project", "estimate_status", "vista_job", "tax_rate", "crew_hours", "crew_cost",
                 "equipment_cost", "cmu_gray_labor_base", "cmu_gray_material_base"]
        self.bm = write(d / "b.csv", bcols, [
            dict(metrics_id="tm-9", project="Big School", estimate_status="final", vista_job="2601", tax_rate="0.0825",
                 crew_hours="1600", crew_cost="60000", equipment_cost="5000", cmu_gray_labor_base="48000",
                 cmu_gray_material_base="80000"),
            dict(metrics_id="tm-1", project="Small Shop", estimate_status="final", vista_job="", tax_rate="0",
                 crew_hours="160", crew_cost="6000", equipment_cost="0", cmu_gray_labor_base="4000",
                 cmu_gray_material_base="2000")])
        self.billing = write(d / "bill.csv", ["vista_no", "total_cost", "billed"],
                             [dict(vista_no="2601", total_cost=70000, billed=100000), dict(vista_no="2602", total_cost=4000, billed=3000)])
        # 2602 has 0 in Vista (as every job did on Nov 16); the Jobs list's ContractValue is never read.
        self.contracts = contracts_xlsx(d / "c.xlsx", {"2601": 200000, "2602": 0, "2604": 5000})
        self.pm = write(d / "pm.csv", ["job", "pct_complete", "as_of", "by"], [dict(job="2602", pct_complete="40", as_of="2026-11-13", by="Jason")])
        self.laid = write(d / "laid.csv", ["job", "job_pct_complete"], [dict(job="2601", job_pct_complete="45.0")])

    def tearDown(self):
        self.t.cleanup()

    def run_(self, *extra):
        out = self.d / "o.csv"
        oj.main(["--actuals", self.actuals, "--jobs", self.jobs, "--bid-metrics", self.bm, "--billing", self.billing,
                 "--contracts", self.contracts, "--pm", self.pm, "--laid", self.laid, "--out", str(out), "--summary", str(self.d / "s.json"), *extra])
        return {r["job"]: r for r in csv.DictReader(open(out))}

    def test_every_open_job_and_flags(self):
        r = self.run_()
        self.assertEqual(sorted(r), ["2601", "2602", "2603", "2604"])            # closed left out, 2999 not on the list
        self.assertEqual(r["2603"]["flag"], "no estimate entered")
        self.assertEqual(r["2601"]["match"], "Bid Metrics vista_job")
        self.assertEqual(r["2602"]["match"], "Jobs row EstimateID")
        self.assertIn("subtotals", r["2604"]["flag"])
        self.assertEqual(r["2604"]["proj_total"], "")                             # never compared

    def test_billed_default_vista_contract_only(self):
        r = self.run_()
        self.assertEqual(float(r["2601"]["pct_used"]), 0.5)                       # 100k billed / 200k Vista contract
        self.assertEqual(r["2601"]["pct_basis"], "billed / contract")
        self.assertEqual(float(r["2601"]["pct_laid"]), 0.45)
        self.assertEqual(r["2602"]["pct_used"], "")                               # PM % no longer overrides billed
        self.assertEqual(float(r["2602"]["pct_pm"]), 0.4)
        self.assertIn("no contract amount in Vista", r["2602"]["notes"])
        self.assertEqual(r["2604"]["contract"], "5000.0")

    def test_jobs_list_contract_ignored(self):
        oj.main(["--actuals", self.actuals, "--jobs", self.jobs, "--bid-metrics", self.bm, "--billing", self.billing,
                 "--out", str(self.d / "o2.csv")])
        r = {x["job"]: x for x in csv.DictReader(open(self.d / "o2.csv"))}
        self.assertEqual(r["2601"]["pct_used"], "")                               # 200000 on the Jobs row, not used
        self.assertEqual(r["2601"]["contract"], "")

    def test_cost_and_revenue_cross_check(self):
        rows = list(csv.DictReader(open(self.billing)))
        rows[0]["total_cost"] = "65000"                                           # JC Detail says 70,000
        write(self.d / "bill.csv", list(rows[0].keys()), rows)
        r = self.run_()
        self.assertIn("Vista reports disagree", r["2601"]["flag"])
        self.assertEqual(r["2601"]["proj_total"], "")
        self.assertNotIn("disagree", r["2602"]["flag"])

    def test_projection_and_over(self):
        r = self.run_("--pct-source", "pm")
        r["2601"] = self.run_()["2601"]
        b = r["2601"]   # 100 md at 50% -> 200 vs 200 est; labor 24k/0.5=48k vs 48k; material 30k/.5=60k vs 86.6k
        self.assertEqual(float(b["proj_mandays"]), 200.0)
        self.assertAlmostEqual(float(b["mandays_over_pct"]), 0.0)
        self.assertEqual(float(b["est_material"]), 86600.0)
        self.assertNotIn("over", b["flag"])
        s = r["2602"]   # 10 md at 40% -> 25 vs 20 est = +25%
        self.assertAlmostEqual(float(s["mandays_over_pct"]), 0.25)
        self.assertTrue(s["flag"].startswith("over"))

    def test_switch_to_pm_and_laid(self):
        r = self.run_("--pct-source", "pm")
        self.assertEqual(r["2601"]["pct_used"], "")                                # no PM % for 2601
        self.assertEqual(float(r["2602"]["pct_used"]), 0.4)
        r = self.run_("--pct-source", "laid")
        self.assertEqual(float(r["2601"]["pct_used"]), 0.45)
        self.assertEqual(r["2601"]["pct_basis"], "laid %")

    def test_two_estimates_claim_a_job(self):
        rows = list(csv.DictReader(open(self.bm)))
        rows.append({**rows[0], "metrics_id": "in-x", "estimate_status": "final"})
        write(self.d / "b.csv", list(rows[0].keys()), rows)
        r = self.run_()
        self.assertIn("two estimates claim this job", r["2601"]["flag"])


if __name__ == "__main__":
    unittest.main()
