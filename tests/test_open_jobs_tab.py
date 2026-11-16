"""The Open Jobs tab and Dashboard section appear only when --open-jobs is passed."""
import csv
import sys
import tempfile
import unittest
from argparse import Namespace
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "dashboard"))
import build_dashboard as bd  # noqa: E402


def write(path, cols, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    return str(path)


class OpenJobsTab(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        d = Path(self.tmp.name)
        self.bl = write(d / "bl.csv",
                        ["bid_id", "project", "city", "job_type", "gc", "bid_date", "status", "submitted_amount"],
                        [dict(bid_id="b1", project="Harbor Warehouse", city="Austin", job_type="education",
                              gc="Acme GC", bid_date="2026-03-01", status="awarded", submitted_amount="100000")])
        self.bm = write(d / "bm.csv",
                        ["metrics_id", "project", "estimate_status", "bid_id", "bid_price", "job_number",
                         "wages", "wages_vs_standard", "notes"],
                        [])
        row = {k: "" for k in bd.OPEN_COLS}
        row.update(job="4101", project="Harbor Warehouse", status="Active", estimate_id="tm-9",
                   match="Jobs row EstimateID", through_date="2026-11-13", flag="over",
                   pct_used="0.5", pct_basis="billed / contract")
        self.oj = write(d / "oj.csv", bd.OPEN_COLS, [row])
        self.out = str(d / "Dashboard.xlsx")
        self.data = str(d / "Dashboard Data.txt")

    def tearDown(self):
        self.tmp.cleanup()

    def args(self, open_jobs=""):
        return Namespace(bid_log=self.bl, bid_metrics=self.bm, prices="", job_cost="", units="",
                         open_jobs=open_jobs, out=self.out, data=self.data, previous="",
                         period="YTD", stamp="2026-11-16 12:00", no_recalc=True)

    def test_switch_off_has_no_tab_and_no_section(self):
        a = self.args("")
        P = bd.prepare(a)
        self.assertFalse(P["show_open_jobs"])
        self.assertNotIn("tblOpenJobs", P["T"])
        bd.build(a, P)
        wb = load_workbook(self.out)
        self.assertNotIn("Open Jobs", wb.sheetnames)
        dash = wb["Dashboard"]
        labels = [dash[f"A{r}"].value for r in range(1, 80)]
        self.assertNotIn("Open jobs — cost to date vs estimate (Open Jobs tab)", labels)
        self.assertNotIn("Open jobs", labels)
        self.assertFalse(any(isinstance(v, str) and "Open Jobs.csv" in v for v in P["about"]))

    def test_switch_on_adds_tab_and_section(self):
        a = self.args(self.oj)
        P = bd.prepare(a)
        self.assertTrue(P["show_open_jobs"])
        self.assertIn("tblOpenJobs", P["T"])
        self.assertEqual(len(P["oj"]), 1)
        bd.build(a, P)
        wb = load_workbook(self.out)
        self.assertIn("Open Jobs", wb.sheetnames)
        dash = wb["Dashboard"]
        labels = [dash[f"A{r}"].value for r in range(1, 80)]
        self.assertIn("Open jobs — cost to date vs estimate (Open Jobs tab)", labels)
        self.assertIn("Open jobs", labels)
        self.assertTrue(any(isinstance(v, str) and "Open Jobs.csv (1 open jobs)" in v for v in P["about"]))


class PublicRepo(unittest.TestCase):
    """The public print kit must not ship real employee or job names in this feature's files."""
    BANNED = (
        "Craig Manufacturing", "Rowlett Creek", "Meacham",
        "Leon, Abel", "Banuelos, Oscar", "Borja II", "Angulo-Nagy",
        "Rico-Robles", "Crissinger", "Israel Gonzalez", "John Cody",
        'by="Jason"', "by=\"Jason\"",
    )
    FILES = (
        "jobcost/jc_detail_multi.py", "jobcost/jc_detail_summary.py", "jobcost/foremen.py",
        "jobcost/open_jobs.py", "dashboard/build_dashboard.py",
        "tests/test_jc_detail_multi.py", "tests/test_open_jobs.py",
    )

    def test_no_real_names_in_open_jobs_files(self):
        hits = []
        for rel in self.FILES:
            text = (ROOT / rel).read_text(encoding="utf-8")
            for s in self.BANNED:
                if s in text:
                    hits.append(f"{rel}: {s}")
        self.assertEqual(hits, [])


if __name__ == "__main__":
    unittest.main()
