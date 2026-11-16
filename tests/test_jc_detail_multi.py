"""Multi-job JC Detail reader: layout, Vista subtotal checks, and the B14 mapping kept identical to vista_jc_detail.py."""
import csv
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "jobcost"))
sys.path.insert(0, str(ROOT / "tests"))
import jc_detail_multi as m  # noqa: E402
import jc_fixture as fx  # noqa: E402

REAL = os.environ.get("JC_DETAIL_REAL", "")   # optional: path to a real export (never committed)


def craig_like():
    """Job 2530-style: masons, a foreman, a laborer, an operator-rate laborer, burden, trucking, AP lines."""
    lab = [fx.pr(101, "Mason, Ann", "1/Bricklayer", 30, 1, "Regular Earnings", 40, 1200),
           fx.pr(102, "Banuelos, Oscar", "1/Journeyman", 40, 1, "Regular Earnings", 32, 1280),
           fx.pr(103, "Helper, Bob", "2/Experience", 20, 1, "Regular Earnings", 40, 800),
           fx.pr(104, "Lift, Carl", "2/Experience", 27, 1, "Regular Earnings", 20, 540),
           fx.pr(101, "Mason, Ann", "1/Bricklayer", 30, 10, "Payroll Tax", 0, 92.5),
           fx.pr(103, "Helper, Bob", "2/Experience", 20, 12, "Workmen's Comp", 0, 61.2),
           fx.pr(104, "Lift, Carl", "2/Experience", 27, 13, "401K Match", 0, 10.8)]
    truck = [fx.pr(103, "Helper, Bob", "2/Experience", 20, 1, "Regular Earnings", 4, 80),
             fx.pr(103, "Helper, Bob", "2/Experience", 20, 10, "Payroll Tax", 0, 6.1)]
    return [("2530", "Craig Manufacturing", [
        ("01650", "Trucking", "1", "L", truck),
        ("04200", "Masonry", "1", "L", lab),
        ("04200", "Masonry", "2", "M", [fx.ap(55, "Acme Block", "INV1", 5000.0), fx.ap(55, "Acme Block", "INV2", -250.0)]),
        ("04200", "Masonry", "8", "I", [fx.ap(56, "Freight Co", "F1", 300.0)]),
        ("04200", "Masonry", "5", "R", [fx.ap(57, "Lift Rentals", "R1", 2100.0)]),
        ("04200", "Masonry", "6", "O", [fx.ap(58, "Misc Co", "O1", 75.0)]),
        ("04200", "Masonry", "11", "SL", [fx.ap(59, "Labor Sub", "S1", 400.0)]),
    ])]


def second_job():
    return [("2622", "Rowlett Creek CMAR", [
        ("04200", "Masonry", "1", "L", [fx.pr(201, "Leon, Abel", "1/Journeyman", 45, 1, "Regular Earnings", 10, 450),
                                         fx.pr(202, "Brick, Dan", "1/Bricklayer", 30, 1, "Regular Earnings", 16, 480)]),
        ("04200", "Masonry", "3", "S", [fx.ap(60, "Some Sub", "X1", 999.0)]),
    ])]


class Reader(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def xls(self, jobs, name="JC Detail 2026-11-13.xls", **kw):
        p = self.dir / name
        fx.write(str(p), jobs, **kw)
        return str(p)

    def test_many_jobs_two_sheets_and_buckets(self):
        p = self.xls(craig_like() + second_job(), sheets=2)
        d = m.read(p)
        self.assertEqual(d["failures"], [])
        self.assertEqual(d["problems"], [])
        self.assertEqual(sorted(d["jobs"]), ["2530", "2622"])
        rows = {r["job"]: r for r in m.summarize(d, m.through_from_name(p))}
        c = rows["2530"]
        self.assertEqual(c["through_date"], "2026-11-13")
        self.assertEqual(c["recon_ok"], "y")
        self.assertEqual(c["labor_hours"], 132.0)               # trucking hours left out
        self.assertEqual(c["act_mandays"], 16.5)                # whole-crew hours / 8
        self.assertEqual(c["act_labor_base"], 3820.0)           # base wage only
        self.assertEqual(c["act_burden"], 164.5)                # masonry burden, not trucking
        self.assertEqual(c["act_trucking"], 86.1)               # all 01650 L lines
        self.assertEqual(c["act_material"], 5050.0)             # M + I (with tax, as Vista books it)
        self.assertEqual(c["act_equipment"], 2100.0)
        self.assertEqual(c["act_other"], 75.0 + 400.0 + 86.1)   # O + SL + trucking
        self.assertIn("11-SL $400", c["notes"])
        self.assertEqual((c["hours_bricklayer"], c["hours_foreman"], c["hours_laborer"], c["hours_operator"]),
                         (40.0, 32.0, 40.0, 20.0))
        self.assertAlmostEqual(c["act_total"], 3820 + 164.5 + 86.1 + 5050 + 2100 + 75 + 400, places=2)
        r = rows["2622"]
        self.assertEqual((r["hours_foreman"], r["hours_bricklayer"], r["act_sub"]), (10.0, 16.0, 999.0))

    def test_subtotal_mismatch_marks_job_not_ok(self):
        p = self.xls(craig_like() + second_job(), break_total=("2622", "cost type"))
        d = m.read(p)
        self.assertTrue(d["failures"])
        rows = {r["job"]: r for r in m.summarize(d)}
        self.assertEqual(rows["2622"]["recon_ok"], "n")
        self.assertEqual(rows["2530"]["recon_ok"], "y")
        self.assertEqual(m.main([p, str(self.dir / "out.csv")]), 1)

    def test_same_lines_as_the_ledger_script(self):
        """Layout parity with jc_detail_ledger.py (the reader used for the Nov 4 history ledgers)."""
        p = self.xls(craig_like() + second_job(), sheets=2)
        led, js = self.dir / "l.csv", self.dir / "l.json"
        subprocess.run([sys.executable, str(ROOT / "jobcost/jc_detail_ledger.py"), p, str(led), str(js)],
                       check=True, capture_output=True)
        ledger = list(csv.DictReader(open(led)))
        mine = m.read(p)["rows"]
        self.assertEqual(len(ledger), len(mine))
        for a, b in zip(ledger, mine):
            self.assertEqual((a["job"], a["phase"], a["ct_code"], float(a["hours"]), float(a["cost"]), a["employee"]),
                             (b["job"], b["phase"], b["ct"], b["hours"], b["cost"], b["employee"]))

    def test_b14_parity_with_vista_jc_detail(self):
        """Single job: the same actuals rows, amounts and buckets as the live B14 converter (final aside)."""
        p = self.xls(craig_like())
        old = self.dir / "old.csv"
        subprocess.run([sys.executable, str(ROOT / "jobcost/vista_jc_detail.py"), p, str(old), "2530", "Craig",
                        "2026-11-13", "Banuelos"], check=True, capture_output=True)
        a = list(csv.DictReader(open(old)))
        b, _ = m.actuals_rows(m.read(p), "2530", "2026-11-13", "Craig")
        key = lambda r: (r["cost_type"], r["labor_class"], str(float(r["hours"])) if r["hours"] != "" else "", r["amount"], r["burdened"])
        self.assertEqual(sorted(map(key, a)), sorted(map(key, b)))
        self.assertTrue(all(r["final"] == "n" for r in b))

    @unittest.skipUnless(REAL and os.path.exists(REAL), "set JC_DETAIL_REAL to a real multi-job export")
    def test_real_export(self):
        d = m.read(REAL)
        self.assertEqual(d["failures"], [])
        rows = {r["job"]: r for r in m.summarize(d)}
        self.assertGreater(len(rows), 10)
        if "2530" in rows:   # Craig Manufacturing, B14 Nov 1: base-wage labor $8,128
            self.assertEqual(rows["2530"]["act_labor_base"], 8128.0)


if __name__ == "__main__":
    unittest.main()
