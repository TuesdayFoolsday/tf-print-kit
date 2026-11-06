"""Bid Metrics v2 + v2b (hours, standard-wage labor, more conditions, vista_job) on top of PR #15.
Synthetic numbers only (this repo is public). Run: python3 -m unittest discover tests"""
import csv
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

KIT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KIT / "metrics"))
import bid_metrics as m  # noqa: E402
import conditions as C  # noqa: E402
import metrics_log as ml  # noqa: E402
import units as U  # noqa: E402

COST_MAP = """class,pattern,bucket,note
CMU,.,cmu_gray,
BRI,.,brick,
GRC,.,support:cmu,
REB,.,support:cmu,
RES,.,support:cmu,
WIR,.,support:cmu,
ANC,.,support:veneer,
CLN,brick,support:brick,
SCA,.,scaffold,
"""


def export(wages=None, grout_units="CbcYds", grout_qty=20.0, by_wall=False):
    """A small two-size CMU job (+ a little brick) as compact JSON. Labor totals reconcile with the totals block."""
    w = "Wall A" if by_wall else ""
    rows = [
        # cls, id, desc, amount(SF), units, qty, qty_units, labor, material, mandays, wall
        ["CMU", "08", '8" REGULAR', 1000.0, "SF", 1150.0, "Pieces", 6000.0, 2500.0, 8.0, w],
        ["CMU", "06", '6" REGULAR', 500.0, "SF", 575.0, "Pieces", 3500.0, 1100.0, 5.0, w],
        ["BRI", "MOD", "Modular brick", 300.0, "SF", 2100.0, "Pieces", 1800.0, 1500.0, 4.0, w],
        ["GRC", "G08", 'GROUT 8" CMU', grout_qty, "CY", grout_qty, grout_units, 1500.0, 3000.0, 2.0, w],
        ["REB", "#4", "#4 vertical", 2400.0, "LF", 2400.0, "LinFt", 600.0, 1000.0, 1.0, w],
        ["RES", "POS", "Positioners", 300.0, "Ea", 300.0, "Each", 0.0, 150.0, 0.0, w],
        ["WIR", "LAD", "Ladder wire", 1200.0, "LF", 1200.0, "LinFt", 0.0, 400.0, 0.0, w],
        ["CLN", "CLB", "Brick cleaning crew", 300.0, "SF", 300.0, "SF", 300.0, 30.0, 0.5, w],
        ["SCA", "SC", "Scaffold", 1500.0, "SF", 30.0, "Pieces", 700.0, 500.0, 1.5, w],
    ]
    lab = sum(r[7] for r in rows)
    mat = sum(r[8] for r in rows)
    wages = wages or {"Bricklayer": 30.0, "Forman": 65.0, "Labor": 20.0, "Operator": 27.0}
    hours = {"Bricklayer": 160.0, "Forman": 60.0, "Labor": 200.0, "Operator": 40.0}
    crew_base = sum(hours[k] * wages[k] for k in hours)
    crew_total, mat_total = crew_base * 1.5, mat * 1.2
    return {"job_title": "Synthetic Job", "job_number": "99001", "rows": rows,
            "totals": {"material": mat, "misc_taxable": 0, "tax_rate": 0.0825, "mat_total": mat_total,
                       "crew_total": crew_total, "equip_total": 0, "misc_total": 0, "sub_total": 0, "clean_total": 0,
                       "oh_pct": 0.1, "profit_pct": 0.1, "bid_price": round(crew_total + mat_total, 2),
                       "labor_base_total": lab, "crew_base": crew_base, "crew_cost": crew_base * 1.19, "equip_cost": 0},
            "wages": wages, "hours": hours}


class Env:
    def __init__(self, **kw):
        self.d = tempfile.TemporaryDirectory()
        self.p = Path(self.d.name)
        (self.p / "map.csv").write_text(COST_MAP)
        (self.p / "x.json").write_text(json.dumps(export(**kw)))

    def units(self, grout_spec="8:full;6:full", wall_grout=None):
        r = m.compute(str(self.p / "x.json"), str(self.p / "map.csv"), wall_grout=wall_grout)
        raw, (lf, mf), v2 = r.pop("_units_raw"), r.pop("_factors"), r.pop("_v2")
        return r, {u["unit_key"]: u for u in U.finish(raw, "tm-99001", r["project"], "tradesmen", "final", lf, mf,
                                                       grout_spec=grout_spec, v2=v2)}


def hist_header(e):
    r = m.compute(str(e.p / "x.json"), str(e.p / "map.csv"))
    keys = [k for k in r if not k.startswith("_") and k != "walls"]
    return ",".join(["metrics_id", "estimate_status", "estimator", "source_detail", "notes", "matches_bid_log"]
                    + [k for k in keys if k != "metrics_id"]) + "\n"


class StandardWages(unittest.TestCase):
    def test_v2_factors(self):
        f = m.v2_factors({"Bricklayer": 10, "Forman": 10, "Labor": 10, "Operator": 10}, 10 * (30 + 65 + 20 + 27))
        self.assertEqual(f["std_ratio"], 1.0)
        self.assertAlmostEqual(f["hours_per_labor"], 40 / 1420)
        self.assertEqual(m.v2_factors({}, 100), {})
        self.assertEqual(m.v2_factors({"Bricklayer": 5}, 0), {})

    def test_wage_bump_moves_priced_not_standard(self):
        _, std = Env().units()
        r2, bumped = Env(wages={"Bricklayer": 35.0, "Forman": 75.0, "Labor": 23.0, "Operator": 31.5}).units()
        self.assertLess(float(r2["labor_std_ratio"]), 1)
        for k in ("tm-99001|cmu_gray|8|full", "tm-99001|brick|modular|n/a"):
            self.assertGreater(bumped[k]["labor_per_unit"], std[k]["labor_per_unit"])
            self.assertAlmostEqual(bumped[k]["labor_std_per_unit"], std[k]["labor_std_per_unit"], places=2)
            self.assertAlmostEqual(bumped[k]["hours_per_unit"], std[k]["hours_per_unit"], places=3)

    def test_hours_are_the_units_share_of_the_crew(self):
        _, u = Env().units()
        unit_h = sum(x["hours"] for x in u.values() if x["hours"] != "")
        self.assertLess(unit_h, 460)                 # scaffold carries the rest of the 460 crew hours
        self.assertGreater(unit_h, 460 * 0.8)

    def test_metric_std_falls_back_to_priced(self):
        self.assertEqual(U.metric_std({"material": "cmu_gray", "labor_per_unit": "7.5", "labor_std_per_unit": ""}), 7.5)
        self.assertEqual(U.metric_std({"material": "cmu_gray", "labor_per_unit": "8.0", "labor_std_per_unit": "7.2"}), 7.2)

    def test_by_wall_runs_and_carries_hours(self):
        _, u = Env(by_wall=True).units(grout_spec=None, wall_grout={"wall a": "full"})
        self.assertTrue(u["tm-99001|cmu_gray|8|full"]["hours"])


class Split(unittest.TestCase):
    def test_parts_sum_exactly_to_labor_per_unit(self):
        _, u = Env().units()
        for k, x in u.items():
            if x["labor_lay"] == "":
                continue
            tot = x["labor_per_unit"] if x["material"] in ("cmu_gray", "cmu_arch", "brick") else x["labor_per_sf"]
            self.assertAlmostEqual(x["labor_lay"] + x["labor_rebar"] + x["labor_grout"] + x["labor_clean"], tot, places=6, msg=k)

    def test_fill_and_rebar(self):
        _, u = Env().units()
        x = u["tm-99001|cmu_gray|8|full"]
        self.assertTrue(120 <= x["grout_fill_pct"] <= 130)
        self.assertEqual(x["rebar_lf_per_sf"], 1.6)


class Conditions(unittest.TestCase):
    def _read(self, text):
        with tempfile.NamedTemporaryFile("w", suffix=".md", delete=False) as f:
            f.write(text)
        try:
            return C.read(f.name)
        finally:
            os.unlink(f.name)

    def test_fields_block_with_new_keys(self):
        c = self._read("talked\n## Fields\n- access: tight\n- secure_site: yes\n- travel: n\n- prevailing_wage: y\n"
                       "- difficulty: 4 (hard)\n- chase_walls: no\n")
        self.assertEqual((c["access"], c["secure_site"], c["travel"], c["prevailing_wage"], c["difficulty"], c["chase_walls"]),
                         ("tight", "y", "n", "y", "4", "n"))
        f = C.fields(c, 1000)
        self.assertEqual(f["difficulty"], "4")
        self.assertEqual(f["secure_site"], "y")

    def test_answers_heading_works_too(self):
        self.assertEqual(self._read("## Answers\n- access: difficult\n")["access"], "difficult")

    def test_absent_is_not_no(self):
        c = self._read("Grout full, 8 inch walls.\n")
        self.assertIsNone(c["chase_walls"])

    def test_bad_values_dropped(self):
        c = self._read("## Fields\n- difficulty: 9\n- access: sideways\n- travel: maybe\n")
        self.assertIsNone(c["difficulty"]); self.assertIsNone(c["access"]); self.assertIsNone(c["travel"])


class Merge(unittest.TestCase):
    def test_v2b_columns_accepted_and_appended(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            (d / "base.csv").write_text("metrics_id,project,source,estimate_status,notes\ntm-1,A,tradesmen,final,x\n")
            with redirect_stdout(io.StringIO()):
                ml.write_patch({"metrics_id": "tm-1", "set": {"vista_job": "2507", "difficulty": "4", "labor_std_ratio": "0.86"}},
                               ml.with_v2(["metrics_id", "project", "source", "estimate_status", "notes"]), {"tm-1"}, str(d / "p"))
            header, rows, applied, warns = ml.merge(str(d / "base.csv"), str(d / "p"))
            self.assertEqual(warns, [])
            row = dict(zip(header, rows[0]))
            self.assertEqual((row["notes"], row["vista_job"], row["difficulty"]), ("x", "2507", "4"))
            uh = ml.with_v2(["unit_key", "metrics_id"])
            for c in ("hours", "hours_per_unit", "labor_std_per_unit"):
                self.assertIn(c, uh)


class BidCheck(unittest.TestCase):
    def test_end_to_end(self):
        e = Env()
        p = e.p
        (p / "hist.csv").write_text(hist_header(e))
        (p / "units.csv").write_text(",".join(U.UNITS_COLS) + "\n")
        (p / "cond.md").write_text("## Fields\n- access: tight\n- difficulty: 4\n- secure_site: y\n")
        out = subprocess.run([sys.executable, str(KIT / "metrics/bid_check.py"), str(p / "x.json"), str(p / "map.csv"),
                              str(p / "hist.csv"), "--units", str(p / "units.csv"), "--conditions", str(p / "cond.md"),
                              "--grout", "8:full;6:full", "--status", "final", "--pending-out", str(p / "pm"),
                              "--units-pending-out", str(p / "pu"), "--by", "test"], capture_output=True, text=True)
        self.assertEqual(out.returncode, 0, out.stderr)
        self.assertIn("at standard wages", out.stdout)
        self.assertIn("crew-hrs/block", out.stdout)
        pm = [json.loads(f.read_text()) for f in (p / "pm").iterdir()]
        self.assertEqual((pm[0]["set"]["access"], pm[0]["set"]["difficulty"], pm[0]["set"]["secure_site"]), ("tight", "4", "y"))
        self.assertEqual(pm[0]["set"]["labor_std_ratio"], "1")
        pu = {json.loads(f.read_text())["unit_key"]: json.loads(f.read_text())["set"] for f in (p / "pu").iterdir()}
        self.assertTrue(pu["tm-99001|cmu_gray|8|full"]["labor_std_per_unit"])
        self.assertTrue(pu["tm-99001|cmu_gray|8|full"]["hours_per_unit"])

    def test_bumped_wages_print_the_premium(self):
        e = Env(wages={"Bricklayer": 35.0, "Forman": 75.0, "Labor": 23.0, "Operator": 31.5})
        p = e.p
        (p / "hist.csv").write_text(hist_header(e))
        (p / "units.csv").write_text(",".join(U.UNITS_COLS) + "\n")
        out = subprocess.run([sys.executable, str(KIT / "metrics/bid_check.py"), str(p / "x.json"), str(p / "map.csv"),
                              str(p / "hist.csv"), "--units", str(p / "units.csv"), "--grout", "8:full;6:full"],
                             capture_output=True, text=True)
        self.assertIn("wages +", out.stdout)


if __name__ == "__main__":
    unittest.main()
