"""System (/takeoff) estimates as a Bid Metrics source (System Rules A35, v1.67)."""
import json, os, subprocess, sys, tempfile, unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "metrics"))
import bid_metrics as m  # noqa: E402
import metrics_log as ml  # noqa: E402

COST_MAP = """class,pattern,bucket,note
CMU,.,cmu_gray,
GRC,.,support:cmu,
REB,.,support:cmu,
RES,.,support:cmu,
WIR,.,support:cmu,
MOR,gray,support:cmu,
FEE,.,support:all,
SCA,.,scaffold,
DF,.,other_items,
CLN,.,support:cmu,
"""


def synthetic():
    """Made-up system estimate (no real job data). Parts reconcile to the bid price."""
    rows = [  # cls, id, desc, amount(SF), units, qty, qty_units, labor, material, mandays, wall
        ["CMU", "8-SYS", '8" field', 900.0, "SF", 1000.0, "Pieces", 4000.0, 2200.0, 6.0, ""],
        ["CMU", "8-SYS", '8" KO / bond beam', 100.0, "SF", 110.0, "Pieces", 500.0, 260.0, 0.8, ""],
        ["GRC", "GROUT 8", "grout", 6.0, "CY", 6.0, "CbcYds", 600.0, 1500.0, 3.0, ""],
        ["MOR", "MORTAR", "mortar type S gray", 3.0, "bags", 3.0, "bags", 0.0, 900.0, 0.0, ""],
        ["REB", "#5 8", "rebar", 900.0, "LF", 900.0, "LinFt", 120.0, 850.0, 0.5, ""],
        ["RES", "RB POS 8", "positioners", 150.0, "Ea", 150.0, "Each", 0.0, 60.0, 0.0, ""],
        ["WIR", "LADDER 8", "ladder", 700.0, "LF", 700.0, "LinFt", 0.0, 160.0, 0.0, ""],
        ["FEE", "DEL BLOCK", "block delivery", 1, "Ea", 1, "Each", 0.0, 375.0, 0.0, ""],
        ["SCA", "SCAF 8", "scaffold", 25.0, "frames", 25.0, "Pieces", 220.0, 500.0, 1.2, ""],
        ["DF", "DF", "door frames", 2, "Ea", 2, "Each", 100.0, 0.0, 0.4, ""],
        ["CLN", "CLEAN CMU", "CMU cleaning", 1000.0, "SF", 1000.0, "SF", 350.0, 30.0, 0.0, ""],
    ]
    lab = sum(r[7] for r in rows); mat = sum(r[8] for r in rows); tax = mat * 0.0825; mk = 1.1 * 1.08
    T = {"material": mat, "misc_taxable": 0, "tax_rate": 0.0825, "mat_total": (mat + tax) * mk,
         "crew_total": lab * 1.19 * mk, "equip_total": 1000 * mk, "misc_total": 500 * mk, "sub_total": 0,
         "clean_total": 0, "oh_pct": 0.1, "profit_pct": 0.088, "labor_base_total": lab, "crew_base": lab,
         "crew_cost": lab * 1.19, "equip_cost": 1000}
    T["bid_price"] = round(T["mat_total"] + T["crew_total"] + T["equip_total"] + T["misc_total"], 2)
    return {"job_title": "Synthetic System Job", "job_number": "", "rows": rows, "totals": T,
            "wages": {"Bricklayer": 32.0, "Forman": 65.0, "Labor": 22.0, "Operator": 26.0},
            "hours": {"Bricklayer": 100.0, "Forman": 50.0, "Labor": 100.0, "Operator": 50.0},
            "source": "system", "source_detail": "synthetic test"}


class SystemSource(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.TemporaryDirectory(); self.p = Path(self.d.name)
        (self.p / "map.csv").write_text(COST_MAP)
        self.c = synthetic(); (self.p / "sys.json").write_text(json.dumps(self.c))

    def test_ids_and_sources(self):
        self.assertEqual(ml.make_id("system", "", "Sample Job"), "sy-sample-job")
        self.assertEqual(ml.make_id("tradesmen", "26067", "x"), "tm-26067")
        self.assertTrue(ml.make_id("india", "", "Hobby Lobby").startswith("in-"))
        hdr = ["metrics_id", "project", "source", "estimate_status"]
        ok = {"metrics_id": "sy-x", "set": {"project": "X", "source": "system", "estimate_status": "final"}}
        self.assertEqual(ml.validate(ok, hdr, set()), [])
        bad = {"metrics_id": "sy-x", "set": {"project": "X", "source": "systm", "estimate_status": "final"}}
        self.assertTrue(ml.validate(bad, hdr, set()))

    def test_compute_reconciles_and_keeps_source(self):
        r = m.compute(str(self.p / "sys.json"), str(self.p / "map.csv"), source="system")
        self.assertEqual(r["source"], "system")
        self.assertAlmostEqual(r["bid_price"], self.c["totals"]["bid_price"], delta=1)
        self.assertEqual(r.get("unmapped_items", ""), "")
        self.assertEqual(m.compute(str(self.p / "sys.json"), str(self.p / "map.csv"))["source"], "tradesmen")  # default unchanged

    def test_bid_check_system_writes_sy_rows(self):
        hist = self.p / "hist.csv"
        r = m.compute(str(self.p / "sys.json"), str(self.p / "map.csv"))
        keys = [k for k in r if not k.startswith("_") and k != "walls"]
        hist.write_text(",".join(["metrics_id", "estimate_status", "estimator", "source_detail", "notes",
                                  "matches_bid_log"] + [k for k in keys if k != "metrics_id"]) + "\n")
        out = subprocess.run([sys.executable, str(ROOT / "metrics" / "bid_check.py"), str(self.p / "sys.json"), str(self.p / "map.csv"),
                              str(hist), "--source", "system", "--status", "final", "--project", "Sample Test",
                              "--pending-out", str(self.p)], capture_output=True, text=True)
        self.assertEqual(out.returncode, 0, out.stderr)
        pend = [json.loads(f.read_text()) for f in self.p.glob("*sy-sample-test*.json")]
        self.assertEqual(len(pend), 1)
        self.assertEqual(pend[0]["set"]["source"], "system")
        self.assertEqual(pend[0]["set"]["estimator"], "FB system (/takeoff)")

    def test_system_needs_system_json(self):
        bad = self.p / "x.json"; c = synthetic(); c.pop("source"); bad.write_text(json.dumps(c))
        (self.p / "h.csv").write_text("metrics_id\n")
        out = subprocess.run([sys.executable, str(ROOT / "metrics" / "bid_check.py"), str(bad), str(self.p / "map.csv"),
                              str(self.p / "h.csv"), "--source", "system"], capture_output=True, text=True)
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("not made by system_metrics.py", out.stderr + out.stdout)


if __name__ == "__main__":
    unittest.main()
