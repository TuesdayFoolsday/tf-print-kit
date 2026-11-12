"""write_data splits tables into PARTS files; they must reassemble exactly (synthetic data, no Excel/LibreOffice)."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "dashboard"))
import build_dashboard as bd  # noqa: E402


def join_parts(texts):
    """Python twin of joinParts() in load_dashboard_data.ts."""
    parts = [json.loads(t) for t in texts if t.strip()]
    expected = parts[0].get("parts") or len(parts)
    if len(parts) != expected:
        raise ValueError(f"Got {len(parts)} data parts, expected {expected}")
    parts.sort(key=lambda p: p["part"])
    out = {"stamp": "", "about": [], "tables": {}}
    for i, p in enumerate(parts):
        if p["set"] != parts[0]["set"]:
            raise ValueError("different builds")
        if p["part"] != i + 1:
            raise ValueError("missing or duplicated part")
        out["stamp"] = p.get("stamp") or out["stamp"]
        out["about"] = p.get("about") or out["about"]
        for name, t in p["tables"].items():
            out["tables"].setdefault(name, {"columns": t["columns"], "rows": []})["rows"].extend(t["rows"])
    return out


def make(n_bid_rows, n_gc_rows=20):
    bid_cols = ["id", "gc", "note", "amount"]
    bids = [[i, f"GC {i % 37}", "é-note " + "x" * 60 + str(i), i * 1.5] for i in range(n_bid_rows)]
    gcs = [[f"GC {i}", i] for i in range(n_gc_rows)]
    return {"T": {"tblBidLog": (bid_cols, bids), "tblGCs": (["gc", "n"], gcs), "tblEmpty": (["a"], [])},
            "stamp": "2026-11-12 19:20 CT", "about": ["line one", "line two"]}


class PartsTest(unittest.TestCase):
    def run_case(self, n_rows):
        P = make(n_rows)
        with tempfile.TemporaryDirectory() as d:
            paths, problems = bd.write_data(str(Path(d) / "Dashboard Data.txt"), P)
            self.assertEqual(problems, [])
            self.assertEqual([Path(p).name for p in paths], [f"Dashboard Data {k}.txt" for k in range(1, bd.PARTS + 1)])
            raw = [Path(p).read_bytes() for p in paths]
            for b in raw:
                self.assertLessEqual(len(b), bd.PART_MAX)
                self.assertTrue(b.endswith(b"\n"))
            docs = [json.loads(b) for b in raw]
            self.assertEqual(len({x["set"] for x in docs}), 1)
            self.assertTrue(all(x["parts"] == len(paths) == bd.PARTS for x in docs))
            self.assertEqual([x["part"] for x in docs], list(range(1, bd.PARTS + 1)))
            self.assertTrue(all("stamp" not in x for x in docs[:-1]))
            got = join_parts([b.decode("utf-8") for b in raw])
            self.assertEqual(got["stamp"], P["stamp"])
            self.assertEqual(got["about"], P["about"])
            self.assertEqual(list(got["tables"]), list(P["T"]))
            for name, (cols, rows) in P["T"].items():
                self.assertEqual(got["tables"][name]["columns"], cols)
                self.assertEqual(got["tables"][name]["rows"], json.loads(json.dumps(rows)))
            return docs

    def test_needs_five_parts(self):
        docs = self.run_case(900)
        self.assertTrue(all(x["tables"] for x in docs), "every part used")   # 900 rows ~ 83 KB, needs all 5

    def test_small_data_writes_empty_valid_parts(self):
        docs = self.run_case(50)
        self.assertEqual(len(docs), bd.PARTS)
        for x in docs[1:-1]:
            self.assertEqual(x["tables"], {})

    def test_too_big_is_an_error_not_a_truncation(self):
        paths, problems = bd.write_data("unused.txt", make(2000))
        self.assertEqual(paths, [])
        self.assertIn("add a part (script + Flow 3)", problems[0])

    def test_single_oversize_row_is_an_error(self):
        P = make(3)
        P["T"]["tblBidLog"][1][1][2] = "y" * (bd.PART_MAX + 10)
        paths, problems = bd.write_data("unused.txt", P)
        self.assertEqual(paths, [])
        self.assertTrue(problems)

    def test_loader_refuses_four_of_five(self):
        with tempfile.TemporaryDirectory() as d:
            paths, _ = bd.write_data(str(Path(d) / "x.txt"), make(900))
            texts = [Path(p).read_text(encoding="utf-8") for p in paths]
        with self.assertRaises(ValueError):
            join_parts(texts[:4])


if __name__ == "__main__":
    unittest.main()
