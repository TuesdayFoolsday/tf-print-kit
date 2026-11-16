"""Pure allow-list tests for courier/staging_lib.py (no network)."""
import os
import unittest
from pathlib import Path
import sys

HERE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(HERE / "courier"))
import staging_lib as sl  # noqa: E402

# A small Graph-like tree: ROOT / _Pending Bids / job / 02+04, plus a real job folder beside it.
TREE = {
    "ROOT": {"id": "ROOT", "name": "Documents", "folder": {}, "root": {}, "parentReference": {}},
    "PEND": {"id": "PEND", "name": "_Pending Bids", "folder": {}, "parentReference": {"id": "ROOT"}},
    "PJOB": {"id": "PJOB", "name": "Red Oak Fire Station - Red Oak", "folder": {}, "parentReference": {"id": "PEND"}},
    "PSUB": {"id": "PSUB", "name": "02 Plans - Full", "folder": {}, "parentReference": {"id": "PJOB"}},
    "PADD": {"id": "PADD", "name": "04 Addenda", "folder": {}, "parentReference": {"id": "PJOB"}},
    "PEST": {"id": "PEST", "name": "07 Estimate", "folder": {}, "parentReference": {"id": "PJOB"}},
    "PFILE": {"id": "PFILE", "name": "Set.pdf", "file": {}, "parentReference": {"id": "PSUB"}},
    "PUNDER": {"id": "PUNDER", "name": "_Nope", "folder": {}, "parentReference": {"id": "PEND"}},
    "JOB": {"id": "JOB", "name": "Alpha Build - Dallas", "folder": {}, "parentReference": {"id": "ROOT"}},
    "SUB": {"id": "SUB", "name": "02 Plans - Full", "folder": {}, "parentReference": {"id": "JOB"}},
    "ADD": {"id": "ADD", "name": "04 Addenda", "folder": {}, "parentReference": {"id": "JOB"}},
    "EST": {"id": "EST", "name": "07 Estimate", "folder": {}, "parentReference": {"id": "JOB"}},
    "SUBM": {"id": "SUBM", "name": "_Submitted", "folder": {}, "parentReference": {"id": "ROOT"}},
    "JOB2": {"id": "JOB2", "name": "Beta Build - Austin", "folder": {}, "parentReference": {"id": "SUBM"}},
    "SUB2": {"id": "SUB2", "name": "04 Addenda", "folder": {}, "parentReference": {"id": "JOB2"}},
}


def meta(i):
    return TREE.get(str(i), f"couldn't look up the folder ({i})")


ON = {
    "STAGING_ENABLED": "true",
    "STAGING_ROOT_NAME": "_Pending Bids",
    "STAGING_ROOT_ID": "PEND",
    "UPLOAD_DRIVE": "b!drive",
    "UPLOAD_JOB_SUBFOLDERS": "02 Plans - Full,04 Addenda",
    "UPLOAD_STATUS_FOLDERS": "_Submitted,_Awarded,_Lost,_No Bid",
    "STAGING_MOVE_DEST_SUBFOLDERS": "02 Plans - Full,04 Addenda,07 Estimate",
    "STAGING_CONFIG": "/tmp/v138/no-such-staging-config.json",
}


class StagingLibTests(unittest.TestCase):
    def setUp(self):
        self._old = {k: os.environ.get(k) for k in ON}
        os.environ.update(ON)

    def tearDown(self):
        for k, v in self._old.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    def s(self):
        return sl.settings()

    def test_switch_off_refuses_everything(self):
        os.environ["STAGING_ENABLED"] = "false"
        s = self.s()
        self.assertFalse(s["enabled"])
        self.assertIn("switched off", sl.upload_allowed("b!drive", "PSUB", meta, s))
        self.assertIn("switched off", sl.mkdir_allowed("b!drive", "PEND", "New Job - City", meta, s))
        self.assertIn("switched off", sl.delete_allowed("b!drive", "PFILE", meta, s))
        self.assertIn("switched off", sl.move_allowed("b!drive", "PFILE", "SUB", meta, s))

    def test_upload_allowed_pending_job_and_plans_only(self):
        s = self.s()
        for fid in ("PJOB", "PSUB", "PADD"):
            with self.subTest(fid):
                self.assertEqual(sl.upload_allowed("b!drive", fid, meta, s), "")
        for fid in ("PEND", "PEST", "PUNDER", "JOB", "SUB", "ROOT", "PFILE", ""):
            with self.subTest(fid):
                self.assertNotEqual(sl.upload_allowed("b!drive", fid, meta, s), "")

    def test_other_drive_refused(self):
        self.assertNotEqual(sl.upload_allowed("b!other", "PSUB", meta, self.s()), "")

    def test_mkdir_rules(self):
        s = self.s()
        self.assertEqual(sl.mkdir_allowed("b!drive", "ROOT", "_Pending Bids", meta, s), "")
        self.assertEqual(sl.mkdir_allowed("b!drive", "PEND", "New Job - City", meta, s), "")
        self.assertEqual(sl.mkdir_allowed("b!drive", "PJOB", "02 Plans - Full", meta, s), "")
        self.assertIn("must not start with _", sl.mkdir_allowed("b!drive", "PEND", "_Nope", meta, s))
        self.assertNotEqual(sl.mkdir_allowed("b!drive", "PJOB", "07 Estimate", meta, s), "")
        self.assertNotEqual(sl.mkdir_allowed("b!drive", "JOB", "02 Plans - Full", meta, s), "")
        self.assertNotEqual(sl.mkdir_allowed("b!drive", "ROOT", "Random", meta, s), "")
        self.assertIn("not allowed", sl.mkdir_allowed("b!drive", "PEND", "bad:name", meta, s))

    def test_delete_carve_out(self):
        s = self.s()
        for iid in ("PFILE", "PSUB", "PJOB", "PUNDER"):
            with self.subTest(iid):
                self.assertEqual(sl.delete_allowed("b!drive", iid, meta, s), "")
        self.assertIn("never deleted", sl.delete_allowed("b!drive", "PEND", meta, s))
        for iid in ("SUB", "JOB", "ROOT", "PFILE-NO"):
            with self.subTest(iid):
                self.assertIn("not inside the staging root", sl.delete_allowed("b!drive", iid, meta, s))

    def test_move_from_pending_into_job_subfolder(self):
        s = self.s()
        self.assertEqual(sl.move_allowed("b!drive", "PFILE", "SUB", meta, s), "")
        self.assertEqual(sl.move_allowed("b!drive", "PFILE", "EST", meta, s), "")
        self.assertEqual(sl.move_allowed("b!drive", "PFILE", "SUB2", meta, s), "")
        self.assertNotEqual(sl.move_allowed("b!drive", "PFILE", "PEND", meta, s), "")
        self.assertNotEqual(sl.move_allowed("b!drive", "PFILE", "JOB", meta, s), "")
        self.assertIn("never moved", sl.move_allowed("b!drive", "PEND", "SUB", meta, s))
        self.assertIn("not inside the staging root", sl.move_allowed("b!drive", "SUB", "ADD", meta, s))


if __name__ == "__main__":
    unittest.main()
