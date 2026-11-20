"""send.py sends a content fingerprint for Office files so the courier's upload robot can check them part by part
after SharePoint stamps its properties in (courier/verify.py, Nov 20 2026)."""
import io
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "courier"))
import send  # noqa: E402

COURIER = Path(os.environ.get("TF_COURIER", "/home/claude/tf-courier"))


def docx_file() -> Path:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("word/document.xml", "<d/>")
    p = Path(tempfile.mkdtemp()) / "Form.docx"
    p.write_bytes(b.getvalue())
    return p


class Fingerprint(unittest.TestCase):
    def test_pdf_gets_none(self):
        p = Path(tempfile.mkdtemp()) / "Plan.pdf"
        p.write_bytes(b"%PDF-1")
        self.assertIsNone(send._office_fingerprint(COURIER, str(p), 6))

    def test_old_courier_clone_gets_none(self):
        p = docx_file()
        self.assertIsNone(send._office_fingerprint(Path(tempfile.mkdtemp()), str(p), p.stat().st_size))

    @unittest.skipUnless((COURIER / "courier" / "verify.py").exists(), "needs a courier clone with verify.py")
    def test_office_file_gets_its_parts(self):
        p = docx_file()
        fp = send._office_fingerprint(COURIER, str(p), p.stat().st_size)
        self.assertEqual(fp["kind"], "zip")
        self.assertIn("word/document.xml", fp["parts"])

    def test_too_big_gets_none(self):
        p = docx_file()
        self.assertIsNone(send._office_fingerprint(COURIER, str(p), 600 * 1024 * 1024))


if __name__ == "__main__":
    unittest.main()
