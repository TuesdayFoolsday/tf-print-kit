"""Tests for courier/receive.py: Office read-back after SharePoint stamps its properties, and runs that
start from an older snapshot (synthetic files only)."""
import io
import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "courier"))
import receive  # noqa: E402

RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="officeDocument" Target="xl/workbook.xml"/>{extra}</Relationships>')
CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="xml" ContentType="application/xml"/>{extra}</Types>')


def book(sheet="<sheet>1</sheet>", rels_extra="", ct_extra="", core="<core>a</core>", added=(), drop=()):
    parts = {"[Content_Types].xml": CT.format(extra=ct_extra), "_rels/.rels": RELS.format(extra=rels_extra),
             "docProps/core.xml": core, "xl/workbook.xml": "<workbook/>", "xl/worksheets/sheet1.xml": sheet}
    parts.update(dict(added))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for n, v in parts.items():
            if n not in drop:
                z.writestr(n, v)
    return buf.getvalue()


SENT = book()
STAMPED = dict(rels_extra='<Relationship Id="rId9" Type="customXml" Target="customXml/item1.xml"/>',
               ct_extra='<Override PartName="/customXml/itemProps1.xml" ContentType="x"/>',
               core="<core>b</core>",
               added=[("customXml/item1.xml", "<x/>"), ("docProps/custom.xml", "<c/>"), ("[trash]/0000.dat", "z")])


def test_sharepoint_stamps_accepted():
    assert receive.office_same(SENT, book(**STAMPED)) == ""


def test_identical_accepted():
    assert receive.office_same(SENT, SENT) == ""


def test_changed_sheet_refused():
    assert "content changed" in receive.office_same(SENT, book(sheet="<sheet>2</sheet>", **{k: v for k, v in STAMPED.items()}))


def test_missing_part_refused():
    assert "missing" in receive.office_same(SENT, book(drop=("xl/worksheets/sheet1.xml",)))


def test_unexpected_part_refused():
    assert "unexpected" in receive.office_same(SENT, book(added=[("xl/worksheets/sheet2.xml", "<s/>")]))


def test_dropped_relationship_refused():
    sent = book(rels_extra='<Relationship Id="rId2" Type="styles" Target="xl/styles.xml"/>')
    assert "entries changed" in receive.office_same(sent, SENT)


def test_not_a_zip_refused():
    assert receive.office_same(b"plain", b"plain2") == "not an Office file"


class FakeResp:
    def __init__(self, status, content=b"", js=None):
        self.status_code, self.content, self._js, self.text = status, content, js, ""

    def json(self):
        return self._js


def run_deliver(monkeypatch, name, sent, stored):
    import base64
    import hashlib
    monkeypatch.setattr(receive.requests, "put", lambda *a, **k: FakeResp(201, js={"id": "ITEM"}))
    monkeypatch.setattr(receive.requests, "get", lambda *a, **k: FakeResp(200, stored))
    f = {"index": 0, "bytes": len(sent), "sha256": hashlib.sha256(sent).hexdigest(), "name": name,
         "folder_id": "F", "b64": base64.b64encode(sent).decode()}
    return receive.deliver({"drive_id": "D", "files": [f]}, "tok")[0]


def test_deliver_xlsx_stamped_is_verified(monkeypatch):
    row = run_deliver(monkeypatch, "Dashboard.xlsx", SENT, book(**STAMPED))
    assert row["verified"] and "note" in row and row["stored_sha256"] != row["sha256"]


def test_deliver_xlsx_changed_fails(monkeypatch):
    row = run_deliver(monkeypatch, "Dashboard.xlsx", SENT, book(sheet="<sheet>9</sheet>"))
    assert not row["verified"] and "content changed" in row["error"]


def test_deliver_text_mismatch_still_fails(monkeypatch):
    row = run_deliver(monkeypatch, "notes.md", b"abc", b"abd")
    assert not row["verified"] and "not an Office file" in row["error"]


def test_already_delivered_inbox_file_is_skipped(tmp_path, monkeypatch):
    (tmp_path / "inbox").mkdir()
    (tmp_path / "results").mkdir()
    (tmp_path / "inbox" / "X1.tfc").write_bytes(b"sealed")
    ok = {"id": "X1", "status": "ok", "files": []}
    (tmp_path / "results" / "X1.json").write_text(json.dumps(ok))
    monkeypatch.setattr(sys, "argv", ["receive.py", "--repo", str(tmp_path)])
    monkeypatch.setattr(receive, "token", lambda: (_ for _ in ()).throw(AssertionError("must not sign in")))
    try:
        receive.main()
    except SystemExit as e:
        assert e.code == 0
    assert not (tmp_path / "inbox" / "X1.tfc").exists()
    assert json.loads((tmp_path / "results" / "X1.json").read_text()) == ok
