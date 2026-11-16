#!/usr/bin/env python3
"""Courier, upload side: big files into SharePoint with the bytes never passing through git.

A box runs print kit `send.py upload`, which pushes <repo>/upload/<id>.tfc: requests sealed with courier.pub.
Two kinds, both handled here (COURIER_PRIVATE_KEY, Microsoft Graph sign-in as in receive.py):

  session  {kind, id, drive_id, folder_id, reply_pub, files[{index, name, bytes, sha256}]}
           Refused unless folder_allowed(): folder_id is one of UPLOAD_ALLOWED_FOLDERS (comma-separated item ids), or a
           folder named one of UPLOAD_JOB_SUBFOLDERS (comma-separated names) directly inside a job folder (a folder not
           starting with `_`, in the drive root or in one of the UPLOAD_STATUS_FOLDERS, which sit at the root), in the
           drive UPLOAD_DRIVE only. Empty settings allow nothing. When STAGING_ENABLED is true, a pending job folder directly under the configured staging root (or a plans/addenda subfolder inside one) is also allowed. Per file: the name must be SharePoint-safe and not already exist in the folder (nothing is ever
           replaced), then a Graph upload session is created (conflictBehavior fail). The upload URLs are sealed with
           the box's one-time key into <repo>/pickup/up-<id>.tfc; the box PUTs the bytes straight to SharePoint.
  verify   {kind, id, drive_id, folder_id, files[{index, item_id, name, bytes, sha256, quickxor?}]}
           Per file: the item is in that folder, has that name and size, and its content (streamed, never held in
           memory) has that SHA-256; when the request carries quickxor it must match Graph's quickXorHash too.

Writes <repo>/results/upload-<id>.json ({id, stage session|verify, status ok|failed, finished, files[...]}: no names,
no URLs; the verify result replaces the session one) and deletes the request. Pickups older than 24 hours are deleted.
Nothing in SharePoint is ever deleted or overwritten. Exit 1 if anything failed, so GitHub emails the owner.

    python3 courier/upload.py --repo .
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from cryptography.exceptions import InvalidTag

sys.path.insert(0, str(Path(__file__).parent))
from box import open_box, seal  # noqa: E402
from receive import GRAPH, OFFICE, token  # noqa: E402
import staging_lib as sl  # noqa: E402

ALLOWED_FOLDERS = {x.strip() for x in os.environ.get("UPLOAD_ALLOWED_FOLDERS", "").split(",") if x.strip()}
UPLOAD_DRIVE = os.environ.get("UPLOAD_DRIVE", "").strip()
JOB_SUBFOLDERS = {x.strip() for x in os.environ.get("UPLOAD_JOB_SUBFOLDERS", "").split(",") if x.strip()}
STATUS_FOLDERS = {x.strip() for x in os.environ.get("UPLOAD_STATUS_FOLDERS", "").split(",") if x.strip()}
_META: dict = {}   # (drive, item id) -> Graph metadata (or an error string), for this run
STALE = timedelta(hours=24)
CHUNK = 8 * 1024 * 1024


def name_problem(name) -> str:
    """'' when `name` is safe for SharePoint (same rule as send.py), else why not."""
    if not isinstance(name, str) or not name:
        return "no file name"
    if any(c in name for c in '\\/:*?"<>|#%') or name != name.strip():
        return "file name not allowed in SharePoint"
    return ""


def _is_sha(s) -> bool:
    return isinstance(s, str) and len(s) == 64 and all(c in "0123456789abcdef" for c in s)


def stream_sha256(url: str, hdr: dict) -> tuple[str, int]:
    """SHA-256 and byte count of a download, read in 8 MB pieces (a 1 GB plan set is never in memory)."""
    h, n = hashlib.sha256(), 0
    with requests.get(url, headers=hdr, stream=True, timeout=300) as r:
        if r.status_code != 200:
            raise RuntimeError(f"couldn't read the file back ({r.status_code})")
        for piece in r.iter_content(CHUNK):
            h.update(piece)
            n += len(piece)
    return h.hexdigest(), n


def _meta(drive: str, item_id, hdr: dict):
    """Folder metadata from Graph (cached for the run): a dict, or a str saying why it couldn't be read."""
    key = (drive, item_id)
    if key not in _META:
        try:
            r = requests.get(f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(str(item_id), safe='')}"
                             "?$select=id,name,folder,parentReference", headers=hdr, timeout=60)
            _META[key] = r.json() if r.status_code == 200 else f"couldn't look up the folder ({r.status_code})"
        except Exception as e:
            _META[key] = f"couldn't look up the folder ({type(e).__name__})"
    return _META[key]


def _is_root(m: dict) -> bool:
    """The drive root: carries the `root` facet, or has no parent item."""
    return "root" in m or not (m.get("parentReference") or {}).get("id")


def _parent(m: dict, drive: str, hdr: dict):
    return _meta(drive, (m.get("parentReference") or {}).get("id"), hdr)


def folder_allowed(drive: str, folder_id: str, hdr: dict) -> str:
    """'' when uploads may go to this folder, else the reason. Fixed ids, or a subfolder named in UPLOAD_JOB_SUBFOLDERS directly
    inside a job folder (root/<job> or root/<status folder>/<job>) of the one configured drive. When staging is on, a pending job folder under the staging root (or its plans/addenda subfolder) is also allowed."""
    if folder_id in ALLOWED_FOLDERS:
        return ""
    if sl.upload_allowed(drive, folder_id, lambda i: _meta(drive, i, hdr), sl.settings()) == "":
        return ""
    no = "that folder is not on the upload allow-list — refused"
    if not JOB_SUBFOLDERS or not UPLOAD_DRIVE or drive != UPLOAD_DRIVE or not folder_id:
        return no
    m = _meta(drive, folder_id, hdr)
    if isinstance(m, str):
        return m
    if "folder" not in m or m.get("name") not in JOB_SUBFOLDERS or _is_root(m):
        return no
    chain = [m]
    for _ in range(3):   # subfolder -> job folder -> (status folder ->) root
        up = _parent(chain[-1], drive, hdr)
        if isinstance(up, str):
            return up
        chain.append(up)
        if _is_root(up):
            break
    else:
        return no
    names = [c.get("name", "") for c in chain[1:-1]]   # between the subfolder and the root
    if any("folder" not in c for c in chain[1:-1]) or not names or names[0].startswith("_"):
        return no
    if len(names) == 1:
        return ""
    if len(names) == 2 and names[1] in STATUS_FOLDERS:
        return ""
    return no


def make_sessions(req: dict, tok: str) -> tuple[list[dict], list[dict]]:
    """-> (rows for the sealed pickup (ready files only), result rows without names or URLs)"""
    hdr = {"Authorization": f"Bearer {tok}"}
    drive, folder = req["drive_id"], req["folder_id"]
    why = folder_allowed(drive, folder, hdr)
    sealed_rows, rows = [], []
    for f in req["files"]:
        row = {"index": f.get("index"), "bytes": f.get("bytes"), "sha256": f.get("sha256"), "ready": False}
        rows.append(row)
        try:
            if why:
                raise RuntimeError(why)
            if name_problem(f.get("name")):
                raise RuntimeError(name_problem(f.get("name")))
            if not isinstance(f.get("bytes"), int) or f["bytes"] < 1 or not _is_sha(f.get("sha256")):
                raise RuntimeError("bad size or sha256 in the request")
            name = urllib.parse.quote(f["name"], safe="")
            base = f"{GRAPH}/drives/{drive}/items/{folder}:/{name}"
            ex = requests.get(base, headers=hdr, timeout=60)
            if ex.status_code == 200:
                raise RuntimeError("already in the folder (not replaced)")
            if ex.status_code != 404:
                raise RuntimeError(f"couldn't check the folder ({ex.status_code})")
            r = requests.post(base + ":/createUploadSession", headers=hdr, timeout=60,
                              json={"item": {"@microsoft.graph.conflictBehavior": "fail"}})
            if r.status_code == 409:
                raise RuntimeError("already in the folder (not replaced)")
            if r.status_code not in (200, 201):
                raise RuntimeError(f"upload session refused ({r.status_code}): {r.text[:300]}")
            s = r.json()
            sealed_rows.append({"index": f["index"], "name": f["name"], "bytes": f["bytes"], "sha256": f["sha256"],
                                "upload_url": s["uploadUrl"], "expires": s.get("expirationDateTime", "")})
            row["ready"] = True
        except Exception as e:  # one bad file doesn't stop the others
            row["error"] = str(e)
    return sealed_rows, rows


def verify_files(req: dict, tok: str) -> list[dict]:
    hdr = {"Authorization": f"Bearer {tok}"}
    drive, folder = req["drive_id"], req["folder_id"]
    why = folder_allowed(drive, folder, hdr)
    rows = []
    for f in req["files"]:
        row = {"index": f.get("index"), "bytes": f.get("bytes"), "sha256": f.get("sha256"),
               "item_id": f.get("item_id"), "verified": False}
        rows.append(row)
        try:
            if why:
                raise RuntimeError(why)
            iid = urllib.parse.quote(str(f["item_id"]), safe="")
            m = requests.get(f"{GRAPH}/drives/{drive}/items/{iid}", headers=hdr, timeout=60)
            if m.status_code != 200:
                raise RuntimeError(f"item not found in SharePoint ({m.status_code})")
            meta = m.json()
            if meta.get("parentReference", {}).get("id") != folder:
                raise RuntimeError("the item is not in the expected folder")
            if meta.get("name") != f["name"]:
                raise RuntimeError("the item's name is not the expected one")
            if int(meta.get("size", -1)) != f["bytes"]:
                raise RuntimeError(f"size in SharePoint is {meta.get('size')} bytes, expected {f['bytes']}")
            sha, n = stream_sha256(f"{GRAPH}/drives/{drive}/items/{iid}/content", hdr)
            if n != f["bytes"] or sha != f["sha256"]:
                why = ("; an Office file: SharePoint may have rewritten its properties — not accepted here"
                       if f["name"].lower().endswith(OFFICE) else "")
                raise RuntimeError(f"content mismatch: {n} bytes, sha256 {sha}{why}")
            want = f.get("quickxor")
            got = meta.get("file", {}).get("hashes", {}).get("quickXorHash")
            if want and got and got != want:
                raise RuntimeError(f"quickXorHash differs: SharePoint {got}, box {want}")
            row["verified"] = True
        except Exception as e:
            row["error"] = str(e)
    return rows


def _age(stem: str, now: datetime) -> timedelta | None:
    try:
        return now - datetime.strptime(stem[:15], "%Y%m%d-%H%M%S").replace(tzinfo=timezone.utc)
    except ValueError:
        return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=".")
    repo = Path(ap.parse_args().repo)
    requests_dir, pickup, results = repo / "upload", repo / "pickup", repo / "results"
    results.mkdir(exist_ok=True)
    pickup.mkdir(exist_ok=True)
    failed, tok = False, None
    for sealed in sorted(requests_dir.glob("*.tfc")) if requests_dir.exists() else []:
        res = {"id": sealed.stem, "stage": "session", "status": "failed", "files": []}
        try:
            req = json.loads(open_box(sealed.read_bytes(), os.environ["COURIER_PRIVATE_KEY"]))
            res["id"] = req["id"]
            kind = req.get("kind")
            if kind not in ("session", "verify") or not req.get("drive_id") or not req.get("folder_id") \
                    or not isinstance(req.get("files"), list) or not req["files"]:
                raise RuntimeError("not a valid upload request")
            res["stage"] = kind
            if kind == "session" and not req.get("reply_pub"):
                raise RuntimeError("not a valid upload request")
            tok = tok or token()
            if kind == "session":
                sealed_rows, res["files"] = make_sessions(req, tok)
                if sealed_rows:
                    manifest = {"id": req["id"], "drive_id": req["drive_id"], "folder_id": req["folder_id"],
                                "files": sealed_rows}
                    (pickup / f"up-{req['id']}.tfc").write_bytes(seal(json.dumps(manifest).encode(), req["reply_pub"]))
                ok = all(f["ready"] for f in res["files"])
            else:
                res["files"] = verify_files(req, tok)
                ok = all(f["verified"] for f in res["files"])
            if ok:
                res["status"] = "ok"
            else:
                res["error"] = "one or more files failed"
        except InvalidTag:
            res["error"] = "can't open the request (wrong COURIER_PRIVATE_KEY, or the file was altered)"
        except Exception as e:
            res["error"] = f"{type(e).__name__}: {e}" if not isinstance(e, RuntimeError) else str(e)
        res["finished"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        (results / f"upload-{res['id']}.json").write_text(json.dumps(res, indent=1) + "\n")
        sealed.unlink()
        failed |= res["status"] != "ok"
        print(f"{'✓' if res['status'] == 'ok' else '✗'} upload {res['stage']} {res['id']} {res.get('error', '')}")
    now = datetime.now(timezone.utc)
    for old in sorted(pickup.glob("up-*.tfc")):
        age = _age(old.stem[3:], now)
        if age is None or age > STALE:
            old.unlink()
            print(f"- removed unclaimed upload pickup {old.stem}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
