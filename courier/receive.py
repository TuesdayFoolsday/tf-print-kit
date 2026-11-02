#!/usr/bin/env python3
"""Courier, receiving side — runs in the private courier repo's GitHub Action (see courier/README.md).

For every sealed file in <repo>/inbox: open it (COURIER_PRIVATE_KEY), check each file's SHA-256 against the
sealed manifest, upload it with Microsoft Graph (app-only token: TENANT_ID, CLIENT_ID, CLIENT_SECRET; the app
holds Sites.Selected write on one site), download it back and compare the fingerprint again. Then write
<repo>/results/<id>.json ({id, status ok|failed, finished, files[index, bytes, sha256, item_id, verified, error]},
no file names) and delete the inbox file. Exit 1 if anything failed, so GitHub emails the owner.

    python3 courier/receive.py --repo .
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

import requests
from cryptography.exceptions import InvalidTag

sys.path.insert(0, str(Path(__file__).parent))
from box import open_box  # noqa: E402

GRAPH = os.environ.get("GRAPH_BASE", "https://graph.microsoft.com/v1.0")


def token() -> str:
    url = os.environ.get("TOKEN_URL") or f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token"
    r = requests.post(url, data={"grant_type": "client_credentials", "client_id": os.environ["CLIENT_ID"],
                                 "client_secret": os.environ["CLIENT_SECRET"],
                                 "scope": "https://graph.microsoft.com/.default"}, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"sign-in failed ({r.status_code}): {r.text[:300]}")
    return r.json()["access_token"]


def deliver(payload: dict, tok: str) -> list[dict]:
    hdr = {"Authorization": f"Bearer {tok}"}
    drive, conflict, out = payload["drive_id"], payload.get("conflict", "fail"), []
    for f in payload["files"]:
        row = {"index": f["index"], "bytes": f["bytes"], "sha256": f["sha256"], "verified": False}
        out.append(row)
        try:
            data = base64.b64decode(f["b64"])
            if len(data) != f["bytes"] or hashlib.sha256(data).hexdigest() != f["sha256"]:
                raise RuntimeError("sealed copy doesn't match its own fingerprint")
            name = urllib.parse.quote(f["name"], safe="")
            url = (f"{GRAPH}/drives/{drive}/items/{f['folder_id']}:/{name}:/content"
                   f"?@microsoft.graph.conflictBehavior={conflict}")
            r = requests.put(url, data=data, headers={**hdr, "Content-Type": "application/octet-stream"}, timeout=120)
            if r.status_code == 409:
                raise RuntimeError("a file with this name already exists (not replaced)")
            if r.status_code not in (200, 201):
                raise RuntimeError(f"upload failed ({r.status_code}): {r.text[:300]}")
            item = r.json()
            row["item_id"] = item["id"]
            back = requests.get(f"{GRAPH}/drives/{drive}/items/{item['id']}/content", headers=hdr, timeout=120)
            if back.status_code != 200:
                raise RuntimeError(f"read-back failed ({back.status_code})")
            got = hashlib.sha256(back.content).hexdigest()
            if len(back.content) != f["bytes"] or got != f["sha256"]:
                raise RuntimeError(f"read-back mismatch: {len(back.content)} bytes, sha256 {got}")
            row["verified"] = True
        except Exception as e:  # one bad file doesn't stop the others
            row["error"] = str(e)
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=".")
    repo = Path(ap.parse_args().repo)
    inbox, results = repo / "inbox", repo / "results"
    results.mkdir(exist_ok=True)
    failed, tok = False, None
    for sealed in sorted(inbox.glob("*.tfc")) if inbox.exists() else []:
        res = {"id": sealed.stem, "status": "failed", "files": []}
        try:
            payload = json.loads(open_box(sealed.read_bytes(), os.environ["COURIER_PRIVATE_KEY"]))
            res["id"] = payload["id"]
            tok = tok or token()
            res["files"] = deliver(payload, tok)
            if res["files"] and all(f["verified"] for f in res["files"]):
                res["status"] = "ok"
            else:
                res["error"] = "one or more files not delivered"
        except InvalidTag:
            res["error"] = "can't open the sealed file (wrong COURIER_PRIVATE_KEY, or the file was altered)"
        except Exception as e:
            res["error"] = f"{type(e).__name__}: {e}"
        res["finished"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        (results / f"{res['id']}.json").write_text(json.dumps(res, indent=1) + "\n")
        sealed.unlink()
        failed |= res["status"] != "ok"
        print(f"{'✓' if res['status'] == 'ok' else '✗'} {res['id']} {res.get('error', '')}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
