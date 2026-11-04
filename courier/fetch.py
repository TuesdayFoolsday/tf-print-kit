#!/usr/bin/env python3
"""Courier, fetching side: an exact copy of one SharePoint file, sealed for the chat that asked for it.

The mirror of receive.py. A chat runs print kit `send.py fetch`, which pushes <repo>/fetch/<id>.tfc: a request
sealed with courier.pub naming the file (drive + item id, or drive + path inside the library) and carrying a
ONE-TIME public key the chat just made. For each request this script opens it (COURIER_PRIVATE_KEY), signs in to
Microsoft Graph (same app as receive.py), downloads the file twice and checks both copies match each other and the
size SharePoint reports, then seals the file (name, ids, size, SHA-256, bytes) with the one-time key into
<repo>/pickup/<id>.tfc. Only the chat holds the matching private key, and it throws it away after pickup, so the
sealed copy left in git history can never be opened again. Writes <repo>/results/fetch-<id>.json ({id, status ok|
failed, finished, bytes, sha256, item_id, error} — no file names) and deletes the request. Pickups older than
24 hours (a chat that never collected) are deleted. Exit 1 if anything failed, so GitHub emails the owner.

    python3 courier/fetch.py --repo .
"""
from __future__ import annotations

import argparse
import base64
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
from receive import GRAPH, token  # noqa: E402

MAX_BYTES = 25 * 1024 * 1024
STALE = timedelta(hours=24)


def _get(url: str, hdr: dict) -> requests.Response:
    r = requests.get(url, headers=hdr, timeout=120)
    if r.status_code == 404:
        raise RuntimeError("file not found in SharePoint")
    if r.status_code != 200:
        raise RuntimeError(f"SharePoint read failed ({r.status_code}): {r.text[:300]}")
    return r


def fetch_one(req: dict, tok: str) -> tuple[dict, dict]:
    """-> (sealed manifest for the chat, result row without names)"""
    hdr = {"Authorization": f"Bearer {tok}"}
    drive = req["drive_id"]
    if req.get("item_id"):
        meta_url = f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(req['item_id'], safe='')}"
    elif req.get("path"):
        meta_url = f"{GRAPH}/drives/{drive}/root:/{urllib.parse.quote(req['path'].strip('/'), safe='/')}"
    else:
        raise RuntimeError("request names no file (item_id or path)")
    meta = _get(meta_url, hdr).json()
    if "file" not in meta:
        raise RuntimeError("that item is a folder, not a file")
    size = int(meta.get("size", -1))
    if not 0 < size <= MAX_BYTES:
        raise RuntimeError(f"file is {size} bytes (must be 1 … {MAX_BYTES})")
    content_url = f"{GRAPH}/drives/{drive}/items/{meta['id']}/content"
    first = _get(content_url, hdr).content
    second = _get(content_url, hdr).content
    if len(first) != size:
        raise RuntimeError(f"download was {len(first)} bytes but SharePoint says {size}")
    sha = hashlib.sha256(first).hexdigest()
    if hashlib.sha256(second).hexdigest() != sha:
        raise RuntimeError("two downloads of the file differ (it changed mid-fetch?) — fetch again")
    manifest = {"id": req["id"], "drive_id": drive, "item_id": meta["id"], "name": meta["name"],
                "parent_id": meta.get("parentReference", {}).get("id", ""), "bytes": size, "sha256": sha,
                "b64": base64.b64encode(first).decode()}
    return manifest, {"bytes": size, "sha256": sha, "item_id": meta["id"]}


def _age(stem: str, now: datetime) -> timedelta | None:
    try:
        return now - datetime.strptime(stem[:15], "%Y%m%d-%H%M%S").replace(tzinfo=timezone.utc)
    except ValueError:
        return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=".")
    repo = Path(ap.parse_args().repo)
    requests_dir, pickup, results = repo / "fetch", repo / "pickup", repo / "results"
    results.mkdir(exist_ok=True)
    pickup.mkdir(exist_ok=True)
    failed, tok = False, None
    for sealed in sorted(requests_dir.glob("*.tfc")) if requests_dir.exists() else []:
        res = {"id": sealed.stem, "status": "failed"}
        try:
            req = json.loads(open_box(sealed.read_bytes(), os.environ["COURIER_PRIVATE_KEY"]))
            res["id"] = req["id"]
            if req.get("kind") != "fetch" or not req.get("reply_pub") or not req.get("drive_id"):
                raise RuntimeError("not a valid fetch request")
            tok = tok or token()
            manifest, row = fetch_one(req, tok)
            (pickup / f"{req['id']}.tfc").write_bytes(seal(json.dumps(manifest).encode(), req["reply_pub"]))
            res.update(row, status="ok")
        except InvalidTag:
            res["error"] = "can't open the request (wrong COURIER_PRIVATE_KEY, or the file was altered)"
        except Exception as e:
            res["error"] = f"{type(e).__name__}: {e}" if not isinstance(e, RuntimeError) else str(e)
        res["finished"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        (results / f"fetch-{res['id']}.json").write_text(json.dumps(res, indent=1) + "\n")
        sealed.unlink()
        failed |= res["status"] != "ok"
        print(f"{'✓' if res['status'] == 'ok' else '✗'} fetch {res['id']} {res.get('error', '')}")
    now = datetime.now(timezone.utc)
    for old in sorted(pickup.glob("*.tfc")):
        age = _age(old.stem, now)
        if age is None or age > STALE:
            old.unlink()
            print(f"- removed uncollected pickup {old.stem}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
