#!/usr/bin/env python3
"""Courier, sending side: get binary files (proposal PDFs) into SharePoint without typing their bytes.

The run seals the files (box.py) into one file in the private courier repo's inbox/ and pushes it. A GitHub
Action in that repo (courier/receive.py) opens it, uploads each file through Microsoft Graph, downloads it back,
compares the SHA-256 fingerprint and writes results/<id>.json. Nothing readable sits in the repo: inbox files are
sealed and removed after delivery; results carry only sizes, fingerprints and item ids (no names).

    python3 courier/send.py keygen                                  # once, in chat: private -> GitHub secret
    python3 courier/send.py send --repo /tmp/courier --drive <driveId> --folder <folderItemId> \
        [--replace] <local file>[=<SharePoint file name>] ...
    python3 courier/send.py wait --repo /tmp/courier <id> [--timeout 300]

`send` prints the id and one line per file (index, bytes, sha256, name). `wait` exits 0 when every file was
uploaded and verified, 1 on a failed delivery (reason printed), 2 if no result arrived within the timeout
(the next run checks again with `wait --timeout 0`). The recipient public key is <repo>/courier.pub.
Without --replace an existing file with the same name is never overwritten (the delivery fails instead).
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import secrets
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from box import new_keypair, seal  # noqa: E402

MAX_BYTES = 25 * 1024 * 1024


def git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(repo), *args], check=check, capture_output=True, text=True)


def push(repo: Path, message: str) -> None:
    git(repo, "commit", "-q", "-m", message)
    for attempt in range(4):
        r = git(repo, "push", "-q", "origin", "HEAD:main", check=False)
        if r.returncode == 0:
            return
        time.sleep(3 + attempt * 3)
        git(repo, "pull", "-q", "--rebase", "origin", "main", check=False)
    sys.exit(f"✗ push failed: {r.stderr.strip()}")


def cmd_keygen(_: argparse.Namespace) -> None:
    priv, pub = new_keypair()
    print(f"private (GitHub secret COURIER_PRIVATE_KEY): {priv}")
    print(f"public  (file courier.pub in the courier repo): {pub}")


def cmd_send(a: argparse.Namespace) -> None:
    repo = Path(a.repo)
    pub = (repo / "courier.pub").read_text().strip()
    files, lines = [], []
    for i, spec in enumerate(a.files):
        local, _, name = spec.partition("=")
        p = Path(local)
        data = p.read_bytes()
        name = name or p.name
        if not data or len(data) > MAX_BYTES:
            sys.exit(f"✗ {p}: {len(data)} bytes (must be 1 … {MAX_BYTES})")
        if any(c in name for c in '\\/:*?"<>|#%') or name != name.strip():
            sys.exit(f"✗ SharePoint name not allowed: {name!r}")
        sha = hashlib.sha256(data).hexdigest()
        files.append({"index": i, "folder_id": a.folder, "name": name, "bytes": len(data), "sha256": sha,
                      "b64": base64.b64encode(data).decode()})
        lines.append(f"  {i}  {len(data)} bytes  sha256 {sha}  {name}")
    cid = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-") + secrets.token_hex(3)
    payload = {"id": cid, "drive_id": a.drive, "conflict": "replace" if a.replace else "fail", "files": files}
    git(repo, "pull", "-q", "--rebase", "origin", "main", check=False)
    (repo / "inbox").mkdir(exist_ok=True)
    (repo / "inbox" / f"{cid}.tfc").write_bytes(seal(json.dumps(payload).encode(), pub))
    git(repo, "add", f"inbox/{cid}.tfc")
    push(repo, f"deliver {cid}")
    print(f"sent {cid} ({len(files)} file{'s' if len(files) != 1 else ''})")
    print("\n".join(lines))


def cmd_wait(a: argparse.Namespace) -> None:
    repo, rel = Path(a.repo), f"results/{a.id}.json"
    deadline = time.time() + a.timeout
    while True:
        git(repo, "fetch", "-q", "origin", "main", check=False)
        r = git(repo, "show", f"origin/main:{rel}", check=False)
        if r.returncode == 0:
            res = json.loads(r.stdout)
            for f in res.get("files", []):
                mark = "✓" if f.get("verified") else "✗"
                print(f"  {mark} {f['index']}  {f.get('bytes')} bytes  sha256 {f.get('sha256')}  item {f.get('item_id', '')}"
                      + (f"  {f['error']}" if f.get("error") else ""))
            if res.get("status") == "ok":
                print(f"✓ delivered {a.id} — every file uploaded and verified")
                sys.exit(0)
            print(f"✗ delivery {a.id} failed: {res.get('error', 'see file lines')}")
            sys.exit(1)
        if time.time() >= deadline:
            print(f"… no result yet for {a.id}")
            sys.exit(2)
        time.sleep(15)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("keygen").set_defaults(fn=cmd_keygen)
    s = sub.add_parser("send")
    s.add_argument("--repo", required=True)
    s.add_argument("--drive", required=True)
    s.add_argument("--folder", required=True)
    s.add_argument("--replace", action="store_true")
    s.add_argument("files", nargs="+")
    s.set_defaults(fn=cmd_send)
    w = sub.add_parser("wait")
    w.add_argument("--repo", required=True)
    w.add_argument("id")
    w.add_argument("--timeout", type=int, default=300)
    w.set_defaults(fn=cmd_wait)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
