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
    python3 courier/send.py directory --repo /tmp/courier <contacts.csv>   # System Rules B3: new Directory contacts
    python3 courier/send.py fetch --repo /tmp/courier --drive <driveId> (--item <itemId> | --path "<path in library>") \
        --out <local file> [--timeout 300]                          # exact copy of a SharePoint file
    python3 courier/send.py fetch --repo /tmp/courier --collect <id> --out <local file>   # after a timeout

`send` prints the id and one line per file (index, bytes, sha256, name). `wait` exits 0 when every file was
uploaded and verified, 1 on a failed delivery (reason printed), 2 if no result arrived within the timeout
(the next run checks again with `wait --timeout 0`). The recipient public key is <repo>/courier.pub.
Without --replace an existing file with the same name is never overwritten (the delivery fails instead).
`--if-unchanged <sha256>` (with --replace, one file): the courier replaces the file only if the copy now in
SharePoint still has that fingerprint — use the one `fetch` printed, so an edit never overwrites a change made in
between (the delivery fails instead: fetch again and redo the edit).

`fetch` is the mirror of `send`: it makes a one-time key, pushes a sealed request to <repo>/fetch/, and the courier's
`fetch` Action downloads the file (twice, compared), seals it with that key into <repo>/pickup/ and writes
results/fetch-<id>.json. `fetch` then opens it, checks size and SHA-256, writes the exact bytes to --out (never
over an existing file), deletes the pickup from the repo and throws the key away. It prints the file's name, item
id, folder id, size, sha256 and the matching `send --replace --if-unchanged` command for sending the edit back.
Exits 0 ok, 1 failed (reason printed), 2 no result within the timeout (finish later with --collect <id>).
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import secrets
import shlex
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from box import new_keypair, open_box, seal  # noqa: E402

MAX_BYTES = 25 * 1024 * 1024
FETCH_KEYS = Path(os.environ.get("TF_COURIER_KEYS", "/tmp/tf-courier-fetch-keys"))


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
    if a.if_unchanged:
        if not a.replace or len(a.files) != 1:
            sys.exit("✗ --if-unchanged needs --replace and exactly one file")
        if len(a.if_unchanged) != 64 or any(c not in "0123456789abcdef" for c in a.if_unchanged):
            sys.exit("✗ --if-unchanged takes the 64-character sha256 that fetch printed")
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
        if a.if_unchanged:
            files[-1]["if_unchanged"] = a.if_unchanged
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


def _new_id() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-") + secrets.token_hex(3)


def cmd_fetch(a: argparse.Namespace) -> None:
    """Exact copy of one SharePoint file, through the courier's `fetch` Action (see the module notes)."""
    repo, out = Path(a.repo), Path(a.out)
    if out.exists():
        sys.exit(f"✗ {out} already exists — pick a new --out path (fetch never overwrites a local file)")
    FETCH_KEYS.mkdir(mode=0o700, parents=True, exist_ok=True)
    if a.collect:
        cid = a.collect
        keyfile = FETCH_KEYS / f"{cid}.key"
        if not keyfile.exists():
            sys.exit(f"✗ no one-time key for {cid} on this machine — run a new fetch")
    else:
        if not a.drive or not (a.item or a.path) or (a.item and a.path):
            sys.exit("✗ fetch needs --drive and one of --item or --path (or --collect <id>)")
        priv, pub_once = new_keypair()
        cid = _new_id()
        keyfile = FETCH_KEYS / f"{cid}.key"
        keyfile.write_text(priv)
        keyfile.chmod(0o600)
        req = {"id": cid, "kind": "fetch", "drive_id": a.drive, "item_id": a.item or "", "path": a.path or "",
               "reply_pub": pub_once}
        git(repo, "pull", "-q", "--rebase", "origin", "main", check=False)
        (repo / "fetch").mkdir(exist_ok=True)
        (repo / "fetch" / f"{cid}.tfc").write_bytes(seal(json.dumps(req).encode(),
                                                          (repo / "courier.pub").read_text().strip()))
        git(repo, "add", f"fetch/{cid}.tfc")
        push(repo, f"fetch {cid}")
        print(f"requested {cid}")
    deadline = time.time() + a.timeout
    while True:
        git(repo, "fetch", "-q", "origin", "main", check=False)
        r = git(repo, "show", f"origin/main:results/fetch-{cid}.json", check=False)
        if r.returncode == 0:
            break
        if time.time() >= deadline:
            print(f"… no result yet for {cid} — finish later with: fetch --repo {a.repo} --collect {cid} --out {a.out}")
            sys.exit(2)
        time.sleep(10)
    res = json.loads(r.stdout)
    if res.get("status") != "ok":
        keyfile.unlink(missing_ok=True)
        print(f"✗ fetch {cid} failed: {res.get('error', 'no reason given')}")
        sys.exit(1)
    blob = subprocess.run(["git", "-C", str(repo), "show", f"origin/main:pickup/{cid}.tfc"],
                          capture_output=True, check=False)
    if blob.returncode != 0:
        sys.exit(f"✗ fetch {cid}: result says ok but the sealed copy is gone (collected already, or older than 24 h)")
    m = json.loads(open_box(blob.stdout, keyfile.read_text()))
    data = base64.b64decode(m["b64"])
    sha = hashlib.sha256(data).hexdigest()
    if m["id"] != cid or len(data) != m["bytes"] or sha != m["sha256"] or sha != res.get("sha256") \
            or len(data) != res.get("bytes"):
        sys.exit(f"✗ fetch {cid}: the copy doesn't match its fingerprint — not saved")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    if hashlib.sha256(out.read_bytes()).hexdigest() != sha:
        sys.exit(f"✗ fetch {cid}: the saved file doesn't match — don't use {out}")
    # pickup done: remove the sealed copy from the repo and throw the one-time key away
    git(repo, "pull", "-q", "--rebase", "origin", "main", check=False)
    if (repo / "pickup" / f"{cid}.tfc").exists():
        git(repo, "rm", "-q", f"pickup/{cid}.tfc")
        push(repo, f"picked up {cid}")
    keyfile.unlink(missing_ok=True)
    print(f"✓ fetched {cid} — exact copy saved to {out}")
    print(f"  name       {m['name']}")
    print(f"  bytes      {m['bytes']}")
    print(f"  sha256     {sha}")
    print(f"  item       {m['item_id']}")
    print(f"  folder     {m['parent_id']}")
    print(f"  drive      {m['drive_id']}")
    print("  send the edit back with:")
    print(f"    python3 {shlex.quote(sys.argv[0])} send --repo {shlex.quote(a.repo)} --drive {shlex.quote(m['drive_id'])} "
          f"--folder {m['parent_id']} --replace --if-unchanged {sha} {shlex.quote(str(out) + '=' + m['name'])}")


def cmd_directory(a: argparse.Namespace) -> None:
    """Hand new contacts to the courier repo's `directory` Action (System Rules B3) by pushing trigger/directory_add.csv."""
    import csv
    src = Path(a.csv)
    rows = list(csv.reader(src.read_text(encoding="utf-8-sig").splitlines()))
    if len(rows) < 2 or "firm_id" not in rows[0] or "contact_name" not in rows[0]:
        sys.exit("✗ not written: the CSV needs Directory.csv's header row (firm_id, contact_name, …) and at least one contact")
    repo = Path(a.repo)
    git(repo, "pull", "-q", "--rebase", "origin", "main", check=False)
    dest = repo / "trigger" / "directory_add.csv"
    dest.parent.mkdir(exist_ok=True)
    dest.write_text(src.read_text(encoding="utf-8-sig"), encoding="utf-8")
    git(repo, "add", "trigger/directory_add.csv")
    if git(repo, "diff", "--cached", "--quiet", check=False).returncode == 0:
        print("= same contacts as the last hand-off; nothing pushed")
        return
    push(repo, f"directory: {len(rows) - 1} contact(s)")
    print(f"✓ {len(rows) - 1} contact(s) handed to the directory Action; result in results/directory-latest.json")


def _push_trigger(a: argparse.Namespace, name: str, required: tuple[str, ...], label: str) -> None:
    """Validate a small JSON request and hand it to the courier repo's `operations` Action."""
    src = Path(a.json)
    try:
        req = json.loads(src.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        sys.exit(f"✗ not written: {src} isn't readable JSON ({e})")
    missing = [k for k in required if not req.get(k)]
    if missing:
        sys.exit(f"✗ not written: missing {', '.join(missing)}")
    repo = Path(a.repo)
    git(repo, "pull", "-q", "--rebase", "origin", "main", check=False)
    dest = repo / "trigger" / name
    dest.parent.mkdir(exist_ok=True)
    dest.write_text(json.dumps(req, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    git(repo, "add", f"trigger/{name}")
    if git(repo, "diff", "--cached", "--quiet", check=False).returncode == 0:
        print(f"= same {label} as the last hand-off; nothing pushed")
        return
    push(repo, f"operations: {label} {req.get('bid_id', '')}")
    result = name.replace("job_update", "update").replace(".json", "")
    print(f"✓ {label} handed to the operations Action; result in results/operations-{result}-latest.json")


def cmd_award(a: argparse.Namespace) -> None:
    """Hand an award to the Operations robot (Award Handoff plan): pushes trigger/award.json."""
    _push_trigger(a, "award.json", ("bid_id", "project"), "award")


def cmd_job_update(a: argparse.Namespace) -> None:
    """Hand David's answers for a Jobs row to the Operations robot: pushes trigger/job_update.json."""
    _push_trigger(a, "job_update.json", ("bid_id", "set"), "job update")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("keygen").set_defaults(fn=cmd_keygen)
    s = sub.add_parser("send")
    s.add_argument("--repo", required=True)
    s.add_argument("--drive", required=True)
    s.add_argument("--folder", required=True)
    s.add_argument("--replace", action="store_true")
    s.add_argument("--if-unchanged", metavar="SHA256", default="")
    s.add_argument("files", nargs="+")
    s.set_defaults(fn=cmd_send)
    f = sub.add_parser("fetch")
    f.add_argument("--repo", required=True)
    f.add_argument("--drive", default="")
    f.add_argument("--item", default="")
    f.add_argument("--path", default="")
    f.add_argument("--collect", default="", metavar="ID")
    f.add_argument("--out", required=True)
    f.add_argument("--timeout", type=int, default=300)
    f.set_defaults(fn=cmd_fetch)
    w = sub.add_parser("wait")
    w.add_argument("--repo", required=True)
    w.add_argument("id")
    w.add_argument("--timeout", type=int, default=300)
    w.set_defaults(fn=cmd_wait)
    d = sub.add_parser("directory")
    d.add_argument("--repo", required=True)
    d.add_argument("csv")
    d.set_defaults(fn=cmd_directory)
    for nm, fn in (("award", cmd_award), ("job-update", cmd_job_update)):
        x = sub.add_parser(nm)
        x.add_argument("--repo", required=True)
        x.add_argument("json")
        x.set_defaults(fn=fn)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
