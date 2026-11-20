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
    python3 courier/send.py award --repo /tmp/courier <award.json>            # System Rules A16: award hand-off
    python3 courier/send.py job-update --repo /tmp/courier <job_update.json>  # B18: David's answers (bid_id or vista)
    python3 courier/send.py job-load --repo /tmp/courier <job_load.json>      # once: jobs open before the award robot
    python3 courier/send.py tick --repo /tmp/courier [--nightly]              # B18: start the Operations robot now
    python3 courier/send.py fetch --repo /tmp/courier --drive <driveId> (--item <itemId> | --path "<path in library>") \
        --out <local file> [--timeout 300]                          # exact copy of a SharePoint file
    python3 courier/send.py fetch --repo /tmp/courier --collect <id> --out <local file>   # after a timeout
    python3 courier/send.py list --repo /tmp/courier --drive <driveId> (--item <folderId> | --path "<path>") \
        [--depth 2] [--branch main] --out <listing.json>            # what is in a SharePoint folder
    python3 courier/send.py upload --repo /tmp/courier --drive <driveId> --folder <folderId> [--branch main] \
        [--timeout 3600] [--chunk-mb 10] [--summary <file.json>] <local file>[=<SharePoint name>] ...   # big files
    python3 courier/send.py upload-wait --repo /tmp/courier [--branch main] [--timeout 3600] <id>   # after exit 2
    python3 courier/send.py upload-resume --repo /tmp/courier [--branch main] [--timeout 300] <id>  # after a crash

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

`list` is `fetch` for a folder: the courier's `fetch` Action lists the folder (depth 0 = its children, each level adds
the children of the sub-folders; 0 … 4, at most 5000 entries) and seals {id, drive_id, root_id, entries[path, name, id,
size, is_folder, quickxor, sha256, modified]} for the one-time key. `list` writes that JSON to --out (never over an
existing file), deletes the pickup from the repo and throws the key away.

`upload` sends files of any size (30 MB - 1 GB plan sets) without their bytes passing through git. It hashes each file
(SHA-256 and quickxor.py's QuickXorHash, streamed), pushes a sealed `session` request to <repo>/upload/, collects the
sealed upload URLs from <repo>/pickup/up-<id>.tfc (then deletes the pickup from git), PUTs the bytes straight to
SharePoint in chunks (resuming after errors through the session's nextExpectedRanges; the URLs are pre-authenticated, no
token is sent, and they are never printed), then pushes a sealed `verify` request and prints one line per file (index,
bytes, sha256, item id, ✓/✗). Exits 0 all verified, 1 any failed, 2 no result in time (finish with `upload-wait`).
Nothing existing is overwritten: a name already in the folder fails that file (a name "held by an unfinished upload" is an
empty placeholder an earlier, abandoned upload left; it goes away when its session expires: retry later). All pushes go
to --branch (default main; the clone must have that branch checked out).

If `upload` can't finish a file (an error, SIGTERM, Ctrl-C) it cancels that file's unused upload session (HTTP DELETE to
the session URL) so SharePoint keeps no hidden 0-byte placeholder. The one-time key and the state file stay in the keys
directory until the upload ends. If the process was killed outright (nothing could cancel), `upload-resume <id>` reopens
the sealed pickup (from the branch tip, or from git history if it was already removed), finishes the PUT of each file not
yet sent when its session is still alive (from nextExpectedRanges) or reports it expired, then sends the verify request.

The keys directory is ~/.local/state/tf-courier-keys (0700; survives a reboot), or $TF_COURIER_KEYS. Every `git pull
--rebase` uses --autostash, so a dirty clone doesn't silently stop it. If the pickup is already gone from the branch
tip (`list`, `fetch`, `upload`), it is read from git history before giving up.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import secrets
import shlex
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from box import new_keypair, open_box, seal  # noqa: E402

MAX_BYTES = 25 * 1024 * 1024
FETCH_KEYS = Path(os.environ.get("TF_COURIER_KEYS") or Path.home() / ".local" / "state" / "tf-courier-keys")
_LAST_PULL_ERROR = ""   # stderr of the most recent failed `git pull`, shown if a push gives up


def _keys_dir() -> Path:
    """The one-time-keys directory, created private (0700) if needed. It must survive a reboot: a lost key means
    the box can neither resume nor cancel an upload session."""
    FETCH_KEYS.mkdir(mode=0o700, parents=True, exist_ok=True)
    FETCH_KEYS.chmod(0o700)
    return FETCH_KEYS


def git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(repo), *args], check=check, capture_output=True, text=True)


def pull(repo: Path, branch: str = "main") -> subprocess.CompletedProcess:
    """`git pull --rebase --autostash` (a dirty clone is stashed, not silently refused). A failed pull is remembered
    for push()'s error message, and a half-done rebase is aborted so the clone isn't left stuck."""
    global _LAST_PULL_ERROR
    r = git(repo, "pull", "-q", "--rebase", "--autostash", "origin", branch, check=False)
    if r.returncode != 0:
        _LAST_PULL_ERROR = r.stderr.strip() or r.stdout.strip()
        git(repo, "rebase", "--abort", check=False)
    return r


def push(repo: Path, message: str, branch: str = "main") -> None:
    git(repo, "commit", "-q", "-m", message)
    for attempt in range(4):
        r = git(repo, "push", "-q", "origin", f"HEAD:{branch}", check=False)
        if r.returncode == 0:
            return
        time.sleep(3 + attempt * 3)
        pull(repo, branch)
    pulled = f"\n  last pull error: {_LAST_PULL_ERROR}" if _LAST_PULL_ERROR else ""
    sys.exit(f"✗ push failed: {r.stderr.strip()}{pulled}")


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
    pull(repo, "main")
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
    _keys_dir()
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
        pull(repo, "main")
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
    blob = _pickup_blob(repo, "main", f"pickup/{cid}.tfc")
    if blob is None:
        sys.exit(f"✗ fetch {cid}: result says ok but the sealed copy is gone (collected already, or older than 24 h)")
    m = json.loads(open_box(blob, keyfile.read_text()))
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
    pull(repo, "main")
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


def _check_branch(repo: Path, branch: str) -> None:
    """Pushes go to origin/<branch> from the clone's HEAD, so the clone must have that branch checked out (a detached
    HEAD, as in an attached-repo run, is allowed for main). Stops a test-branch clone from pushing its commits to main."""
    here = git(repo, "rev-parse", "--abbrev-ref", "HEAD", check=False).stdout.strip()
    if here != branch and not (branch == "main" and here == "HEAD"):
        sys.exit(f"✗ the clone has {here!r} checked out but --branch is {branch!r}; check out {branch} first")


def _send_request(repo: Path, folder: str, cid: str, req: dict, branch: str, label: str) -> None:
    """Seal `req` with courier.pub into <folder>/<cid>.tfc and push it to `branch`."""
    pull(repo, branch)
    (repo / folder).mkdir(exist_ok=True)
    (repo / folder / f"{cid}.tfc").write_bytes(seal(json.dumps(req).encode(), (repo / "courier.pub").read_text().strip()))
    git(repo, "add", f"{folder}/{cid}.tfc")
    push(repo, f"{label} {cid}", branch)


def _show(repo: Path, branch: str, rel: str, binary: bool = False):
    """Contents of <rel> on origin/<branch> after a fetch, or None."""
    git(repo, "fetch", "-q", "origin", branch, check=False)
    r = subprocess.run(["git", "-C", str(repo), "show", f"origin/{branch}:{rel}"], capture_output=True, check=False)
    if r.returncode != 0:
        return None
    return r.stdout if binary else r.stdout.decode("utf-8")


def _pickup_blob(repo: Path, branch: str, rel: str):
    """The sealed pickup <rel>: from the tip of origin/<branch>, else (someone removed it already, e.g. an earlier run of
    the same job) from the last commit in its history that had it. None if neither has it."""
    blob = _show(repo, branch, rel, binary=True)
    if blob is not None:
        return blob
    sha = git(repo, "log", "-n", "1", "--diff-filter=AM", "--format=%H", f"origin/{branch}", "--", rel,
              check=False).stdout.strip()
    if not sha:
        return None
    r = subprocess.run(["git", "-C", str(repo), "show", f"{sha}:{rel}"], capture_output=True, check=False)
    return r.stdout if r.returncode == 0 else None


def _wait_json(repo: Path, branch: str, rel: str, timeout: float, poll: float | None = None, accept=lambda d: True):
    """Poll origin/<branch>:<rel> until it holds JSON that `accept`s; None on timeout."""
    deadline = time.time() + timeout
    while True:
        txt = _show(repo, branch, rel)
        if txt is not None:
            d = json.loads(txt)
            if accept(d):
                return d
        if time.time() >= deadline:
            return None
        time.sleep(POLL if poll is None else poll)


def _pickup_done(repo: Path, branch: str, rel: str, label: str) -> None:
    """Remove a collected pickup from the repo (commit + push)."""
    pull(repo, branch)
    if (repo / rel).exists():
        git(repo, "rm", "-q", rel)
        push(repo, label, branch)


def cmd_list(a: argparse.Namespace) -> None:
    """What is in a SharePoint folder, through the courier's `fetch` Action (kind "list"; see the module notes)."""
    repo, out = Path(a.repo), Path(a.out)
    if out.exists():
        sys.exit(f"✗ {out} already exists — pick a new --out path (list never overwrites a local file)")
    if not a.drive or bool(a.item) == bool(a.path):
        sys.exit("✗ list needs --drive and exactly one of --item or --path")
    if not 0 <= a.depth <= 4:
        sys.exit("✗ --depth must be 0 … 4")
    _check_branch(repo, a.branch)
    _keys_dir()
    priv, pub_once = new_keypair()
    cid = _new_id()
    keyfile = FETCH_KEYS / f"{cid}.key"
    keyfile.write_text(priv)
    keyfile.chmod(0o600)
    req = {"id": cid, "kind": "list", "drive_id": a.drive, "item_id": a.item or "", "path": a.path or "",
           "depth": a.depth, "reply_pub": pub_once}
    _send_request(repo, "fetch", cid, req, a.branch, "list")
    print(f"requested {cid}")
    res = _wait_json(repo, a.branch, f"results/fetch-{cid}.json", a.timeout)
    if res is None:
        keyfile.unlink(missing_ok=True)
        print(f"… no result yet for {cid} — the listing can't be collected later; run list again")
        sys.exit(2)
    if res.get("status") != "ok":
        keyfile.unlink(missing_ok=True)
        print(f"✗ list {cid} failed: {res.get('error', 'no reason given')}")
        sys.exit(1)
    blob = _pickup_blob(repo, a.branch, f"pickup/{cid}.tfc")
    if blob is None:
        keyfile.unlink(missing_ok=True)
        sys.exit(f"✗ list {cid}: result says ok but the sealed listing is gone (not on the branch, not in its history)")
    m = json.loads(open_box(blob, keyfile.read_text()))
    if m.get("id") != cid or len(m.get("entries", [])) != res.get("count"):
        keyfile.unlink(missing_ok=True)
        sys.exit(f"✗ list {cid}: the listing doesn't match its result — not saved")
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("x", encoding="utf-8") as fh:
        json.dump(m, fh, indent=1, ensure_ascii=False)
        fh.write("\n")
    _pickup_done(repo, a.branch, f"pickup/{cid}.tfc", f"picked up {cid}")
    keyfile.unlink(missing_ok=True)
    print(f"✓ listed {cid} — {len(m['entries'])} entries saved to {out}")


# ---- upload (big files; the bytes go straight to SharePoint, never through git)
UNIT = 320 * 1024                 # Graph wants upload chunks in multiples of 320 KiB
POLL = float(os.environ.get("TF_COURIER_POLL", "10"))              # seconds between looks at the courier repo
RETRY_SLEEP = float(os.environ.get("TF_UPLOAD_RETRY_SLEEP", "2"))   # seconds; doubles each retry


def name_problem(name: str) -> str:
    if not name or any(c in name for c in '\\/:*?"<>|#%') or name != name.strip():
        return f"SharePoint name not allowed: {name!r}"
    return ""


def hash_file(path: Path) -> tuple[str, str, int]:
    """-> (sha256 hex, quickxor base64, size), one streaming pass."""
    from quickxor import QuickXor
    sha, qx = hashlib.sha256(), QuickXor()
    with path.open("rb") as f:
        while piece := f.read(8 * 1024 * 1024):
            sha.update(piece)
            qx.update(piece)
    return sha.hexdigest(), qx.b64(), qx.n


def _next_expected(url: str, requests) -> int | None:
    """Where the session wants the next byte (from nextExpectedRanges), or None if it can't be read."""
    try:
        r = requests.get(url, timeout=60)
        if r.status_code == 200:
            ranges = r.json().get("nextExpectedRanges") or []
            if ranges:
                return int(str(ranges[0]).split("-")[0])
    except Exception:
        pass
    return None


def cancel_session(url: str) -> bool:
    """Drop an upload session (HTTP DELETE to its URL: no Authorization header, the URL is never printed) so SharePoint
    keeps no hidden 0-byte placeholder under the name. True when the session is gone (cancelled, or already unknown)."""
    import requests
    try:
        return requests.delete(url, timeout=60).status_code in (200, 202, 204, 404)
    except requests.RequestException:
        return False


def put_file(url: str, path: Path, total: int, chunk: int, start: int = 0) -> dict:
    """PUT the file to a pre-authenticated upload URL in chunks (no Authorization header; the URL is never printed),
    beginning at byte `start` (0, or where a live session says it expects the next byte).
    -> the driveItem of the final response. Raises RuntimeError (without the URL) when it can't finish."""
    import requests
    offset, tries = start, 0
    with path.open("rb") as f:
        while True:
            end = min(offset + chunk, total) - 1
            f.seek(offset)
            data = f.read(end - offset + 1)
            hdr = {"Content-Length": str(len(data)), "Content-Range": f"bytes {offset}-{end}/{total}"}
            code, r = None, None
            try:
                r = requests.put(url, data=data, headers=hdr, timeout=300)
                code = r.status_code
            except requests.RequestException as e:
                code = type(e).__name__          # never str(e): it contains the URL
            if code in (200, 201) and end == total - 1:
                return r.json()
            if code == 202 and end < total - 1:
                nxt = None
                try:
                    nxt = int(str(r.json().get("nextExpectedRanges", [""])[0]).split("-")[0])
                except Exception:
                    pass
                offset, tries = (nxt if nxt is not None and offset < nxt <= total - 1 else end + 1), 0
                continue
            if code == 409:
                raise RuntimeError("name conflict: the name already exists in the folder (not replaced)")
            if isinstance(code, int) and 400 <= code < 500 and code not in (408, 416, 429):
                raise RuntimeError(f"SharePoint refused a chunk ({code})")
            tries += 1
            if tries > 5:
                raise RuntimeError(f"gave up after 5 retries at byte {offset} (last: {code})")
            time.sleep(RETRY_SLEEP * 2 ** (tries - 1))
            nxt = _next_expected(url, requests)
            if nxt is not None and 0 <= nxt < total:
                offset = nxt
            # a 200/201/202 that doesn't fit (e.g. lost reply) falls through to a retry from the server's position


def _print_rows(rows: list[dict]) -> None:
    for f in rows:
        print(f"  {'✓' if f.get('verified') else '✗'} {f.get('index')}  {f.get('bytes')} bytes  sha256 {f.get('sha256')}"
              f"  item {f.get('item_id') or ''}" + (f"  {f['error']}" if f.get("error") else ""))


def _state_path(cid: str) -> Path:
    return FETCH_KEYS / f"up-{cid}.state.json"


def _save_state(cid: str, state: dict) -> None:
    _state_path(cid).write_text(json.dumps(state))   # no URLs, no keys


def _end_upload(cid: str) -> None:
    """The upload is over (verified, failed, or nothing sent): only now do the state and the one-time key go."""
    _state_path(cid).unlink(missing_ok=True)
    (FETCH_KEYS / f"up-{cid}.key").unlink(missing_ok=True)


def _raise_interrupt(signum, frame):
    raise KeyboardInterrupt   # SIGTERM is handled like Ctrl-C, so the unfinished sessions get cancelled


def _held(error: str) -> bool:
    """True for the courier's "held by an unfinished upload (empty placeholder)" refusal."""
    return "held by an unfinished upload" in (error or "")


def _upload_finish(a: argparse.Namespace, cid: str, state: dict, timeout: float) -> None:
    """Wait for the verify result, print one line per file, write --summary, exit 0/1/2."""
    res = _wait_json(Path(a.repo), a.branch, f"results/upload-{cid}.json", timeout,
                     accept=lambda d: d.get("stage") == "verify")
    if res is None:
        print(f"… no verify result yet for {cid} — finish later with: upload-wait --repo {a.repo} --branch {a.branch} {cid}")
        sys.exit(2)
    by_index = {f["index"]: f for f in res.get("files", [])}
    for f in state["files"]:
        row = by_index.get(f["index"])
        if row:
            f["verified"], f["item_id"] = bool(row.get("verified")), row.get("item_id") or f.get("item_id", "")
            f["error"] = row.get("error", f.get("error", ""))
            f["office_rewritten"] = bool(row.get("office_rewritten"))   # SharePoint stamped its properties: accepted
        else:
            f["verified"] = False
            f["error"] = f.get("error") or "not verified"
    _print_rows([{**f} for f in state["files"]])
    _write_summary(getattr(a, "summary", ""), cid, state)
    ok = res.get("status") == "ok" and all(f["verified"] for f in state["files"])
    print(f"{'✓ uploaded and verified' if ok else '✗ upload failed'} {cid}")
    _end_upload(cid)
    sys.exit(0 if ok else 1)


def _write_summary(path: str, cid: str, state: dict) -> None:
    """--summary file: one row per file. held_placeholder is true when the name is held by an earlier unfinished
    upload (retry later); office_rewritten when SharePoint rewrote an Office file's properties (still verified)."""
    if not path:
        return
    keys = ("index", "local", "name", "bytes", "sha256", "quickxor", "item_id", "verified", "error",
            "existing_bytes", "existing_item_id")
    rows = [{**{k: f.get(k, "") for k in keys}, "held_placeholder": _held(f.get("error", "")),
             "office_rewritten": bool(f.get("office_rewritten"))} for f in state["files"]]
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"id": cid, "files": rows}, fh, indent=1, ensure_ascii=False)
        fh.write("\n")


def _chunk_bytes(mb: float) -> int:
    chunk = int(mb * 1024 * 1024) // UNIT * UNIT
    if chunk < UNIT:
        sys.exit("✗ --chunk-mb is too small (one chunk is at least 320 KiB)")
    return chunk


def _put_all(cid: str, state: dict, sessions: dict, chunk: int, resume: bool = False, prelude=None) -> None:
    """PUT every file that has a session and isn't sent yet (`resume`: start where a live session says; an expired one
    is reported). A file that can't be finished has its session cancelled, so no empty placeholder stays behind. On
    SIGTERM / Ctrl-C (or if `prelude` dies) every unfinished session is cancelled before the interrupt goes on.
    The state is saved after each file."""
    files = state["files"]
    todo = [f for f in files if not f.get("item_id") and sessions.get(f["index"])]
    for f in files:
        if not f.get("item_id") and not sessions.get(f["index"]):
            f["error"] = f.get("error") or "no upload session"
    finished: set = set()
    try:
        if prelude:
            prelude()
        for f in todo:
            url = sessions[f["index"]]["upload_url"]
            try:
                print(f"  → {f['index']}  {f['bytes']} bytes")
                start = 0
                if resume:
                    import requests
                    start = _next_expected(url, requests)
                    if start is None:
                        raise RuntimeError("upload session expired or cancelled — run upload again")
                f["item_id"] = put_file(url, Path(f["local"]), f["bytes"], chunk, start)["id"]
                f["error"] = ""
            except Exception as e:
                f["error"] = str(e) if isinstance(e, RuntimeError) else f"upload failed ({type(e).__name__})"
                cancel_session(url)
            finished.add(f["index"])
            _save_state(cid, state)
    except BaseException:
        for f in todo:
            if f["index"] not in finished:
                f["error"] = "interrupted — its upload session was cancelled"
                cancel_session(sessions[f["index"]]["upload_url"])
        _save_state(cid, state)
        raise


def _send_files_then_verify(a: argparse.Namespace, repo: Path, cid: str, state: dict, sessions: dict, drive: str,
                            folder: str, collected, resume: bool = False) -> None:
    """PUT the files, then send the verify request for those that went through and wait for its result."""
    _put_all(cid, state, sessions, _chunk_bytes(a.chunk_mb), resume, prelude=collected)
    sent = [f for f in state["files"] if f["item_id"]]
    if not sent:
        _print_rows(state["files"])
        _write_summary(a.summary, cid, state)
        print(f"✗ upload failed {cid}")
        _end_upload(cid)
        sys.exit(1)
    vfiles = []
    for f in sent:
        row = {k: f[k] for k in ("index", "item_id", "name", "bytes", "sha256", "quickxor")}
        fp = _office_fingerprint(repo, f["local"], f["bytes"])
        if fp:
            row["fingerprint"] = fp   # SharePoint stamps Office files; the robot checks every content part against this
        vfiles.append(row)
    vreq = {"kind": "verify", "id": cid, "drive_id": drive, "folder_id": folder, "files": vfiles}
    _send_request(repo, "upload", cid, vreq, a.branch, "upload verify")
    _upload_finish(a, cid, state, a.timeout)


OFFICE_EXT = (".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".doc", ".dot", ".xls", ".xlt", ".ppt", ".pot")


def _office_fingerprint(repo: Path, local: str, size: int) -> dict | None:
    """Content fingerprint of an Office file (courier/verify.py, read from the courier clone so both sides use the
    same code). None for other files, files over 500 MB, or an old courier clone without verify.py."""
    if not local.lower().endswith(OFFICE_EXT) or size > 500 * 1024 * 1024:
        return None
    src = Path(repo) / "courier" / "verify.py"
    if not src.is_file():
        return None
    import importlib.util
    spec = importlib.util.spec_from_file_location("courier_verify", src)
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
        fingerprint = mod.fingerprint
    except Exception:  # noqa: BLE001
        return None
    try:
        return fingerprint(Path(local).read_bytes())
    except Exception:  # noqa: BLE001 - no fingerprint; the robot then refuses a changed Office file and says why
        return None


def _note_refused(state: dict, res: dict) -> None:
    """Copy the session result's per-file refusals (and what already sits under a taken name) into the state."""
    for row in res.get("files", []):
        if not row.get("ready"):
            f = state["files"][row["index"]]
            f["error"] = row.get("error", "session refused")
            for k in ("existing_bytes", "existing_item_id"):
                if k in row:
                    f[k] = row[k]


def cmd_upload(a: argparse.Namespace) -> None:
    repo = Path(a.repo)
    _check_branch(repo, a.branch)
    if a.summary and Path(a.summary).exists():
        sys.exit(f"✗ {a.summary} already exists — pick a new --summary path")
    _chunk_bytes(a.chunk_mb)
    files, names = [], set()
    for i, spec in enumerate(a.files):
        local, _, name = spec.partition("=")
        p = Path(local)
        if not p.is_file():
            sys.exit(f"✗ {p}: not a file")
        name = name or p.name
        if name_problem(name):
            sys.exit(f"✗ {name_problem(name)}")
        if name in names:
            sys.exit(f"✗ two files would be named {name!r}")
        names.add(name)
        sha, qx, size = hash_file(p)
        if size < 1:
            sys.exit(f"✗ {p}: empty file")
        files.append({"index": i, "local": str(p), "name": name, "bytes": size, "sha256": sha, "quickxor": qx,
                      "item_id": "", "verified": False, "error": ""})
    cid = _new_id()
    print(f"upload {cid} ({len(files)} file{'s' if len(files) != 1 else ''})")
    _keys_dir()
    signal.signal(signal.SIGTERM, _raise_interrupt)
    state = {"files": files}
    _save_state(cid, state)   # for upload-wait / upload-resume
    priv, pub_once = new_keypair()
    keyfile = FETCH_KEYS / f"up-{cid}.key"
    keyfile.write_text(priv)
    keyfile.chmod(0o600)
    req = {"kind": "session", "id": cid, "drive_id": a.drive, "folder_id": a.folder, "reply_pub": pub_once,
           "files": [{k: f[k] for k in ("index", "name", "bytes", "sha256")} for f in files]}
    _send_request(repo, "upload", cid, req, a.branch, "upload session")
    res = _wait_json(repo, a.branch, f"results/upload-{cid}.json", min(a.timeout, 1800),
                     accept=lambda d: d.get("stage") in ("session", "verify"))
    if res is None:
        # the key and state stay: the courier may still answer, and then the session is theirs to finish or cancel
        print(f"… no session result for {cid} yet — finish or cancel it later with: "
              f"upload-resume --repo {a.repo} --branch {a.branch} {cid}")
        sys.exit(2)
    _note_refused(state, res)
    sessions, blob = {}, _pickup_blob(repo, a.branch, f"pickup/up-{cid}.tfc")
    if blob is not None:
        sessions = {f["index"]: f for f in json.loads(open_box(blob, keyfile.read_text()))["files"]}
    _send_files_then_verify(
        a, repo, cid, state, sessions, a.drive, a.folder,
        collected=lambda: _pickup_done(repo, a.branch, f"pickup/up-{cid}.tfc", f"picked up upload {cid}"))


def cmd_upload_resume(a: argparse.Namespace) -> None:
    """After a crash / reboot / kill: reopen the sealed pickup (branch tip, or git history if already removed), finish
    the PUT of each file not yet sent while its session is alive (else report it expired and cancel it), then send the
    verify request as `upload` does."""
    repo, cid = Path(a.repo), a.id
    _check_branch(repo, a.branch)
    keyfile = FETCH_KEYS / f"up-{cid}.key"
    if not _state_path(cid).exists():
        sys.exit(f"✗ no saved state for {cid} on this machine — run upload again")
    state = json.loads(_state_path(cid).read_text())
    res = _wait_json(repo, a.branch, f"results/upload-{cid}.json", a.timeout,
                     accept=lambda d: d.get("stage") in ("session", "verify"))
    if res is None:
        print(f"… the courier hasn't answered {cid} yet — try upload-resume again later")
        sys.exit(2)
    if res.get("stage") == "verify":
        _upload_finish(a, cid, state, 0)
    signal.signal(signal.SIGTERM, _raise_interrupt)
    _note_refused(state, res)
    blob = _pickup_blob(repo, a.branch, f"pickup/up-{cid}.tfc")
    if blob is None and any(r.get("ready") for r in res.get("files", [])):
        _end_upload(cid)
        sys.exit(f"✗ upload {cid}: the sealed upload sessions are gone — run upload again")
    if blob is not None and not keyfile.exists():
        sys.exit(f"✗ no one-time key for {cid} on this machine — its sessions can't be opened; run upload again")
    manifest = json.loads(open_box(blob, keyfile.read_text())) if blob is not None else {}
    sessions = {f["index"]: f for f in manifest.get("files", [])}
    drive, folder = manifest.get("drive_id", ""), manifest.get("folder_id", "")
    _send_files_then_verify(
        a, repo, cid, state, sessions, drive, folder, resume=True,
        collected=lambda: _pickup_done(repo, a.branch, f"pickup/up-{cid}.tfc", f"picked up upload {cid}"))


def cmd_upload_wait(a: argparse.Namespace) -> None:
    sf = _state_path(a.id)
    state = json.loads(sf.read_text()) if sf.exists() else None
    res = _wait_json(Path(a.repo), a.branch, f"results/upload-{a.id}.json", a.timeout,
                     accept=lambda d: d.get("stage") == "verify")
    if state is None and res is not None:
        state = {"files": [{"index": f["index"], "bytes": f.get("bytes"), "sha256": f.get("sha256")}
                           for f in res.get("files", [])]}
    if state is None:
        print(f"… no result yet for {a.id}")
        sys.exit(2)
    a.summary = getattr(a, "summary", "")
    _upload_finish(a, a.id, state, 0)


def cmd_directory(a: argparse.Namespace) -> None:
    """Hand new contacts to the courier repo's `directory` Action (System Rules B3) by pushing trigger/directory_add.csv."""
    import csv
    src = Path(a.csv)
    rows = list(csv.reader(src.read_text(encoding="utf-8-sig").splitlines()))
    if len(rows) < 2 or "firm_id" not in rows[0] or "contact_name" not in rows[0]:
        sys.exit("✗ not written: the CSV needs Directory.csv's header row (firm_id, contact_name, …) and at least one contact")
    repo = Path(a.repo)
    pull(repo, "main")
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
    """Validate a small JSON request and hand it to the courier repo's `operations` Action.
    A required key written "a|b" is satisfied by either key."""
    src = Path(a.json)
    try:
        req = json.loads(src.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        sys.exit(f"✗ not written: {src} isn't readable JSON ({e})")
    missing = [k for k in required if not any(req.get(x) for x in k.split("|"))]
    if missing:
        sys.exit(f"✗ not written: missing {', '.join(missing)}")
    repo = Path(a.repo)
    pull(repo, "main")
    dest = repo / "trigger" / name
    dest.parent.mkdir(exist_ok=True)
    dest.write_text(json.dumps(req, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    git(repo, "add", f"trigger/{name}")
    if git(repo, "diff", "--cached", "--quiet", check=False).returncode == 0:
        print(f"= same {label} as the last hand-off; nothing pushed")
        return
    push(repo, f"operations: {label} {req.get('bid_id') or req.get('vista') or ''}".rstrip())
    result = name.replace("job_update", "update").replace("job_load", "load").replace(".json", "")
    print(f"✓ {label} handed to the operations Action; result in results/operations-{result}-latest.json")


def cmd_award(a: argparse.Namespace) -> None:
    """Hand an award to the Operations robot (Award Handoff plan): pushes trigger/award.json."""
    _push_trigger(a, "award.json", ("bid_id", "project"), "award")


def cmd_job_update(a: argparse.Namespace) -> None:
    """Hand David's answers for a Jobs row to the Operations robot: pushes trigger/job_update.json.
    The row is found by bid_id (award questions) or vista (jobs loaded from Vista; "done or demobilized?")."""
    _push_trigger(a, "job_update.json", ("bid_id|vista", "set"), "job update")


def cmd_job_load(a: argparse.Namespace) -> None:
    """One time, from chat with David's OK: Jobs rows for jobs open before the award robot (trigger/job_load.json)."""
    _push_trigger(a, "job_load.json", ("jobs",), "job load")


def cmd_tick(a: argparse.Namespace) -> None:
    """Start the Operations robot now (System Rules B18): the hourly run ticks at its start, the nightly run with
    --nightly. GitHub's own schedule is only a backup; it comes hours late or not at all."""
    repo = Path(a.repo)
    pull(repo, "main")
    dest = repo / "trigger" / "tick.txt"
    dest.parent.mkdir(exist_ok=True)
    kind = "nightly" if a.nightly else "hourly"
    dest.write_text(f"{kind} {datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}\n", encoding="utf-8")
    git(repo, "add", "trigger/tick.txt")
    push(repo, f"operations: {kind} tick")
    print(f"✓ {kind} tick: the operations robot starts in about a minute")



def cmd_staging(a: argparse.Namespace) -> None:
    """mkdir, move, or (narrowly) delete through the courier's `staging` Action. Switched off
    until the deployment turns it on. Delete is refused unless the item sits inside the
    configured staging root; the root itself is never deleted."""
    repo = Path(a.repo)
    _check_branch(repo, a.branch)
    if not a.drive:
        sys.exit("✗ staging needs --drive")
    cid = _new_id()
    req = {"id": cid, "kind": a.staging_cmd, "drive_id": a.drive}
    if a.staging_cmd == "mkdir":
        if not a.parent or not a.name:
            sys.exit("✗ mkdir needs --parent and --name")
        req.update(parent_id=a.parent, name=a.name)
    elif a.staging_cmd == "move":
        if not a.item or not a.dest:
            sys.exit("✗ move needs --item and --dest")
        req.update(item_id=a.item, dest_folder_id=a.dest)
        if a.name:
            req["name"] = a.name
    else:  # delete
        ids = list(a.items or [])
        if not ids:
            sys.exit("✗ delete needs one or more --item ids")
        if a.purpose not in ("pass", "cleanup"):
            sys.exit("✗ delete needs --purpose pass or cleanup")
        req.update(item_ids=ids, purpose=a.purpose)
    if a.staging_cmd in ("mkdir", "move") and req.get("name") and name_problem(req["name"]):
        sys.exit(f"✗ {name_problem(req['name'])} — nothing was sent")   # the Action would refuse it after a full round trip
    if a.out and Path(a.out).exists():
        sys.exit(f"✗ {a.out} already exists — pick a new --out path")
    print(f"staging {a.staging_cmd} {cid}")
    _send_request(repo, "staging", cid, req, a.branch, f"staging {a.staging_cmd}")
    res = _wait_json(repo, a.branch, f"results/staging-{cid}.json", a.timeout)
    if res is None:
        print(f"… no result yet for {cid} — look at results/staging-{cid}.json once the Action finishes")
        sys.exit(2)
    if a.out:
        Path(a.out).parent.mkdir(parents=True, exist_ok=True)
        with open(a.out, "x", encoding="utf-8") as fh:
            json.dump(res, fh, indent=1, ensure_ascii=False)
            fh.write("\n")
    if res.get("status") != "ok":
        print(f"✗ staging {cid} failed: {res.get('error', 'no reason given')}")
        if res.get("item_id"):
            print(f"  item {res['item_id']}")
        for row in res.get("items") or []:
            print(f"  {'✓' if row.get('ok') else '✗'} item {row.get('item_id','')}"
                  + (f"  {row['error']}" if row.get("error") else ""))
        sys.exit(1)
    print(f"✓ staging {a.staging_cmd} {cid}"
          + (f"  item {res['item_id']}" if res.get("item_id") else "")
          + (f"  {len(res.get('items') or [])} item(s)" if res.get("items") else ""))
    for row in res.get("items") or []:
        print(f"  ✓ item {row.get('item_id','')}")
    sys.exit(0)


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
    ls = sub.add_parser("list")
    ls.add_argument("--repo", required=True)
    ls.add_argument("--drive", default="")
    ls.add_argument("--item", default="")
    ls.add_argument("--path", default="")
    ls.add_argument("--depth", type=int, default=2)
    ls.add_argument("--branch", default="main")
    ls.add_argument("--out", required=True)
    ls.add_argument("--timeout", type=int, default=300)
    ls.set_defaults(fn=cmd_list)
    up = sub.add_parser("upload")
    up.add_argument("--repo", required=True)
    up.add_argument("--drive", required=True)
    up.add_argument("--folder", required=True)
    up.add_argument("--branch", default="main")
    up.add_argument("--timeout", type=int, default=3600)
    up.add_argument("--chunk-mb", type=float, default=10)
    up.add_argument("--summary", default="")
    up.add_argument("files", nargs="+")
    up.set_defaults(fn=cmd_upload)
    ur = sub.add_parser("upload-resume")
    ur.add_argument("--repo", required=True)
    ur.add_argument("--branch", default="main")
    ur.add_argument("--timeout", type=int, default=300)
    ur.add_argument("--chunk-mb", type=float, default=10)
    ur.add_argument("--summary", default="")
    ur.add_argument("id")
    ur.set_defaults(fn=cmd_upload_resume)
    uw = sub.add_parser("upload-wait")
    uw.add_argument("--repo", required=True)
    uw.add_argument("--branch", default="main")
    uw.add_argument("--timeout", type=int, default=3600)
    uw.add_argument("--summary", default="")
    uw.add_argument("id")
    uw.set_defaults(fn=cmd_upload_wait)
    w = sub.add_parser("wait")
    w.add_argument("--repo", required=True)
    w.add_argument("id")
    w.add_argument("--timeout", type=int, default=300)
    w.set_defaults(fn=cmd_wait)
    d = sub.add_parser("directory")
    d.add_argument("--repo", required=True)
    d.add_argument("csv")
    d.set_defaults(fn=cmd_directory)
    t = sub.add_parser("tick")
    t.add_argument("--repo", required=True)
    t.add_argument("--nightly", action="store_true")
    t.set_defaults(fn=cmd_tick)
    for nm, fn in (("award", cmd_award), ("job-update", cmd_job_update), ("job-load", cmd_job_load)):
        x = sub.add_parser(nm)
        x.add_argument("--repo", required=True)
        x.add_argument("json")
        x.set_defaults(fn=fn)
    st = sub.add_parser("staging")
    stg = st.add_subparsers(dest="staging_cmd", required=True)
    def _st(name):
        p = stg.add_parser(name)
        p.add_argument("--repo", required=True)
        p.add_argument("--drive", required=True)
        p.add_argument("--branch", default="main")
        p.add_argument("--timeout", type=int, default=600)
        p.add_argument("--out", default="")
        p.set_defaults(fn=cmd_staging)
        return p
    mk = _st("mkdir")
    mk.add_argument("--parent", required=True)
    mk.add_argument("--name", required=True)
    mv = _st("move")
    mv.add_argument("--item", required=True)
    mv.add_argument("--dest", required=True)
    mv.add_argument("--name", default="")
    dl = _st("delete")
    dl.add_argument("--item", dest="items", action="append", required=True)
    dl.add_argument("--purpose", required=True, choices=["pass", "cleanup"])
    a = ap.parse_args()
    try:
        a.fn(a)
    except KeyboardInterrupt:
        sys.exit("✗ interrupted — unfinished upload sessions were cancelled; the state is kept (see upload-resume)")


if __name__ == "__main__":
    main()
