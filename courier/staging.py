#!/usr/bin/env python3
"""Courier, staging side: create, move, or (narrowly) delete items in a configured staging root.

A box runs print kit `send.py staging …`, which pushes <repo>/staging/<id>.tfc (sealed with courier.pub).
Three kinds, all refused unless STAGING_ENABLED is true and staging_lib allows the path:

  mkdir   {kind, id, drive_id, parent_id, name}
          Creates one folder (conflictBehavior fail). Result includes the new item id (no name).
  move    {kind, id, drive_id, item_id, dest_folder_id, name?}
          Moves an item that sits inside the staging root into an allowed job subfolder
          (plans, addenda, or estimate). Never overwrites. Optional new name.
  delete  {kind, id, drive_id, item_ids, purpose: pass|cleanup}
          Deletes only items that sit inside the staging root. The staging root itself is
          never deleted. purpose must be pass or cleanup. This is the only SharePoint
          delete this courier performs.

Writes <repo>/results/staging-<id>.json ({id, kind, status, finished, item_id? / items[..]}:
item ids only, no names). Deletes the request. Exit 1 if anything failed, so GitHub emails
the owner.

    python3 courier/staging.py --repo .
"""
from __future__ import annotations

import argparse
import json
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

import requests
from cryptography.exceptions import InvalidTag

sys.path.insert(0, str(Path(__file__).parent))
from box import open_box  # noqa: E402
from receive import GRAPH, token  # noqa: E402
import staging_lib as sl  # noqa: E402

PURPOSES = {"pass", "cleanup"}


def _hdr(tok: str) -> dict:
    return {"Authorization": f"Bearer {tok}"}


def _meta_fn(drive: str, tok: str, cache: dict):
    def meta(item_id):
        key = (drive, str(item_id))
        if key not in cache:
            try:
                r = requests.get(
                    f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(str(item_id), safe='')}"
                    "?$select=id,name,folder,file,parentReference",
                    headers=_hdr(tok), timeout=60,
                )
                cache[key] = r.json() if r.status_code == 200 else f"couldn't look up the folder ({r.status_code})"
            except Exception as e:
                cache[key] = f"couldn't look up the folder ({type(e).__name__})"
        return cache[key]
    return meta


def do_mkdir(req: dict, tok: str) -> dict:
    s, drive, parent, name = sl.settings(), req.get("drive_id", ""), req.get("parent_id", ""), req.get("name", "")
    meta = _meta_fn(drive, tok, {})
    why = sl.mkdir_allowed(drive, parent, name, meta, s)
    if why:
        raise RuntimeError(why)
    body = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    r = requests.post(f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(str(parent), safe='')}/children",
                      headers=_hdr(tok), json=body, timeout=60)
    if r.status_code == 409:
        raise RuntimeError("already in the folder (not replaced)")
    if r.status_code not in (200, 201):
        raise RuntimeError(f"couldn't create the folder ({r.status_code})")
    item = r.json()
    return {"item_id": item.get("id", "")}


def do_move(req: dict, tok: str) -> dict:
    s = sl.settings()
    drive, item, dest = req.get("drive_id", ""), req.get("item_id", ""), req.get("dest_folder_id", "")
    new_name = req.get("name") or ""
    if new_name and sl.name_problem(new_name):
        raise RuntimeError(sl.name_problem(new_name))
    meta = _meta_fn(drive, tok, {})
    why = sl.move_allowed(drive, item, dest, meta, s)
    if why:
        raise RuntimeError(why)
    src = meta(item)
    if isinstance(src, str):
        raise RuntimeError(src)
    dest_name = new_name or src.get("name", "")
    # refuse if that name is already in the destination
    ex = requests.get(
        f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(str(dest), safe='')}:/{urllib.parse.quote(dest_name, safe='')}",
        headers=_hdr(tok), timeout=60,
    )
    if ex.status_code == 200:
        raise RuntimeError("already in the folder (not replaced)")
    if ex.status_code != 404:
        raise RuntimeError(f"couldn't check the destination ({ex.status_code})")
    patch = {"parentReference": {"id": dest}}
    if new_name:
        patch["name"] = new_name
    r = requests.patch(
        f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(str(item), safe='')}?@microsoft.graph.conflictBehavior=fail",
        headers=_hdr(tok), json=patch, timeout=60,
    )
    if r.status_code == 409:
        raise RuntimeError("already in the folder (not replaced)")
    if r.status_code not in (200, 201):
        raise RuntimeError(f"couldn't move the item ({r.status_code})")
    return {"item_id": r.json().get("id", item)}


def do_delete(req: dict, tok: str) -> dict:
    s = sl.settings()
    drive = req.get("drive_id", "")
    purpose = req.get("purpose", "")
    if purpose not in PURPOSES:
        raise RuntimeError("delete requires purpose pass or cleanup")
    ids = req.get("item_ids") or []
    if isinstance(req.get("item_id"), str) and req["item_id"]:
        ids = [req["item_id"], *ids]
    if not ids:
        raise RuntimeError("delete requires at least one item id")
    meta = _meta_fn(drive, tok, {})
    rows = []
    ok_all = True
    for iid in ids:
        row = {"item_id": iid, "ok": False}
        why = sl.delete_allowed(drive, iid, meta, s)
        if why:
            row["error"] = why
            ok_all = False
            rows.append(row)
            continue
        r = requests.delete(
            f"{GRAPH}/drives/{drive}/items/{urllib.parse.quote(str(iid), safe='')}",
            headers=_hdr(tok), timeout=60,
        )
        if r.status_code in (200, 204):
            row["ok"] = True
        elif r.status_code == 404:
            row["ok"] = True  # already gone: a leftover cleanup is still a success
        else:
            row["error"] = f"couldn't delete the item ({r.status_code})"
            ok_all = False
        rows.append(row)
    if not ok_all:
        raise _Partial(rows)
    return {"items": rows, "purpose": purpose}


class _Partial(RuntimeError):
    def __init__(self, rows):
        super().__init__("one or more items failed")
        self.rows = rows


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=".")
    repo = Path(ap.parse_args().repo)
    requests_dir, results = repo / "staging", repo / "results"
    results.mkdir(exist_ok=True)
    failed, tok = False, None
    if not requests_dir.is_dir():
        sys.exit(0)
    for sealed in sorted(requests_dir.glob("*.tfc")):
        cid = sealed.stem
        res = {"id": cid, "kind": "", "status": "failed", "error": ""}
        try:
            req = json.loads(open_box(sealed.read_bytes(), os_priv()))
            if not isinstance(req, dict) or req.get("id") != cid:
                raise RuntimeError("not a valid staging request")
            kind = req.get("kind")
            if kind not in ("mkdir", "move", "delete"):
                raise RuntimeError("not a valid staging request")
            res["kind"] = kind
            tok = tok or token()
            if kind == "mkdir":
                res.update(do_mkdir(req, tok))
            elif kind == "move":
                res.update(do_move(req, tok))
            else:
                res.update(do_delete(req, tok))
            res["status"] = "ok"
        except _Partial as e:
            res["items"] = e.rows
            res["error"] = str(e)
        except InvalidTag:
            res["error"] = "can't open the request (wrong COURIER_PRIVATE_KEY, or the file was altered)"
        except Exception as e:
            res["error"] = f"{type(e).__name__}: {e}" if not isinstance(e, RuntimeError) else str(e)
        res["finished"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        # never write names into the result
        if "name" in res:
            del res["name"]
        (results / f"staging-{res['id']}.json").write_text(json.dumps(res, indent=1) + "\n")
        sealed.unlink()
        failed |= res["status"] != "ok"
        print(f"{'✓' if res['status'] == 'ok' else '✗'} staging {res.get('kind') or '?'} {res['id']} {res.get('error', '')}")
    sys.exit(1 if failed else 0)


def os_priv() -> str:
    import os
    key = os.environ.get("COURIER_PRIVATE_KEY", "")
    if not key:
        raise RuntimeError("COURIER_PRIVATE_KEY is missing")
    return key


if __name__ == "__main__":
    main()
