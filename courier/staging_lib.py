#!/usr/bin/env python3
"""Allow-list checks for a staging root folder (pending bid sets).

Switched off unless STAGING_ENABLED is true. Every function is a check only — no
SharePoint writes live here. `meta(item_id)` returns a Graph-like dict (id, name,
folder, parentReference, optional root) or a string error.
"""
from __future__ import annotations

import json
import os
from pathlib import Path


def flag(name: str, default: str = "") -> bool:
    return os.environ.get(name, default).strip().lower() in {"1", "true", "yes", "on"}


def _csv(name: str) -> set[str]:
    return {x.strip() for x in os.environ.get(name, "").split(",") if x.strip()}


def _config_file() -> dict:
    """Optional courier/staging_config.json (or STAGING_CONFIG). Missing or unreadable → {}."""
    path = Path(os.environ.get("STAGING_CONFIG") or "courier/staging_config.json")
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text())
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def settings() -> dict:
    """Current staging settings. A set environment variable wins over the config file."""
    cfg = _config_file()
    raw = os.environ.get("STAGING_ENABLED")
    if raw is not None and str(raw).strip() != "":
        enabled = flag("STAGING_ENABLED")
    else:
        enabled = bool(cfg.get("enabled"))
    jobs = _csv("UPLOAD_JOB_SUBFOLDERS")
    if not jobs and isinstance(cfg.get("job_subfolders"), list):
        jobs = {str(x).strip() for x in cfg["job_subfolders"] if str(x).strip()}
    move = _csv("STAGING_MOVE_DEST_SUBFOLDERS")
    if not move:
        names = cfg.get("move_dest_subfolders")
        move = {str(x).strip() for x in names if str(x).strip()} if isinstance(names, list) else (jobs | {"07 Estimate"})
    return {
        "enabled": enabled,
        "root_name": (os.environ.get("STAGING_ROOT_NAME") or cfg.get("root_name") or "").strip(),
        "root_id": (os.environ.get("STAGING_ROOT_ID") or cfg.get("root_id") or "").strip(),
        "drive": os.environ.get("UPLOAD_DRIVE", "").strip(),
        "job_subfolders": jobs,
        "move_dest_subfolders": move,
        "status_folders": _csv("UPLOAD_STATUS_FOLDERS"),
    }


def name_problem(name) -> str:
    if not isinstance(name, str) or not name:
        return "no file name"
    if any(c in name for c in '\\/:*?"<>|#%') or name != name.strip():
        return "file name not allowed in SharePoint"
    return ""


def _is_root(m: dict) -> bool:
    return "root" in m or not (m.get("parentReference") or {}).get("id")


def walk(item_id, meta, limit: int = 6) -> list | str:
    """Chain [item, parent, …, root], or an error string."""
    if not item_id:
        return "couldn't look up the folder (no id)"
    chain = []
    cur = item_id
    for _ in range(limit):
        m = meta(cur)
        if isinstance(m, str):
            return m
        if not isinstance(m, dict):
            return "couldn't look up the folder"
        chain.append(m)
        if _is_root(m):
            return chain
        cur = (m.get("parentReference") or {}).get("id")
        if not cur:
            return chain
    return "that folder is not on the upload allow-list — refused"


def _is_staging_root(m: dict, s: dict) -> bool:
    if s["root_id"] and m.get("id") == s["root_id"]:
        return True
    if s["root_name"] and m.get("name") == s["root_name"] and "folder" in m:
        parent_id = (m.get("parentReference") or {}).get("id")
        # a staging root sits at the drive root
        if not parent_id:
            return True
        # we may not have the parent in this dict; callers also check the next chain item
        return False
    return False


def staging_root_in(chain: list, s: dict):
    """The staging-root metadata in `chain`, or None."""
    for i, m in enumerate(chain):
        if s["root_id"] and m.get("id") == s["root_id"] and "folder" in m:
            nxt = chain[i + 1] if i + 1 < len(chain) else None
            if nxt is None or _is_root(nxt):
                return m
        if s["root_name"] and m.get("name") == s["root_name"] and "folder" in m:
            nxt = chain[i + 1] if i + 1 < len(chain) else None
            if nxt is not None and _is_root(nxt):
                return m
            if nxt is None and not (m.get("parentReference") or {}).get("id"):
                return m
    return None


def in_staging_tree(item_id, meta, s: dict) -> tuple[bool, list]:
    """(True, chain) when the item sits in the staging tree (including the root)."""
    if not s["enabled"] or not (s["root_name"] or s["root_id"]):
        return False, []
    chain = walk(item_id, meta)
    if isinstance(chain, str):
        return False, []
    return (staging_root_in(chain, s) is not None), chain


def _drive_ok(drive: str, s: dict) -> bool:
    return not s["drive"] or drive == s["drive"]


def upload_allowed(drive: str, folder_id: str, meta, s: dict | None = None) -> str:
    """'' when a staging upload may go to this folder, else the reason.

    Allowed only when the switch is on, and only for a direct child of the staging
    root (the pending job folder) or a job-subfolder (plans/addenda) inside one.
    The staging root itself is not an upload target.
    """
    s = s or settings()
    no = "that folder is not on the upload allow-list — refused"
    if not s["enabled"] or not (s["root_name"] or s["root_id"]):
        return "staging uploads are switched off"
    if not _drive_ok(drive, s) or not folder_id:
        return no
    ok, chain = in_staging_tree(folder_id, meta, s)
    if not ok:
        return no
    root = staging_root_in(chain, s)
    if root is None or chain[0].get("id") == root.get("id"):
        return no
    if "folder" not in chain[0]:
        return no
    # direct child of the staging root: the pending job folder (name must not start with _)
    parent_id = (chain[0].get("parentReference") or {}).get("id")
    if parent_id == root.get("id"):
        if chain[0].get("name", "").startswith("_"):
            return no
        return ""
    # plans/addenda directly inside that pending job folder
    if (chain[0].get("name") in s["job_subfolders"] and len(chain) >= 3
            and chain[1].get("id") and (chain[1].get("parentReference") or {}).get("id") == root.get("id")
            and not chain[1].get("name", "").startswith("_") and "folder" in chain[1]):
        return ""
    return no


def mkdir_allowed(drive: str, parent_id: str, name: str, meta, s: dict | None = None) -> str:
    """'' when a folder named `name` may be created under parent_id."""
    s = s or settings()
    if not s["enabled"]:
        return "staging is switched off"
    if not _drive_ok(drive, s):
        return "that folder is not on the upload allow-list — refused"
    if name_problem(name):
        return name_problem(name)
    parent = meta(parent_id)
    if isinstance(parent, str):
        return parent
    # create the staging root at the drive root, once
    if _is_root(parent) and s["root_name"] and name == s["root_name"]:
        return ""
    ok, chain = in_staging_tree(parent_id, meta, s)
    if not ok:
        return "mkdir is only allowed inside the staging root"
    root = staging_root_in(chain, s)
    if root is None:
        return "mkdir is only allowed inside the staging root"
    # a pending job folder directly under the staging root
    if parent.get("id") == root.get("id"):
        if name.startswith("_"):
            return "a pending job folder name must not start with _"
        return ""
    # 02 / 04 inside a pending job folder
    if ((parent.get("parentReference") or {}).get("id") == root.get("id")
            and not parent.get("name", "").startswith("_") and name in s["job_subfolders"]):
        return ""
    return "mkdir is not allowed in that folder"


def delete_allowed(drive: str, item_id: str, meta, s: dict | None = None) -> str:
    """'' when this item may be deleted (it sits inside the staging root; the root itself may not)."""
    s = s or settings()
    if not s["enabled"]:
        return "staging is switched off"
    if not _drive_ok(drive, s) or not item_id:
        return "that item is not inside the staging root — refused"
    ok, chain = in_staging_tree(item_id, meta, s)
    if not ok:
        return "that item is not inside the staging root — refused"
    root = staging_root_in(chain, s)
    if root is None or chain[0].get("id") == root.get("id"):
        return "the staging root is never deleted"
    return ""


def _is_job_subfolder(drive: str, folder_id: str, meta, s: dict, names: set[str]) -> str:
    """'' when folder_id is a named subfolder of a real job folder (root or status folder)."""
    no = "that destination is not an allowed job folder — refused"
    if s["drive"] and drive != s["drive"]:
        return no
    chain = walk(folder_id, meta)
    if isinstance(chain, str):
        return chain
    if "folder" not in chain[0] or chain[0].get("name") not in names or _is_root(chain[0]):
        return no
    # subfolder -> job -> (status ->) root
    if len(chain) < 3 and not (len(chain) >= 2 and _is_root(chain[-1])):
        return no
    # find root at end
    if not _is_root(chain[-1]):
        return no
    mid = chain[1:-1]  # job, or job + status
    if not mid or any("folder" not in c for c in mid) or mid[0].get("name", "").startswith("_"):
        return no
    if len(mid) == 1:
        return ""
    if len(mid) == 2 and mid[1].get("name") in s["status_folders"]:
        return ""
    return no


def move_allowed(drive: str, item_id: str, dest_id: str, meta, s: dict | None = None) -> str:
    """'' when the item (inside staging) may be moved into dest (a job subfolder)."""
    s = s or settings()
    if not s["enabled"]:
        return "staging is switched off"
    if not _drive_ok(drive, s):
        return "that folder is not on the upload allow-list — refused"
    why = delete_allowed(drive, item_id, meta, s)  # same path rule: in tree, not the root
    if why:
        return why if "staging root is never deleted" not in why else "the staging root is never moved"
    dest_names = s["move_dest_subfolders"] or s["job_subfolders"]
    return _is_job_subfolder(drive, dest_id, meta, s, dest_names)
