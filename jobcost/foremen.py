"""Foreman roster for JC Detail role classification.

The public print kit never ships a list of employee names. The courier (or a local
operator) passes a roster file at runtime with --foremen.

Accepted name forms, matched case-insensitively:
  * Vista payroll:  "Last, First" or "Last, First M." (suffixes like Jr / II ignored)
  * Everyday form:  "First Last"  (as SharePoint / M365 display names often appear)

A hyphenated payroll last name matches a shorter everyday last name on the first
part ("Cole-Hayes, Morgan A." matches "Morgan Cole").
"""
from __future__ import annotations

import json
import re
from pathlib import Path

_SUFFIX = re.compile(r"\b(jr|sr|ii|iii|iv|v)\b\.?", re.I)
_INITIAL = re.compile(r"\b([a-z])\b\.?", re.I)
_PUNCT = re.compile(r"[^a-z0-9\s-]+")


def _tokens(s: str) -> list[str]:
    s = _SUFFIX.sub(" ", (s or "").lower())
    s = _PUNCT.sub(" ", s)
    s = _INITIAL.sub(" ", s)
    return [t for t in s.split() if t]


def name_parts(raw: str) -> tuple[str, str]:
    """Return (first, last) where last may still contain a hyphen."""
    s = (raw or "").strip()
    if not s:
        return "", ""
    if "," in s:
        last, _, rest = s.partition(",")
        toks = _tokens(rest)
        first = toks[0] if toks else ""
        last_toks = _tokens(last)
        last = "-".join(last_toks)
        return first, last
    toks = _tokens(s)
    if not toks:
        return "", ""
    if len(toks) == 1:
        return "", toks[0]
    return toks[0], toks[-1]


def names_match(a: str, b: str) -> bool:
    fa, la = name_parts(a)
    fb, lb = name_parts(b)
    if not fa or not fb or not la or not lb:
        return False
    if fa != fb:
        return False
    if la == lb:
        return True
    # hyphenated last name: "cole-hayes" matches "cole"
    a_parts, b_parts = la.split("-"), lb.split("-")
    return a_parts[0] == b_parts[0] and (la.startswith(lb + "-") or lb.startswith(la + "-") or
                                         a_parts[0] == la or b_parts[0] == lb)


def is_foreman(payroll_name: str, roster: list[str] | set[str] | None) -> bool:
    if not payroll_name or not roster:
        return False
    return any(names_match(payroll_name, r) for r in roster if r and r.strip())


def load_roster(path: str | Path) -> list[str]:
    """Load names from a text file (one per line, # comments ok) or a JSON list."""
    p = Path(path)
    text = p.read_text(encoding="utf-8")
    if p.suffix.lower() == ".json":
        data = json.loads(text)
        if not isinstance(data, list):
            raise ValueError(f"{p}: roster JSON must be a list of names")
        return [str(x).strip() for x in data if str(x).strip()]
    out = []
    for ln in text.splitlines():
        s = ln.split("#", 1)[0].strip()
        if s:
            out.append(s)
    return out
