#!/usr/bin/env python3
"""Render a Fenimore-Blythe masonry proposal PDF from a bid JSON.

    python render/render.py samples/unt-jazz.json            # -> out/FB_Proposal_<slug>_<date>_R<n>.pdf
    python render/render.py samples/unt-jazz.json --html     # also keep the rendered HTML for inspection

Pipeline: JSON -> derived fields -> Jinja2 -> HTML -> Chromium (Playwright) + Paged.js -> PDF
          -> (optional) append NCTRCA SBE certificate PDF.
Two render passes: pass 1 learns the page count, pass 2 prints "Page X of N".
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, StrictUndefined
from playwright.sync_api import sync_playwright
from pypdf import PdfReader, PdfWriter

sys.path.insert(0, str(Path(__file__).parent))
from bond import bond_add  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = ROOT / "template"
OUT_DIR = ROOT / "out"

# ---- company constants (change here, not in the template) -----------------------------------
COMPANY = {
    "name_caps": "FENIMORE-BLYTHE",
    "short": "FB",
    "legal_name": "Fenimore-Blythe Commercial Masonry, LLC",
    "est": "1945",
    "address": "2300 Higgins Lane, Fort Worth, TX 76111",
    "phone": "(817) 834-6301",
    "website": "fenimoreblythe.com",
    "sbe_cert_no": "WMSB62323N1127",
    "sbe_cert_exp": "11/30/2027",
    "sbe_cert_exp_iso": "2027-11-30",
    "sbe_cert_pdf": TEMPLATE_DIR / "assets" / "nctrca_sbe_cert.pdf",
}
SIGNER_DEFAULT = {
    "name": "David Grzesiak",
    "title": "President",
    "cell": "(262) 391-7319",
    "email": "davidg@fenimoreblythe.com",
}
TERMS_DEFAULT = {"remob_fee": 3500, "validity_days": 30}
CERT_WARN_DAYS = 90


# ---- number formatting ------------------------------------------------------------------------
ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
        "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
SCALE = ["", "Thousand", "Million", "Billion"]


def _chunk(n: int) -> str:
    out = []
    if n >= 100:
        out.append(f"{ONES[n // 100]} Hundred")
        n %= 100
    if n >= 20:
        out.append(TENS[n // 10] + (f"-{ONES[n % 10]}" if n % 10 else ""))
    elif n:
        out.append(ONES[n])
    return " ".join(out)


def words(n) -> str:
    """Whole-dollar amount in words, title case, no cents: 243643 -> 'Two Hundred Forty-Three Thousand Six Hundred Forty-Three'."""
    n = int(round(float(n)))
    if n == 0:
        return "Zero"
    parts, i = [], 0
    while n:
        n, r = divmod(n, 1000)
        if r:
            parts.append((_chunk(r) + (" " + SCALE[i] if SCALE[i] else "")).strip())
        i += 1
    return " ".join(reversed(parts))


def money(n) -> str:
    return f"{int(round(float(n))):,}"


def signed_money(n) -> str:
    n = int(round(float(n)))
    return f"− $ {abs(n):,}" if n < 0 else f"$ {n:,}"


# ---- derived fields -------------------------------------------------------------------------------
def mdy(d: date) -> str:
    return f"{d.month}/{d.day}/{d.year}"


def parse_date(s: str) -> date:
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unrecognized date: {s!r} (use YYYY-MM-DD)")


def prepare(raw: dict) -> dict:
    d = json.loads(json.dumps(raw))  # deep copy
    bid = d["bid"]
    terms = {**TERMS_DEFAULT, **d.get("terms", {})}
    bid_date = parse_date(bid["date"])
    bid["date"] = mdy(bid_date)
    bid["plans_dated"] = mdy(parse_date(bid["plans_dated"])) if bid.get("plans_dated") else "—"
    bid["valid_through"] = mdy(bid_date + timedelta(days=terms["validity_days"]))
    bid["rev"] = int(bid.get("rev", 0))
    bid["supersedes"] = bid["rev"] - 1 if bid["rev"] > 0 else None
    bid["sales_tax_included"] = bool(bid.get("sales_tax_included", True))
    bid["attach_cert"] = bool(bid.get("attach_cert", False))
    bid["price_bond"] = bool(bid.get("price_bond", False))  # P&P bond shown only when the GC explicitly asks
    bid["bond_amount"] = bond_add(bid["base_bid"])
    addenda = bid.get("addenda") or []
    bid["addenda_text"] = ", ".join(str(a) for a in addenda) if addenda else "None received"
    bid["alternates"] = [{"no": i + 1, **a} for i, a in enumerate(bid.get("alternates") or [])]
    # Optional breakout of the Base Bid by building/area, shown on page 1 ("included in Base Bid"); must add up.
    bid["breakout"] = [{"label": str(b["label"]), "amount": int(round(float(b["amount"])))} for b in (bid.get("breakout") or [])]
    if bid["breakout"] and sum(b["amount"] for b in bid["breakout"]) != int(bid["base_bid"]):
        raise SystemExit(f"breakout adds to {sum(b['amount'] for b in bid['breakout']):,}, not the base bid {int(bid['base_bid']):,}")
    for party in ("owner", "architect"):
        p = d[party]
        p["address_lines"] = p.get("address_lines") or [ln.strip() for ln in re.split(r"\s*\|\s*|\n", p.get("address", "")) if ln.strip()]
    d["terms"] = terms
    d["company"] = COMPANY
    d["signer"] = {**SIGNER_DEFAULT, **d.get("signer", {})}
    d["bid_date_iso"] = bid_date.isoformat()
    return d


def check_cert(attach: bool) -> None:
    exp = date.fromisoformat(COMPANY["sbe_cert_exp_iso"])
    today = date.today()
    if attach and today > exp:
        sys.exit(f"REFUSED: NCTRCA SBE certificate expired {exp}. Renew before attaching.")
    if today > exp - timedelta(days=CERT_WARN_DAYS):
        print(f"WARNING: NCTRCA SBE certificate expires {exp} ({(exp - today).days} days). Start renewal.", file=sys.stderr)


def slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-")


# ---- render ---------------------------------------------------------------------------------------
def _env(autoescape: bool) -> Environment:
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), undefined=StrictUndefined, autoescape=autoescape)
    env.filters.update(money=money, words=words, signed_money=signed_money)
    return env


def render_html(ctx: dict) -> str:
    """Full HTML with the CSS inlined (Paged.js XHRs <link> stylesheets, which file:// blocks).
    The CSS is itself a Jinja template so page totals and dates can land in @page margin boxes."""
    html = _env(True).get_template("proposal.html.j2").render(**ctx)
    css = _env(False).get_template("proposal.css").render(**ctx)
    return html.replace('<link rel="stylesheet" href="proposal.css">', f"<style>\n{css}\n</style>")


PAGED_HOOK = """
<script>
// Page 1 is a fixed sheet. Long titles, GC blocks or the exempt-tax statement can push the signature off the
// bottom; before pagination, tighten spacing in up to three steps until the cover fits.
window.PagedConfig = {
  before: async () => {
    if (document.fonts && document.fonts.ready) { await document.fonts.ready; }
    const sheet = document.querySelector('.sheet.cover');
    for (let lvl = 1; sheet && lvl <= 3 && sheet.scrollHeight > sheet.clientHeight + 1; lvl++) { sheet.classList.add('fit' + lvl); }
    window.__coverFit = sheet ? (sheet.scrollHeight > sheet.clientHeight + 1 ? 'overflow' : sheet.className) : 'none';
  },
  after: (flow) => { window.__pagedTotal = flow.total; }
};
</script>
"""


def to_pdf(html: str, out_pdf: Path) -> int:
    """Write HTML to disk beside the template (so relative asset URLs resolve), print to PDF, return page count."""
    tmp_html = TEMPLATE_DIR / "_render.html"
    html = html.replace('<script src="vendor/paged.polyfill.js"></script>', PAGED_HOOK + '<script src="vendor/paged.polyfill.js"></script>')
    tmp_html.write_text(html, encoding="utf-8")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(tmp_html.resolve().as_uri(), wait_until="load")
        page.wait_for_function("window.__pagedTotal !== undefined", timeout=60_000)
        total = page.evaluate("window.__pagedTotal")
        fit = page.evaluate("window.__coverFit")
        if fit == "overflow":
            print("⚠ page 1 still overflows after tightening — check the signature block", file=sys.stderr)
        page.pdf(path=str(out_pdf), prefer_css_page_size=True, print_background=True)
        browser.close()
    tmp_html.unlink(missing_ok=True)
    return int(total)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("bid_json")
    ap.add_argument("--html", action="store_true", help="keep rendered HTML next to the PDF")
    ap.add_argument("--out", default=None, help="output directory (default: out/)")
    args = ap.parse_args()

    ctx = prepare(json.loads(Path(args.bid_json).read_text(encoding="utf-8")))
    check_cert(ctx["bid"]["attach_cert"])
    out_dir = Path(args.out) if args.out else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    name = f"FB_Proposal_{slug(ctx['project']['name'])}_{ctx['bid_date_iso']}_R{ctx['bid']['rev']}"
    body_pdf = out_dir / f"{name}.body.pdf"

    # pass 1: learn the page count
    ctx.update(page_total="?", scope_pages=1)
    pages = to_pdf(render_html(ctx), body_pdf)
    extra = 1 if ctx["bid"]["attach_cert"] else 0
    # pass 2: real totals
    ctx.update(page_total=pages + extra, scope_pages=max(pages - 2, 1))
    html = render_html(ctx)
    pages2 = to_pdf(html, body_pdf)
    assert pages2 == pages, "page count changed between passes"

    final_pdf = out_dir / f"{name}.pdf"
    if extra:
        w = PdfWriter()
        for pg in PdfReader(str(body_pdf)).pages:
            w.add_page(pg)
        for pg in PdfReader(str(COMPANY["sbe_cert_pdf"])).pages:
            w.add_page(pg)
        with final_pdf.open("wb") as f:
            w.write(f)
        body_pdf.unlink()
    else:
        body_pdf.rename(final_pdf)
    if args.html:
        (out_dir / f"{name}.html").write_text(html, encoding="utf-8")
    print(f"{final_pdf}  ({pages + extra} pages; bond add ${ctx['bid']['bond_amount']:,}; valid through {ctx['bid']['valid_through']})")


if __name__ == "__main__":
    main()
