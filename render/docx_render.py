#!/usr/bin/env python3
"""Editable Word copy of the proposal, from the same bid JSON as the PDF.

    python3 render/docx_render.py <bid.json> --out <dir>      # -> <dir>/<same base name as the PDF>.docx

Content comes from the same prepare() and the same HTML template as the PDF (terms clauses are read
out of the rendered HTML), so the two never drift. Layout is plain Word: real numbered lists, simple
tables, logo header, Page X of Y footer — built to be edited, then saved as PDF.
The SBE certificate is not embedded (Word can't hold the PDF page); send it as its own attachment.
"""
from __future__ import annotations

import argparse
import html as htmllib
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).parent))
import render as R  # noqa: E402

GREEN = RGBColor(0x18, 0x9E, 0x4C)
GREEN_DEEP = RGBColor(0x0F, 0x6B, 0x35)
MUTED = RGBColor(0x6B, 0x6B, 0x6B)
INK = RGBColor(0x11, 0x11, 0x11)
FONT = "Arial"
LOGO = R.TEMPLATE_DIR / "assets" / "fb-logo.png"


# ---- small helpers ------------------------------------------------------------------------------
def run(p, text, size=10, bold=False, color=INK, italic=False, caps=False):
    r = p.add_run(text)
    r.font.name, r.font.size, r.font.bold, r.font.italic = FONT, Pt(size), bold, italic
    r.font.color.rgb = color
    r.font.all_caps = caps
    return r


def para(doc_or_cell, text="", size=10, bold=False, color=INK, align=None, before=0, after=4, caps=False):
    p = doc_or_cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if align:
        p.alignment = align
    if text:
        run(p, text, size, bold, color, caps=caps)
    return p


def shade(cell_or_par, hex_fill):
    el = cell_or_par._tc.get_or_add_tcPr() if hasattr(cell_or_par, "_tc") else cell_or_par._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_fill)
    el.append(shd)


def bottom_rule(p, color="189E4C", size=18):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", str(size)), ("space", "4"), ("color", color)):
        b.set(qn(f"w:{k}"), v)
    bdr.append(b); pPr.append(bdr)


def field(p, instr, size=8, color=MUTED):
    for kind, text in (("begin", None), (None, instr), ("separate", None), (None, "1"), ("end", None)):
        r = p.add_run()
        r.font.name, r.font.size = FONT, Pt(size)
        r.font.color.rgb = color
        if kind:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), kind); r._r.append(fc)
        elif text == instr:
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "; r._r.append(it)
        else:
            r.text = text


def no_borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "nil"); b.append(e)
    tblPr.append(b)


def widths(table, inches):
    table.autofit = False
    for col, w in zip(table.columns, inches):
        col.width = Inches(w)
    for gc, w in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), inches):
        gc.set(qn("w:w"), str(int(w * 1440)))
    for row in table.rows:
        for c, w in zip(row.cells, inches):
            c.width = Inches(w)
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(1)


class Numbering:
    """Real Word numbered lists that each restart at 1 (so deleting a line renumbers itself)."""

    def __init__(self, doc):
        self.doc = doc
        self.numbering = doc.part.numbering_part.element
        self.abstract_id = self._make_abstract()

    def _make_abstract(self):
        ids = [int(a.get(qn("w:abstractNumId"))) for a in self.numbering.findall(qn("w:abstractNum"))]
        aid = max(ids + [0]) + 1
        a = OxmlElement("w:abstractNum"); a.set(qn("w:abstractNumId"), str(aid))
        lvl = OxmlElement("w:lvl"); lvl.set(qn("w:ilvl"), "0")
        for tag, val in (("w:start", "1"), ("w:numFmt", "decimal"), ("w:lvlText", "%1."), ("w:lvlJc", "left")):
            e = OxmlElement(tag); e.set(qn("w:val"), val); lvl.append(e)
        pPr = OxmlElement("w:pPr"); ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "360"); ind.set(qn("w:hanging"), "360"); pPr.append(ind); lvl.append(pPr)
        a.append(lvl)
        self.numbering.insert(0, a)
        return aid

    def new_list(self):
        ids = [int(n.get(qn("w:numId"))) for n in self.numbering.findall(qn("w:num"))]
        nid = max(ids + [0]) + 1
        n = OxmlElement("w:num"); n.set(qn("w:numId"), str(nid))
        an = OxmlElement("w:abstractNumId"); an.set(qn("w:val"), str(self.abstract_id)); n.append(an)
        ov = OxmlElement("w:lvlOverride"); ov.set(qn("w:ilvl"), "0")
        so = OxmlElement("w:startOverride"); so.set(qn("w:val"), "1"); ov.append(so); n.append(ov)
        self.numbering.append(n)
        return nid

    def item(self, nid, text, size=9.5):
        p = para(self.doc, after=1.5)
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        il = OxmlElement("w:ilvl"); il.set(qn("w:val"), "0"); numPr.append(il)
        ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(nid)); numPr.append(ni)
        pPr.append(numPr)
        run(p, text, size)
        return p


def clauses_from_html(ctx):
    """Terms & Conditions exactly as the PDF prints them."""
    html = R.render_html(ctx)
    out = []
    for label, body in re.findall(r'<p class="clause" data-label="([^"]+)">(.*?)</p>', html, flags=re.S):
        text = htmllib.unescape(re.sub(r"<[^>]+>", "", body))
        out.append((htmllib.unescape(label), re.sub(r"\s+", " ", text).strip()))
    return out


def section_head(doc, title, note):
    p = para(doc, title, size=13, bold=True, color=INK, before=10, after=1, caps=True)
    bottom_rule(p, size=12)
    if note:
        para(doc, note, size=8.5, color=MUTED, after=4)


ORDER = {
    "pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs "
           "suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd "
           "snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment "
           "textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange",
    "tblPr": "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd "
             "tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription",
    "tcPr": "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark",
}


def schema_order(root):
    """python-docx appends raw elements at the end; Word wants schema order or it reports unreadable content."""
    for tag, seq in ORDER.items():
        rank = {qn(f"w:{n}"): i for i, n in enumerate(seq.split())}
        for el in root.iter(qn(f"w:{tag}")):
            kids = list(el)
            for k in kids:
                el.remove(k)
            for k in sorted(kids, key=lambda k: rank.get(k.tag, 999)):
                el.append(k)


# ---- document -----------------------------------------------------------------------------------
def build(ctx, out_path: Path):
    b, co, sg, t = ctx["bid"], ctx["company"], ctx["signer"], ctx["terms"]
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = Inches(0.75)
    sec.top_margin, sec.bottom_margin = Inches(0.6), Inches(0.7)
    sec.header_distance = sec.footer_distance = Inches(0.35)

    # header: logo + run line; footer: forms-part note + Page X of Y
    h = sec.header
    ht = h.add_table(1, 2, Inches(7)); no_borders(ht); widths(ht, [3.2, 3.8])
    ht.cell(0, 0).paragraphs[0].add_run().add_picture(str(LOGO), width=Inches(2.3))
    hp = ht.cell(0, 1).paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(hp, f"{ctx['project']['name']} · Rev {b['rev']} · Bid {b['date']}", 8, color=MUTED)
    h.paragraphs[0].paragraph_format.space_after = Pt(0)
    rule = h.add_paragraph(); bottom_rule(rule, size=24); rule.paragraph_format.space_after = Pt(0)
    f = sec.footer.paragraphs[0]
    run(f, f"Forms part of the {co['name_caps'].title()} proposal dated {b['date']} · to be included in its entirety in any "
           f"contract for masonry work · NCTRCA SBE # {co['sbe_cert_no']}    ", 7.5, color=MUTED)
    run(f, "Page ", 7.5, color=MUTED); field(f, "PAGE", 7.5); run(f, " of ", 7.5, color=MUTED); field(f, "NUMPAGES", 7.5)

    # ---- page 1
    p = para(doc, after=2)
    run(p, f"Est. {co['est']} · {co['address']} · {co['phone']} · {sg['email']} · {co['website']}", 8.5, color=MUTED)
    p = para(doc, after=6)
    run(p, "NCTRCA CERTIFIED SBE ", 9, bold=True, color=GREEN_DEEP)
    run(p, f"· Small Business Enterprise · Cert # {co['sbe_cert_no']} · Exp. {co['sbe_cert_exp']}", 8.5, color=MUTED)

    band = doc.add_table(1, 2); no_borders(band); widths(band, [3.5, 3.5])
    for c in band.rows[0].cells:
        shade(c, "111111")
    bp = band.cell(0, 0).paragraphs[0]; run(bp, "MASONRY PROPOSAL & SCOPE OF WORK", 10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    bp2 = band.cell(0, 1).paragraphs[0]; bp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(bp2, f"Rev {b['rev']} · Bid {b['date']} · {b['time']} · Valid through {b['valid_through']}", 8.5, color=RGBColor(0xFF, 0xFF, 0xFF))

    para(doc, ctx["project"]["name"], size=20, bold=True, before=10, after=0)
    sub = ctx["project"]["address"] + (f" · Architect Proj. # {ctx['project']['arch_project_no']}" if ctx["project"].get("arch_project_no") else "")
    para(doc, sub, size=9.5, color=MUTED, after=6)

    docs = doc.add_table(3, 2); no_borders(docs); widths(docs, [1.3, 5.7])
    for i, (k, v) in enumerate((("Plans dated", b["plans_dated"]), ("Addenda", b["addenda_text"]), ("Divisions", b["divisions"]))):
        run(docs.cell(i, 0).paragraphs[0], k.upper(), 8, bold=True, color=GREEN_DEEP)
        run(docs.cell(i, 1).paragraphs[0], v, 9.5)

    para(doc, "We propose to complete the Masonry Work per the Contract Documents for", size=10, before=10, after=0)
    p = para(doc, after=0); run(p, f"$ {R.money(b['base_bid'])}", 30, bold=True)
    para(doc, f"{R.words(b['base_bid'])} Dollars", size=10, bold=True, color=GREEN_DEEP, after=4)
    if b["sales_tax_included"]:
        p = para(doc, after=6); run(p, "Texas sales tax INCLUDED.", 9.5, bold=True)
    else:
        p = para(doc, after=6)
        run(p, "Sales tax on incorporated materials EXCLUDED", 9.5, bold=True)
        run(p, " — Owner tax-exempt; materials purchased tax-free upon receipt of the Owner's Texas Sales and Use Tax "
               "Exemption Certificate. Tax on non-incorporated items (equipment, tools, consumables) INCLUDED.", 9.5)

    extras = [(f"Alternate {a['no']}" + (f" — {a['title']}" if a.get("title") else ""), R.signed_money(a["amount"])) for a in b["alternates"]]
    if b["price_bond"]:
        extras.append(("P&P Bond (on Base Bid)", f"$ {R.money(b['bond_amount'])}"))
    if extras:
        at = doc.add_table(len(extras), 2); at.style = "Table Grid"; widths(at, [5.0, 2.0])
        for i, (k, v) in enumerate(extras):
            run(at.cell(i, 0).paragraphs[0], k, 9.5, bold=True)
            vp = at.cell(i, 1).paragraphs[0]; vp.alignment = WD_ALIGN_PARAGRAPH.RIGHT; run(vp, v, 10, bold=True)

    para(doc, after=2)
    pt = doc.add_table(1, 3); no_borders(pt); widths(pt, [2.4, 2.3, 2.3])
    blocks = [
        ("Submitted to (CM/GC)", ctx["gc"]["name"], [f"Attn: {ctx['gc']['contact']}" if ctx["gc"].get("contact") else "",
                                                     ctx["gc"].get("address", ""), ctx["gc"].get("phone", ""), ctx["gc"].get("email", "")]),
        ("Owner", ctx["owner"]["name"], ctx["owner"]["address_lines"]),
        ("Architect", ctx["architect"]["name"], ctx["architect"]["address_lines"] + [ctx["architect"].get("phone", "")]),
    ]
    for c, (lbl, name, lines) in zip(pt.rows[0].cells, blocks):
        run(c.paragraphs[0], lbl.upper(), 8, bold=True, color=GREEN_DEEP)
        para(c, name, size=10, bold=True, after=0)
        for ln in lines:
            if ln:
                para(c, ln, size=9, after=0)

    p = para(doc, before=10, after=2); run(p, "KEY TERMS — SUMMARY  ", 8, bold=True, color=GREEN_DEEP)
    run(p, "Full Terms & Conditions on page 2 govern and form part of this proposal.", 8, color=MUTED)
    bottom_rule(p, color="CFCFCF", size=6)
    kt = [
        ("Mobilization", f"One (1) included. Each additional, through no fault of FB: ${R.money(t['remob_fee'])} change order."),
        ("Payment", "Monthly by the 15th. Public work: Tex. Gov't Code §2251.022–.023, 12% interest (§2251.025). Private work: "
                    "Tex. Prop. Code Ch. 28, 1.5%/month (§28.004). Final within 45 days."),
        ("Retainage", "Only if specified in bid documents or agreed in contract; not more than 5%."),
        ("Validity", f"Void if not accepted by subcontract or LOI by {b['valid_through']} (30 days from bid date)."),
        ("Equipment", "FB equipment and scaffold for FB personnel only; use by others requires signed agreement and COI."),
        ("Warranty", "One (1) year, labor and materials. No overtime included. Breakouts for accounting only."),
    ]
    tt = doc.add_table(len(kt), 2); no_borders(tt); widths(tt, [1.1, 5.9])
    for i, (k, v) in enumerate(kt):
        run(tt.cell(i, 0).paragraphs[0], k, 9, bold=True)
        run(tt.cell(i, 1).paragraphs[0], v, 9)

    para(doc, sg["name"], size=11, bold=True, before=14, after=0)
    para(doc, f"{sg['title']}, {co['legal_name']}", size=9, after=0)
    para(doc, f"{sg['cell']} · {sg['email']}", size=9, after=4)
    para(doc, "This proposal is to be included in its entirety in any contract for masonry work.", size=8, color=MUTED, after=0)
    if b["supersedes"] is not None:
        para(doc, f"Rev {b['rev']} supersedes Rev {b['supersedes']} in its entirety.", size=8, color=MUTED, after=0)

    # ---- page 2: terms
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    p = para(doc, "Terms & Conditions", size=16, bold=True, after=6); bottom_rule(p, size=12)
    for label, text in clauses_from_html(ctx):
        p = para(doc, after=5); p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run(p, label.upper() + ". ", 8, bold=True, color=GREEN_DEEP); run(p, text, 8.5)
    para(doc, "We are grateful for the opportunity to submit our quotation to you. Should you need further information, please, contact us.", size=9, after=10)
    p = para(doc, "Sincerely,", size=9.5, after=8); p.paragraph_format.keep_with_next = True
    p = para(doc, sg["name"], size=11, bold=True, after=0); p.paragraph_format.keep_with_next = True
    para(doc, f"{sg['title']}, {co['legal_name']} · {co['phone']} · {sg['email']}", size=9, after=0)

    # ---- page 3+: scope
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    num = Numbering(doc)
    s = ctx["scope"]
    section_head(doc, "Masonry Inclusions", "Furnished and installed by Fenimore-Blythe." + (" All applicable taxes & insurance included." if b["sales_tax_included"] else ""))
    nid = num.new_list()
    for it in s["inclusions"]:
        num.item(nid, it)
    if b["alternates"]:
        section_head(doc, "Alternates", "Priced on page 1. Each alternate modifies the Base Bid scope only as stated below; all other "
                                        "Inclusions, Exclusions and Qualifications remain in force.")
        for a in b["alternates"]:
            p = para(doc, after=1); run(p, f"ALT {a['no']}  ", 9.5, bold=True, color=GREEN_DEEP); run(p, a.get("title", ""), 9.5, bold=True)
            for k in ("add", "delete", "qualify"):
                if a.get(k):
                    p = para(doc, after=1); p.paragraph_format.left_indent = Inches(0.3)
                    run(p, f"{k.title()}: ", 9, bold=True); run(p, a[k], 9)
    section_head(doc, "Masonry Exclusions", "Our proposal DOES NOT INCLUDE the following items and/or activities.")
    nid = num.new_list()
    for it in s["exclusions"]:
        num.item(nid, it)
    section_head(doc, "Qualifications", "Based on our Company being PROVIDED by others the following services and/or items at no cost to us.")
    nid = num.new_list()
    for it in s["qualifications"]:
        num.item(nid, it)

    doc.core_properties.title = f"{ctx['project']['name']} — Masonry Proposal — Rev {b['rev']}"
    doc.core_properties.author = sg["name"]
    for part in [doc.part] + [sec.header.part, sec.footer.part]:
        schema_order(part.element)
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None:
        zoom.set(qn("w:percent"), "100")
    doc.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("bid_json"); ap.add_argument("--out", default=None)
    a = ap.parse_args()
    ctx = R.prepare(json.loads(Path(a.bid_json).read_text(encoding="utf-8")))
    ctx.update(page_total="?", scope_pages=1)
    out_dir = Path(a.out) if a.out else R.OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    name = f"FB_Proposal_{R.slug(ctx['project']['name'])}_{ctx['bid_date_iso']}_R{ctx['bid']['rev']}"
    print(build(ctx, out_dir / f"{name}.docx"))


if __name__ == "__main__":
    main()
