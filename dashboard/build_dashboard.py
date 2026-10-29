"""Build the preconstruction Dashboard workbook from the bid CSVs. Data tabs hold values; every summary is a live formula.

    python3 dashboard/build_dashboard.py --bid-log "Bid Log 2026.csv" --bid-metrics "Bid Metrics 2026.csv" \
        [--prices "Price Book Latest.csv"] [--job-cost "Job Cost History.csv"] [--out Dashboard.xlsx] [--no-recalc]

Price Book Latest.csv comes from `pricebook/price_book.py export <folder> <out.csv>`.
After building, LibreOffice recalculates the file (so the stored values are current for anything that reads them),
then the script checks every formula for errors and re-checks the headline numbers against plain Python.
Exit 1 = do not upload.

Headline: labor $/SF by unit type (Crew Page base wage, before burden and OH&P). Hit rate = won / (won + lost);
submitted-awaiting and closed are shown beside it, never in it.
"""
import argparse, csv, datetime as dt, os, re, shutil, statistics, subprocess, sys, tempfile
from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

UNIT_TYPES = [("cmu_gray", "Gray CMU"), ("cmu_arch", "Architectural CMU"), ("brick", "Brick"),
              ("stone", "Stone"), ("cast_stone", "Cast stone")]
STD_WAGES = "Bricklayer $30 · Foreman $65 · Laborer $20 · Operator $27 /hr"
PERIODS = ["All", "YTD", "Trailing 12 months"]
GC_SUFFIX = {"llc", "inc", "lp", "i", "ltd", "co", "corp", "company", "construction", "constructors", "contractors",
             "contracting", "general", "group", "builders", "national", "services", "the"}
# Different companies that the suffix rule would otherwise merge, or one company spelled two ways
GC_KEEP_APART = {"arco murray", "rudick joeris", "evans consulting"}

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True)
F_SECTION = Font(name=FONT, size=12, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="666666")
F_INPUT = Font(name=FONT, size=10, bold=True, color="0000FF")
FILL_HEAD = PatternFill("solid", fgColor="1F2937")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_BAND = PatternFill("solid", fgColor="F3F4F6")
THIN = Border(bottom=Side(style="thin", color="D1D5DB"))
MONEY = '$#,##0;($#,##0);"-"'
RATE = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
INT = '#,##0;(#,##0);"-"'
DATE = "yyyy-mm-dd"


# ---------- input ----------
def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def num(v):
    try:
        return float(v) if str(v).strip() != "" else None
    except ValueError:
        return None


def day(v):
    try:
        return dt.date.fromisoformat(str(v).strip()[:10])
    except ValueError:
        return None


def gc_key(name):
    s = re.sub(r"\(.*?\)", " ", name or "").lower().replace("&", " and ")
    toks = [t for t in re.split(r"[^a-z0-9]+", s) if t]
    core = " ".join(t for t in toks if t not in GC_SUFFIX) or " ".join(toks)
    for k in GC_KEEP_APART:
        if core.startswith(k):
            return k
    return core


def gc_groups(rows):
    """raw GC name -> display group. Display = the most common spelling in the group (ties: shortest)."""
    by = defaultdict(Counter)
    for r in rows:
        g = (r["gc"] or "").strip()
        if not g or g.lower() == "unknown":
            continue
        by[gc_key(g)][g] += 1
    raw2grp, merges = {}, []
    for k, c in by.items():
        clean = Counter()
        for raw, n in c.items():
            clean[re.sub(r"\s*\(.*?\)\s*", " ", raw).strip()] += n
        disp = sorted(clean.items(), key=lambda t: (-t[1], len(t[0])))[0][0]
        for raw in c:
            raw2grp[raw] = disp
        if len(c) > 1 or disp != next(iter(c)):
            merges.append((disp, sorted(c)))
    return raw2grp, sorted(merges)


# ---------- helpers ----------
def head(ws, row, col, labels, widths=None):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font, c.fill = F_HEAD, FILL_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[L(col + i)].width = w


def put(ws, ref, value, fmt=None, font=F_BASE):
    ws[ref] = value
    ws[ref].font = font
    if fmt:
        ws[ref].number_format = fmt
    return ws[ref]


def arr(ws, ref, formula, fmt=None):
    ws[ref] = ArrayFormula(ref, formula)
    ws[ref].font = F_BASE
    if fmt:
        ws[ref].number_format = fmt


def style_body(ws, r0, r1, c0, c1):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(r, c)
            if cell.font != F_BOLD:
                cell.font = F_BASE
            cell.border = THIN


# ---------- build ----------
def build(a):
    bl = read_csv(a.bid_log)
    bm = read_csv(a.bid_metrics)
    prices = read_csv(a.prices) if a.prices and os.path.exists(a.prices) else []
    jc = read_csv(a.job_cost) if a.job_cost and os.path.exists(a.job_cost) else []
    raw2grp, merges = gc_groups(bl)
    warn = []

    wb = Workbook()
    dash = wb.active; dash.title = "Dashboard"
    bygc = wb.create_sheet("By GC")
    lh = wb.create_sheet("Labor History")
    pr = wb.create_sheet("Prices")
    jcs = wb.create_sheet("Job Cost")
    bls = wb.create_sheet("Bid Log")
    bms = wb.create_sheet("Bid Metrics")
    about = wb.create_sheet("About")

    # --- Bid Log data tab (values + one live in_period flag) ---
    cols = ["bid_id", "project", "city", "job_type", "gc", "gc_group", "bid_date", "status", "submitted_amount",
            "base_bid", "alternates_total", "tax", "submitted_date", "winner", "winning_amount", "in_period"]
    head(bls, 1, 1, cols, [26, 40, 16, 12, 30, 26, 11, 11, 13, 13, 12, 14, 11, 20, 13, 9])
    n_bl = len(bl)
    for i, r in enumerate(bl, start=2):
        gc = (r["gc"] or "").strip()
        vals = [r["bid_id"], r["project"], r["city"], r["job_type"] or "unknown", gc,
                raw2grp.get(gc, "(unknown)"), day(r["bid_date"]), r["status"], num(r["submitted_amount"]),
                num(r["base_bid"]), num(r["alternates_total"]), r["tax"], day(r["submitted_date"]), r["winner"],
                num(r["winning_amount"])]
        for j, v in enumerate(vals, start=1):
            c = bls.cell(i, j, v)
            if isinstance(v, dt.date):
                c.number_format = DATE
            elif isinstance(v, float):
                c.number_format = MONEY
        bls.cell(i, 16, f'=IF(G{i}="",IF(Dashboard!$C$4="All",1,0),IF(G{i}>=Dashboard!$C$5,1,0))')
    bls.freeze_panes = "C2"
    bls.auto_filter.ref = f"A1:P{n_bl + 1}"
    last = n_bl + 1
    BL = lambda col: f"'Bid Log'!${col}$2:${col}${last}"
    ST, AMT, INP, JT, GG, BD = BL("H"), BL("I"), BL("P"), BL("D"), BL("F"), BL("G")

    # --- Bid Metrics raw tab ---
    mcols = list(bm[0].keys()) if bm else []
    head(bms, 1, 1, mcols)
    for i, r in enumerate(bm, start=2):
        for j, k in enumerate(mcols, start=1):
            v = r[k]
            nv = num(v) if k not in ("bid_id", "project", "job_number", "source", "wages", "reconciles",
                                     "reconcile_notes", "unmapped_items", "notes", "matches_bid_log") else None
            bms.cell(i, j, nv if nv is not None else (v if v != "" else None))
    bms.freeze_panes = "C2"

    # --- Labor History (one row per estimate; linked to the Bid Log by bid_id, else by bid price) ---
    lcols = ["Project", "Job #", "Job type", "Bid date", "Bid price", "In Bid Log", "Wages", "Wages vs std", "Notes"]
    per = [("sf", "SF", INT), ("labor_per_sf", "labor $/SF", RATE), ("sf_per_manday", "SF/man-day", "#,##0.0"),
           ("material_per_sf", "material $/SF", RATE), ("sell_per_sf", "sell $/SF", RATE)]
    for u, name in UNIT_TYPES:
        lcols += [f"{name} {p[1]}" for p in per]
    head(lh, 1, 1, lcols, [24, 8, 12, 11, 12, 8, 34, 9, 24] + [10] * (len(per) * len(UNIT_TYPES)))
    lh.row_dimensions[1].height = 42
    by_id = {r["bid_id"]: r for r in bl if r["bid_id"]}
    priced = [(num(r["submitted_amount"]) or num(r["base_bid"]), r) for r in bl]
    priced = [(p, r) for p, r in priced if p]

    def by_price(price, name):
        """Bid Log rows within 0.05% of the estimate's bid price (Tradesmen rounds); best name overlap first."""
        toks = set(re.findall(r"[a-z0-9]+", name.lower()))
        hits = [r for p, r in priced if abs(p - price) <= max(5, price * 0.0005)]
        hits.sort(key=lambda r: -len(toks & set(re.findall(r"[a-z0-9]+", r["project"].lower()))))
        return hits[0] if hits else None
    unit_col = {}
    for i, r in enumerate(bm, start=2):
        link = by_id.get(r.get("bid_id", ""))
        if not link and num(r["bid_price"]) is not None:
            link = by_price(num(r["bid_price"]), r["project"])
        how = "y" if link else "n"
        if not link:  # iterations: link by name only when every name match is the same project
            toks = [t for t in re.findall(r"[a-z0-9]+", r["project"].lower()) if len(t) >= 4]
            cands = [x for x in bl if toks and all(t in x["project"].lower() for t in toks)]
            if cands and len({x["project"] for x in cands}) == 1:
                link, how = cands[0], "name"
        if not link:
            warn.append(f"Labor History: {r['project']} not linked to a Bid Log row (job type/date blank)"
                        + (" — iteration" if "iteration" in r.get("notes", "") else ""))
        vals = [r["project"], r["job_number"], (link or {}).get("job_type") or None, day((link or {}).get("bid_date", "")),
                num(r["bid_price"]), how, r["wages"], num(r["wages_vs_standard"]), r["notes"] or None]
        fmts = [None, None, None, DATE, MONEY, None, None, PCT, None]
        for u, _ in UNIT_TYPES:
            sf = num(r.get(f"{u}_sf"))
            for p in per:
                v = num(r.get(f"{u}_{p[0]}"))
                vals.append(v if sf else None)
                fmts.append(p[2])
        for j, (v, f) in enumerate(zip(vals, fmts), start=1):
            c = lh.cell(i, j, v)
            if f:
                c.number_format = f
    for k, (u, _) in enumerate(UNIT_TYPES):
        base = 10 + k * len(per)
        unit_col[u] = {p[0]: L(base + n) for n, p in enumerate(per)}
    n_lh = max(len(bm), 1)
    lh.freeze_panes = "B2"
    lh.auto_filter.ref = f"A1:{L(len(lcols))}{n_lh + 1}"
    LH = lambda col: f"'Labor History'!${col}$2:${col}${n_lh + 1}"
    style_body(lh, 2, n_lh + 1, 1, len(lcols))

    # --- Prices tab (values from Price Book Latest.csv) ---
    pcols = [("item", "Item", 22), ("supplier", "Supplier", 26), ("product", "Product", 36), ("size_color", "Size / color", 20),
             ("unit", "Unit", 6), ("unit_price", "Unit price", 11), ("price_per_sf", "$/SF", 8), ("quote_date", "Quote date", 11),
             ("project", "Project", 28), ("hold_until", "Price hold", 11), ("prior_price", "Prior price", 11),
             ("prior_date", "Prior date", 11), ("change_pct", "Change", 8), ("mover", "Mover", 6), ("notes", "Notes", 30)]
    head(pr, 1, 1, [p[1] for p in pcols], [p[2] for p in pcols])
    pfmt = {"unit_price": RATE, "price_per_sf": RATE, "prior_price": RATE, "change_pct": PCT,
            "quote_date": DATE, "prior_date": DATE, "hold_until": DATE}
    for i, r in enumerate(prices, start=2):
        for j, (k, _, _) in enumerate(pcols, start=1):
            v = r.get(k, "")
            if k in ("unit_price", "price_per_sf", "prior_price", "change_pct"):
                v = num(v)
            elif k in ("quote_date", "prior_date", "hold_until"):
                v = day(v) or (v or None)
            else:
                v = v or None
            c = pr.cell(i, j, v)
            if k in pfmt:
                c.number_format = pfmt[k]
    n_pr = len(prices)
    pr.freeze_panes = "C2"
    pr.auto_filter.ref = f"A1:{L(len(pcols))}{max(n_pr, 1) + 1}"
    if not prices:
        warn.append("Prices: no Price Book Latest.csv given — Prices tab empty")

    # --- Job Cost tab (values from Job Cost History.csv) ---
    jcols = ["job_number", "estimate", "scope", "through_date", "status", "est_sf", "est_mandays", "act_mandays",
             "mandays_var_pct", "est_labor_base", "act_labor_base", "labor_var_pct", "est_labor_per_sf", "act_labor_per_sf",
             "est_sf_per_manday", "act_sf_per_manday", "est_material", "act_material", "material_var_pct", "labor_basis",
             "hours_by_class", "notes"]
    head(jcs, 1, 1, jcols, [10, 22, 10, 11, 8] + [11] * 15 + [30, 30])
    jcs.row_dimensions[1].height = 30
    for i, r in enumerate(jc, start=2):
        for j, k in enumerate(jcols, start=1):
            v = r.get(k, "")
            if k == "through_date":
                v = day(v)
            elif k not in ("job_number", "estimate", "scope", "status", "labor_basis", "hours_by_class", "notes"):
                v = num(v)
            c = jcs.cell(i, j, v if v != "" else None)
            c.font = F_BASE
            c.number_format = (DATE if k == "through_date" else PCT if k.endswith("var_pct")
                               else RATE if k.endswith("per_sf") else MONEY if k.startswith(("est_l", "act_l", "est_m", "act_m")) and "manday" not in k
                               else "#,##0.0" if "manday" in k else "General")
    n_jc = max(len(jc), 1)
    jcs.freeze_panes = "C2"
    jcs.auto_filter.ref = f"A1:{L(len(jcols))}{n_jc + 1}"
    JC = lambda col: f"'Job Cost'!${col}$2:${col}${n_jc + 1}"

    # --- Dashboard ---
    ws = dash
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJ", [22, 13, 13, 13, 30, 13, 30, 11, 12, 13]):
        ws.column_dimensions[col].width = w
    put(ws, "A1", "Preconstruction Dashboard", font=F_TITLE)
    put(ws, "A2", f"Rebuilt {a.stamp} from Bid Log 2026.csv ({n_bl} rows), Bid Metrics 2026.csv ({len(bm)} estimates)"
        + (f", Price Book Latest.csv ({n_pr} prices)" if prices else "") + ". Don't edit — it is rebuilt nightly.", font=F_NOTE)
    put(ws, "A4", "Period (pick)", font=F_BOLD)
    c = put(ws, "C4", a.period, font=F_INPUT); c.fill = FILL_INPUT
    dv = DataValidation(type="list", formula1='"' + ",".join(PERIODS) + '"', allow_blank=False)
    ws.add_data_validation(dv); dv.add("C4")
    put(ws, "A5", "From"); put(ws, "C5", '=IF(C4="YTD",DATE(YEAR(TODAY()),1,1),IF(C4="Trailing 12 months",EDATE(TODAY(),-12),DATE(2000,1,1)))', DATE)
    put(ws, "E4", "Period drives hit rates (by bid date, from the date above onward). Labor $/SF uses every estimate.", font=F_NOTE)

    # Labor $/SF block
    r0 = 8
    put(ws, f"A{r0}", "Labor $/SF by unit type — Crew Page base wage, before burden and OH&P", font=F_SECTION)
    head(ws, r0 + 1, 1, ["Unit type", "Median $/SF", "Low $/SF", "Low job", "", "High $/SF", "High job", "Estimates",
                         "Total SF", "Median SF/man-day"])
    ws.merge_cells(start_row=r0 + 1, start_column=4, end_row=r0 + 1, end_column=5)
    proj, notes = LH("A"), LH("I")
    for k, (u, name) in enumerate(UNIT_TYPES):
        r = r0 + 2 + k
        lp, sf, spm = LH(unit_col[u]["labor_per_sf"]), LH(unit_col[u]["sf"]), LH(unit_col[u]["sf_per_manday"])
        put(ws, f"A{r}", name, font=F_BOLD)
        arr(ws, f"B{r}", f'=IF(COUNT({lp})=0,"",MEDIAN(IF({lp}<>"",{lp})))', RATE)
        put(ws, f"C{r}", f'=IF(COUNT({lp})=0,"",_xlfn.MINIFS({lp},{lp},">0"))', RATE)
        job = lambda ref: (f'=IF({ref}="","",INDEX({proj},MATCH({ref},{lp},0))&IF(INDEX({notes},MATCH({ref},{lp},0))="",""," ("&INDEX({notes},MATCH({ref},{lp},0))&")"))')
        put(ws, f"D{r}", job(f"C{r}"))
        ws.merge_cells(f"D{r}:E{r}")
        put(ws, f"F{r}", f'=IF(COUNT({lp})=0,"",_xlfn.MAXIFS({lp},{lp},">0"))', RATE)
        put(ws, f"G{r}", job(f"F{r}"))
        put(ws, f"H{r}", f"=COUNT({lp})", INT)
        put(ws, f"I{r}", f"=SUM({sf})", INT)
        arr(ws, f"J{r}", f'=IF(COUNT({spm})=0,"",MEDIAN(IF({spm}<>"",{spm})))', "#,##0.0")
        for col in "ABCDEFGHIJ":
            ws[f"{col}{r}"].border = THIN
    rl = r0 + 2 + len(UNIT_TYPES)
    put(ws, f"A{rl}", f'="Estimates with wages above standard ({STD_WAGES}): "&COUNTIF({LH("H")},">0.01")&" of "&COUNTA({proj})&". '
        f'Iterations count toward rates; a job is flagged when labor $/SF is >20% outside this range."', font=F_NOTE)

    chart = BarChart()
    chart.type, chart.style = "bar", 10
    chart.title, chart.y_axis.title = "Median labor $/SF", None
    chart.add_data(Reference(ws, min_col=2, min_row=r0 + 1, max_row=r0 + 1 + len(UNIT_TYPES)), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=r0 + 2, max_row=r0 + 1 + len(UNIT_TYPES)))
    chart.legend = None
    chart.height, chart.width = 6.5, 12
    ws.add_chart(chart, "L8")

    # Hit rate block
    def hit_row(r, label, crit=""):
        put(ws, f"A{r}", label, font=F_BOLD if not crit else F_BASE)
        cnt = lambda s: f'COUNTIFS({ST},"{s}",{INP},1{crit})'
        sm = lambda s: f'SUMIFS({AMT},{ST},"{s}",{INP},1{crit})'
        put(ws, f"B{r}", f"={cnt('submitted')}+{cnt('awarded')}+{cnt('lost')}+{cnt('closed')}", INT)
        put(ws, f"C{r}", f"={cnt('awarded')}", INT)
        put(ws, f"D{r}", f"={cnt('lost')}", INT)
        put(ws, f"E{r}", f'=IF(C{r}+D{r}=0,"",C{r}/(C{r}+D{r}))', PCT)
        put(ws, f"F{r}", f"={sm('awarded')}", MONEY)
        put(ws, f"G{r}", f"={sm('lost')}", MONEY)
        put(ws, f"H{r}", f'=IF(F{r}+G{r}=0,"",F{r}/(F{r}+G{r}))', PCT)
        put(ws, f"I{r}", f"={cnt('submitted')}", INT)
        put(ws, f"J{r}", f"={cnt('closed')}", INT)
        for col in "ABCDEFGHIJ":
            ws[f"{col}{r}"].border = THIN

    h0 = rl + 3
    put(ws, f"A{h0}", '="Hit rate — "&C4&" (won ÷ won + lost)"', font=F_SECTION)
    hdr = ["", "Bids", "Won", "Lost", "Hit rate #", "Won $", "Lost $", "Hit rate $", "Awaiting", "Closed"]
    head(ws, h0 + 1, 1, hdr)
    hit_row(h0 + 2, "All bids")
    jts = [j for j, _ in Counter(r["job_type"] or "unknown" for r in bl).most_common()]
    put(ws, f"A{h0 + 3}", "By job type", font=F_BOLD)
    for k, jt in enumerate(jts):
        hit_row(h0 + 4 + k, jt, f',{JT},"{jt}"')
    hr = h0 + 4 + len(jts)
    missing = sum(1 for r in bl if r["status"] in ("awarded", "lost") and num(r["submitted_amount"]) is None)
    put(ws, f"A{hr}", f"Bids = submitted + won + lost + closed. Awaiting and closed are not in the rate. "
        f"$ rates use submitted amounts ({missing} won/lost bids have no amount and count in # only). "
        "All GCs: By GC tab.", font=F_NOTE)

    # Top GCs (live; ordered by all-time bid count)
    g0 = hr + 2
    put(ws, f"A{g0}", '="Top GCs by bids — "&C4', font=F_SECTION)
    head(ws, g0 + 1, 1, ["GC"] + hdr[1:])
    grp_count = Counter(raw2grp.get((r["gc"] or "").strip(), "(unknown)") for r in bl
                        if r["status"] in ("submitted", "awarded", "lost", "closed"))
    grps = [g for g, _ in sorted(grp_count.items(), key=lambda t: (-t[1], t[0]))]
    top = [g for g in grps if g != "(unknown)"][:a.top]
    for k, g in enumerate(top):
        hit_row(g0 + 2 + k, g, f',{GG},"{g}"')
        ws[f"A{g0 + 2 + k}"].font = F_BASE

    # Price movers (values, refreshed Fridays)
    m0 = g0 + 3 + len(top)
    movers = sorted([p for p in prices if p.get("mover") == "y"], key=lambda p: -abs(num(p["change_pct"]) or 0))
    put(ws, f"A{m0}", "Price movers — same supplier, same product, ≥ 3% vs a quote 30+ days earlier, latest quote in the last 90 days",
        font=F_SECTION)
    head(ws, m0 + 1, 1, ["Item", "Change", "Now", "Before", "Product", "Unit", "Supplier", "Quoted", "Prior", ""])
    for k, p in enumerate(movers[:a.movers]):
        r = m0 + 2 + k
        put(ws, f"A{r}", p["item"]); put(ws, f"B{r}", num(p["change_pct"]), PCT)
        put(ws, f"C{r}", num(p["unit_price"]), RATE); put(ws, f"D{r}", num(p["prior_price"]), RATE)
        put(ws, f"E{r}", p["product"]); put(ws, f"F{r}", p["unit"]); put(ws, f"G{r}", p["supplier"])
        put(ws, f"H{r}", day(p["quote_date"]), DATE); put(ws, f"I{r}", day(p["prior_date"]), DATE)
        for col in "ABCDEFGHIJ":
            ws[f"{col}{r}"].border = THIN
    if not movers:
        put(ws, f"A{m0 + 2}", "None." if prices else "Price Book Latest.csv not supplied.", font=F_NOTE)
    put(ws, f"A{m0 + 3 + max(len(movers[:a.movers]), 1)}",
        f"{len(movers)} movers of {n_pr} latest prices; full list on the Prices tab (filter Mover = y). "
        "Prices refresh with the Friday sweep.", font=F_NOTE)
    j0 = m0 + 5 + max(len(movers[:a.movers]), 1)
    put(ws, f"A{j0}", "Actual vs estimate — closed-out jobs (Job Cost tab)", font=F_SECTION)
    head(ws, j0 + 1, 1, ["", "Jobs", "Median man-day var", "Median labor $ var", "", "", "", "", "", ""])
    sc, stt, mdv, lv = JC("C"), JC("E"), JC("I"), JC("L")
    put(ws, f"A{j0 + 2}", "Final, job level", font=F_BOLD)
    put(ws, f"B{j0 + 2}", f'=COUNTIFS({sc},"job",{stt},"final")', INT)
    arr(ws, f"C{j0 + 2}", f'=IF(COUNTIFS({sc},"job",{stt},"final",{mdv},"<>")=0,"",MEDIAN(IF(({sc}="job")*({stt}="final")*({mdv}<>""),{mdv})))', PCT)
    arr(ws, f"D{j0 + 2}", f'=IF(COUNTIFS({sc},"job",{stt},"final",{lv},"<>")=0,"",MEDIAN(IF(({sc}="job")*({stt}="final")*({lv}<>""),{lv})))', PCT)
    put(ws, f"A{j0 + 3}", "Positive = actual over estimate. Labor $ compared only where the accounting export is base wage; "
        "unit-type rows exist only for single-unit-type jobs or a manual split." if jc else
        "No closeout actuals loaded yet (Job Cost History.csv).", font=F_NOTE)
    ws.freeze_panes = "A7"

    # --- By GC tab (every GC; live) ---
    ws = bygc
    ws.sheet_view.showGridLines = False
    put(ws, "A1", '="Hit rate by GC — "&Dashboard!C4&" (change the period on the Dashboard)"', font=F_SECTION)
    head(ws, 2, 1, ["GC", "Bids", "Won", "Lost", "Hit rate #", "Won $", "Lost $", "Hit rate $", "Awaiting", "Closed",
                    "Last bid date"], [34, 8, 8, 8, 10, 13, 13, 10, 9, 8, 12])
    for k, g in enumerate(grps, start=3):
        cnt = lambda s: f'COUNTIFS({ST},"{s}",{INP},1,{GG},A{k})'
        sm = lambda s: f'SUMIFS({AMT},{ST},"{s}",{INP},1,{GG},A{k})'
        put(ws, f"A{k}", g)
        put(ws, f"B{k}", f"={cnt('submitted')}+{cnt('awarded')}+{cnt('lost')}+{cnt('closed')}", INT)
        put(ws, f"C{k}", f"={cnt('awarded')}", INT); put(ws, f"D{k}", f"={cnt('lost')}", INT)
        put(ws, f"E{k}", f'=IF(C{k}+D{k}=0,"",C{k}/(C{k}+D{k}))', PCT)
        put(ws, f"F{k}", f"={sm('awarded')}", MONEY); put(ws, f"G{k}", f"={sm('lost')}", MONEY)
        put(ws, f"H{k}", f'=IF(F{k}+G{k}=0,"",F{k}/(F{k}+G{k}))', PCT)
        put(ws, f"I{k}", f"={cnt('submitted')}", INT); put(ws, f"J{k}", f"={cnt('closed')}", INT)
        put(ws, f"K{k}", f'=IF(COUNTIFS({GG},A{k},{BD},"<>")=0,"",_xlfn.MAXIFS({BD},{GG},A{k}))', DATE)
        for col in "ABCDEFGHIJK":
            ws[f"{col}{k}"].border = THIN
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:K{len(grps) + 2}"

    # --- About ---
    ws = about
    ws.column_dimensions["A"].width = 120
    lines = [("About this workbook", F_SECTION),
             (f"Rebuilt {a.stamp} by dashboard/build_dashboard.py (print kit). Rebuilt nightly; edits are overwritten.", F_BASE),
             ("Sources: _Knowledge/Bid Log 2026.csv · _Knowledge/Bid Metrics 2026.csv · _Knowledge/Price Book/Price Book Latest.csv (written by the Friday sweep).", F_BASE),
             ("Hit rate = won ÷ (won + lost) by bid date. Submitted-awaiting and closed bids are shown, never counted in the rate.", F_BASE),
             ("Labor $/SF = Crew Page base wage spread across unit types by line-item labor (before burden and OH&P). "
              f"Standard wages: {STD_WAGES}.", F_BASE),
             ("", F_BASE), ("GC names merged (display ← spellings in the Bid Log)", F_BOLD)]
    lines += [(f"{d} ← " + " | ".join(raws), F_BASE) for d, raws in merges] or [("none", F_BASE)]
    lines += [("", F_BASE), ("⚠ items", F_BOLD)] + ([(w, F_BASE) for w in warn] or [("none", F_BASE)])
    for i, (t, f) in enumerate(lines, start=1):
        put(ws, f"A{i}", t, font=f)
        ws[f"A{i}"].alignment = Alignment(wrap_text=True)

    for s in (bls, bms, pr, jcs):
        for row in s.iter_rows(min_row=2):
            for c in row:
                c.font = F_BASE
    for s_ in (dash, bygc):
        s_.page_setup.orientation = "landscape"
        s_.sheet_properties.pageSetUpPr.fitToPage = True
        s_.page_setup.fitToWidth, s_.page_setup.fitToHeight = 1, 0
    wb.active = 0
    wb.save(a.out)
    return bl, bm, prices, grps, warn


__all__ = ["build"]


# ---------- recalc + checks ----------
MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub RecalculateAndSave()
  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>"""


def recalc(path, timeout=120):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return "LibreOffice (soffice) not found — formulas were not recalculated"
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory(prefix="lo-profile-") as prof:
        url = "file://" + prof
        subprocess.run([soffice, "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                       capture_output=True, timeout=timeout, env=env)
        mdir = os.path.join(prof, "user", "basic", "Standard")
        if not os.path.isdir(mdir):
            return "LibreOffice profile not created — formulas were not recalculated"
        open(os.path.join(mdir, "Module1.xba"), "w").write(MACRO)
        before = os.stat(path).st_mtime_ns
        p = subprocess.run([soffice, "--headless", "--norestore", f"-env:UserInstallation={url}",
                            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                            os.path.abspath(path)], capture_output=True, text=True, timeout=timeout, env=env)
        if p.returncode != 0 or os.stat(path).st_mtime_ns == before:
            return f"LibreOffice did not rewrite the file (exit {p.returncode}) {p.stderr.strip()[:200]}"
    return None


ERRORS = ("#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A", "Err:")


def check(path, bl, bm, grps, period):
    wb = load_workbook(path, data_only=True)
    problems, nformula = [], 0
    wf = load_workbook(path)
    for ws in wf.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula) or (isinstance(v, str) and v.startswith("=")):
                    nformula += 1
                    cv = wb[ws.title][c.coordinate].value
                    if isinstance(cv, str) and any(cv.startswith(e) for e in ERRORS):
                        problems.append(f"{ws.title}!{c.coordinate} = {cv}")
    if nformula == 0:
        problems.append("no formulas found")
    d = wb["Dashboard"]
    # Python cross-check of the headline numbers (period = All only; other periods depend on TODAY())
    if period == "All":
        won = sum(r["status"] == "awarded" for r in bl)
        lost = sum(r["status"] == "lost" for r in bl)
        row = next(r for r in range(1, 80) if d[f"A{r}"].value == "All bids")
        got = (d[f"C{row}"].value, d[f"D{row}"].value)
        if got != (won, lost):
            problems.append(f"hit rate cross-check: sheet {got} vs python {(won, lost)}")
    for k, (u, name) in enumerate(UNIT_TYPES):
        vals = [num(r.get(f"{u}_labor_per_sf")) for r in bm if num(r.get(f"{u}_sf"))]
        vals = [v for v in vals if v is not None]
        row = next(r for r in range(1, 40) if d[f"A{r}"].value == name)
        got = d[f"B{row}"].value
        exp = round(statistics.median(vals), 2) if vals else None
        if (got in ("", None)) != (exp is None) or (exp is not None and abs(float(got) - exp) > 0.005):
            problems.append(f"labor $/SF cross-check {name}: sheet {got} vs python {exp}")
    return nformula, problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--bid-log", required=True); ap.add_argument("--bid-metrics", required=True)
    ap.add_argument("--prices", default=""); ap.add_argument("--job-cost", default="")
    ap.add_argument("--out", default="Dashboard.xlsx")
    ap.add_argument("--period", default="YTD", choices=PERIODS)
    ap.add_argument("--top", type=int, default=12); ap.add_argument("--movers", type=int, default=10)
    ap.add_argument("--stamp", default=dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
    ap.add_argument("--no-recalc", action="store_true")
    a = ap.parse_args()

    # Cross-check against python with period=All, then rebuild with the requested default period
    tmp = a.out + ".check.xlsx"
    out, period = a.out, a.period
    a.out, a.period = tmp, "All"
    bl, bm, prices, grps, warn = build(a)
    problems = []
    if not a.no_recalc:
        err = recalc(tmp)
        if err:
            problems.append(err)
        else:
            nf, problems = check(tmp, bl, bm, grps, "All")
    os.remove(tmp)
    a.out, a.period = out, period
    build(a)
    nf = 0
    if not a.no_recalc and not problems:
        err = recalc(out)
        if err:
            problems.append(err)
        else:
            nf, p2 = check(out, bl, bm, grps, period)
            problems += p2
    size = os.path.getsize(out)
    print(f"Dashboard: {out} ({size // 1024} KB) · {len(bl)} bids · {len(bm)} estimates · {len(prices)} prices · "
          f"{len(grps)} GCs · {nf} formulas · period {period}")
    for w in warn:
        print("⚠ " + w)
    if size > 900_000:
        problems.append(f"file is {size // 1024} KB — over the 900 KB upload margin")
    for p in problems:
        print("✗ " + p)
    print("Upload: NO — fix the ✗ lines" if problems else "Upload: yes (base64, replace Dashboard.xlsx)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
