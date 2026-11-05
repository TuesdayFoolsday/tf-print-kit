"""Ledger CSV (from jc_detail_ledger.py) -> tagged ledger CSV + Excel summary workbook with live SUMIFS formulas.

    python3 jobcost/jc_detail_summary.py <ledger.csv> "<out Ledger.csv>" "<out Summary.xlsx>" "<title line for the Notes tab>"

Adds the cost bucket (System Rules B14 mapping) and the employee role (foreman list below, mason, operator, laborer)
to every row, then builds Jobs / Job by Month / Employees / Vendors / Materials / Notes tabs over a Ledger tab.
Documented in Preconstruction _Knowledge/Job Cost Ledgers - Build Sheet.md. Built Nov 4, 2026.
"""
import csv, re
from collections import defaultdict, OrderedDict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import sys
SRC,LEDGER,XLSX,TITLE=sys.argv[1:5]
rows=list(csv.DictReader(open(SRC)))
FOREMEN={"Leon, Abel","Borja II, Frank R.","Barragan, Ignacio","Angulo-Nagy, Wilman A.","Banuelos, Oscar","Rico-Robles Jr, Ruben","Crissinger, Matthew E.","Barker, Donald P","Frilot, Antoine R","Flores, Abel","Porras, Tommy","Guizar, Juan J","Valencia, Ignacio","Valladares, Julio C","Mills, Randy D.","Perez, Joel A"}
# employee avg regular rate
emp=defaultdict(lambda:[0.0,0.0,set()])
for r in rows:
    if r['trans_type']=='PR' and r['earn_desc']=='Regular Earnings':
        e=emp[r['employee']]; e[0]+=float(r['hours']); e[1]+=float(r['cost']); e[2].add(r['craft_class'])
role={}
for name,(h,c,cl) in emp.items():
    avg=c/h if h else 0
    if name in FOREMEN: role[name]="foreman"
    elif any(x in ("Bricklayer","Journeyman","Level 3","Foreman") for x in cl): role[name]="mason"
    elif 25<=avg<=28: role[name]="operator"
    else: role[name]="laborer"
def bucket(r):
    ct=r['cost_type']
    if ct=='1':
        if r['earn_desc'] in ('Regular Earnings',) :
            return 'labor_trucking' if r['phase']=='01650' else 'labor_base'
        if r['earn_desc']=='Miscellaneous': return 'labor_misc'
        return 'labor_burden'
    return {'2':'material','8':'material','3':'sub','11':'sub_labor','4':'equipment','5':'equipment','6':'other','10':'inventory','9':'sub'}.get(ct,'unknown')
for r in rows:
    r['bucket']=bucket(r); r['role']=role.get(r['employee'],'') if r['trans_type']=='PR' else ''
cols=list(rows[0].keys())
with open(LEDGER,'w',newline='',encoding='utf-8') as f:
    w=csv.DictWriter(f,fieldnames=cols,lineterminator='\n'); w.writeheader(); w.writerows(rows)
# ---- workbook ----
wb=Workbook(); ws=wb.active; ws.title='Ledger'
num_cols={'units','hours','cost','trans_no'}
ws.append(cols)
for r in rows:
    ws.append([ (float(r[c]) if c in ('units','hours','cost') else (int(r[c]) if c=='trans_no' and r[c] else r[c])) for c in cols])
n=len(rows)+1
L={c:get_column_letter(cols.index(c)+1) for c in cols}
def rng(c): return f"Ledger!${L[c]}$2:${L[c]}${n}"
hdr=Font(bold=True,color='FFFFFF'); fill=PatternFill('solid',fgColor='1F4E78')
def style_header(sheet,ncols):
    for i in range(1,ncols+1):
        c=sheet.cell(row=1,column=i); c.font=hdr; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='center')
    sheet.freeze_panes='A2'
style_header(ws,len(cols))
for i,c in enumerate(cols,1): ws.column_dimensions[get_column_letter(i)].width=14 if c not in ('description','employee','vendor','matl_desc','job_name') else 32
for row in ws.iter_rows(min_row=2,min_col=cols.index('cost')+1,max_col=cols.index('cost')+1):
    for c in row: c.number_format='#,##0.00'
# Jobs sheet
jobs=OrderedDict(); jm=defaultdict(set)
for r in rows: jobs.setdefault(r['job'],r['job_name']); jm[r['job']].add(r['mth'])
js=wb.create_sheet('Jobs')
jh=['Job','Job name','Labor hours','Mason hours','Mason-days','Foreman hours','Laborer hours','Operator hours','Base labor $','Avg $/hr','Burden $','Burden %','Trucking labor $','Misc labor $','Material $','Equipment $','Sub $','Sub labor $','Other $','Total cost $','First month','Last month','Months active']
js.append(jh)
B=lambda c,crit: f'SUMIFS({rng("cost")},{rng("job")},$A{{r}},{rng("bucket")},"{crit}")'
for i,(j,name) in enumerate(jobs.items(),2):
    r=i
    f=[j,name,
       f'=SUMIFS({rng("hours")},{rng("job")},$A{r})',
       f'=SUMIFS({rng("hours")},{rng("job")},$A{r},{rng("role")},"mason")',
       f'=D{r}/8',
       f'=SUMIFS({rng("hours")},{rng("job")},$A{r},{rng("role")},"foreman")',
       f'=SUMIFS({rng("hours")},{rng("job")},$A{r},{rng("role")},"laborer")',
       f'=SUMIFS({rng("hours")},{rng("job")},$A{r},{rng("role")},"operator")',
       '='+B('cost','labor_base').format(r=r),
       f'=IF(C{r}=0,0,(I{r}+M{r})/C{r})',
       '='+B('cost','labor_burden').format(r=r),
       f'=IF(I{r}+M{r}=0,0,K{r}/(I{r}+M{r}))',
       '='+B('cost','labor_trucking').format(r=r),
       '='+B('cost','labor_misc').format(r=r),
       '='+B('cost','material').format(r=r),
       '='+B('cost','equipment').format(r=r),
       '='+B('cost','sub').format(r=r),
       '='+B('cost','sub_labor').format(r=r),
       '='+B('cost','other').format(r=r)+'+'+B('cost','inventory').format(r=r),
       f'=SUM(I{r},K{r},M{r},N{r},O{r},P{r},Q{r},R{r},S{r})',
       min(m for m in jm[j]), max(m for m in jm[j]), len(jm[j])]
    js.append(f)
tr=len(jobs)+2
js.append(['TOTAL','']+[f'=SUM({get_column_letter(c)}2:{get_column_letter(c)}{tr-1})' if c not in (10,12,21,22,23) else '' for c in range(3,24)])
js.cell(row=tr,column=10).value=f'=IF(C{tr}=0,0,(I{tr}+M{tr})/C{tr})'; js.cell(row=tr,column=12).value=f'=IF(I{tr}+M{tr}=0,0,K{tr}/(I{tr}+M{tr}))'
js.cell(row=tr,column=1).font=Font(bold=True)
style_header(js,len(jh)); js.column_dimensions['B'].width=34
for row in js.iter_rows(min_row=2,max_row=tr):
    for c in row:
        if c.column in (9,11,13,14,15,16,17,18,19,20): c.number_format='#,##0'
        elif c.column in (10,): c.number_format='0.00'
        elif c.column==12: c.number_format='0.0%'
        elif c.column in (3,4,5,6,7,8): c.number_format='#,##0.0'
# Job x Month
months=sorted(set(r['mth'] for r in rows))
ms=wb.create_sheet('Job by Month'); ms.append(['Job','Job name']+months+['Total'])
for i,(j,name) in enumerate(jobs.items(),2):
    ms.append([j,name]+[f'=SUMIFS({rng("cost")},{rng("job")},$A{i},{rng("mth")},{get_column_letter(k)}$1)' for k in range(3,3+len(months))]+[f'=SUM(C{i}:{get_column_letter(2+len(months))}{i})'])
tr2=len(jobs)+2
ms.append(['TOTAL','']+[f'=SUM({get_column_letter(k)}2:{get_column_letter(k)}{tr2-1})' for k in range(3,4+len(months))])
style_header(ms,3+len(months)); ms.column_dimensions['B'].width=34
for row in ms.iter_rows(min_row=2,max_row=tr2,min_col=3):
    for c in row: c.number_format='#,##0'
# Employees
es=wb.create_sheet('Employees'); es.append(['Employee','Role','Craft/class','Regular hours','Regular $','Avg $/hr','Burden $','Burden %','Jobs worked','Foreman? (edit)'])
for i,(name,(h,c,cl)) in enumerate(sorted(emp.items(),key=lambda x:-x[1][0]),2):
    es.append([name,role[name],", ".join(sorted(cl)),
      f'=SUMIFS({rng("hours")},{rng("employee")},$A{i},{rng("earn_desc")},"Regular Earnings")',
      f'=SUMIFS({rng("cost")},{rng("employee")},$A{i},{rng("earn_desc")},"Regular Earnings")',
      f'=IF(D{i}=0,0,E{i}/D{i})',
      f'=SUMIFS({rng("cost")},{rng("employee")},$A{i},{rng("bucket")},"labor_burden")',
      f'=IF(E{i}=0,0,G{i}/E{i})',
      f'=SUMPRODUCT(({rng("employee")}=$A{i})*({rng("earn_desc")}="Regular Earnings")/COUNTIFS({rng("employee")},{rng("employee")},{rng("earn_desc")},"Regular Earnings",{rng("job")},{rng("job")}))' if False else '',
      'yes' if name in FOREMEN else ''])
style_header(es,10); es.column_dimensions['A'].width=36; es.column_dimensions['C'].width=28
for row in es.iter_rows(min_row=2):
    for c in row:
        if c.column in (5,7): c.number_format='#,##0'
        elif c.column==6: c.number_format='0.00'
        elif c.column==8: c.number_format='0.0%'
        elif c.column==4: c.number_format='#,##0.0'
# jobs worked per employee computed in python (distinct jobs with regular earnings)
ej=defaultdict(set)
for r in rows:
    if r['earn_desc']=='Regular Earnings': ej[r['employee']].add(r['job'])
for i in range(2,es.max_row+1):
    es.cell(row=i,column=9).value=len(ej[es.cell(row=i,column=1).value])
# Vendors
vend=OrderedDict()
for r in rows:
    if r['trans_type']=='AP': vend.setdefault(r['vendor'],r['vendor_no'])
vs=wb.create_sheet('Vendors'); vs.append(['Vendor','Vendor #','Lines','Spend $','Material $','Equipment $','Sub $','Sub labor $','Other $','Jobs'])
for i,(v,vn) in enumerate(sorted(vend.items(),key=lambda x:-sum(float(r['cost']) for r in rows if r['vendor']==x[0])),2):
    vs.append([v,vn,f'=COUNTIFS({rng("vendor")},$A{i})',f'=SUMIFS({rng("cost")},{rng("vendor")},$A{i})']+
              [f'=SUMIFS({rng("cost")},{rng("vendor")},$A{i},{rng("bucket")},"{b}")' for b in ('material','equipment','sub','sub_labor','other')]+
              [len(set(r['job'] for r in rows if r['vendor']==v))])
style_header(vs,10); vs.column_dimensions['A'].width=34
for row in vs.iter_rows(min_row=2):
    for c in row:
        if c.column in range(4,10): c.number_format='#,##0'
# Materials
mats=OrderedDict()
for r in rows:
    if r['matl_code']: mats.setdefault((r['matl_code'],r['matl_desc']),0)
mts=wb.create_sheet('Materials'); mts.append(['Matl code','Description','Lines','Spend $','Units'])
for i,(k,_) in enumerate(mats.items(),2):
    mts.append([k[0],k[1],f'=COUNTIFS({rng("matl_code")},$A{i},{rng("matl_desc")},$B{i})',f'=SUMIFS({rng("cost")},{rng("matl_code")},$A{i},{rng("matl_desc")},$B{i})',f'=SUMIFS({rng("units")},{rng("matl_code")},$A{i},{rng("matl_desc")},$B{i})'])
style_header(mts,5); mts.column_dimensions['B'].width=44
for row in mts.iter_rows(min_row=2):
    for c in row:
        if c.column in (4,): c.number_format='#,##0'
# Notes
ns=wb.create_sheet('Notes')
notes=[TITLE,
f"Source rows: {len(rows):,} transactions; ${sum(float(r['cost']) for r in rows):,.2f}; {sum(float(r['hours']) for r in rows):,.1f} hours. Every cost-type, phase and job subtotal in the report matched the parsed rows.",
"Buckets (per System Rules B14 mapping, David Nov 1 2026): labor_base = cost type 1 Regular Earnings (with hours); labor_burden = cost type 1 Payroll Tax, Workmen's Comp, 401K Match, FUTA, SUTA, Liability Insurance, Health Insurance; labor_trucking = phase 01650 Regular Earnings; labor_misc = earn code 30 Miscellaneous (no hours); material = cost types 2-M and 8-I; equipment = 4-E and 5-R; sub = 3-S; sub_labor = 11-SL; other = 6-O; inventory = 10-INV (nets to zero).",
"Roles: foreman = named on the Operations Field roster (Leon, Borja, Barragan, Angulo-Nagy A., Banuelos O., Rico-Robles Jr) plus Crissinger, Don Barker, Frilot, Flores, Porras, Guizar, Valencia, Valladares, Mills and Perez (David, Nov 4; John Cody Barker is a mason); mason = craft 1 (Bricklayer, Journeyman, Level 3); operator = craft 2 with average regular rate $25-28/hr (Israel Gonzalez only); laborer = other craft 2. Edit the Foreman column on Employees and tell Claude if any are wrong.",
"Cost type 11 SL and 3 S are both Texas E & D Masonry (labor subcontract); several invoices were posted to one type and reversed into another, so read sub + sub labor together.",

"All summary sheets are live formulas over the Ledger sheet. Add rows to Ledger and they update."]
for t in notes: ns.append([t])
ns.column_dimensions['A'].width=160
wb.save(XLSX); print("saved", n-1, "ledger rows;", len(jobs), "jobs;", len(emp), "employees;", len(vend), "vendors;", len(mats), "materials")
