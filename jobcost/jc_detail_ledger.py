"""Vista JC Detail (.xls, the printed report saved as Excel) -> one-row-per-transaction CSV + reconciliation summary.

    python3 jobcost/jc_detail_ledger.py "<JC Detail.xls>" <out ledger.csv> <out summary.json>

Reads every sheet (Excel splits the report at 65k rows), follows the job / phase / cost-type headers, parses payroll
(craft, employee, earnings code) and AP (vendor, invoice, material code, PO line) descriptions, and checks every
"Total for Cost Type / For Phase / For Job" line against the rows read under it. Any mismatch is listed in the JSON.
Documented in Preconstruction _Knowledge/Job Cost Ledgers - Build Sheet.md. Built Nov 4, 2026.
"""
import sys, re, csv, json, xlrd
from collections import Counter, defaultdict
wb = xlrd.open_workbook(sys.argv[1], on_demand=True)
out_csv, out_json = sys.argv[2], sys.argv[3]
def d(v):
    return xlrd.xldate_as_datetime(v, wb.datemode).date().isoformat() if isinstance(v,(int,float)) and v>1000 else ""
JOB=re.compile(r"^\s*(\d{4})-\s{2}(.*?)\s*(- Continued)?\s*$")
PH=re.compile(r"^\s*(\d{5})-\s+-\s+(.*?)\s{2,}(\d+)\s+([A-Z]+)\s*$")
TOT=re.compile(r"^Total (for Cost Type|For Phase|For Job):\s*(.*)$", re.I)
PRD=re.compile(r"^(?:(\d+)/(.+?)\s+)?([\d.]+)\s+/\s+(\d+)\s+/\s+(.*?)\s*(\d{1,3})\s+(\S.*)$")
APD=re.compile(r"^(\d+)\s+(.*?)\s+(\S+)\s+/\s+TR#\s*(\d+)\s*/\s*(\d+)\s*/\s*(?:Matl:\s*([^/]*?)-\s*(.*?)\s*/\s*)?APCo:\s*(\d+)\s*(?:/\s*PO#-Line\s*(\S+)\s*)?(.*)$")
rows=[]; recon=[]; problems=[]
job=jobname=None; ph=None
sum_ct=[0.0,0.0]; sum_ph=[0.0,0.0]; sum_job=[0.0,0.0]; jobs_seen=[]
earn=Counter(); unparsed_pr=[]; unparsed_ap=[]
for sh in wb.sheets():
    for r in range(sh.nrows):
        row=sh.row(r); c0=row[0].value; c0s=str(c0).strip()
        if isinstance(c0,str) and c0s:
            m=JOB.match(c0)
            if m and not m.group(3):
                job,jobname=m.group(1),m.group(2).strip(); jobs_seen.append(job); sum_job=[0.0,0.0]; continue
            if m: continue
            m=PH.match(c0)
            if m:
                ph=(m.group(1),m.group(2).strip(),m.group(3),m.group(4)); sum_ct=[0.0,0.0]
                continue
            if re.match(r"^\s*\d{5}-",c0): problems.append(("unparsed phase",sh.name,r,c0s)); continue
            continue
        t15=str(row[15].value).strip() if len(row)>15 else ""
        m=TOT.match(t15)
        if m:
            kind=m.group(1).lower(); hrs=row[33].value or 0.0; cost=row[36].value or 0.0
            if kind.startswith("for cost"):
                recon.append(("cost type",job,ph[0] if ph else "",ph[2] if ph else "",round(sum_ct[0],2),round(hrs,2),round(sum_ct[1],2),round(cost,2)))
                sum_ph[0]+=sum_ct[0]; sum_ph[1]+=sum_ct[1]; sum_ct=[0.0,0.0]
            elif kind.startswith("for phase"):
                recon.append(("phase",job,ph[0] if ph else "","",round(sum_ph[0],2),round(hrs,2),round(sum_ph[1],2),round(cost,2)))
                sum_job[0]+=sum_ph[0]; sum_job[1]+=sum_ph[1]; sum_ph=[0.0,0.0]
            else:
                recon.append(("job",job,"","",round(sum_job[0],2),round(hrs,2),round(sum_job[1],2),round(cost,2)))
            continue
        if isinstance(c0,(int,float)) and c0s:
            tt=str(row[7].value).strip(); desc=str(row[14].value).strip()
            units=row[31].value or 0.0; hrs=row[34].value or 0.0; cost=row[36].value or 0.0
            rec={"job":job,"job_name":jobname,"phase":ph[0],"phase_desc":ph[1],"cost_type":ph[2],"ct_code":ph[3],
                 "mth":d(c0)[:7],"trans_no":int(row[1].value) if row[1].value!="" else "","posted_date":d(row[2].value),"actual_date":d(row[4].value),
                 "trans_type":tt,"source":str(row[9].value).strip(),"description":desc,"units":units,"hours":hrs,"cost":cost,
                 "craft":"","craft_class":"","emp_no":"","employee":"","earn_code":"","earn_desc":"","vendor_no":"","vendor":"","invoice":"","matl_code":"","matl_desc":"","po_line":"","ap_note":""}
            if tt=="PR":
                m=PRD.match(desc)
                if m: rec.update(craft=m.group(1) or "",craft_class=m.group(2) or "",emp_no=m.group(4),employee=re.sub(r"\s+"," ",m.group(5)).replace(" ,",",").strip(),earn_code=m.group(6),earn_desc=m.group(7).strip()); earn[(m.group(6),m.group(7).strip())]+=1
                else: unparsed_pr.append(desc)
            elif tt=="AP":
                m=APD.match(desc)
                if m: rec.update(vendor_no=m.group(1),vendor=m.group(2).strip(),invoice=m.group(3),matl_code=m.group(6) or "",matl_desc=(m.group(7) or "").strip(),po_line=m.group(9) or "",ap_note=m.group(10).strip())
                else: unparsed_ap.append(desc)
            else: problems.append(("trans type",tt,desc))
            rows.append(rec); sum_ct[0]+=hrs; sum_ct[1]+=cost
bad=[x for x in recon if abs(x[4]-x[5])>0.011 or abs(x[6]-x[7])>0.011]
with open(out_csv,"w",newline="",encoding="utf-8") as f:
    w=csv.DictWriter(f,fieldnames=list(rows[0].keys()),lineterminator="\n"); w.writeheader(); w.writerows(rows)
summary={"rows":len(rows),"jobs_with_lines":sorted(set(r["job"] for r in rows)),"jobs_in_header_order":jobs_seen,
 "total_cost":round(sum(r["cost"] for r in rows),2),"total_hours":round(sum(r["hours"] for r in rows),2),
 "recon_checks":len(recon),"recon_failures":bad,"job_totals":[x for x in recon if x[0]=="job"],
 "earn_codes":sorted([(k[0],k[1],v) for k,v in earn.items()]),"unparsed_pr":unparsed_pr[:10],"unparsed_pr_n":len(unparsed_pr),
 "unparsed_ap":unparsed_ap[:10],"unparsed_ap_n":len(unparsed_ap),"problems":problems[:20],
 "cost_types":sorted(Counter((r["cost_type"],r["ct_code"]) for r in rows).items()),
 "months":sorted(Counter(r["mth"] for r in rows).items())}
json.dump(summary,open(out_json,"w"),indent=1)
print(json.dumps({k:v for k,v in summary.items() if k not in ("job_totals",)},indent=1)[:6000])
