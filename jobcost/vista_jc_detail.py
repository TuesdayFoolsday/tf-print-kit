"""Vista JC Detail (.xls, JCDetail.rpt) -> job_cost.py actuals.csv, per the FB mapping in jobcost/job_cost.py.

    python3 jobcost/vista_jc_detail.py <JC Detail.xls> <actuals.csv> <estimate job_number> "<estimate name>" <through YYYY-MM-DD> <foreman last name>
(needs: pip install xlrd). Final = y; ask David before using it on a job that is not closed out."""
import csv, re, sys, collections, xlrd
path, out, job, estimate, thru, foreman = sys.argv[1:7]
s = xlrd.open_workbook(path).sheet_by_index(0)
phase = None; recs = []
for r in range(s.nrows):
    vals = [str(x).strip() for x in s.row_values(r) if str(x).strip() != '']
    if not vals: continue
    if len(vals) == 1 and re.match(r'^\d{5}-', vals[0]): phase = re.sub(r'\s+', ' ', vals[0]); continue
    if re.match(r'^\d{5}\.0$', vals[0]) and len(vals) >= 9:
        recs.append(dict(phase=phase, ct=phase.split()[-1], src=vals[4], desc=vals[6], hours=float(vals[-2]), cost=float(vals[-1])))
def person(d): return d.split(' / ')[2].split('  ,')[0].strip() + ', ' + re.sub(r'\d+\s+\D+$', '', d.split('  ,')[1]).strip() if '  ,' in d else d
rate = collections.defaultdict(lambda: [0.0, 0.0])
for r in recs:
    if r['ct'] == 'L' and r['hours'] > 0:
        rate[person(r['desc'])][0] += r['hours']; rate[person(r['desc'])][1] += r['cost']
def cls(r):
    craft = r['desc'].split(' ')[0]
    p = person(r['desc'])
    if foreman.lower() in p.lower(): return 'foreman'
    if craft.startswith('2/'):
        h, c = rate[p]; return 'operator' if 25 <= c / h <= 28.0001 else 'laborer'
    return 'bricklayer'
rows = []; burden = 0.0
for r in recs:
    base = dict(job_number=job, estimate=estimate, through_date=thru, final='y', labor_class='', hours='', burdened='')
    if r['ct'] == 'L':
        if r['phase'].startswith('01650'):
            rows.append({**base, 'cost_type': 'other', 'amount': f"{r['cost']:.2f}"})   # trucking: shown, not compared
        elif r['hours'] > 0:
            rows.append({**base, 'cost_type': 'labor', 'labor_class': cls(r), 'hours': r['hours'], 'amount': f"{r['cost']:.2f}", 'burdened': 'n'})
        else:
            burden += r['cost']
    else:
        ct = {'M': 'material', 'I': 'material', 'E': 'equipment', 'R': 'equipment', 'O': 'other', 'S': 'sub'}.get(r['ct'], 'other')
        rows.append({**base, 'cost_type': ct, 'amount': f"{r['cost']:.2f}"})
with open(out, 'w', newline='') as f:
    w = csv.DictWriter(f, fieldnames=["job_number","estimate","through_date","final","cost_type","labor_class","hours","amount","burdened"], lineterminator="\n")
    w.writeheader(); w.writerows(rows)
tot = sum(r['cost'] for r in recs)
print(f"{len(rows)} actual lines; JC Detail total ${tot:,.2f}; masonry burden left out ${burden:,.2f}")
for p, (h, c) in sorted(rate.items()): print(f"  {p}: {h:g} h, ${c/h:.2f}/h -> {cls(next(r for r in recs if r['ct']=='L' and r['hours']>0 and person(r['desc'])==p))}")
