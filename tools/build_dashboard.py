#!/usr/bin/env python3
"""
VPS Lakeshore (LHRC) — Daily Revenue Dashboard builder (v2, granular).

Scans the "Daily MIS Reports" folder for:
  - Daily Revenue Flash_New_DD-MM-YYYY*.xlsx   (Burjeel flash: revenue, budget,
    collections, dept/doctor day detail, discharge lists)
  - Daily MIS  (DD-MM-YYYY)- VPS Lakeshore*.xlsx (internal master: doctor x day
    revenue / OP visits / admissions / discharges, bed days, ALOS)

and regenerates LHRC_Revenue_Dashboard.html (self-contained).
Doctor/dept month aggregates persist in _dashboard_tools/history.json so past
months survive even if the source files are removed.

Usage:  python3 build_dashboard.py [folder]
"""
import sys, os, re, json, glob, datetime
import openpyxl

FY_MONTHS = ["APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER",
             "NOVEMBER","DECEMBER","JANUARY","FEBRUARY","MARCH"]

def file_date(path):
    m = re.search(r'(\d{2})-(\d{2})-(\d{4})', os.path.basename(path))
    if not m: return None
    d, mo, y = map(int, m.groups())
    try: return datetime.date(y, mo, d)
    except ValueError: return None

def num(v): return float(v) if isinstance(v, (int, float)) else 0.0

def norm(s): return re.sub(r'\s+', ' ', str(s).strip()).upper()

def title(s):
    return ' '.join(w if ('.' in w and len(w) <= 4) else w.capitalize()
                    for w in str(s).strip().split())

# ---------------------------------------------------------------- flash parsers
def parse_year_tables(ws):
    rows = list(ws.iter_rows(values_only=True, max_col=25))
    out, cur = {}, None
    for r in rows:
        a = str(r[0]).strip() if r[0] is not None else ""
        up = a.upper()
        if up.startswith("YEAR 20") or up.startswith("FY 20"):
            cur = re.sub(r'^(YEAR|FY)\s*', 'FY ', a, flags=re.I).strip()
            out[cur] = {"months": []}
        elif cur and up in FY_MONTHS and any(isinstance(x,(int,float)) for x in r[1:7]):
            out[cur]["months"].append(dict(
                month=up.title(), opVisits=num(r[1]), ipDisch=num(r[2]),
                revOP=num(r[3]), revIP=num(r[4]), revPH=num(r[5]), revTot=num(r[6]),
                budOPv=num(r[7]), budIPd=num(r[8]),
                budOP=num(r[9]), budIP=num(r[10]), budPH=num(r[11]), budTot=num(r[12]),
                collCash=num(r[13]), collCredit=num(r[14]), collTot=num(r[15]),
                budCollTot=num(r[18]),
                census=num(r[23]) if len(r) > 24 else 0,
                occDays=num(r[24]) if len(r) > 24 else 0))
    return out

def parse_month_sheet(ws):
    days = []
    for r in ws.iter_rows(values_only=True, max_col=17):
        if isinstance(r[1], (datetime.datetime, datetime.date)):
            days.append(dict(
                date=r[1].strftime("%Y-%m-%d"), dow=str(r[0] or ""),
                opVisits=num(r[2]), ipDisch=num(r[3]),
                revOP=num(r[4]), revIP=num(r[5]), revPH=num(r[6]), revTot=num(r[7]),
                budTot=num(r[13]), collCash=num(r[14]), collCredit=num(r[15]),
                collTot=num(r[16])))
    return days

def parse_dept_sheet(ws):
    header, hrow = None, None
    for i, r in enumerate(ws.iter_rows(values_only=True, max_row=10)):
        vals = [str(v).strip() if v else "" for v in r]
        if "Department" in vals and "Doctor" in vals:
            header, hrow = vals, i + 1; break
    if not header: return {}, {}, {}, {}
    ix = {n: header.index(n) for n in ("Department","Doctor","Revenue","New Type") if n in header}
    if "Revenue" not in ix: return {}, {}, {}, {}
    dept, doc, typ, doctype = {}, {}, {}, {}
    for r in ws.iter_rows(values_only=True, min_row=hrow + 1):
        d = r[ix["Department"]] if ix["Department"] < len(r) else None
        if not d or not str(d).strip(): continue
        rev = num(r[ix["Revenue"]]) if ix["Revenue"] < len(r) else 0.0
        dept[str(d).strip()] = dept.get(str(d).strip(), 0) + rev
        dn, t = None, None
        if "Doctor" in ix and ix["Doctor"] < len(r) and r[ix["Doctor"]]:
            dn = norm(r[ix["Doctor"]]); doc[dn] = doc.get(dn, 0) + rev
        if "New Type" in ix and ix["New Type"] < len(r) and r[ix["New Type"]]:
            t = norm(r[ix["New Type"]])
            if t in ("IP","OP","PH"): typ[t] = typ.get(t, 0) + rev
        if dn and t in ("IP","OP","PH"):
            dd = doctype.setdefault(dn, {"IP":0,"OP":0,"PH":0}); dd[t] += rev
    return dept, doc, typ, doctype

def parse_op_sheet(ws):
    """Flash 'OP' sheet -> {doctor: {new, free, renew, tot}} for that day."""
    header, hrow = None, None
    for i, r in enumerate(ws.iter_rows(values_only=True, max_row=8)):
        vals = [norm(v) if v else "" for v in r]
        if "DOCTOR NAME" in vals and "TOTAL VISIT" in vals:
            header, hrow = vals, i + 1; break
    if not header: return {}
    ix = {n: header.index(n) for n in ("DOCTOR NAME","NEW","FREE","RENEW","TOTAL VISIT") if n in header}
    out = {}
    for r in ws.iter_rows(values_only=True, min_row=hrow + 1):
        dn = r[ix["DOCTOR NAME"]] if ix["DOCTOR NAME"] < len(r) else None
        if not dn or norm(dn) in ("TOTAL","GRAND TOTAL"): continue
        rec = out.setdefault(norm(dn), {"new":0,"free":0,"renew":0,"tot":0})
        rec["new"] += num(r[ix.get("NEW", -1)]) if "NEW" in ix else 0
        rec["free"] += num(r[ix.get("FREE", -1)]) if "FREE" in ix else 0
        rec["renew"] += num(r[ix.get("RENEW", -1)]) if "RENEW" in ix else 0
        rec["tot"] += num(r[ix.get("TOTAL VISIT", -1)]) if "TOTAL VISIT" in ix else 0
    return out

def classify_scheme(s):
    u = norm(s)
    if "CASH" in u: return "Cash"
    if "ECHS" in u: return "ECHS"
    if any(k in u for k in ("INSURANCE","TPA","HEALTH","ASSIST","MEDI","STAR",
                            "NIVA","ICICI","HDFC","BAJAJ","CIGNA","ALLIANZ")):
        return "Insurance"
    return "Corporate/Other"

def classify_status(s):
    u = norm(s)
    if "RECOVER" in u or "IMPROV" in u or "CURED" in u: return "Recovered"
    if "EXPIR" in u or "DEATH" in u or "DECEAS" in u: return "Expired"
    if "DAMA" in u or "LAMA" in u or "AGAINST" in u or "REQUEST" in u: return "DAMA/LAMA"
    if "REFER" in u or "TRANSFER" in u: return "Referred/Transferred"
    return "Other"

def parse_dis_sheet(ws):
    """Flash 'Dis' sheet -> list of discharge dicts."""
    header, hrow = None, None
    for i, r in enumerate(ws.iter_rows(values_only=True, max_row=8)):
        vals = [norm(v) if v else "" for v in r]
        if "DISCHARGE DATE" in vals and "DOCTOR" in vals:
            header, hrow = vals, i + 1; break
    if not header: return []
    ix = {n: header.index(n) for n in
          ("DISCHARGE DATE","ADMISSION DATE","PATIENT NO","DOCTOR","STATUS","SCHEME")
          if n in header}
    out = []
    for r in ws.iter_rows(values_only=True, min_row=hrow + 1):
        dd = r[ix["DISCHARGE DATE"]] if "DISCHARGE DATE" in ix else None
        if not isinstance(dd, (datetime.datetime, datetime.date)): continue
        ad = r[ix.get("ADMISSION DATE", -1)] if "ADMISSION DATE" in ix else None
        los = (dd.date() if isinstance(dd, datetime.datetime) else dd)
        alos = None
        if isinstance(ad, (datetime.datetime, datetime.date)):
            a = ad.date() if isinstance(ad, datetime.datetime) else ad
            alos = max((los - a).days, 0)
        out.append(dict(
            date=los.strftime("%Y-%m-%d"),
            pid=str(r[ix["PATIENT NO"]]) if "PATIENT NO" in ix else "",
            doctor=norm(r[ix["DOCTOR"]]) if "DOCTOR" in ix and r[ix["DOCTOR"]] else "",
            status=classify_status(r[ix["STATUS"]]) if "STATUS" in ix and r[ix["STATUS"]] else "Other",
            scheme=classify_scheme(r[ix["SCHEME"]]) if "SCHEME" in ix and r[ix["SCHEME"]] else "Corporate/Other",
            alos=alos))
    return out

# ------------------------------------------------------------ Daily MIS parsers
def parse_doctor_day_matrix(ws, value_name):
    """Sheets laid out as [Sl, Department, Doctor, <date cols>...] -> {(dept,doc): total}."""
    rows = list(ws.iter_rows(values_only=True))
    hrow, datecols = None, []
    for i, r in enumerate(rows[:8]):
        dc = [(j, c.date() if isinstance(c, datetime.datetime) else c)
              for j, c in enumerate(r) if isinstance(c, (datetime.datetime, datetime.date))]
        if len(dc) >= 5:
            # dominant month only, first occurrence of each date (a trailing
            # 'Total' column reuses day-1's date and would double-count)
            months = {}
            for _, c in dc: months[(c.year, c.month)] = months.get((c.year, c.month), 0) + 1
            dom = max(months, key=months.get)
            seen = set()
            for j, c in dc:
                if (c.year, c.month) == dom and c not in seen:
                    seen.add(c); datecols.append(j)
            hrow = i; break
    if hrow is None: return {}
    # find dept/doctor columns from a header row at/above hrow
    dept_ix, doc_ix = 1, 2
    for r in rows[:hrow + 2]:
        vals = [str(v).strip() if v else "" for v in r]
        if "Department" in vals and "Doctor" in vals:
            dept_ix, doc_ix = vals.index("Department"), vals.index("Doctor")
    out, last_dept = {}, ""
    for r in rows[hrow + 1:]:
        dept = str(r[dept_ix]).strip() if len(r) > dept_ix and r[dept_ix] else ""
        if dept: last_dept = dept
        doc = str(r[doc_ix]).strip() if len(r) > doc_ix and r[doc_ix] else ""
        if not doc: continue
        if norm(doc) in ("TOTAL", "GRAND TOTAL"):
            break  # bottom summary block (TOTAL / category rows) — not doctors
        tot = sum(num(r[j]) for j in datecols if j < len(r))
        key = (norm(last_dept), norm(doc))
        out[key] = out.get(key, 0) + tot
    return out

def parse_mom_fy(ws):
    """'MoM FY 26-27' -> {month 'YYYY-MM': {bedCap, occDays, alos}}"""
    rows = list(ws.iter_rows(values_only=True, max_col=16))
    mcols = {}
    for r in rows[:10]:
        dc = [(j, c) for j, c in enumerate(r) if isinstance(c, (datetime.datetime, datetime.date))]
        if len(dc) >= 6:
            mcols = {j: c.strftime("%Y-%m") for j, c in dc}; break
    if not mcols: return {}
    out = {m: {} for m in mcols.values()}
    for r in rows:
        lbl = str(r[1]).strip() if len(r) > 1 and r[1] else ""
        key = None
        if lbl.startswith("Bed Capacity"): key = "bedCap"
        elif lbl.startswith("Bed Occupancy (D"): key = "occDays"
        elif lbl == "ALOS": key = "alos"
        if key:
            for j, m in mcols.items():
                if j < len(r) and isinstance(r[j], (int, float)):
                    out[m][key] = float(r[j])
    return {m: v for m, v in out.items() if v}

# ---------------------------------------------------------------------- main
def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    tools = os.path.join(folder, "_dashboard_tools")
    os.makedirs(tools, exist_ok=True)
    hist_path = os.path.join(tools, "history.json")
    history = {}
    if os.path.exists(hist_path):
        try: history = json.load(open(hist_path))
        except Exception: history = {}

    # ---- flash files ----
    by_date = {}
    for f in glob.glob(os.path.join(folder, "Daily Revenue Flash_New_*.xlsx")):
        d = file_date(f)
        if d and (d not in by_date or os.path.getmtime(f) > os.path.getmtime(by_date[d])):
            by_date[d] = f
    if not by_date:
        print("No flash files found in", folder); sys.exit(1)
    dates = sorted(by_date)
    latest_date, latest_file = dates[-1], by_date[dates[-1]]
    print(f"{len(dates)} flash dates; latest {latest_date}")

    wb = openpyxl.load_workbook(latest_file, read_only=True, data_only=True)
    year_tables = {}
    for sn in wb.sheetnames:
        if re.match(r'YEAR .*-D', sn): year_tables = parse_year_tables(wb[sn])
    summary = {}
    if "Summary_New" in wb.sheetnames:
        rows = list(wb["Summary_New"].iter_rows(values_only=True, max_col=13))
        def stat(lbl):
            for r in rows:
                if r[0] and str(r[0]).strip().startswith(lbl):
                    return [num(x) for x in r[2:5]]
            return [0,0,0]
        summary = {"bedOcc": stat("Bed Occup."), "surgeries": stat("No. of Surgery"),
                   "admissions": stat("Admissions")}
        for r in rows:
            lbl = str(r[6]).strip() if r[6] else ""
            if lbl in ("OP","IP","PH","IP Conv. Rate"):
                summary.setdefault("real", {})[lbl] = [num(r[7]), num(r[8])]
    wb.close()

    # daily series: newest flash per month carries the whole month
    month_files = {}
    for d in dates: month_files[(d.year, d.month)] = by_date[d]
    daily = {}
    for (y, mo), f in sorted(month_files.items()):
        w = openpyxl.load_workbook(f, read_only=True, data_only=True)
        for sn in w.sheetnames:
            m = re.match(r'([A-Za-z]+)-(\d{4})-D$', sn)
            if m and int(m.group(2)) == y and m.group(1)[:3].upper() in \
                    {mm[:3] for mm in FY_MONTHS} and \
                    FY_MONTHS.index([mm for mm in FY_MONTHS if mm.startswith(m.group(1)[:3].upper())][0]) == (mo - 4) % 12:
                rows = [r for r in parse_month_sheet(w[sn]) if r["revTot"] > 0 or r["budTot"] > 0]
                daily[f"{y}-{mo:02d}"] = rows
        w.close()

    # dept/doctor day cuts + discharge lists from every flash
    dept_tot, doc_tot, dept_dates, discharges, seen_pid = {}, {}, [], [], set()
    doc_type_mix, op_agg = {}, {}
    for d in dates:
        w = openpyxl.load_workbook(by_date[d], read_only=True, data_only=True)
        if "Dept wise" in w.sheetnames:
            dept, doc, _, doctype = parse_dept_sheet(w["Dept wise"])
            if dept:
                dept_dates.append(d.strftime("%Y-%m-%d"))
                for k, v in dept.items(): dept_tot[k] = dept_tot.get(k, 0) + v
                for k, v in doc.items():  doc_tot[k] = doc_tot.get(k, 0) + v
                for k, v in doctype.items():
                    dd = doc_type_mix.setdefault(k, {"IP":0,"OP":0,"PH":0})
                    for t in ("IP","OP","PH"): dd[t] += v[t]
        if "OP" in w.sheetnames:
            for k, v in parse_op_sheet(w["OP"]).items():
                rec = op_agg.setdefault(k, {"new":0,"free":0,"renew":0,"tot":0})
                for f in rec: rec[f] += v[f]
        if "Dis" in w.sheetnames:
            for rec in parse_dis_sheet(w["Dis"]):
                key = (rec["pid"], rec["date"])
                if key not in seen_pid:
                    seen_pid.add(key); discharges.append(rec)
        w.close()

    # ---- Daily MIS files: doctor x month granularity ----
    mis_by_date = {}
    for f in glob.glob(os.path.join(folder, "Daily MIS*.xlsx")):
        d = file_date(f)
        if d and (d not in mis_by_date or os.path.getmtime(f) > os.path.getmtime(mis_by_date[d])):
            mis_by_date[d] = f
    mis_month_files = {}
    for d in sorted(mis_by_date): mis_month_files[f"{d.year}-{d.month:02d}"] = (d, mis_by_date[d])

    mom_fy = {}
    for mkey, (d, f) in sorted(mis_month_files.items()):
        print("Parsing Daily MIS", os.path.basename(f))
        w = openpyxl.load_workbook(f, read_only=True, data_only=True)
        sheets = {norm(s): s for s in w.sheetnames}
        def get(name): return w[sheets[norm(name)]] if norm(name) in sheets else None
        docs = {}
        pairs = [("Doc wise revenue date conso.", "rev"), ("Patient visits", "opv"),
                 ("No. Admissions", "adm"), ("No. discharges", "dis")]
        for sn, field in pairs:
            ws = get(sn)
            if ws is None: continue
            for (dept, doc), val in parse_doctor_day_matrix(ws, field).items():
                rec = docs.setdefault(doc, {"dept": dept, "rev": 0, "opv": 0, "adm": 0, "dis": 0})
                if dept: rec["dept"] = dept
                rec[field] += val
        ws = get("MoM FY 26-27") or get("MoM FY 27-28")
        if ws is not None: mom_fy.update(parse_mom_fy(ws))
        w.close()
        history[mkey] = {"asOf": d.strftime("%Y-%m-%d"),
                         "daysElapsed": d.day,
                         "doctors": {k: v for k, v in docs.items() if any(v[f] for f in ("rev","opv","adm","dis"))}}
    json.dump(history, open(hist_path, "w"))
    print("History months:", sorted(history))

    # discharge aggregates per doctor + overall
    dis_by_doc, status_mix, payer_mix = {}, {}, {}
    for rec in discharges:
        status_mix[rec["status"]] = status_mix.get(rec["status"], 0) + 1
        payer_mix[rec["scheme"]] = payer_mix.get(rec["scheme"], 0) + 1
        dd = dis_by_doc.setdefault(rec["doctor"], {"n": 0, "losSum": 0, "losN": 0, "status": {}, "cash": 0})
        dd["n"] += 1
        if rec["alos"] is not None: dd["losSum"] += rec["alos"]; dd["losN"] += 1
        dd["status"][rec["status"]] = dd["status"].get(rec["status"], 0) + 1
        if rec["scheme"] == "Cash": dd["cash"] += 1

    data = dict(
        generated=datetime.datetime.now().strftime("%d %b %Y %H:%M"),
        latestDate=latest_date.strftime("%d %b %Y"),
        filesParsed=[os.path.basename(by_date[d]) for d in dates] +
                    [os.path.basename(f) for _, f in mis_month_files.values()],
        yearTables=year_tables, daily=daily, summary=summary,
        deptTop=[{"name": k, "rev": v} for k, v in sorted(dept_tot.items(), key=lambda x: -x[1])[:14]],
        deptDates=dept_dates,
        history=history, momFY=mom_fy,
        disByDoc=dis_by_doc, statusMix=status_mix, payerMix=payer_mix,
        docTypeMix=doc_type_mix, opAgg=op_agg,
        disDates=sorted({r["date"] for r in discharges}),
        nDischarges=len(discharges),
    )
    out = os.path.join(folder, "LHRC_Revenue_Dashboard.html")
    open(out, "w", encoding="utf-8").write(TEMPLATE.replace("__DATA__", json.dumps(data)))
    print("Wrote", out)

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>VPS Lakeshore — Daily Revenue Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
:root{--blue:#2B7CBE;--maroon:#8B1A4A;--gray:#7F8C9B;--bg:#f6f8fa;--card:#fff;--good:#1a7f4e;--bad:#c0392b}
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:#243342}
header{background:linear-gradient(90deg,var(--blue),#1b5e94);color:#fff;padding:18px 28px;display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap}
header h1{margin:0;font-size:21px;font-weight:600}
header .sub{font-size:12.5px;opacity:.85}
.wrap{max-width:1320px;margin:0 auto;padding:20px 24px 48px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:12px;margin:6px 0 22px}
.card{background:var(--card);border-radius:10px;padding:13px 15px;box-shadow:0 1px 3px rgba(36,51,66,.08);border-top:3px solid var(--blue)}
.card.m{border-top-color:var(--maroon)}
.card .lbl{font-size:11px;color:var(--gray);text-transform:uppercase;letter-spacing:.05em}
.card .val{font-size:21px;font-weight:650;margin-top:3px}
.card .delta{font-size:12px;margin-top:2px}
.up{color:var(--good)}.dn{color:var(--bad)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:950px){.grid{grid-template-columns:1fr}}
.panel{background:var(--card);border-radius:10px;padding:16px 18px;box-shadow:0 1px 3px rgba(36,51,66,.08);margin-bottom:18px}
.panel h2{margin:0 0 4px;font-size:14.5px;color:var(--maroon);font-weight:650}
.panel .note{font-size:11.5px;color:var(--gray);margin-bottom:10px}
.panel canvas{max-height:330px}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;color:var(--gray);font-weight:600;border-bottom:2px solid #e3e8ee;padding:6px 7px;cursor:pointer;white-space:nowrap;user-select:none}
th:hover{color:var(--blue)}
td{padding:5px 7px;border-bottom:1px solid #eef1f5;white-space:nowrap}
td.r,th.r{text-align:right;font-variant-numeric:tabular-nums}
td.doc{max-width:190px;overflow:hidden;text-overflow:ellipsis}
.pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;background:rgba(43,124,190,.1);color:var(--blue);margin:0 6px 6px 0}
.mbtn{border:1px solid #d4dbe3;background:#fff;color:#243342;border-radius:6px;padding:4px 12px;margin-right:6px;cursor:pointer;font-size:12.5px}
.mbtn.on{background:var(--blue);color:#fff;border-color:var(--blue)}
.tag{font-size:10.5px;padding:1px 7px;border-radius:10px}
.tag.g{background:rgba(26,127,78,.12);color:var(--good)}.tag.r{background:rgba(192,57,43,.1);color:var(--bad)}.tag.y{background:rgba(127,140,155,.15);color:#5a6875}
footer{font-size:11px;color:var(--gray);margin-top:26px;line-height:1.6}
.scroll{overflow-x:auto}
</style></head><body>
<header><div><h1>VPS Lakeshore · Daily Revenue Dashboard</h1>
<div class="sub">Lakeshore Hospital &amp; Research Centre Ltd, Kochi · Global Lifecare</div></div>
<div class="sub" id="asof"></div></header>
<div class="wrap">
<div class="cards" id="cards"></div>

<div class="panel"><h2>Daily Gross Revenue — Actual vs Budget</h2>
<div class="note">OP / IP / Pharmacy stacked; line = budgeted total.</div>
<div id="mbtns" style="margin-bottom:8px"></div><canvas id="dailyChart"></canvas></div>

<div class="grid">
<div class="panel"><h2>Monthly Revenue — FY 26-27 vs Budget vs FY 25-26</h2>
<div class="note">₹ Cr. Current-month bar is month-to-date.</div><canvas id="monthlyChart"></canvas></div>
<div class="panel"><h2>Efficiency — ARPOB · Occupancy · ALOS</h2>
<div class="note">ARPOB = gross revenue ÷ occupied bed-days (₹ ’000/day). Bars = occupancy %, dotted = ALOS (days).</div><canvas id="effChart"></canvas></div>
<div class="panel"><h2>Collections — Cash + Credit vs Budget</h2>
<div class="note">₹ Cr per month, FY 26-27. Line marker = collections ÷ gross revenue (cash conversion).</div><canvas id="collChart"></canvas></div>
<div class="panel"><h2>Volumes — OP Visits &amp; IP Discharges</h2>
<div class="note">FY 26-27 monthly vs budget (dotted).</div><canvas id="volChart"></canvas></div>
</div>

<div class="panel"><h2>Doctor League Table — Month on Month</h2>
<div class="note" id="dlnote"></div>
<div style="margin-bottom:8px"><input id="docFilter" placeholder="Filter doctor / department…" style="padding:5px 10px;border:1px solid #d4dbe3;border-radius:6px;width:260px;font-size:12.5px"></div>
<div class="scroll"><table id="league"><thead><tr>
<th data-k="doc">Doctor</th><th data-k="dept">Department</th>
<th class="r" data-k="revC">Rev cur (₹L)</th><th class="r" data-k="rrC">₹L/day cur</th>
<th class="r" data-k="revP">Rev prev (₹L)</th><th class="r" data-k="rrP">₹L/day prev</th>
<th class="r" data-k="mom">Δ run-rate</th>
<th class="r" data-k="share">Share</th><th class="r" data-k="rkD">Rank Δ</th>
<th class="r" data-k="opC">OP cur</th><th class="r" data-k="opP">OP prev</th>
<th class="r" data-k="admC">Adm cur</th><th class="r" data-k="disC">Disch cur</th>
<th class="r" data-k="conv">Conv %</th><th class="r" data-k="ipmix">IP mix*</th>
<th class="r" data-k="newp">New OP*</th><th class="r" data-k="freep">Free OP*</th>
<th class="r" data-k="rpd">₹L/disch</th><th class="r" data-k="alos">ALOS*</th>
<th class="r" data-k="recov">Recovered*</th><th class="r" data-k="dama">DAMA*</th><th class="r" data-k="exp">Expired*</th>
<th class="r" data-k="cash">Cash mix*</th>
</tr></thead><tbody></tbody></table></div></div>

<div class="panel"><h2>Department League Table — Month on Month</h2>
<div class="note" id="dtnote"></div>
<div class="scroll"><table id="deptTable"><thead><tr>
<th data-k="dept">Department</th>
<th class="r" data-k="revC">Rev cur (₹L)</th><th class="r" data-k="rrC">₹L/day cur</th>
<th class="r" data-k="revP">Rev prev (₹L)</th><th class="r" data-k="rrP">₹L/day prev</th>
<th class="r" data-k="mom">Δ run-rate</th><th class="r" data-k="share">Share</th>
<th class="r" data-k="opC">OP cur</th><th class="r" data-k="admC">Adm cur</th>
<th class="r" data-k="disC">Disch cur</th><th class="r" data-k="conv">Conv %</th>
<th class="r" data-k="rpd">₹L/disch</th><th class="r" data-k="alos">ALOS*</th>
<th class="r" data-k="arpob">ARPOB† ₹k</th><th class="r" data-k="cash">Cash mix*</th>
</tr></thead><tbody></tbody></table></div></div>

<div class="grid">
<div class="panel"><h2>Doctor Concentration — Cumulative Revenue Share</h2>
<div class="note" id="concnote"></div><canvas id="concChart"></canvas></div>
<div class="panel"><h2>Department Revenue — Current vs Previous Month (run-rate)</h2>
<div class="note">₹ Lakhs per day, doctor-attributed revenue from Daily MIS.</div><canvas id="deptMoM" style="max-height:400px"></canvas></div>
<div class="panel"><h2>Discharge Outcomes &amp; Payer Mix</h2>
<div class="note" id="disnote"></div>
<div style="display:flex;gap:10px"><div style="width:50%"><canvas id="statusChart"></canvas></div>
<div style="width:50%"><canvas id="payerChart"></canvas></div></div></div>
<div class="panel"><h2>Department Revenue — Captured Flash Days</h2>
<div class="note" id="deptnote"></div><canvas id="deptChart" style="max-height:400px"></canvas></div>
</div>
<footer id="foot"></footer>
</div>
<script>
const D = __DATA__;
const CR=1e7, L=1e5;
const fmtCr=v=>'₹'+(v/CR).toFixed(2)+' Cr';
const pct=(a,b)=>b? ((a/b-1)*100):0;
const BLUE='#2B7CBE',MAROON='#8B1A4A',GRAY='#7F8C9B',LT='rgba(43,124,190,.35)',GOLD='#c8952b';
const tc=s=>s.split(' ').map(w=>w.includes('.')&&w.length<=5?w:w.charAt(0)+w.slice(1).toLowerCase()).join(' ');
document.getElementById('asof').textContent='Data through '+D.latestDate+' · generated '+D.generated;

const fyKeys=Object.keys(D.yearTables);
const cur=D.yearTables[fyKeys[0]]||{months:[]}, prev=D.yearTables[fyKeys[1]]||{months:[]};
const mwd=cur.months.filter(m=>m.revTot>0);
const ytd={rev:0,bud:0,coll:0,budColl:0,opv:0,ipd:0};
mwd.forEach(m=>{ytd.rev+=m.revTot;ytd.bud+=m.budTot;ytd.coll+=m.collTot;ytd.budColl+=m.budCollTot;ytd.opv+=m.opVisits;ytd.ipd+=m.ipDisch;});
const mtd=mwd[mwd.length-1]||{};
const prevYtdRev=prev.months.slice(0,mwd.length).reduce((a,m)=>a+m.revTot,0);

// history months
const hMonths=Object.keys(D.history).sort();
const curM=hMonths[hMonths.length-1], prevM=hMonths[hMonths.length-2];
const hCur=curM? D.history[curM]:null, hPrev=prevM? D.history[prevM]:null;
function daysIn(mk,h){const [y,m]=mk.split('-').map(Number);
 const last=new Date(y,m,0).getDate();
 const asOfDay=h&&h.daysElapsed? h.daysElapsed:last;
 return Math.min(asOfDay,last);}
const dCur=curM? daysIn(curM,hCur):1, dPrev=prevM? daysIn(prevM,hPrev):1;
const mName=mk=>{const [y,m]=mk.split('-');return new Date(y,m-1,1).toLocaleString('en',{month:'short'})+'-'+y.slice(2)};

// efficiency series (FY26-27 from MoM sheet + fallback to flash occDays for prior FY)
const effM=Object.keys(D.momFY).sort();
const eff=effM.map(mk=>{const e=D.momFY[mk];
 const ym=mk; let rev=0;
 const mi=cur.months[(parseInt(mk.split('-')[1])-4+12)%12];
 if(mi) rev=mi.revTot;
 return {mk, arpob: e.occDays? rev/e.occDays/1000:null,
         occ: e.bedCap? e.occDays/e.bedCap*100:null, alos:e.alos||null};});

// KPI cards
function card(lbl,val,delta,cls,m){return `<div class="card${m?' m':''}"><div class="lbl">${lbl}</div><div class="val">${val}</div><div class="delta ${cls}">${delta}</div></div>`}
const dv=(a,b,sfx)=>{const p=pct(a,b);return [(p>=0?'▲ +':'▼ ')+p.toFixed(1)+'% '+sfx,p>=0?'up':'dn']};
let [d1,c1]=dv(mtd.revTot,mtd.budTot,'vs budget');
let [d2,c2]=dv(ytd.rev,ytd.bud,'vs budget');
let [d3,c3]=dv(ytd.rev,prevYtdRev,'YoY');
let [d4,c4]=dv(ytd.coll,ytd.budColl,'vs budget');
const bo=(D.summary.bedOcc||[0,0,0]);
const lastEff=eff.filter(e=>e.arpob).slice(-1)[0]||{};
const prevEff=eff.filter(e=>e.arpob).slice(-2)[0]||{};
const conv=(D.summary.real&&D.summary.real['IP Conv. Rate'])||[0,0];
document.getElementById('cards').innerHTML=
 card('MTD Revenue ('+(mtd.month||'')+')',fmtCr(mtd.revTot||0),d1,c1)+
 card('YTD Revenue FY 26-27',fmtCr(ytd.rev),d2,c2)+
 card('YTD YoY',fmtCr(prevYtdRev)+' LY',d3,c3,1)+
 card('YTD Collections',fmtCr(ytd.coll),d4,c4,1)+
 card('ARPOB (MTD)','₹'+(lastEff.arpob||0).toFixed(1)+'k','prev mo ₹'+(prevEff.arpob||0).toFixed(1)+'k','')+
 card('ALOS (MTD)',(lastEff.alos||0).toFixed(1)+' d','prev mo '+(prevEff.alos||0).toFixed(1)+' d','')+
 card('Bed Occupancy (MTD)',(bo[1]*100).toFixed(0)+'%','YTD '+(bo[2]*100).toFixed(0)+'%','')+
 card('OP→IP Conversion',(conv[0]*100).toFixed(1)+'%','YTD '+(conv[1]*100).toFixed(1)+'%','');

// daily chart
const dailyKeys=Object.keys(D.daily).sort();
let dailyChart=null;
window.drawDaily=function(key){
 const rows=D.daily[key];
 if(dailyChart)dailyChart.destroy();
 dailyChart=new Chart(document.getElementById('dailyChart'),{data:{labels:rows.map(r=>r.date.slice(8)+' '+r.dow.slice(0,2)),datasets:[
  {type:'bar',label:'OP',data:rows.map(r=>r.revOP/L),backgroundColor:LT,stack:'a'},
  {type:'bar',label:'IP',data:rows.map(r=>r.revIP/L),backgroundColor:BLUE,stack:'a'},
  {type:'bar',label:'Pharmacy',data:rows.map(r=>r.revPH/L),backgroundColor:MAROON,stack:'a'},
  {type:'line',label:'Budget',data:rows.map(r=>r.budTot/L),borderColor:GRAY,borderDash:[5,4],pointRadius:0,tension:.2}]},
  options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
  scales:{x:{stacked:true,ticks:{font:{size:10}}},y:{stacked:true,title:{display:true,text:'₹ Lakhs'}}}}});
 document.querySelectorAll('.mbtn').forEach(b=>b.classList.toggle('on',b.dataset.k===key));
}
document.getElementById('mbtns').innerHTML=dailyKeys.map(k=>`<button class="mbtn" data-k="${k}" onclick="drawDaily('${k}')">${k}</button>`).join('');
if(dailyKeys.length)drawDaily(dailyKeys[dailyKeys.length-1]);

// monthly chart
const mlabels=cur.months.map(m=>m.month.slice(0,3));
new Chart(document.getElementById('monthlyChart'),{data:{labels:mlabels,datasets:[
 {type:'bar',label:fyKeys[0]+' Actual',data:cur.months.map(m=>m.revTot/CR),backgroundColor:BLUE},
 {type:'bar',label:'Budget',data:cur.months.map(m=>m.budTot/CR),backgroundColor:'rgba(127,140,155,.35)'},
 {type:'line',label:fyKeys[1]+' Actual',data:prev.months.map(m=>m.revTot/CR),borderColor:MAROON,pointRadius:2,tension:.25}]},
 options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},scales:{y:{title:{display:true,text:'₹ Cr'}}}}});

// efficiency chart
new Chart(document.getElementById('effChart'),{data:{labels:eff.map(e=>mName(e.mk)),datasets:[
 {type:'bar',label:'Occupancy %',data:eff.map(e=>e.occ),backgroundColor:LT,yAxisID:'y2'},
 {type:'line',label:'ARPOB ₹’000/day',data:eff.map(e=>e.arpob),borderColor:BLUE,pointRadius:3,tension:.25,yAxisID:'y'},
 {type:'line',label:'ALOS (days)',data:eff.map(e=>e.alos),borderColor:MAROON,borderDash:[5,4],pointRadius:3,yAxisID:'y3'}]},
 options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
 scales:{y:{title:{display:true,text:'ARPOB ₹’000'}},y2:{position:'right',grid:{drawOnChartArea:false},max:100,title:{display:true,text:'Occ %'}},y3:{display:false,min:0,max:8}}}});

// collections
const convLine=cur.months.map(m=>m.revTot? m.collTot/m.revTot*100:null);
new Chart(document.getElementById('collChart'),{data:{labels:mlabels,datasets:[
 {type:'bar',label:'Cash',data:cur.months.map(m=>m.collCash/CR),backgroundColor:BLUE,stack:'c'},
 {type:'bar',label:'Credit',data:cur.months.map(m=>m.collCredit/CR),backgroundColor:MAROON,stack:'c'},
 {type:'line',label:'Budgeted collections',data:cur.months.map(m=>m.budCollTot/CR),borderColor:GRAY,borderDash:[5,4],pointRadius:0},
 {type:'line',label:'Coll ÷ Revenue %',data:convLine,borderColor:GOLD,pointRadius:3,yAxisID:'y2'}]},
 options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
 scales:{x:{stacked:true},y:{stacked:true,title:{display:true,text:'₹ Cr'}},y2:{position:'right',grid:{drawOnChartArea:false},min:0,max:120,title:{display:true,text:'%'}}}}});

// volumes
new Chart(document.getElementById('volChart'),{data:{labels:mlabels,datasets:[
 {type:'bar',label:'OP+ER visits',data:cur.months.map(m=>m.opVisits||null),backgroundColor:LT,yAxisID:'y'},
 {type:'line',label:'OP budget',data:cur.months.map(m=>m.budOPv||null),borderColor:GRAY,borderDash:[4,4],pointRadius:0,yAxisID:'y'},
 {type:'line',label:'IP discharges',data:cur.months.map(m=>m.ipDisch||null),borderColor:MAROON,pointRadius:3,yAxisID:'y2'},
 {type:'line',label:'IP budget',data:cur.months.map(m=>m.budIPd||null),borderColor:MAROON,borderDash:[4,4],pointRadius:0,yAxisID:'y2'}]},
 options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
 scales:{y:{title:{display:true,text:'OP visits'}},y2:{position:'right',grid:{drawOnChartArea:false},title:{display:true,text:'IP discharges'}}}}});

// ---------------- doctor league table ----------------
let rowsL=[];
if(hCur){
 const docsC=hCur.doctors, docsP=hPrev? hPrev.doctors:{};
 const totC=Object.values(docsC).reduce((a,v)=>a+(v.rev||0),0);
 const rankOf=docs=>{const o={};Object.entries(docs).sort((a,b)=>(b[1].rev||0)-(a[1].rev||0))
  .forEach(([n],i)=>o[n]=i+1);return o};
 const rkC=rankOf(docsC), rkP=rankOf(docsP);
 const names=new Set([...Object.keys(docsC),...Object.keys(docsP)]);
 names.forEach(n=>{
  const c=docsC[n]||{}, p=docsP[n]||{};
  const dd=D.disByDoc[n]||null;
  const tm=D.docTypeMix[n]||null, oa=D.opAgg[n]||null;
  const revC=(c.rev||0), revP=(p.rev||0);
  const rrC=revC/dCur, rrP=revP/dPrev;
  const st=dd? dd.status:{};
  const tmTot=tm? (tm.IP+tm.OP+tm.PH):0;
  rowsL.push({doc:tc(n),dept:tc(c.dept||p.dept||''),
   revC:revC/L, rrC:rrC/L, revP:revP/L, rrP:rrP/L,
   mom:rrP? (rrC/rrP-1)*100:null,
   share:totC? revC/totC*100:null,
   rkD:(rkC[n]&&rkP[n])? rkP[n]-rkC[n]:null,
   opC:c.opv||0, opP:p.opv||0, admC:c.adm||0, disC:c.dis||0,
   conv:c.opv? (c.adm||0)/c.opv*100:null,
   ipmix:tmTot? tm.IP/tmTot*100:null,
   newp:oa&&oa.tot? oa.new/oa.tot*100:null,
   freep:oa&&oa.tot? oa.free/oa.tot*100:null,
   rpd:c.dis? revC/c.dis/L:null,
   alos:dd&&dd.losN? dd.losSum/dd.losN:null,
   recov:dd&&dd.n? (st['Recovered']||0)/dd.n*100:null,
   dama:dd? (st['DAMA/LAMA']||0):null, exp:dd? (st['Expired']||0):null,
   cash:dd&&dd.n? dd.cash/dd.n*100:null});
 });
 rowsL=rowsL.filter(r=>r.revC>0.01||r.revP>0.01);
 document.getElementById('dlnote').innerHTML=
  `Current = <b>${mName(curM)}</b> (${dCur} days elapsed) vs previous = <b>${prevM?mName(prevM):'—'}</b> (${dPrev} days). `+
  `Revenue/visits/admissions/discharges from the Daily MIS doctor sheets. Conv % = admissions ÷ OP visits (current month). Rank Δ = movement in revenue rank vs previous month. Columns marked * come from the flash detail sheets over ${D.disDates.length} captured days: IP mix (share of doctor revenue billed as IP), New/Free OP visit shares, and from ${D.nDischarges} discharges — status, ALOS and cash mix. Click a header to sort.`;
}
let sortK='revC',sortDir=-1;
function renderLeague(){
 const q=(document.getElementById('docFilter').value||'').toUpperCase();
 let rows=rowsL.filter(r=>!q||r.doc.toUpperCase().includes(q)||r.dept.toUpperCase().includes(q));
 rows.sort((a,b)=>{const x=a[sortK],y=b[sortK];
  if(x==null&&y==null)return 0; if(x==null)return 1; if(y==null)return -1;
  return (x<y?-1:x>y?1:0)*(typeof x==='string'?-sortDir:sortDir);});
 const f=(v,d=1)=>v==null?'—':v.toFixed(d);
 const momCell=v=>v==null?'—':`<span class="tag ${v>=10?'g':v<=-10?'r':'y'}">${v>=0?'+':''}${v.toFixed(0)}%</span>`;
 document.querySelector('#league tbody').innerHTML=rows.slice(0,60).map(r=>
  `<tr><td class="doc" title="${r.doc}">${r.doc}</td><td class="doc">${r.dept}</td>`+
  `<td class="r"><b>${f(r.revC)}</b></td><td class="r">${f(r.rrC,2)}</td>`+
  `<td class="r">${f(r.revP)}</td><td class="r">${f(r.rrP,2)}</td>`+
  `<td class="r">${momCell(r.mom)}</td>`+
  `<td class="r">${r.share==null?'—':r.share.toFixed(1)+'%'}</td>`+
  `<td class="r">${r.rkD==null?'—':(r.rkD>0?'▲'+r.rkD:r.rkD<0?'▼'+(-r.rkD):'=')}</td>`+
  `<td class="r">${r.opC||'—'}</td><td class="r">${r.opP||'—'}</td>`+
  `<td class="r">${r.admC||'—'}</td><td class="r">${r.disC||'—'}</td>`+
  `<td class="r">${r.conv==null?'—':r.conv.toFixed(1)+'%'}</td>`+
  `<td class="r">${r.ipmix==null?'—':r.ipmix.toFixed(0)+'%'}</td>`+
  `<td class="r">${r.newp==null?'—':r.newp.toFixed(0)+'%'}</td>`+
  `<td class="r">${r.freep==null?'—':r.freep.toFixed(0)+'%'}</td>`+
  `<td class="r">${f(r.rpd,2)}</td><td class="r">${f(r.alos)}</td>`+
  `<td class="r">${r.recov==null?'—':r.recov.toFixed(0)+'%'}</td>`+
  `<td class="r">${r.dama==null?'—':r.dama}</td><td class="r">${r.exp==null?'—':r.exp}</td>`+
  `<td class="r">${r.cash==null?'—':r.cash.toFixed(0)+'%'}</td></tr>`).join('');
}
document.querySelectorAll('#league th').forEach(th=>th.onclick=()=>{
 const k=th.dataset.k; if(sortK===k)sortDir*=-1; else {sortK=k;sortDir=-1;} renderLeague();});
document.getElementById('docFilter').oninput=renderLeague;
renderLeague();

// ---------------- department league table ----------------
if(hCur){
 const agg=(h,days)=>{const m={};Object.values(h.doctors).forEach(v=>{
  const k=v.dept||'UNALLOCATED';
  const r=m[k]=m[k]||{rev:0,opv:0,adm:0,dis:0};
  r.rev+=(v.rev||0);r.opv+=(v.opv||0);r.adm+=(v.adm||0);r.dis+=(v.dis||0);});return m};
 const aC=agg(hCur,dCur), aP=hPrev? agg(hPrev,dPrev):{};
 // dept-level discharge stats via doctor->dept map
 const d2d={};Object.entries(hCur.doctors).forEach(([n,v])=>{if(v.dept)d2d[n]=v.dept});
 const dStats={};
 Object.entries(D.disByDoc).forEach(([n,v])=>{
  const k=d2d[n]; if(!k)return;
  const s=dStats[k]=dStats[k]||{n:0,losSum:0,losN:0,cash:0};
  s.n+=v.n;s.losSum+=v.losSum;s.losN+=v.losN;s.cash+=v.cash;});
 const totC=Object.values(aC).reduce((a,v)=>a+v.rev,0);
 let rowsD=Object.keys(aC).map(k=>{
  const c=aC[k],p=aP[k]||{rev:0};
  const rrC=c.rev/dCur, rrP=p.rev? p.rev/dPrev:0;
  const s=dStats[k];
  const alos=s&&s.losN? s.losSum/s.losN:null;
  return {dept:tc(k),revC:c.rev/L,rrC:rrC/L,revP:(p.rev||0)/L,rrP:rrP/L,
   mom:rrP? (rrC/rrP-1)*100:null, share:totC? c.rev/totC*100:null,
   opC:c.opv,admC:c.adm,disC:c.dis,
   conv:c.opv? c.adm/c.opv*100:null,
   rpd:c.dis? c.rev/c.dis/L:null,
   alos:alos,
   arpob:(alos&&c.dis)? c.rev/(c.dis*alos)/1000:null,
   cash:s&&s.n? s.cash/s.n*100:null};
 }).filter(r=>r.revC>0.5);
 document.getElementById('dtnote').innerHTML=
  `Current = <b>${mName(curM)}</b> (${dCur} days) vs <b>${prevM?mName(prevM):'—'}</b> (${dPrev} days), doctor-attributed revenue rolled up to department. `+
  `* from flash discharge lists (captured days only). † ARPOB proxy = revenue ÷ (discharges × ALOS*) — departmental bed-days are not reported, so treat as directional. Click headers to sort.`;
 let sK='revC',sD=-1;
 const fmt=(v,d=1)=>v==null?'—':v.toFixed(d);
 function renderDept(){
  rowsD.sort((a,b)=>{const x=a[sK],y=b[sK];
   if(x==null&&y==null)return 0; if(x==null)return 1; if(y==null)return -1;
   return (x<y?-1:x>y?1:0)*(typeof x==='string'?-sD:sD);});
  const momCell=v=>v==null?'—':`<span class="tag ${v>=10?'g':v<=-10?'r':'y'}">${v>=0?'+':''}${v.toFixed(0)}%</span>`;
  document.querySelector('#deptTable tbody').innerHTML=rowsD.map(r=>
   `<tr><td class="doc">${r.dept}</td><td class="r"><b>${fmt(r.revC)}</b></td><td class="r">${fmt(r.rrC,2)}</td>`+
   `<td class="r">${fmt(r.revP)}</td><td class="r">${fmt(r.rrP,2)}</td>`+
   `<td class="r">${momCell(r.mom)}</td><td class="r">${r.share==null?'—':r.share.toFixed(1)+'%'}</td>`+
   `<td class="r">${r.opC||'—'}</td><td class="r">${r.admC||'—'}</td><td class="r">${r.disC||'—'}</td>`+
   `<td class="r">${r.conv==null?'—':r.conv.toFixed(1)+'%'}</td>`+
   `<td class="r">${fmt(r.rpd,2)}</td><td class="r">${fmt(r.alos)}</td>`+
   `<td class="r">${fmt(r.arpob)}</td><td class="r">${r.cash==null?'—':r.cash.toFixed(0)+'%'}</td></tr>`).join('');
 }
 document.querySelectorAll('#deptTable th').forEach(th=>th.onclick=()=>{
  const k=th.dataset.k; if(sK===k)sD*=-1; else {sK=k;sD=-1;} renderDept();});
 renderDept();
}

// doctor concentration (current month)
if(hCur){
 const arr=Object.entries(hCur.doctors).map(([n,v])=>({n:tc(n),r:v.rev||0}))
  .filter(x=>x.r>0).sort((a,b)=>b.r-a.r);
 const tot=arr.reduce((a,x)=>a+x.r,0);
 let cum=0; const top=arr.slice(0,20).map(x=>{cum+=x.r;return {n:x.n,r:x.r,c:cum/tot*100}});
 const t5=arr.slice(0,5).reduce((a,x)=>a+x.r,0)/tot*100, t10=arr.slice(0,10).reduce((a,x)=>a+x.r,0)/tot*100;
 document.getElementById('concnote').textContent=
  `${mName(curM)}: top-5 doctors = ${t5.toFixed(0)}% and top-10 = ${t10.toFixed(0)}% of doctor-attributed revenue (${arr.length} active doctors). Key-person concentration watch.`;
 new Chart(document.getElementById('concChart'),{data:{labels:top.map(x=>x.n),datasets:[
  {type:'bar',label:'Revenue ₹L',data:top.map(x=>x.r/L),backgroundColor:BLUE,yAxisID:'y'},
  {type:'line',label:'Cumulative %',data:top.map(x=>x.c),borderColor:MAROON,pointRadius:2,yAxisID:'y2'}]},
  options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
  scales:{x:{ticks:{font:{size:9},maxRotation:75,minRotation:60}},y:{title:{display:true,text:'₹ L'}},
  y2:{position:'right',min:0,max:100,grid:{drawOnChartArea:false},title:{display:true,text:'Cum %'}}}}});
}

// dept MoM run-rate
if(hCur){
 const agg=(h,days)=>{const m={};Object.values(h.doctors).forEach(v=>{if(v.dept)m[v.dept]=(m[v.dept]||0)+(v.rev||0)/days});return m};
 const a=agg(hCur,dCur), b=hPrev? agg(hPrev,dPrev):{};
 const keys=Object.keys(a).sort((x,y)=>a[y]-a[x]).slice(0,12);
 new Chart(document.getElementById('deptMoM'),{type:'bar',data:{labels:keys.map(tc),datasets:[
  {label:mName(curM)+' ₹L/day',data:keys.map(k=>(a[k]||0)/L),backgroundColor:BLUE},
  {label:(prevM?mName(prevM):'prev')+' ₹L/day',data:keys.map(k=>(b[k]||0)/L),backgroundColor:'rgba(139,26,74,.55)'}]},
  options:{indexAxis:'y',plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
  scales:{x:{title:{display:true,text:'₹ Lakhs / day'}},y:{ticks:{font:{size:10.5}}}}}});
}

// discharge outcomes + payer
document.getElementById('disnote').textContent=
 `${D.nDischarges} discharges across ${D.disDates.length} captured flash days (${D.disDates[0]||''} → ${D.disDates[D.disDates.length-1]||''}). Left: clinical status. Right: payer scheme mix.`;
const sm=D.statusMix, pm=D.payerMix;
new Chart(document.getElementById('statusChart'),{type:'doughnut',
 data:{labels:Object.keys(sm),datasets:[{data:Object.values(sm),backgroundColor:[ '#1a7f4e',MAROON,'#c8952b',BLUE,GRAY]}]},
 options:{plugins:{legend:{position:'bottom',labels:{boxWidth:11,font:{size:10.5}}}}}});
new Chart(document.getElementById('payerChart'),{type:'doughnut',
 data:{labels:Object.keys(pm),datasets:[{data:Object.values(pm),backgroundColor:[BLUE,MAROON,'#c8952b',GRAY]}]},
 options:{plugins:{legend:{position:'bottom',labels:{boxWidth:11,font:{size:10.5}}}}}});

// dept pareto (flash captured days)
document.getElementById('deptnote').textContent='Net service revenue summed over '+D.deptDates.length+' captured day(s).';
new Chart(document.getElementById('deptChart'),{type:'bar',data:{labels:D.deptTop.map(d=>d.name),
 datasets:[{label:'Revenue (₹ L)',data:D.deptTop.map(d=>d.rev/L),backgroundColor:BLUE}]},
 options:{indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{title:{display:true,text:'₹ Lakhs'}},y:{ticks:{font:{size:10.5}}}}}});

document.getElementById('foot').innerHTML='<b>Files parsed:</b> '+D.filesParsed.map(f=>'<span class="pill">'+f+'</span>').join('')+
 '<br>Gross revenue per Daily Revenue Flash. Doctor-month figures from Daily MIS doctor sheets (newest file per month; current month is MTD — compare on ₹/day run-rate). Discharge status / ALOS / payer mix cover only dates with a flash file on hand. ARPOB = gross revenue ÷ occupied bed-days; ALOS from MIS MoM sheet. Operations include VPSLMC (satellite) figures where the source does.';
</script></body></html>
"""

if __name__ == "__main__":
    main()
