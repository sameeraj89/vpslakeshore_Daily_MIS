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
import sys, os, re, json, glob, datetime, fnmatch
import openpyxl

# Collect files matching `pattern` under `root`, recursing into active month
# subfolders (July/, June/, ...) but SKIPPING the historical archive and the
# tools dir so we don't sweep in years of old MIS files.
_SKIP_DIRS = {"Base Reports Daily & Monthly MIS", "_dashboard_tools",
              "dailyrevenueflashsourcereports"}
def collect_files(root, pattern):
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames
                       if d not in _SKIP_DIRS and not d.startswith('.')]
        for fn in filenames:
            if fnmatch.fnmatch(fn, pattern):
                out.append(os.path.join(dirpath, fn))
    return out

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
    day_of = {j: rows[hrow][j].day if isinstance(rows[hrow][j], datetime.date)
              else rows[hrow][j].date().day for j in datecols}
    out, daily, last_dept = {}, {}, ""
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
        dd = daily.setdefault(norm(doc), {})
        for j in datecols:
            if j < len(r) and num(r[j]):
                dd[day_of[j]] = dd.get(day_of[j], 0) + num(r[j])
    return out, daily

def parse_calicut_input(ws):
    """'Input Calicut' (service group x day) -> ({service: month_total}, {iso_date: day_total})"""
    rows = list(ws.iter_rows(values_only=True))
    hrow, datecols = None, []
    for i, r in enumerate(rows[:6]):
        dc = [(j, c.date() if isinstance(c, datetime.datetime) else c)
              for j, c in enumerate(r) if isinstance(c, (datetime.datetime, datetime.date))]
        if len(dc) >= 5:
            months = {}
            for _, c in dc: months[(c.year, c.month)] = months.get((c.year, c.month), 0) + 1
            dom = max(months, key=months.get)
            seen = set()
            for j, c in dc:
                if (c.year, c.month) == dom and c not in seen:
                    seen.add(c); datecols.append((j, c))
            hrow = i; break
    if hrow is None: return {}, {}
    serv, daily = {}, {}
    for r in rows[hrow + 1:]:
        lbl = str(r[0]).strip() if r[0] else ""
        if not lbl: continue
        if norm(lbl) in ("GRAND TOTAL", "TOTAL"):
            for j, c in datecols:
                if j < len(r) and num(r[j]):
                    daily[c.strftime("%Y-%m-%d")] = num(r[j])
            break
        tot = sum(num(r[j]) for j, _ in datecols if j < len(r))
        if tot: serv[lbl] = serv.get(lbl, 0) + tot
    return serv, daily

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

def parse_pnl_summ(ws):
    """Monthly management P&L ('12P+L Summ') -> (month 'YYYY-MM', {line: {a,aPct,b,bPct,var}}).
    Columns: E(4)=label, G(6)=Actual, H(7)=Actual%, I(8)=Budget, J(9)=Budget%, K(10)=Variance.
    Real audited close per month; one month per monthly MIS file."""
    rows = list(ws.iter_rows(values_only=True, max_col=12))
    # month = first datetime found in the header band
    mkey = None
    for r in rows[:12]:
        for c in r:
            if isinstance(c, (datetime.datetime, datetime.date)):
                mkey = c.strftime("%Y-%m"); break
        if mkey: break
    # ordered (key, label-prefix) — first numeric-actual match wins
    specs = [("revenue", "Total Revenue"), ("directCost", "Direct Costs"),
             ("contribution", "Net Revenue"), ("staff", "Staff Costs"),
             ("overheads", "Overheads"), ("badDebt", "Provision for Bad"),
             ("totalExp", "Total Expenses"), ("ebitda", "Operating Profit"),
             ("nonOp", "Non-Operating"), ("netOp", "Net Operating Profit"),
             ("finance", "Finance Charges"), ("cashProfit", "Cash Profit"),
             ("depreciation", "Depreciation"), ("net", "Net Profit")]
    out = {}
    for key, pref in specs:
        for r in rows:
            lbl = str(r[4]).strip() if len(r) > 4 and r[4] else ""
            if not lbl.startswith(pref):
                continue
            if pref in ("Finance Charges", "Depreciation") and "Right" in lbl:
                continue
            if pref == "Net Profit" and "AFS" in lbl:
                continue
            a = num(r[6]) if len(r) > 6 else 0
            b = num(r[8]) if len(r) > 8 else 0
            if a == 0 and b == 0 and key not in ("nonOp", "finance", "badDebt"):
                continue
            out[key] = {"a": a, "aPct": num(r[7]) if len(r) > 7 else 0,
                        "b": b, "bPct": num(r[9]) if len(r) > 9 else 0,
                        "var": num(r[10]) if len(r) > 10 else 0}
            break
    return mkey, out

_MONTHS3 = {m[:3].lower(): i for i, m in enumerate(
    ["January","February","March","April","May","June","July","August",
     "September","October","November","December"], 1)}

def _pl_month(cell):
    """Normalise a header cell to 'YYYY-MM' (handles dates and 'jun 26' text)."""
    if isinstance(cell, (datetime.datetime, datetime.date)):
        return cell.strftime("%Y-%m")
    s = str(cell).strip().lower()
    m = re.match(r'([a-z]{3,})[\s\-]*(\d{2,4})', s)
    if m and m.group(1)[:3] in _MONTHS3:
        yr = int(m.group(2)); yr += 2000 if yr < 100 else 0
        return f"{yr}-{_MONTHS3[m.group(1)[:3]]:02d}"
    m = re.match(r'(\d{4})-(\d{1,2})', s)
    if m: return f"{int(m.group(1))}-{int(m.group(2)):02d}"
    return None

# (canonical label, lowercase startswith-prefix, group, is_total)
PL_SPEC = [
    ("IP Revenue","ip revenue","rev",False), ("IP Pharmacy","ip pharmacy","rev",False),
    ("OP Revenue","op revenue","rev",False), ("OP Pharmacy","op pharmacy","rev",False),
    ("Ayurveda","aurveda","rev",False), ("F & B Revenue","f & b","rev",False),
    ("Revenue From Operations","revenue from op","rev",True),
    ("Other Income","other income","rev",False), ("Total Revenue","total revenue","rev",True),
    ("Material Cost","material cost","cost",False), ("Doctor Cost","doctor cost","cost",False),
    ("Employee Cost","employee cost","cost",False), ("Utilities (Power, Fuel)","utilities","cost",False),
    ("Repair & Maintenance","repair & main","cost",False), ("Marketing Cost","marketing","cost",False),
    ("Lab Test Charges","lab test","cost",False), ("Rent Paid","rent paid","cost",False),
    ("House Keeping","house keeping","cost",False), ("Security Expenses","security","cost",False),
    ("Printing & Stationery","printing","cost",False), ("Insurance","insurance","cost",False),
    ("Rates & Taxes","rates & tax","cost",False), ("Professional Fee","professional fee","cost",False),
    ("Quality & Infection Control","quality","cost",False),
    ("Provision for Bad Debts","provis","cost",False), ("Communication Costs","communication","cost",False),
    ("Business Travel Cost","business travel","cost",False),
    ("Miscellaneous Expenses","miscellaneous","cost",False), ("CSR Activity","csr","cost",False),
    ("Total Expenses","total expenses","cost",True),
    ("EBITDA","ebitda","profit",True), ("Depreciation","depreciation","profit",False),
    ("EBIT (Operating Profit)","ebit (","profit",True), ("Finance Cost","finance cost","profit",False),
    ("PBT","pbt","profit",True), ("Tax","tax","profit",False), ("PAT","pat","profit",True),
]

def parse_pl_consol(ws):
    """Consolidated monthly P&L ('PL_Consol' in Profit_and_Loss_Statement files).
    Columns shift per file, so locate CM/Budget/SMLY dynamically from the group +
    header rows. Values are in ₹ Lakhs -> stored as absolute INR.
    Returns (month 'YYYY-MM', {'lines':[{label,grp,tot,a,b,smly}], 'src':...})."""
    rows = list(ws.iter_rows(values_only=True))
    hrow = next((i for i, r in enumerate(rows[:14])
                 if any(isinstance(c, str) and c.strip() == "Particulars" for c in r)), None)
    if hrow is None or hrow == 0: return None, None
    g = [str(c).strip().lower() if c else "" for c in rows[hrow - 1]]
    h = rows[hrow]
    cm = g.index("cm") if "cm" in g else None
    bud = g.index("budget") if "budget" in g else None
    if cm is None or bud is None: return None, None
    cm_month = _pl_month(h[cm])
    if not cm_month: return None, None
    cm_mo = int(cm_month[5:])
    smly = {c.month: j for j, c in enumerate(h)
            if isinstance(c, (datetime.datetime, datetime.date)) and c.year < int(cm_month[:4])}
    LK = 1e5
    lines = []
    for label, pref, grp, tot in PL_SPEC:
        for r in rows[hrow + 1:]:
            lbl = str(r[1]).strip().lower() if len(r) > 1 and r[1] else ""
            if not lbl.startswith(pref):
                continue
            a = num(r[cm]) if cm < len(r) else 0
            b = num(r[bud]) if bud < len(r) else 0
            sv = num(r[smly[cm_mo]]) if cm_mo in smly and smly[cm_mo] < len(r) else 0
            lines.append({"label": label, "grp": grp, "tot": tot,
                          "a": a * LK, "b": b * LK, "smly": sv * LK})
            break
    return cm_month, {"lines": lines}

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
    for f in collect_files(folder, "Daily Revenue Flash_New_*.xlsx"):
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
        for r in rows:
            lbl = str(r[4]).strip() if r[4] else ""
            if lbl == "Total No. of Working Days": summary["workingDays"] = num(r[7])
            if lbl == "Current Cumulative days":   summary["cumDays"] = num(r[7])
        # MONTH (Projections) block: ACTUAL / BUDGET rows, Total Revenue col H ('000)
        in_proj = False
        for r in rows:
            a = str(r[0]).strip() if r[0] else ""
            b = str(r[1]).strip() if r[1] else ""
            if a.startswith("MONTH"): in_proj = True; continue
            if in_proj and b == "ACTUAL":  summary["projActual"] = num(r[7]) * 1000
            if in_proj and b == "BUDGET":  summary["monthBudget"] = num(r[7]) * 1000; in_proj = False
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
    for f in collect_files(folder, "Daily MIS*.xlsx"):
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
        docs, doc_daily = {}, {}
        pairs = [("Doc wise revenue date conso.", "rev"), ("Patient visits", "opv"),
                 ("No. Admissions", "adm"), ("No. discharges", "dis")]
        for sn, field in pairs:
            ws = get(sn)
            if ws is None: continue
            totals, per_day = parse_doctor_day_matrix(ws, field)
            if field == "rev": doc_daily = per_day
            for (dept, doc), val in totals.items():
                rec = docs.setdefault(doc, {"dept": dept, "rev": 0, "opv": 0, "adm": 0, "dis": 0})
                if dept: rec["dept"] = dept
                rec[field] += val
        ws = get("MoM FY 26-27") or get("MoM FY 27-28")
        if ws is not None: mom_fy.update(parse_mom_fy(ws))
        # ---- Calicut satellite (VPSLMC) sheets ----
        cal = {}
        ws = get("Input Calicut")
        if ws is not None:
            cserv, cdaily = parse_calicut_input(ws)
            if cdaily:
                cal["serv"] = {k: round(v) for k, v in cserv.items()}
                cal["daily"] = {k: round(v) for k, v in cdaily.items()}
        cdocs = {}
        for sn, field in [("Doc wise revenue Calicut", "rev"), ("OP Visit-Calicut", "opv")]:
            ws = get(sn)
            if ws is None: continue
            parsed = parse_doctor_day_matrix(ws, field)
            if not parsed or not parsed[0]: continue
            for (dept, doc), val in parsed[0].items():
                rec = cdocs.setdefault(doc, {"dept": dept, "rev": 0, "opv": 0})
                if dept: rec["dept"] = dept
                rec[field] += val
        if cdocs:
            cal["doctors"] = {k: {"dept": v["dept"], "rev": round(v["rev"]), "opv": round(v["opv"])}
                              for k, v in cdocs.items() if v["rev"] or v["opv"]}
        w.close()
        history[mkey] = {"calicut": cal, "asOf": d.strftime("%Y-%m-%d"),
                         "daysElapsed": d.day,
                         "doctors": {k: v for k, v in docs.items() if any(v[f] for f in ("rev","opv","adm","dis"))},
                         "docDaily": {k: {str(day): round(v, 0) for day, v in dd.items()}
                                      for k, dd in doc_daily.items() if dd}}
    json.dump(history, open(hist_path, "w"))
    print("History months:", sorted(history))

    # ---- Consolidated monthly P&L from 'Profit_and_Loss_Statement_*.xlsx' ----
    # Rebuilt fresh from the P&L files present (they persist in the folder).
    pnl_path = os.path.join(tools, "pnl_history.json")
    pnl_hist = {}
    for f in collect_files(folder, "Profit_and_Loss_Statement*.xlsx"):
        try:
            w = openpyxl.load_workbook(f, read_only=True, data_only=True)
        except Exception as e:
            print("  skip P&L", os.path.basename(f), e); continue
        sn = next((s for s in w.sheetnames if s.strip().lower() == "pl_consol"), None)
        if sn:
            mkey, pnl = parse_pl_consol(w[sn])
            if mkey and pnl and pnl.get("lines"):
                pnl["_src"] = os.path.basename(f)
                pnl_hist[mkey] = pnl
                print("  Parsed P&L", mkey, "from", os.path.basename(f), "-", len(pnl["lines"]), "lines")
        w.close()
    try: json.dump(pnl_hist, open(pnl_path, "w"))
    except Exception as e: print("  (pnl_history not written:", e, ")")
    print("P&L months:", sorted(k for k in pnl_hist))

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
        pnl=pnl_hist,
        pnlLatest=(sorted(k for k in pnl_hist) or [None])[-1],
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
.tabbar{background:#fff;border-bottom:2px solid #e3e8ee;box-shadow:0 1px 3px rgba(36,51,66,.05)}
.tabs{max-width:1320px;margin:0 auto;padding:0 24px;display:flex;gap:4px}
.tab{padding:13px 22px;font-size:14px;font-weight:600;color:var(--gray);cursor:pointer;border:none;background:none;border-bottom:3px solid transparent;margin-bottom:-2px;transition:color .15s}
.tab:hover{color:var(--blue)}
.tab.on{color:var(--blue);border-bottom-color:var(--blue)}
.pane{display:none}.pane.on{display:block;animation:fade .25s}
@keyframes fade{from{opacity:0}to{opacity:1}}
.banner{background:rgba(43,124,190,.07);border:1px solid rgba(43,124,190,.2);border-radius:8px;padding:10px 14px;font-size:12px;color:#3a4a5a;margin-bottom:18px;line-height:1.55}
.banner b{color:var(--maroon)}
#pnlTable td.lbl0{font-weight:650;color:#243342}
#pnlTable tr.sub td{color:var(--gray);padding-left:20px}
#pnlTable tr.tot td{border-top:2px solid #d4dbe3;font-weight:650;background:rgba(43,124,190,.04)}
#pnlTable tr.grand td{border-top:2px solid var(--maroon);font-weight:700;background:rgba(139,26,74,.05)}
</style></head><body>
<header><div><h1>VPS Lakeshore · Daily Revenue Dashboard</h1>
<div class="sub">Lakeshore Hospital &amp; Research Centre Ltd, Kochi · Global Lifecare</div></div>
<div class="sub" id="asof"></div></header>
<div class="tabbar"><div class="tabs">
<button class="tab on" data-pane="rev">Daily Revenue &amp; Operations</button>
<button class="tab" data-pane="pnl">P&amp;L &amp; Cost Analysis</button>
</div></div>
<div class="wrap">
<div id="pane-rev" class="pane on">
<div class="cards" id="cards"></div>

<div class="panel"><h2>Daily Commentary</h2>
<div class="note" id="cmnote"></div>
<div id="cmtext" style="font-size:13px;line-height:1.7;max-width:1000px"></div></div>

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

<div class="panel"><h2>Doctor × Day Revenue Matrix</h2>
<div class="note" id="mxnote"></div>
<div style="margin-bottom:8px">
<span id="mxbtns"></span>
<input id="mxFilter" placeholder="Filter doctor…" style="padding:5px 10px;border:1px solid #d4dbe3;border-radius:6px;width:220px;font-size:12.5px;margin-left:8px">
</div>
<div class="scroll" style="max-height:520px;overflow:auto"><table id="matrix"><thead></thead><tbody></tbody></table></div></div>

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

<div id="calSection" style="display:none">
<h2 style="color:var(--maroon);font-size:17px;margin:28px 0 12px;border-bottom:2px solid var(--maroon);padding-bottom:6px">VPS Lakeshore Medical Centre · Calicut (Satellite)</h2>
<div class="cards" id="calcards"></div>
<div class="grid">
<div class="panel"><h2>Calicut — Daily Revenue</h2>
<div class="note" id="caldnote"></div>
<div id="calbtns" style="margin-bottom:8px"></div><canvas id="calDaily"></canvas></div>
<div class="panel"><h2>Calicut — Service Group Mix</h2>
<div class="note" id="calsnote"></div><canvas id="calServ" style="max-height:330px"></canvas></div>
</div>
<div class="panel"><h2>Calicut — Doctor League</h2>
<div class="note" id="callnote"></div>
<div class="scroll"><table id="calTable"><thead><tr>
<th>Doctor</th><th>Department</th><th class="r">Rev cur (₹L)</th><th class="r">₹L/day</th>
<th class="r">OP visits</th><th class="r">Rev prev (₹L)</th><th class="r">Δ run-rate</th><th class="r">Share</th>
</tr></thead><tbody></tbody></table></div></div>
</div>

</div><!-- /pane-rev -->

<div id="pane-pnl" class="pane">
<div class="banner" id="pnlbanner"></div>
<div class="cards" id="pnlcards"></div>

<div class="panel"><h2>P&amp;L Commentary — <span id="pnlmlbl"></span></h2>
<div class="note" id="pnlcmnote"></div>
<div id="pnlcmtext" style="font-size:13px;line-height:1.7;max-width:1000px"></div></div>

<div class="panel"><h2>P&amp;L Waterfall — Revenue to Net Profit</h2>
<div class="note">₹ Cr. Blue = revenue &amp; profit subtotals; red = cost deductions; green = net profit. Hover for values.</div>
<canvas id="wfChart" style="max-height:380px"></canvas></div>

<div class="grid">
<div class="panel"><h2>Cost Drivers — % of Revenue: Actual vs Budget</h2>
<div class="note">Each cost line as a share of revenue. Bars over budget (dotted) are adverse.</div>
<canvas id="drvChart"></canvas></div>
<div class="panel"><h2>Cost Mix — Share of Total Cost</h2>
<div class="note" id="mixnote"></div><canvas id="mixChart" style="max-height:330px"></canvas></div>
</div>

<div class="panel"><h2>Profit &amp; Loss Statement — <span id="pnlmlbl2"></span></h2>
<div class="note">Audited monthly close. Variance is favourable (green) when actual beats budget — lower cost or higher revenue/profit. All figures ₹ Cr.</div>
<div class="scroll"><table id="pnlTable"><thead><tr>
<th>Particulars</th><th class="r">Actual ₹Cr</th><th class="r">% Rev</th>
<th class="r">Budget ₹Cr</th><th class="r">% Rev</th><th class="r">Variance ₹Cr</th><th>F/A</th>
</tr></thead><tbody></tbody></table></div></div>

<div class="grid">
<div class="panel"><h2>Monthly Revenue — Actual vs Budget</h2>
<div class="note">₹ Cr per closed month, consolidated management P&amp;L.</div>
<canvas id="revTrend"></canvas></div>
<div class="panel"><h2>Profitability Trend — EBITDA &amp; PAT</h2>
<div class="note" id="mgnnote"></div><canvas id="marginTrend"></canvas></div>
</div>
</div><!-- /pane-pnl -->

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
// month run-rate projection: MTD ÷ elapsed working days × total working days, vs full-month budget
const wdTot=D.summary.workingDays||0, wdCum=D.summary.cumDays||0;
const dailyKeys0=Object.keys(D.daily).sort();
const curKey=dailyKeys0[dailyKeys0.length-1];
const fullBud=D.summary.monthBudget||(curKey? D.daily[curKey].reduce((a,r)=>a+r.budTot,0):0);
const rrDay=wdCum? (mtd.revTot||0)/wdCum:0;
const proj=rrDay*wdTot;
let projCard='';
if(proj&&fullBud){
 const p=pct(proj,fullBud);
 projCard=card('Month Run-Rate → Landing',fmtCr(proj),
  `₹${(rrDay/CR).toFixed(2)} Cr/wk-day × ${wdTot} days · <span class="${p>=0?'up':'dn'}">${p>=0?'▲ +':'▼ '}${p.toFixed(1)}% vs ₹${(fullBud/CR).toFixed(0)} Cr budget</span>`,'',1);
}
document.getElementById('cards').innerHTML=
 projCard+
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

// ---------------- doctor x day revenue matrix ----------------
const mxMonths=hMonths.filter(mk=>D.history[mk].docDaily&&Object.keys(D.history[mk].docDaily).length);
let mxMonth=mxMonths[mxMonths.length-1]||null;
window.drawMatrix=function(mk){
 mxMonth=mk;
 const h=D.history[mk], dd=h.docDaily||{};
 const nDays=daysIn(mk,h);
 const q=(document.getElementById('mxFilter').value||'').toUpperCase();
 // rows sorted by month total desc
 let docs=Object.keys(dd).map(n=>{
  const tot=Object.values(dd[n]).reduce((a,v)=>a+v,0);
  return {n, tot};}).filter(x=>x.tot>1000)
  .filter(x=>!q||x.n.includes(q))
  .sort((a,b)=>b.tot-a.tot).slice(0,50);
 const days=[...Array(nDays).keys()].map(i=>i+1);
 const [y,m]=mk.split('-').map(Number);
 const dows=days.map(d=>'SMTWTFS'[new Date(y,m-1,d).getDay()]);
 document.querySelector('#matrix thead').innerHTML=
  '<tr><th style="position:sticky;left:0;background:#fff;z-index:2">Doctor</th><th class="r">Total ₹L</th>'+
  days.map((d,i)=>`<th class="r" style="min-width:44px${dows[i]==='S'?';color:#c0392b':''}">${d}<br><span style="font-weight:400">${dows[i]}</span></th>`).join('')+'</tr>';
 const maxAll=Math.max(...docs.slice(0,15).map(x=>Math.max(...Object.values(dd[x.n]))));
 document.querySelector('#matrix tbody').innerHTML=docs.map(x=>{
  const row=dd[x.n];
  return '<tr><td class="doc" style="position:sticky;left:0;background:#fff;z-index:1" title="'+tc(x.n)+'">'+tc(x.n)+'</td>'+
   `<td class="r"><b>${(x.tot/L).toFixed(1)}</b></td>`+
   days.map(d=>{
    const v=row[String(d)]||0;
    if(!v) return '<td class="r" style="color:#c8d0d9">·</td>';
    const a=Math.min(v/maxAll,1);
    const bg=`rgba(43,124,190,${(0.06+a*0.5).toFixed(2)})`;
    const txt=v>=1e5? (v/L).toFixed(1) : (v/1000).toFixed(0)+'k';
    return `<td class="r" style="background:${bg}${a>0.65?';color:#fff':''}" title="₹${Math.round(v).toLocaleString('en-IN')}">${txt}</td>`;
   }).join('')+'</tr>';
 }).join('');
 document.getElementById('mxnote').innerHTML=
  `Billed gross revenue attributed to each doctor, by calendar day of <b>${mName(mk)}</b> (₹ Lakhs; values under ₹1 L shown as ₹’000 with “k”). Sundays in red. Top 50 doctors — use the filter for others. Note: this is billed revenue, not cash collected — the MIS does not attribute collections to doctors.`;
 document.querySelectorAll('#mxbtns .mbtn').forEach(b=>b.classList.toggle('on',b.dataset.k===mk));
}
document.getElementById('mxbtns').innerHTML=mxMonths.map(k=>`<button class="mbtn" data-k="${k}" onclick="drawMatrix('${k}')">${k}</button>`).join('');
document.getElementById('mxFilter').oninput=()=>drawMatrix(mxMonth);
if(mxMonth)drawMatrix(mxMonth);

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
 '<br>Gross revenue per Daily Revenue Flash. Doctor-month figures from Daily MIS doctor sheets (newest file per month; current month is MTD — compare on ₹/day run-rate). Discharge status / ALOS / payer mix cover only dates with a flash file on hand. ARPOB = gross revenue ÷ occupied bed-days; ALOS from MIS MoM sheet. Operations include VPSLMC (satellite) figures where the source does. Calicut section from the Daily MIS "Input Calicut" / "Doc wise revenue Calicut" / "OP Visit-Calicut" sheets.';

// ---- daily commentary (auto-generated) ----
(function(){
 const rows=D.daily[curKey]||[]; if(!rows.length)return;
 const last=rows[rows.length-1];
 const dObj=new Date(last.date+'T12:00:00');
 const dName=dObj.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
 document.getElementById('cmnote').textContent='Auto-generated from the '+D.latestDate+' flash and Daily MIS.';
 const fL=v=>'\u20B9'+(v/L).toFixed(1)+' L';
 const workRows=rows.filter(r=>r.revTot>0);
 const prior=workRows.slice(0,-1).slice(-7);
 const avg7=prior.length? prior.reduce((a,r)=>a+r.revTot,0)/prior.length:0;
 const vb=pct(last.revTot,last.budTot), va=avg7?pct(last.revTot,avg7):0;
 const mtdRev=rows.reduce((a,r)=>a+r.revTot,0), mtdBud=rows.reduce((a,r)=>a+r.budTot,0), mtdP=pct(mtdRev,mtdBud);
 const mix=[['IP',last.revIP],['OP',last.revOP],['Pharmacy',last.revPH]].sort((a,b)=>b[1]-a[1]);
 const dd=(D.history[curM]&&D.history[curM].docDaily)||{};
 const dayK=String(dObj.getDate());
 const docs=(D.history[curM]&&D.history[curM].doctors)||{};
 const tops=Object.entries(dd).map(([k,v])=>[k,v[dayK]||0]).filter(x=>x[1]>0&&!/^LHRC/i.test(x[0])).sort((a,b)=>b[1]-a[1]).slice(0,3);
 const topTxt=tops.map(([k,v])=>tc(k)+(docs[k]&&docs[k].dept?' ('+tc(docs[k].dept)+', '+fL(v)+')':' ('+fL(v)+')')).join(', ');
 const conv=last.revTot? last.collTot/last.revTot*100:0;
 const isSun=(last.dow||'').toUpperCase().startsWith('SUN');
 const s=[];
 s.push(`<b>${dName}</b> closed at <b>${fL(last.revTot)}</b> gross revenue against a budget of ${fL(last.budTot)} (<span class="${vb>=0?'up':'dn'}">${vb>=0?'+':''}${vb.toFixed(1)}%</span>)`+(avg7?`, ${Math.abs(va).toFixed(0)}% ${va>=0?'above':'below'} the trailing 7-day average of ${fL(avg7)}.`:'.'));
 s.push(`The day was led by ${mix[0][0]} at ${fL(mix[0][1])}, followed by ${mix[1][0]} (${fL(mix[1][1])}) and ${mix[2][0]} (${fL(mix[2][1])}).`);
 if(topTxt) s.push(`Top contributors: ${topTxt}.`);
 s.push(`Collections came in at ${fL(last.collTot)} \u2014 ${conv.toFixed(0)}% of the day's gross.`);
 s.push(`Month to date stands at ${fmtCr(mtdRev)} vs ${fmtCr(mtdBud)} budgeted (<span class="${mtdP>=0?'up':'dn'}">${mtdP>=0?'+':''}${mtdP.toFixed(1)}%</span>)`+((proj&&fullBud)?`, with the working-day run-rate pointing to a ${fmtCr(proj)} landing against the \u20B9${(fullBud/CR).toFixed(0)} Cr month budget.`:'.'));
 if(isSun) s.push('Sunday posting \u2014 the low absolute number is seasonal, not an anomaly.');
 document.getElementById('cmtext').innerHTML=s.join(' ');
})();

// ---------------- Calicut satellite (VPSLMC) ----------------
(function(){
 const calM=hMonths.filter(mk=>D.history[mk].calicut&&D.history[mk].calicut.daily);
 if(!calM.length)return;
 document.getElementById('calSection').style.display='';
 const ck=calM[calM.length-1], pk=calM[calM.length-2];
 const c=D.history[ck].calicut, p=pk? D.history[pk].calicut:null;
 const dC=daysIn(ck,D.history[ck]), dP=pk? daysIn(pk,D.history[pk]):1;
 const sum=o=>Object.values(o||{}).reduce((a,v)=>a+v,0);
 const mtdRev=sum(c.daily), rr=mtdRev/dC;
 const pRev=p? sum(p.daily):0, prr=p? pRev/dP:0;
 const opTot=Object.values(c.doctors||{}).reduce((a,v)=>a+(v.opv||0),0);
 const active=Object.values(c.doctors||{}).filter(v=>v.rev>0).length;
 const momP=prr? (rr/prr-1)*100:null;
 const share=(mtd.revTot)? mtdRev/mtd.revTot*100:null;
 document.getElementById('calcards').innerHTML=
  card('Calicut MTD Revenue ('+mName(ck)+')','\u20b9'+(mtdRev/L).toFixed(1)+' L',
   share!=null? share.toFixed(1)+'% of unit gross':'','',1)+
  card('Run-rate','\u20b9'+(rr/L).toFixed(2)+' L/day',
   momP==null?'\u2014':(momP>=0?'\u25b2 +':'\u25bc ')+momP.toFixed(1)+'% vs '+(pk?mName(pk):'prev'),momP==null?'':(momP>=0?'up':'dn'))+
  card('OP Visits (MTD)',opTot.toLocaleString('en-IN'),'~'+(opTot/dC).toFixed(0)+' / day','')+
  card('Active Doctors',active,(pk&&p.doctors)? Object.values(p.doctors).filter(v=>v.rev>0).length+' in '+mName(pk):'','');
 // daily chart with month toggle
 let calChart=null;
 window.drawCal=function(mk){
  const h=D.history[mk], cc=h.calicut;
  const nDays=daysIn(mk,h);
  const [y,m]=mk.split('-').map(Number);
  const days=[...Array(nDays).keys()].map(i=>i+1);
  const vals=days.map(d=>{
   const iso=`${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
   return (cc.daily[iso]||0)/L;});
  const dows=days.map(d=>'SMTWTFS'[new Date(y,m-1,d).getDay()]);
  if(calChart)calChart.destroy();
  calChart=new Chart(document.getElementById('calDaily'),{type:'bar',
   data:{labels:days.map((d,i)=>d+' '+dows[i]),datasets:[
    {label:'Gross revenue (\u20b9 L)',data:vals,backgroundColor:days.map((d,i)=>dows[i]==='S'?'rgba(139,26,74,.35)':MAROON)}]},
   options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
   scales:{x:{ticks:{font:{size:10}}},y:{title:{display:true,text:'\u20b9 Lakhs'}}}}});
  document.querySelectorAll('#calbtns .mbtn').forEach(b=>b.classList.toggle('on',b.dataset.k===mk));
 }
 document.getElementById('calbtns').innerHTML=calM.map(k=>`<button class="mbtn" data-k="${k}" onclick="drawCal('${k}')">${k}</button>`).join('');
 drawCal(ck);
 document.getElementById('caldnote').textContent='From the Daily MIS "Input Calicut" sheet. Sundays (maroon-light) \u2014 clinic closed or skeleton OP; zero posting is normal.';
 // service mix
 const sv=Object.entries(c.serv||{}).sort((a,b)=>b[1]-a[1]);
 document.getElementById('calsnote').textContent=mName(ck)+' MTD, \u20b9 Lakhs by collection head.';
 new Chart(document.getElementById('calServ'),{type:'doughnut',
  data:{labels:sv.map(x=>x[0]),datasets:[{data:sv.map(x=>x[1]/L),
   backgroundColor:[MAROON,BLUE,'#c8952b','#1a7f4e',GRAY,LT,'#5a6875','#d4dbe3']}]},
  options:{plugins:{legend:{position:'bottom',labels:{boxWidth:11,font:{size:10.5}}},
   tooltip:{callbacks:{label:t=>t.label+': \u20b9'+t.parsed.toFixed(1)+' L'}}}}});
 // doctor league
 const docsP=(p&&p.doctors)||{};
 let rows=Object.entries(c.doctors||{}).map(([n,v])=>{
  const pv=docsP[n]||{};
  const rrC=(v.rev||0)/dC, rrP=pv.rev? pv.rev/dP:0;
  return {doc:tc(n),dept:tc(v.dept||''),rev:(v.rev||0)/L,rr:rrC/L,opv:v.opv||0,
   revP:(pv.rev||0)/L,mom:rrP?(rrC/rrP-1)*100:null,share:mtdRev? (v.rev||0)/mtdRev*100:null};})
  .filter(r=>r.rev>0.005||r.opv>0).sort((a,b)=>b.rev-a.rev).slice(0,25);
 const momCell=v=>v==null?'\u2014':`<span class="tag ${v>=10?'g':v<=-10?'r':'y'}">${v>=0?'+':''}${v.toFixed(0)}%</span>`;
 document.querySelector('#calTable tbody').innerHTML=rows.map(r=>
  `<tr><td class="doc" title="${r.doc}">${r.doc}</td><td class="doc">${r.dept}</td>`+
  `<td class="r"><b>${r.rev.toFixed(2)}</b></td><td class="r">${r.rr.toFixed(3)}</td>`+
  `<td class="r">${r.opv||'\u2014'}</td><td class="r">${r.revP?r.revP.toFixed(2):'\u2014'}</td>`+
  `<td class="r">${momCell(r.mom)}</td><td class="r">${r.share==null?'\u2014':r.share.toFixed(1)+'%'}</td></tr>`).join('');
 document.getElementById('callnote').innerHTML=
  `Doctor-attributed gross revenue and OP visits at the Calicut medical centre, <b>${mName(ck)}</b> (${dC} days elapsed)`+
  (pk?` vs <b>${mName(pk)}</b> (${dP} days) on \u20b9/day run-rate.`:'.')+' Top 25 by revenue.';
 // append a Calicut line to the daily commentary
 const dKeys=Object.keys(c.daily).sort();
 if(dKeys.length){
  const lastD=dKeys[dKeys.length-1], v=c.daily[lastD];
  const el=document.getElementById('cmtext');
  if(el&&el.innerHTML) el.innerHTML+=` At the <b>Calicut satellite</b>, ${new Date(lastD+'T12:00:00').toLocaleDateString('en-GB',{day:'numeric',month:'long'})} posted \u20b9${(v/L).toFixed(1)} L, taking its MTD to \u20b9${(mtdRev/L).toFixed(1)} L${share!=null?' ('+share.toFixed(1)+'% of unit gross)':''}.`;
 }
})();

// ==================== TAB SWITCHING ====================
document.querySelectorAll('.tab').forEach(t=>t.addEventListener('click',()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('on'));
  document.querySelectorAll('.pane').forEach(x=>x.classList.remove('on'));
  t.classList.add('on');
  document.getElementById('pane-'+t.dataset.pane).classList.add('on');
  if(t.dataset.pane==='pnl') initPnl();
}));

// ==================== P&L & COST ANALYSIS TAB ====================
const RED='#c0392b', GREEN='#1a7f4e', AMBER='#c8952b';
const crc=v=>'₹'+(v/CR).toFixed(2)+' Cr';
let pnlBuilt=false;
function initPnl(){
 if(pnlBuilt) return; pnlBuilt=true;
 const P=D.pnl||{}, keys=Object.keys(P).sort(), mk=D.pnlLatest;
 const mName2=k=>new Date(k+'-01T12:00:00').toLocaleDateString('en-GB',{month:'long',year:'numeric'});
 const mShort=k=>new Date(k+'-01T12:00:00').toLocaleDateString('en-GB',{month:'short',year:'2-digit'});
 if(!mk||!P[mk]||!P[mk].lines){
   document.getElementById('pnlbanner').innerHTML='No consolidated P&amp;L found. Add <b>Profit_and_Loss_Statement_FY27 &lt;Month&gt;.xlsx</b> files and rebuild.';
   return;
 }
 const lines=P[mk].lines, map={}; lines.forEach(l=>map[l.label]=l);
 const g=l=>map[l]||{a:0,b:0,smly:0,grp:'cost'};
 const rev=g('Total Revenue').a, revB=g('Total Revenue').b||1, revS=g('Total Revenue').smly||0;
 const ml=mName2(mk);
 document.getElementById('pnlmlbl').textContent=ml;
 document.getElementById('pnlmlbl2').textContent=ml;
 document.getElementById('pnlbanner').innerHTML=
  'Consolidated management P&amp;L (Kochi + Calicut + F&amp;B), latest close <b>'+ml+'</b>. '+
  'Actual vs Budget vs same-month-last-year (YoY), all ₹ Cr. Source: <span style="color:var(--gray)">'+(P[mk]._src||'')+'</span>. '+
  keys.length+' month(s) loaded — trend extends as more P&amp;L files are added.';

 const eb=g('EBITDA'), pat=g('PAT'), mat=g('Material Cost'), doc=g('Doctor Cost'), emp=g('Employee Cost');
 const man=doc.a+emp.a, texp=g('Total Expenses');
 const dv=v=>(v>=0?'+':'')+(v/CR).toFixed(2)+' Cr';
 const cards=[
  ['Total Revenue',crc(rev),dv(rev-revB)+' vs bud',''],
  ['EBITDA',crc(eb.a)+' · '+(eb.a/rev*100).toFixed(1)+'%',dv(eb.a-eb.b)+' vs bud','m'],
  ['PAT',crc(pat.a)+' · '+(pat.a/rev*100).toFixed(1)+'%',dv(pat.a-pat.b)+' vs bud','m'],
  ['Material Cost',(mat.a/rev*100).toFixed(1)+'% of rev','bud '+(mat.b/revB*100).toFixed(1)+'%',''],
  ['Manpower (Dr+Emp)',(man/rev*100).toFixed(1)+'% of rev','bud '+((doc.b+emp.b)/revB*100).toFixed(1)+'%',''],
  ['Total Opex',(texp.a/rev*100).toFixed(1)+'% of rev','bud '+(texp.b/revB*100).toFixed(1)+'%','']
 ];
 document.getElementById('pnlcards').innerHTML=cards.map(function(c){var l=c[0],v=c[1],d=c[2],cls=c[3];
   var cl=/vs bud/.test(d)?(/^\+/.test(d)?'up':(/^-/.test(d)?'dn':'')):'';
   return '<div class="card '+cls+'"><div class="lbl">'+l+'</div><div class="val">'+v+'</div><div class="delta '+cl+'">'+d+'</div></div>';
 }).join('');

 // Waterfall
 const cr=v=>v/CR, otherOpex=texp.a-mat.a-doc.a-emp.a;
 let run=0,labels=[],ranges=[],colors=[];
 const push=(l,rng,c)=>{labels.push(l);ranges.push(rng);colors.push(c);};
 push('Revenue',[0,cr(rev)],BLUE); run=rev;
 [['Material',mat.a],['Doctor',doc.a],['Employee',emp.a],['Other Opex',otherOpex]].forEach(function(x){push(x[0],[cr(run-x[1]),cr(run)],RED);run-=x[1];});
 push('EBITDA',[0,cr(eb.a)],'#1b5e94'); run=eb.a;
 [['Depreciation',g('Depreciation').a],['Finance',g('Finance Cost').a],['Tax',g('Tax').a]].forEach(function(x){push(x[0],[cr(run-x[1]),cr(run)],RED);run-=x[1];});
 push('PAT',[0,cr(pat.a)],GREEN);
 new Chart(document.getElementById('wfChart'),{type:'bar',data:{labels:labels,datasets:[{data:ranges,backgroundColor:colors,borderRadius:3}]},
  options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>'₹'+Math.abs(c.raw[1]-c.raw[0]).toFixed(2)+' Cr'}}},
   scales:{y:{title:{display:true,text:'₹ Cr'}},x:{ticks:{font:{size:10},maxRotation:45,minRotation:30}}}}});

 // Cost drivers: top operating-cost lines, actual% vs budget% of revenue
 const costLines=lines.filter(l=>l.grp==='cost'&&!l.tot&&l.a>0).sort((a,b)=>b.a-a.a), topc=costLines.slice(0,8);
 new Chart(document.getElementById('drvChart'),{type:'bar',data:{labels:topc.map(l=>l.label),datasets:[
   {label:'Actual % rev',data:topc.map(l=>l.a/rev*100),backgroundColor:BLUE,borderRadius:3},
   {label:'Budget % rev',data:topc.map(l=>l.b/revB*100),backgroundColor:'rgba(139,26,74,.35)',borderRadius:3}
 ]},options:{indexAxis:'y',plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},scales:{x:{title:{display:true,text:'% of revenue'}}}}});

 // Cost mix doughnut (share of total opex)
 const top6=costLines.slice(0,6), restV=costLines.slice(6).reduce((s,l)=>s+l.a,0);
 const mixL=top6.map(l=>l.label).concat(restV>0?['Other']:[]), mixV=top6.map(l=>l.a/CR).concat(restV>0?[restV/CR]:[]);
 new Chart(document.getElementById('mixChart'),{type:'doughnut',data:{labels:mixL,datasets:[{data:mixV,
   backgroundColor:[BLUE,MAROON,AMBER,GRAY,'#7aa9d0','#d08a9a','#b9c2cc']}]},
   options:{plugins:{legend:{position:'right',labels:{boxWidth:11,font:{size:10.5}}}}}});
 document.getElementById('mixnote').textContent='Total operating cost ₹'+(texp.a/CR).toFixed(2)+' Cr in '+ml+'. Material + manpower dominate the base.';

 // P&L statement table
 const fav=l=>{var v=l.a-l.b, good=(l.grp==='cost')?v<=0:v>=0;
   return '<span class="tag '+(good?'g':'r')+'">'+(good?'Fav ':'Adv ')+(v>=0?'+':'')+(v/CR).toFixed(2)+'</span>';};
 const yoy=l=>l.smly>0?(((l.a/l.smly-1)*100)>=0?'+':'')+((l.a/l.smly-1)*100).toFixed(0)+'%':'—';
 document.querySelector('#pnlTable tbody').innerHTML=lines.map(function(l){
   var cls=l.label==='PAT'?'grand':(l.tot?'tot':(l.grp==='cost'?'sub':''));
   return '<tr class="'+cls+'"><td class="'+(cls==='sub'?'':'lbl0')+'">'+l.label+'</td>'+
     '<td class="r">'+(l.a/CR).toFixed(2)+'</td><td class="r">'+(l.a/rev*100).toFixed(1)+'%</td>'+
     '<td class="r">'+(l.b/CR).toFixed(2)+'</td><td class="r">'+(l.b/revB*100).toFixed(1)+'%</td>'+
     '<td class="r">'+fav(l)+'</td><td class="r">'+yoy(l)+'</td></tr>';
 }).join('');

 // Trend across all loaded months
 const cm=keys.filter(k=>P[k]&&P[k].lines);
 const lm=(k,l)=>{var x=P[k].lines.find(z=>z.label===l); return x||{a:0,b:0};};
 new Chart(document.getElementById('revTrend'),{data:{labels:cm.map(mShort),datasets:[
   {type:'bar',label:'Actual ₹Cr',data:cm.map(k=>lm(k,'Total Revenue').a/CR),backgroundColor:BLUE,borderRadius:4},
   {type:'bar',label:'Budget ₹Cr',data:cm.map(k=>lm(k,'Total Revenue').b/CR),backgroundColor:'rgba(139,26,74,.3)',borderRadius:4}
 ]},options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},scales:{y:{title:{display:true,text:'₹ Cr'}}}}});
 new Chart(document.getElementById('marginTrend'),{data:{labels:cm.map(mShort),datasets:[
   {type:'bar',label:'EBITDA ₹Cr',data:cm.map(k=>lm(k,'EBITDA').a/CR),backgroundColor:'rgba(43,124,190,.75)',borderRadius:4,yAxisID:'y'},
   {type:'line',label:'EBITDA %',data:cm.map(k=>lm(k,'EBITDA').a/lm(k,'Total Revenue').a*100),borderColor:MAROON,borderWidth:2,pointRadius:3,yAxisID:'y1'},
   {type:'line',label:'PAT %',data:cm.map(k=>lm(k,'PAT').a/lm(k,'Total Revenue').a*100),borderColor:GREEN,borderWidth:2,borderDash:[5,4],pointRadius:3,yAxisID:'y1'}
 ]},options:{plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},scales:{
   y:{position:'left',title:{display:true,text:'₹ Cr'}},
   y1:{position:'right',title:{display:true,text:'% of revenue'},grid:{drawOnChartArea:false}}}}});
 document.getElementById('mgnnote').innerHTML=cm.length>1?'EBITDA (bars) with EBITDA% and PAT% margins across '+cm.length+' months.':'Add more P&amp;L files to extend the trend.';

 // Commentary
 const prevK=cm.indexOf(mk)>0?cm[cm.indexOf(mk)-1]:null;
 const momRev=prevK?((rev/lm(prevK,'Total Revenue').a-1)*100):null;
 const bigc=costLines.slice(0,6).map(l=>({n:l.label,v:l.a-l.b,pa:l.a/rev*100,pb:l.b/revB*100}));
 const worst=bigc.slice().sort((a,b)=>b.v-a.v)[0], best=bigc.slice().sort((a,b)=>a.v-b.v)[0];
 document.getElementById('pnlcmnote').textContent='Auto-generated from the '+ml+' consolidated P&L close.';
 document.getElementById('pnlcmtext').innerHTML=
   '<b>'+ml+'</b> revenue was <b>'+crc(rev)+'</b>, '+(rev>=revB?'ahead of':'behind')+' budget by ₹'+Math.abs((rev-revB)/CR).toFixed(2)+' Cr'+
   (revS>0?' and up '+((rev/revS-1)*100).toFixed(0)+'% YoY':'')+
   (momRev!=null?', '+(momRev>=0?'up':'down')+' '+Math.abs(momRev).toFixed(1)+'% on '+mName2(prevK):'')+'. '+
   'EBITDA <b>'+crc(eb.a)+' ('+(eb.a/rev*100).toFixed(1)+'%)</b> vs '+(eb.b/revB*100).toFixed(1)+'% budget, and PAT <b>'+crc(pat.a)+' ('+(pat.a/rev*100).toFixed(1)+'%)</b>. '+
   'The largest cost is <b>material at '+(mat.a/rev*100).toFixed(1)+'% of revenue</b> (budget '+(mat.b/revB*100).toFixed(1)+'%). '+
   'Biggest adverse line: <b>'+worst.n+'</b> (₹'+Math.abs(worst.v/CR).toFixed(2)+' Cr over, '+worst.pa.toFixed(1)+'% vs '+worst.pb.toFixed(1)+'%); '+
   'biggest saving: <b>'+best.n+'</b> (₹'+Math.abs(best.v/CR).toFixed(2)+' Cr under, '+best.pa.toFixed(1)+'% vs '+best.pb.toFixed(1)+'%).';
}
</script></body></html>
"""

if __name__ == "__main__":
    main()
