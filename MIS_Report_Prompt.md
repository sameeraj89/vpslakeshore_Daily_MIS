# LHRC Daily Flash — Source-to-Report Prompt

Copy the prompt below into a new Claude chat (or Cowork session with the
**Daily MIS Reports** folder connected), drop in the raw HIS exports for the day,
and Claude will produce the daily revenue report. The source exports are the same
seven files that land in `dailyrevenueflashsourcereports/`:

1. `Dept Wise Category Wise With Doc Name Kochi.xlsx` (+ `...Calicut.xlsx`)
2. `Op Visit - Doctor Wise Kochi.xlsx` (+ `...Calicut.xlsx`)
3. `Admission Count (Dyn).xlsx`
4. `Discharge Patient List.xlsx`
5. `Doctor Wise Bed Occupancy Date Wise.xlsx`

---

## THE PROMPT (copy from here)

You are preparing the daily ACTIVITY & REVENUE FLASH for Lakeshore Hospital &
Research Centre Ltd (VPS Lakeshore), Kochi — the report sent daily to Burjeel Group.
I have attached the raw HIS exports for **[DATE]**. Produce the flash and an
executive commentary. Follow these rules exactly:

**Parsing.** Every export has a 9-row letterhead; the header row is row 10.
- *Dept Wise Category Wise With Doc Name (Kochi & Calicut)*: columns Department,
  Category, Doctor, Type (IP/OP), Package, Gross, Discount, Net, Gross Refund,
  Refund Discount, Net Refund. Day revenue per line = Net − Net Refund.
- *Op Visit - Doctor Wise*: Department, Doctor, NEW, FREE, RENEW, TOTAL VISIT.
- *Admission Count (Dyn)*: Department, Doctor, Admissiondate, Admissioncount.
- *Discharge Patient List*: one row per discharge (with admission date, doctor,
  scheme, status).
- *Doctor Wise Bed Occupancy Date Wise*: one row per occupied bed.

**Compute for the day:**
1. Total net service revenue, split IP vs OP (from Type), and Pharmacy vs
   Laboratory vs Radiology vs Healthcare Services (from Category groupings).
2. Department-wise revenue ranked descending, with IP/OP split — flag the top 5
   and any department >±30% vs its recent daily average if history is available.
3. Doctor-wise revenue top 10 (name, department, revenue, admissions, OP visits).
4. Volumes: OP visits total (NEW / FREE / RENEW split), new-patient share,
   admissions, discharges, occupied beds, and ALOS of the day's discharges
   (discharge date − admission date).
5. Payer signal from the Discharge list Scheme column: cash vs insurance vs
   ECHS/corporate mix of discharges.
6. Kochi vs Calicut (satellite) split for revenue and OP visits.

**Report format:** a one-page flash in this order — headline table (Revenue: OP /
IP / Pharmacy / Total; Volumes: OP visits, admissions, discharges, occupied beds),
then dept table, then doctor table, then 5–8 bullet-free sentences of executive
commentary written like a CFO note: what drove the day, mix shift, payer mix,
anything anomalous (refunds, negative lines, a single large bill), and what to
watch tomorrow. Use ₹ Lakhs with one decimal. If a budget or prior-day file is
also attached, add variance columns and a run-rate view for the month.

**Checks before you answer:** reconcile the sum of dept revenue to the sum of
doctor revenue and to the grand total (they must tie); state the reconciliation.
Flag any negative net lines >₹10,000 (refunds) separately rather than netting
them silently. If Calicut files are missing, say so and report Kochi-only.

## END PROMPT

---

*Tip: also attach yesterday's `Daily Revenue Flash_New_*.xlsx` and the month's
budget if you want DAY vs BUDGET and MTD variance computed in the same pass —
the prompt handles it automatically.*
