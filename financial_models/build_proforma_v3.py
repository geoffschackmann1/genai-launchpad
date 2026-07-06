"""AZALEA HOSPICE PROFORMA - REVISION 3.00 DYNAMIC.

Clean rebuild of the Rev 2.00 SBA budget per its own Handoff Brief (build order:
Control Tower -> Census Waterfall -> Revenue -> Staffing -> Operating Budget ->
P&L -> Balance Sheet (no plugs) -> Cash Flow & Runway -> 3-Year Summary), with:
  - the GRANULAR COA-backed P&L structure (account codes preserved),
  - EVERY calculated cell a formula; only Control Tower cells are values,
  - deal terms defaulted to the CURRENT Refuge structure ($500K; $125K down;
    $25K/mo Sep-Dec; Jan balloon @6% funded by balloon refi; toggles for the
    Sept $500K/6%/36 refi and the SBA $450K),
  - census defaulted to the 7/6 plan (open ~25 transfers; 27.4 EOM M2; 32.4 M4;
    35 M6; -> 50 by M24 -> 56 by M36),
  - post-audit cost settings kept (benefits 22%, workers comp 3%, stepped rent,
    MAX(fixed, % of NPR) G&A lines, contingency 5%, EMR per-patient),
  - cash flow with collections timing (%same/+1/+2) plus a CHOW payment-hold.
36 monthly columns (Jul-26 .. Jun-29).
"""
from openpyxl.utils import get_column_letter as gcl
from openpyxl.formatting.rule import CellIsRule

from .engine import styles as S
from .build_operating_model import OB, MC0, mlet, NM, MONTH_LABELS, MRANGE

ADMITS = [20.0, 11.41, 11.4, 12.32, 13.25, 14.18, 13.6, 13.98, 14.35, 14.72, 15.09,
          15.46, 15.66, 15.98, 16.28, 16.59, 16.9, 17.21, 17.51, 17.83, 18.13, 18.44,
          18.76, 19.06, 19.04, 19.22, 19.41, 19.59, 19.78, 19.96, 20.15, 20.33, 20.52,
          20.7, 20.89, 21.08]

GA_LINES = [  # key, label, acct, fixed $/mo, % of NPR (0 = none), start month
    ("ga_emr",    "EMR - base platform + per-patient (6530-Z)", "6530-Z", None, None, 1),  # special
    ("ga_malp",   "Malpractice / professional liability", "6580-Z", 833.33, 0.0, 1),
    ("ga_gl",     "General liability insurance", "6580-Z", 150.0, 0.0, 1),
    ("ga_phone",  "Phone & internet", "6710-Z", 100.0, 0.0, 1),
    ("ga_mkt",    "Marketing & referral development", "6800-M", 1000.0, 0.015, 1),
    ("ga_acct",   "Accounting & bookkeeping", "6510-Z", 950.0, 0.01, 1),
    ("ga_legal",  "Legal & compliance", "6600-Z", 250.0, 0.003, 4),
    ("ga_office", "Office supplies", "6650-Z", 300.0, 0.005, 1),
    ("ga_ce",     "Continuing education & licensing", "6560-Z", 100.0, 0.001, 1),
    ("ga_dues",   "Dues & subscriptions", "6570/6705", 100.0, 0.0, 1),
    ("ga_post",   "Postage, printing & copier", "6565/6670", 200.0, 0.0, 1),
]
CONT = ("ga_cont", "Contingency reserve (below EBITDA - non-GAAP buffer)", "6640-Z", 500.0, 0.05, 1)

CLIN = [  # key, label, acct, salary-key, caseload-key, start
    ("rn",  "RN Case Manager (4000-0-6-1)", "rn_sal", "rn_case"),
    ("oc",  "RN On-Call / part-time (4000-R-6-3)", "oc_sal", None),
    ("cna", "Hospice Aide / CNA (4000-R-6-5)", "cna_sal", "cna_case"),
    ("msw", "Medical Social Worker (4000-R-6-6)", "msw_sal", "msw_case"),
]
PRN = [  # key, label, req-key, cap-key, rate-key, fte-of (None = fixed visits)
    ("prn_rn",  "PRN RN supplement (4025-0-6-1)", "rn_req", "rn_cap", "rn_prn", "rn"),
    ("prn_cna", "PRN Aide supplement (4025-0-6-5)", "cna_req", "cna_cap", "cna_prn", "cna"),
    ("prn_msw", "PRN MSW supplement (4025-0-6-6)", "msw_req", "msw_cap", "msw_prn", "msw"),
    ("prn_ch",  "PRN Chaplain above base (4025-0-6-7)", "ch_req", "ch_base_v", "ch_prn", None),
]


def yr(i):  # 0-based month -> escalation exponent
    return i // 12


# =====================================================================
def control_tower(bk: OB):
    ws = bk.sheet("Control Tower", tab="1F3864")
    bk.title(ws, "CONTROL TOWER - every number in the model flows from this tab. Yellow = input; everything else is formulas.")
    bk.widths(ws, label=52, n=4, w=13)
    ws.column_dimensions["D"].width = 70
    r = 4
    def item(key, label, val, fmt, note=""):
        nonlocal r
        bk.lbl(ws, r, label); bk.inp(ws, r, 2, val, fmt, key=key)
        if note: bk.lbl(ws, r, note, c=4, italic=True)
        r += 1
    def sec(t):
        nonlocal r
        bk.section(ws, r, t, span=4); r += 1

    sec("COMPANY & DEAL  (current Refuge structure)")
    item("cash0", "Starting capital on deposit (equity)", 250000.0, S.FMT_CUR, "Jim Bullard $195K + owner $55K")
    item("startup", "Pre-opening / startup spend from capital", 0.0, S.FMT_CUR, "0 = funded elsewhere")
    item("price", "Refuge license purchase price (100% membership CHOW)", 500000.0, S.FMT_CUR, "Acct 1980 intangible")
    item("down", "Down payment at closing (Month 1)", 125000.0, S.FMT_CUR)
    item("s_pmt", "Seller monthly payment", 25000.0, S.FMT_CUR)
    item("s_first", "First seller payment month", 3, S.FMT_INT, "3 = September")
    item("s_last", "Last seller payment month", 6, S.FMT_INT, "6 = December")
    item("s_rate", "Seller note simple interest (APR)", 0.06, S.FMT_PCT)
    item("balloon_mo", "Balloon month (36-month rule / January)", 7, S.FMT_INT)
    item("bref_rate", "Balloon refinance rate (APR)", 0.09, S.FMT_PCT, "funds the Jan balloon when full-refi toggle is OFF")
    item("bref_term", "Balloon refinance term (months)", 120, S.FMT_INT)
    item("refi_on", "FULL-REFI TOGGLE (1 = $500K/6%/36 note - BASE CASE)", 1, S.FMT_INT, "base: sellers paid off Sept; company services $15,211/mo x36")
    item("refi_mo", "Full-refi funding month", 3, S.FMT_INT, "3 = September")
    item("refi_amt", "Full-refi amount", 500000.0, S.FMT_CUR)
    item("refi_rate", "Full-refi rate (APR)", 0.06, S.FMT_PCT)
    item("refi_term", "Full-refi term (months)", 36, S.FMT_INT)
    item("sba_on", "SBA TOGGLE (1 = SBA 7(a) funds)", 0, S.FMT_INT)
    item("sba_mo", "SBA funding month", 3, S.FMT_INT)
    item("sba_amt", "SBA amount", 450000.0, S.FMT_CUR)
    item("sba_rate", "SBA rate (APR)", 0.105, S.FMT_PCT)
    item("sba_term", "SBA term (months)", 180, S.FMT_INT)
    item("amort_yrs", "License intangible amortization (yrs, Acct 7200)", 15, S.FMT_INT)
    item("chow_mo", "CHOW / go-live month (amortization starts)", 1, S.FMT_INT)

    sec("CENSUS  (admissions row lives on Census Waterfall - blue)")
    item("days_mo", "Days per month (model uniform)", 30.4, S.FMT_NUM1)
    item("alos", "Average length of stay (days)", 82.0, S.FMT_NUM1, "drives discharges = BOM x days/ALOS")
    item("death_share", "Deaths as share of discharges", 0.55, S.FMT_PCT)

    sec("SCENARIO CASES  (1 = Base, 2 = Downside, 3 = Upside; active values are formulas)")
    item("case", "ACTIVE CASE", 1, S.FMT_INT, "flip 1/2/3 - drives capture, payroll adj, CHOW hold")
    bk.lbl(ws, r, "Case matrix:"); bk.lbl(ws, r, "Base", c=2, bold=True)
    bk.lbl(ws, r, "Downside", c=3, bold=True); bk.lbl(ws, r, "Upside", c=4, bold=True); r += 1
    for key, label, vals, fmt in [("cap", "Census capture", (1.0, 0.85, 1.1), S.FMT_PCT),
                                  ("padj", "Payroll adjustment", (0.0, 0.10, 0.0), S.FMT_PCT),
                                  ("hld", "CHOW hold (months)", (1, 2, 0), S.FMT_INT)]:
        bk.lbl(ws, r, f"  {label}")
        for j, v in enumerate(vals):
            bk.inp(ws, r, 2 + j, v, fmt, key=f"{key}{j+1}")
        r += 1
    bk.lbl(ws, r, "ACTIVE census capture (formula)")
    bk.fml(ws, r, 2, f"=CHOOSE({bk.addr['case']},{bk.addr['cap1']},{bk.addr['cap2']},{bk.addr['cap3']})", S.FMT_PCT)
    bk.addr["capture"] = f"'{ws.title}'!$B${r}"; r += 1
    bk.lbl(ws, r, "ACTIVE payroll adjustment (formula)")
    bk.fml(ws, r, 2, f"=CHOOSE({bk.addr['case']},{bk.addr['padj1']},{bk.addr['padj2']},{bk.addr['padj3']})", S.FMT_PCT)
    bk.addr["pay_adj"] = f"'{ws.title}'!$B${r}"; r += 1
    bk.lbl(ws, r, "ACTIVE CHOW hold (formula)")
    bk.fml(ws, r, 2, f"=CHOOSE({bk.addr['case']},{bk.addr['hld1']},{bk.addr['hld2']},{bk.addr['hld3']})", S.FMT_INT)
    bk.addr["hold_mo"] = f"'{ws.title}'!$B${r}"; r += 1

    sec("REVENUE - CMS FY2026 SMITH COUNTY  (Acct 3000/3010)")
    item("hi_rate", "Medicare RHC HIGH rate, days 1-60 ($/day, net of seq.)", 212.76, S.FMT_RATE, "217.10 x 98 pct")
    item("lo_rate", "Medicare RHC LOW rate, days 61+ ($/day, net of seq.)", 167.70, S.FMT_RATE, "171.12 x 98 pct")
    item("mcaid_rate", "Texas Medicaid rate ($/day)", 155.0, S.FMT_RATE, "verify with TMHP")
    item("mix_mcare", "Medicare payer mix", 1.0, S.FMT_PCT)
    item("hi_pct", "% Medicare days at HIGH rate", 0.15, S.FMT_PCT, "transfer panel mostly days 61+")
    item("bill_fee", "Billing company fee (% of gross, Acct 6510-Z)", 0.015, S.FMT_PCT)
    item("rate_esc", "CMS rate escalation (annual, yrs 2-3)", 0.025, S.FMT_PCT)

    sec("COLLECTIONS TIMING  (cash engine)")
    item("col_m0", "% collected same month", 0.0, S.FMT_PCT)
    item("col_m1", "% collected month +1", 1.0, S.FMT_PCT)
    item("col_m2", "% collected month +2", 0.0, S.FMT_PCT)

    sec("CLINICAL LABOR  (Acct 4000; benefits/WC per audit)")
    item("benefits", "Benefits & payroll tax load (employed)", 0.22, S.FMT_PCT, "post-audit: bottom of 22-28% band")
    item("wcomp", "Workers comp (% of employed payroll)", 0.03, S.FMT_PCT)
    item("sal_esc", "Salary escalation (annual, yrs 2-3)", 0.03, S.FMT_PCT)
    item("rn_sal", "RN Case Manager salary", 80000.0, S.FMT_CUR)
    item("rn_case", "RN caseload (patients : 1 FTE)", 12.0, S.FMT_NUM1)
    item("oc_sal", "RN On-Call salary (per 1.0 FTE)", 80000.0, S.FMT_CUR)
    item("oc_fte_min", "RN On-Call minimum FTE", 0.25, S.FMT_NUM2)
    item("cna_sal", "Hospice Aide / CNA salary", 48000.0, S.FMT_CUR)
    item("cna_case", "CNA caseload (patients : 1 FTE)", 10.0, S.FMT_NUM1)
    item("msw_sal", "MSW salary", 60000.0, S.FMT_CUR)
    item("msw_case", "MSW caseload (patients : 1 FTE)", 30.0, S.FMT_NUM1)
    item("ch_base", "Chaplain base PRN contract ($/mo, Acct 4020-0-6-7)", 600.0, S.FMT_CUR)
    item("medd", "Medical director stipend ($/mo, Acct 4020-0-6-M)", 4000.0, S.FMT_CUR)
    item("medd_start", "Medical director start month", 1, S.FMT_INT)

    sec("PRN PER-VISIT SUPPLEMENT  (1099; fires when FTE capacity < required visits)")
    item("rn_req", "RN required visits / patient / month", 2.0, S.FMT_NUM1, "CoP 418.56 - q14 days")
    item("rn_cap", "RN visits capacity / FTE / month", 24.0, S.FMT_NUM1)
    item("rn_prn", "RN PRN rate ($/visit)", 85.0, S.FMT_RATE)
    item("cna_req", "Aide required visits / patient / month", 8.0, S.FMT_NUM1)
    item("cna_cap", "Aide visits capacity / FTE / month", 80.0, S.FMT_NUM1)
    item("cna_prn", "Aide PRN rate ($/visit)", 45.0, S.FMT_RATE)
    item("msw_req", "MSW required visits / patient / month", 1.0, S.FMT_NUM1)
    item("msw_cap", "MSW visits capacity / FTE / month", 20.0, S.FMT_NUM1)
    item("msw_prn", "MSW PRN rate ($/visit)", 75.0, S.FMT_RATE)
    item("ch_req", "Chaplain required visits / patient / month", 1.0, S.FMT_NUM1)
    item("ch_base_v", "Chaplain visits covered by base contract", 8.0, S.FMT_NUM1)
    item("ch_prn", "Chaplain PRN rate ($/visit)", 60.0, S.FMT_RATE)

    sec("INDIRECT LABOR - ADMIN & LEADERSHIP  (Acct 4000-0-*)")
    item("silas", "Silas Shelton - Executive Director salary", 150000.0, S.FMT_CUR)
    item("silas_start", "  start month", 2, S.FMT_INT)
    item("dana", "Dana Davenport - Director of Nursing salary", 120000.0, S.FMT_CUR)
    item("dana_start", "  start month", 2, S.FMT_INT)
    item("brad", "Brad Woodard - Director of Sales salary", 120000.0, S.FMT_CUR)
    item("brad_start", "  start month", 1, S.FMT_INT)
    item("rhonda", "Rhonda Smith - ADON / Intake salary", 60000.0, S.FMT_CUR)
    item("rhonda_start", "  start month", 1, S.FMT_INT)
    item("volber", "Volunteer + Bereavement Coordinator salary", 50000.0, S.FMT_CUR)
    item("volber_start", "  start month", 1, S.FMT_INT)
    item("qapi", "QAPI / Compliance salary (per 1.0 FTE)", 70000.0, S.FMT_CUR)
    item("qapi_adc", "QAPI 0.5 FTE fires at ADC >=", 40.0, S.FMT_NUM1)

    sec("DIRECT PATIENT CARE - per patient-day  (Acct 5050-5120)")
    item("dme_pd", "DME rental ($/PD, 5050)", 9.0, S.FMT_RATE)
    item("rx_pd", "Medications ($/PD, 5100)", 3.5, S.FMT_RATE)
    item("sup_pd", "Medical supplies ($/PD, 5090)", 3.0, S.FMT_RATE)
    item("mile_pd", "Clinical mileage ($/PD, 5110)", 0.5, S.FMT_RATE)
    item("mobile", "Mobile communications - clinical ($/mo, 5120)", 50.0, S.FMT_CUR)

    sec("FACILITY  (700 N Main St - Acct 6010-6150; rent steps with census)")
    item("rent_y1", "Rent - Year 1 ($/mo)", 3000.0, S.FMT_CUR)
    item("rent_y2", "Rent - Year 2 ($/mo)", 5000.0, S.FMT_CUR)
    item("rent_y3", "Rent - Year 3 ($/mo)", 7000.0, S.FMT_CUR)
    item("util", "Utilities ($/mo, 6150)", 250.0, S.FMT_CUR)
    item("maint", "Maintenance & repairs ($/mo, 6060)", 100.0, S.FMT_CUR)
    item("alarm", "Alarm system ($/mo, 6010)", 100.0, S.FMT_CUR)

    sec("G&A  (MAX(fixed, % of net revenue) per audit pattern)")
    item("emr_base", "EMR base platform ($/mo)", 0.0, S.FMT_CUR)
    item("emr_pp", "EMR per active patient ($/patient/mo)", 75.0, S.FMT_CUR)
    for key, label, acct, fixed, pct, start in GA_LINES + [CONT]:
        if key == "ga_emr":
            continue
        bk.lbl(ws, r, f"{label}  ({acct})")
        bk.inp(ws, r, 2, fixed, S.FMT_CUR, key=key + "_f")
        bk.inp(ws, r, 3, pct, S.FMT_PCT2, key=key + "_p")
        if start > 1:
            bk.lbl(ws, r, f"starts month {start}", c=4, italic=True)
        r += 1

    sec("OTHER")
    item("floor", "Minimum-cash floor (revolver draws to hold this)", 25000.0, S.FMT_CUR)
    item("rev_limit", "Working-capital revolver commitment", 250000.0, S.FMT_CUR, "balancing facility - PE convention: cash never goes negative")
    item("rev_rate", "Revolver rate (APR)", 0.105, S.FMT_PCT)
    item("tx_tax", "TX franchise/margin tax (% of NPR if profitable)", 0.00375, S.FMT_PCT2)
    item("cap_limit", "Medicare aggregate cap / beneficiary (FY2026)", 33900.0, S.FMT_CUR, "monitor on Medicare CAP logic")
    return ws


# =====================================================================
def census(bk: OB):
    ws = bk.sheet("Census Waterfall", tab="2E75B6")
    bk.title(ws, "CENSUS WATERFALL - admissions are the operating lever (blue). Discharges follow ALOS. 36 months.")
    bk.widths(ws)
    A = bk.addr; R = bk.rows["Census Waterfall"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    bk.lbl(ws, r, "Beginning-of-month census (BOM)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{r+3}" if i else "0"
        bk.fml(ws, r, MC0 + i, f"={prev}", S.FMT_NUM1)
    R["bom"] = r; r += 1
    bk.lbl(ws, r, "+ New admissions  [INPUT - 7/6 plan default]")
    for i in range(NM):
        bk.inp(ws, r, MC0 + i, ADMITS[i], S.FMT_NUM1)
    R["admits"] = r; r += 1
    bk.lbl(ws, r, "- Discharges (deaths + revocations) = BOM x days/ALOS")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['bom']}*{A['days_mo']}/{A['alos']}", S.FMT_NUM1)
    R["disch"] = r; r += 1
    bk.lbl(ws, r, "End-of-month census (EOM)", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['bom']}+{c}{R['admits']}-{c}{R['disch']}", S.FMT_NUM1, bold=True)
    R["eom"] = r; r += 1
    bk.lbl(ws, r, "  of which deaths (bereavement trigger)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['disch']}*{A['death_share']}", S.FMT_NUM1)
    R["deaths"] = r; r += 1
    bk.lbl(ws, r, "ADC = (BOM + EOM)/2 x capture", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"=({c}{R['bom']}+{c}{R['eom']})/2*{A['capture']}",
               S.FMT_NUM1, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["adc"] = r; r += 1
    bk.lbl(ws, r, "Patient days (ADC x days/month)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['adc']}*{A['days_mo']}", S.FMT_NUM)
    R["pd"] = r; r += 1
    bk.lbl(ws, r, "Monthly discharge rate (of BOM)", italic=True)
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=IF({mlet(i)}{R['bom']}=0,0,{mlet(i)}{R['disch']}/{mlet(i)}{R['bom']})", S.FMT_PCT)
    r += 1
    return ws


# =====================================================================
def revenue(bk: OB):
    ws = bk.sheet("Revenue Model", tab="2E75B6")
    bk.title(ws, "REVENUE - patient days x CMS rates (high/low split), less billing fee -> NET PATIENT REVENUE (accrual).")
    bk.widths(ws)
    A = bk.addr; C = bk.rows["Census Waterfall"]; R = bk.rows["Revenue Model"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    def row(label, f_tmpl, key, fmt=S.FMT_CUR, bold=False, fill=None):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f_tmpl(i), fmt, bold=bold, fill=fill)
        R[key] = r; r += 1
    esc = lambda i: f"(1+{A['rate_esc']})^{yr(i)}"
    row("Medicare RHC HIGH (days 1-60, 3000-R-6-0)",
        lambda i: f"='Census Waterfall'!{mlet(i)}{C['pd']}*{A['mix_mcare']}*{A['hi_pct']}*{A['hi_rate']}*{esc(i)}", "hi")
    row("Medicare RHC LOW (days 61+, 3000-R-6-0)",
        lambda i: f"='Census Waterfall'!{mlet(i)}{C['pd']}*{A['mix_mcare']}*(1-{A['hi_pct']})*{A['lo_rate']}*{esc(i)}", "lo")
    row("Medicaid RHC (3010-R-6-0)",
        lambda i: f"='Census Waterfall'!{mlet(i)}{C['pd']}*(1-{A['mix_mcare']})*{A['mcaid_rate']}*{esc(i)}", "mcaid")
    row("TOTAL GROSS REVENUE", lambda i: f"={mlet(i)}{R['hi']}+{mlet(i)}{R['lo']}+{mlet(i)}{R['mcaid']}",
        "gross", bold=True)
    row("Less: billing company fee (6510-Z)", lambda i: f"=-{mlet(i)}{R['gross']}*{A['bill_fee']}", "fee")
    row("NET PATIENT REVENUE (accrual)", lambda i: f"={mlet(i)}{R['gross']}+{mlet(i)}{R['fee']}",
        "npr", bold=True, fill=S.fill(S.LIGHTBLUE))
    return ws


# =====================================================================
def staffing(bk: OB):
    ws = bk.sheet("Staffing", tab="548235")
    bk.title(ws, "STAFFING - FTEs step with census (caseload formulas); PRN 1099 supplement fires when FTE capacity < required visits.")
    bk.widths(ws)
    A = bk.addr; C = bk.rows["Census Waterfall"]; R = bk.rows["Staffing"]
    adcref = lambda i: f"'Census Waterfall'!{mlet(i)}{C['adc']}"
    r = 4
    hr = bk.mheader(ws, r); r = hr
    padj = f"(1+{A['pay_adj']})"

    bk.section(ws, r, "CLINICAL IDG - EMPLOYED (Acct 4000; FTE = census/caseload in 0.5 steps)"); r += 1
    fte_rows = {}
    for key, label, salkey, casekey in CLIN:
        bk.lbl(ws, r, f"{label} - FTE")
        for i in range(NM):
            if key == "oc":
                f = f"=MAX({A['oc_fte_min']},ROUNDUP({adcref(i)}/60*2,0)/2)"
            else:
                f = f"=ROUNDUP({adcref(i)}/{A[casekey]}*2,0)/2"
            bk.fml(ws, r, MC0 + i, f, S.FMT_NUM2)
        fte_rows[key] = r; R[key + "_fte"] = r; r += 1
        bk.lbl(ws, r, f"{label} - wages")
        for i in range(NM):
            bk.fml(ws, r, MC0 + i,
                   f"={mlet(i)}{fte_rows[key]}*{A[salkey]}/12*(1+{A['sal_esc']})^{yr(i)}*{padj}", S.FMT_CUR)
        R[key + "_w"] = r; r += 1
    bk.lbl(ws, r, "Employed clinical wages", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['rn_w']}+{c}{R['oc_w']}+{c}{R['cna_w']}+{c}{R['msw_w']}", S.FMT_CUR, bold=True)
    R["clin_w2"] = r; r += 1

    bk.section(ws, r, "PRN 1099 SUPPLEMENT (Acct 4025 - no benefits)"); r += 1
    for key, label, reqk, capk, ratek, ftekey in PRN:
        bk.lbl(ws, r, label)
        for i in range(NM):
            if ftekey:
                cap = f"{mlet(i)}{R[ftekey+'_fte']}*{A[capk]}"
            else:
                cap = f"{A[capk]}"
            bk.fml(ws, r, MC0 + i,
                   f"=MAX(0,{adcref(i)}*{A[reqk]}-{cap})*{A[ratek]}*{padj}", S.FMT_CUR)
        R[key] = r; r += 1
    bk.lbl(ws, r, "Total PRN supplement", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['prn_rn']}+{c}{R['prn_cna']}+{c}{R['prn_msw']}+{c}{R['prn_ch']}", S.FMT_CUR, bold=True)
    R["prn_tot"] = r; r += 1

    bk.section(ws, r, "CONTRACT & INDIRECT (leadership per Control Tower start months)"); r += 1
    bk.lbl(ws, r, "Chaplain base contract (4020-0-6-7)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={A['ch_base']}*(1+{A['sal_esc']})^{yr(i)}", S.FMT_CUR)
    R["chap"] = r; r += 1
    bk.lbl(ws, r, "Medical director (4020-0-6-M)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=IF({i+1}>={A['medd_start']},{A['medd']}*(1+{A['sal_esc']})^{yr(i)},0)", S.FMT_CUR)
    R["medd"] = r; r += 1
    for key, label in [("silas", "Silas Shelton - Executive Director (4000-0-A-0)"),
                       ("dana", "Dana Davenport - Director of Nursing (4000-0-C-0)"),
                       ("brad", "Brad Woodard - Director of Sales (4000-0-M-0)"),
                       ("rhonda", "Rhonda Smith - ADON / Intake (4000-0-I-1)"),
                       ("volber", "Volunteer + Bereavement Coordinator (4000-0-A-0)")]:
        bk.lbl(ws, r, label)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i,
                   f"=IF({i+1}>={A[key+'_start']},{A[key]}/12*(1+{A['sal_esc']})^{yr(i)}*{padj},0)", S.FMT_CUR)
        R[key] = r; r += 1
    bk.lbl(ws, r, "QAPI / Compliance 0.5 FTE (fires at ADC threshold)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=IF({adcref(i)}>={A['qapi_adc']},0.5*{A['qapi']}/12*(1+{A['sal_esc']})^{yr(i)}*{padj},0)", S.FMT_CUR)
    R["qapi"] = r; r += 1
    bk.lbl(ws, r, "Indirect employed wages", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={c}{R['silas']}+{c}{R['dana']}+{c}{R['brad']}+{c}{R['rhonda']}+{c}{R['volber']}+{c}{R['qapi']}",
               S.FMT_CUR, bold=True)
    R["ind_w2"] = r; r += 1

    bk.section(ws, r, "LOADS (employed wages only - Acct 4110/4120/4130)"); r += 1
    for key, label, base, ratekey in [
        ("clin_ben", "Clinical benefits load", "clin_w2", "benefits"),
        ("clin_wc", "Clinical workers comp", "clin_w2", "wcomp"),
        ("ind_ben", "Indirect benefits load", "ind_w2", "benefits"),
        ("ind_wc", "Indirect workers comp", "ind_w2", "wcomp")]:
        bk.lbl(ws, r, label)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R[base]}*{A[ratekey]}", S.FMT_CUR)
        R[key] = r; r += 1
    bk.lbl(ws, r, "TOTAL CLINICAL LABOR (employed + loads + contract + PRN)", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={c}{R['clin_w2']}+{c}{R['clin_ben']}+{c}{R['clin_wc']}+{c}{R['chap']}+{c}{R['medd']}+{c}{R['prn_tot']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["clin_tot"] = r; r += 1
    bk.lbl(ws, r, "TOTAL INDIRECT LABOR (employed + loads)", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['ind_w2']}+{c}{R['ind_ben']}+{c}{R['ind_wc']}", S.FMT_CUR, bold=True,
               fill=S.fill(S.LIGHTBLUE))
    R["ind_tot"] = r; r += 1
    bk.lbl(ws, r, "Total employed payroll cash (wages + loads)", italic=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={c}{R['clin_w2']}+{c}{R['ind_w2']}+{c}{R['clin_ben']}+{c}{R['clin_wc']}+{c}{R['ind_ben']}+{c}{R['ind_wc']}",
               S.FMT_CUR)
    R["payroll_cash"] = r; r += 1
    return ws


# =====================================================================
def opbudget(bk: OB):
    ws = bk.sheet("Operating Budget", tab="7F7F7F")
    bk.title(ws, "OPERATING BUDGET - direct patient care per PD; facility (stepped rent); G&A = MAX(fixed, % of NPR) per audit pattern.")
    bk.widths(ws)
    A = bk.addr; C = bk.rows["Census Waterfall"]; V = bk.rows["Revenue Model"]; R = bk.rows["Operating Budget"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    pd = lambda i: f"'Census Waterfall'!{mlet(i)}{C['pd']}"
    npr = lambda i: f"'Revenue Model'!{mlet(i)}{V['npr']}"

    bk.section(ws, r, "DIRECT PATIENT CARE (contracted variable - Acct 5050-5120)"); r += 1
    for key, label, ratekey in [("dme", "DME rental (5050)", "dme_pd"), ("rx", "Medications (5100)", "rx_pd"),
                                ("sup", "Medical supplies (5090)", "sup_pd"), ("mile", "Clinical mileage (5110)", "mile_pd")]:
        bk.lbl(ws, r, label)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f"={pd(i)}*{A[ratekey]}", S.FMT_CUR)
        R[key] = r; r += 1
    bk.lbl(ws, r, "Mobile communications - clinical (5120)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={A['mobile']}", S.FMT_CUR)
    R["mobile"] = r; r += 1
    bk.lbl(ws, r, "TOTAL DIRECT PATIENT CARE", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['dme']}+{c}{R['rx']}+{c}{R['sup']}+{c}{R['mile']}+{c}{R['mobile']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["dpc"] = r; r += 1

    bk.section(ws, r, "FACILITY (Acct 6010-6150; rent steps by year)"); r += 1
    bk.lbl(ws, r, "Rent (6045-Z)")
    for i in range(NM):
        rent = [A["rent_y1"], A["rent_y2"], A["rent_y3"]][yr(i)]
        bk.fml(ws, r, MC0 + i, f"={rent}", S.FMT_CUR)
    R["rent"] = r; r += 1
    for key, label in [("util", "Utilities (6150)"), ("maint", "Maintenance & repairs (6060)"), ("alarm", "Alarm (6010)")]:
        bk.lbl(ws, r, label)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f"={A[key]}", S.FMT_CUR)
        R[key] = r; r += 1
    bk.lbl(ws, r, "TOTAL FACILITY", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['rent']}+{c}{R['util']}+{c}{R['maint']}+{c}{R['alarm']}", S.FMT_CUR, bold=True)
    R["fac"] = r; r += 1

    bk.section(ws, r, "G&A (MAX(fixed, % of NPR); legal starts M4)"); r += 1
    bk.lbl(ws, r, "EMR - base + per active patient (6530-Z)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={A['emr_base']}+'Census Waterfall'!{mlet(i)}{C['eom']}*{A['emr_pp']}", S.FMT_CUR)
    R["ga_emr"] = r; r += 1
    ga_keys = ["ga_emr"]
    for key, label, acct, fixed, pct, start in GA_LINES:
        if key == "ga_emr":
            continue
        bk.lbl(ws, r, f"{label} ({acct})")
        for i in range(NM):
            base = f"MAX({A[key+'_f']},{npr(i)}*{A[key+'_p']})"
            f = f"=IF({i+1}>={start},{base},0)" if start > 1 else f"={base}"
            bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
        R[key] = r; ga_keys.append(key); r += 1
    bk.lbl(ws, r, "TOTAL G&A (excl. contingency)", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, "=" + "+".join(f"{c}{R[k]}" for k in ga_keys), S.FMT_CUR, bold=True,
               fill=S.fill(S.LIGHTBLUE))
    R["ga"] = r; r += 1
    key, label, acct, fixed, pct, start = CONT
    bk.lbl(ws, r, f"{label} ({acct})")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=MAX({A[key+'_f']},{npr(i)}*{A[key+'_p']})", S.FMT_CUR)
    R["cont"] = r; r += 1
    return ws


# =====================================================================
def pl(bk: OB):
    ws = bk.sheet("P&L", tab="C00000")
    bk.title(ws, "P&L - COA-backed, accrual. Clinical labor above gross profit. Every cell is a link or formula.")
    bk.widths(ws)
    A = bk.addr
    C = bk.rows["Census Waterfall"]; V = bk.rows["Revenue Model"]
    ST = bk.rows["Staffing"]; OB_ = bk.rows["Operating Budget"]; R = bk.rows["P&L"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    def link(label, sheet, srow, key, bold=False, fill=None, fmt=S.FMT_CUR):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f"='{sheet}'!{mlet(i)}{srow}", fmt, bold=bold, fill=fill, link=True)
        R[key] = r; r += 1
    def calc(label, f_tmpl, key, bold=False, fill=None, fmt=S.FMT_CUR):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f_tmpl(i), fmt, bold=bold, fill=fill)
        R[key] = r; r += 1

    link("Average Daily Census", "Census Waterfall", C["adc"], "adc", fmt=S.FMT_NUM1)
    bk.section(ws, r, "REVENUE"); r += 1
    link("Medicare RHC high / low / Medicaid -> gross (3000/3010)", "Revenue Model", V["gross"], "gross")
    link("Less: billing company fee (6510-Z)", "Revenue Model", V["fee"], "fee")
    link("NET PATIENT REVENUE", "Revenue Model", V["npr"], "npr", bold=True, fill=S.fill(S.LIGHTBLUE))
    bk.section(ws, r, "COST OF SERVICES"); r += 1
    link("Total direct patient care (5050-5120)", "Operating Budget", OB_["dpc"], "dpc")
    link("Total direct clinical labor (4000/4025/4020)", "Staffing", ST["clin_tot"], "clin")
    calc("TOTAL COST OF SERVICES", lambda i: f"={mlet(i)}{R['dpc']}+{mlet(i)}{R['clin']}", "cos", bold=True)
    calc("GROSS PROFIT", lambda i: f"={mlet(i)}{R['npr']}-{mlet(i)}{R['cos']}", "gp", bold=True,
         fill=S.fill(S.LIGHTBLUE))
    calc("Gross margin %", lambda i: f"=IF({mlet(i)}{R['npr']}=0,0,{mlet(i)}{R['gp']}/{mlet(i)}{R['npr']})",
         "gm", fmt=S.FMT_PCT)
    bk.section(ws, r, "OPERATING EXPENSES"); r += 1
    link("Total indirect labor (4000-0-*)", "Staffing", ST["ind_tot"], "ind")
    link("Total facility (6010-6150)", "Operating Budget", OB_["fac"], "fac")
    link("Total G&A (6510-6810)", "Operating Budget", OB_["ga"], "ga")
    calc("License intangible amortization (7200)",
         lambda i: f"=IF({i+1}>={A['chow_mo']},{A['price']}/{A['amort_yrs']}/12,0)", "amort")
    calc("TOTAL OPERATING EXPENSES",
         lambda i: f"={mlet(i)}{R['ind']}+{mlet(i)}{R['fac']}+{mlet(i)}{R['ga']}+{mlet(i)}{R['amort']}",
         "opex", bold=True)
    calc("EBIT (NOI)", lambda i: f"={mlet(i)}{R['gp']}-{mlet(i)}{R['opex']}", "ebit", bold=True)
    calc("EBITDA", lambda i: f"={mlet(i)}{R['ebit']}+{mlet(i)}{R['amort']}", "ebitda", bold=True,
         fill=S.fill(S.LIGHTBLUE))
    calc("EBITDA margin %", lambda i: f"=IF({mlet(i)}{R['npr']}=0,0,{mlet(i)}{R['ebitda']}/{mlet(i)}{R['npr']})",
         "em", fmt=S.FMT_PCT)
    link("Contingency provision (below EBITDA, 6640-Z)", "Operating Budget", OB_["cont"], "cont")
    return ws


# =====================================================================
def debt_and_cash(bk: OB):
    ws = bk.sheet("Cash Flow & Runway", tab="C00000")
    bk.title(ws, "CASH FLOW - collections timing + CHOW hold; payroll in month; DPC net-30; full debt schedules with toggles. Direct method.")
    bk.widths(ws)
    A = bk.addr
    V = bk.rows["Revenue Model"]; ST = bk.rows["Staffing"]; OB_ = bk.rows["Operating Budget"]
    P = bk.rows["P&L"]; R = bk.rows["Cash Flow & Runway"]
    r = 4
    R["_hdr"] = r
    hr = bk.mheader(ws, r); r = hr
    npr = lambda i: f"'Revenue Model'!{mlet(i)}{V['npr']}"

    bk.section(ws, r, "COLLECTIONS ENGINE"); r += 1
    bk.lbl(ws, r, "Collection month for service month (lag + hold on first claims)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={i+1}+1+MAX(0,{A['hold_mo']}-{i})", S.FMT_INT)
    R["cmo"] = r; r += 1
    bk.lbl(ws, r, "NPR earned (link)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={npr(i)}", S.FMT_CUR, link=True)
    R["earned"] = r; r += 1
    bk.lbl(ws, r, "CASH COLLECTIONS", bold=True)
    for i in range(NM):
        # same-month share ignores hold only if hold=0; +1/+2 shares via collection-month map
        f = (f"={A['col_m0']}*SUMPRODUCT(({MRANGE(R['cmo'])}-1={i+1})*{MRANGE(R['earned'])})"
             f"+{A['col_m1']}*SUMPRODUCT(({MRANGE(R['cmo'])}={i+1})*{MRANGE(R['earned'])})"
             f"+{A['col_m2']}*SUMPRODUCT(({MRANGE(R['cmo'])}={i})*{MRANGE(R['earned'])})")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["coll"] = r; r += 1
    bk.lbl(ws, r, "A/R balance (earned - collected, cumulative)", italic=True)
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUM($C${R['earned']}:{mlet(i)}{R['earned']})-SUM($C${R['coll']}:{mlet(i)}{R['coll']})", S.FMT_CUR)
    R["ar"] = r; r += 1

    bk.section(ws, r, "DEBT - SELLER NOTE (accepted terms; auto-payoff if full-refi ON)"); r += 1
    bk.lbl(ws, r, "Seller beg balance")
    for i in range(NM):
        first = f"={A['price']}-{A['down']}"
        bk.fml(ws, r, MC0 + i, first if i == 0 else f"={mlet(i-1)}{r+4}", S.FMT_CUR)
    R["s_beg"] = r; r += 1
    bk.lbl(ws, r, "Seller interest accrued (6% simple)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['s_beg']}*{A['s_rate']}/12", S.FMT_CUR)
    R["s_int"] = r; r += 1
    bk.lbl(ws, r, "Seller principal paid")
    for i in range(NM):
        f = (f"=IF({mlet(i)}{R['s_beg']}<=0,0,"
             f"IF(AND({A['refi_on']}=1,{i+1}={A['refi_mo']}),{mlet(i)}{R['s_beg']},"
             f"IF(AND({i+1}>={A['s_first']},{i+1}<={A['s_last']}),{A['s_pmt']},"
             f"IF({i+1}={A['balloon_mo']},{mlet(i)}{R['s_beg']},0))))")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["s_prin"] = r; r += 1
    bk.lbl(ws, r, "Seller interest paid (at payoff)")
    for i in range(NM):
        f = (f"=IF(OR(AND({A['refi_on']}=1,{i+1}={A['refi_mo']}),AND({A['refi_on']}=0,{i+1}={A['balloon_mo']})),"
             f"SUM($C${R['s_int']}:{mlet(i)}{R['s_int']}),0)")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["s_intpaid"] = r; r += 1
    bk.lbl(ws, r, "Seller end balance")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['s_beg']}-{mlet(i)}{R['s_prin']}", S.FMT_CUR)
    R["s_end"] = r; r += 1

    bk.section(ws, r, "DEBT - BALLOON REFI (auto when full-refi OFF) / FULL REFI / SBA (toggles)"); r += 1
    bk.lbl(ws, r, "Balloon refi amount")
    bk.fml(ws, r, 2,
           f"=IF({A['refi_on']}=1,0,SUMPRODUCT(($C${R['_hdr']}:$AL${R['_hdr']}={A['balloon_mo']})"
           f"*($C${R['s_prin']}:$AL${R['s_prin']}+$C${R['s_intpaid']}:$AL${R['s_intpaid']})))", S.FMT_CUR)
    bref = f"$B${r}"; R["bref_amt"] = r; r += 1
    specs = [("bref", "Balloon refi", bref, A["bref_rate"], A["bref_term"], f"{A['balloon_mo']}", f"{A['refi_on']}=0"),
             ("refi", "Full refi", A["refi_amt"], A["refi_rate"], A["refi_term"], f"{A['refi_mo']}", f"{A['refi_on']}=1"),
             ("sba", "SBA 7(a)", A["sba_amt"], A["sba_rate"], A["sba_term"], f"{A['sba_mo']}", f"{A['sba_on']}=1")]
    for key, label, amt, rate, term, fmo, cond in specs:
        bk.lbl(ws, r, f"{label} - payment (P&I)")
        for i in range(NM):
            f = (f"=IF(AND({cond},{i+1}>{fmo},{i+1}<={fmo}+{term}),-PMT({rate}/12,{term},{amt}),0)")
            bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
        R[key + "_pmt"] = r; r += 1
        bk.lbl(ws, r, f"{label} - balance (end)")
        for i in range(NM):
            prev = f"{mlet(i-1)}{r}" if i else "0"
            f = (f"=IF(NOT({cond}),0,IF({i+1}<{fmo},0,IF({i+1}={fmo},{amt},"
                 f"MAX(0,{prev}*(1+{rate}/12)-{mlet(i)}{R[key+'_pmt']}))))")
            bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
        R[key + "_bal"] = r; r += 1
        bk.lbl(ws, r, f"{label} - interest")
        for i in range(NM):
            prev = f"{mlet(i-1)}{R[key+'_bal']}" if i else "0"
            bk.fml(ws, r, MC0 + i, f"={prev}*{rate}/12*({mlet(i)}{R[key+'_pmt']}>0)", S.FMT_CUR)
        R[key + "_int"] = r; r += 1

    # ---- precompute row layout so earlier rows can reference later ones ----
    r0 = r
    L = {}
    seq = ["_sec_ni", "pretax", "tax", "ni", "_sec_cash", "cin", "cpay", "cdpc", "cga",
           "ctax", "csell", "cfin", "rev_int", "cbefore", "rev_draw", "rev_repay",
           "rev_bal", "cash", "dscr", "mincash"]
    for k in seq:
        L[k] = r0; r0 += 1
    R.update({k: v for k, v in L.items() if not k.startswith("_sec")})

    bk.section(ws, L["_sec_ni"], "NET INCOME (after interest, revolver interest & TX margin tax)")
    bk.lbl(ws, L["pretax"], "Pre-tax income")
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, L["pretax"], MC0 + i,
               f"='P&L'!{c}{P['ebit']}-'P&L'!{c}{P['cont']}-{c}{R['s_int']}-{c}{R['bref_int']}"
               f"-{c}{R['refi_int']}-{c}{R['sba_int']}-{c}{L['rev_int']}", S.FMT_CUR)
    bk.lbl(ws, L["tax"], "TX margin tax")
    for i in range(NM):
        bk.fml(ws, L["tax"], MC0 + i,
               f"=IF({mlet(i)}{L['pretax']}>0,{mlet(i)}{R['earned']}*{A['tx_tax']},0)", S.FMT_CUR)
    bk.lbl(ws, L["ni"], "NET INCOME", bold=True)
    for i in range(NM):
        bk.fml(ws, L["ni"], MC0 + i, f"={mlet(i)}{L['pretax']}-{mlet(i)}{L['tax']}", S.FMT_CUR, bold=True)

    bk.section(ws, L["_sec_cash"], "CASH (direct method; revolver is the balancing facility)")
    bk.lbl(ws, L["cin"], "Cash collections")
    for i in range(NM):
        bk.fml(ws, L["cin"], MC0 + i, f"={mlet(i)}{R['coll']}", S.FMT_CUR, link=True)
    bk.lbl(ws, L["cpay"], "Payroll + contract labor (in month)")
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, L["cpay"], MC0 + i,
               f"=-'Staffing'!{c}{ST['payroll_cash']}-'Staffing'!{c}{ST['chap']}-'Staffing'!{c}{ST['medd']}"
               f"-'Staffing'!{c}{ST['prn_tot']}", S.FMT_CUR)
    bk.lbl(ws, L["cdpc"], "Direct patient care (net-30: prior month)")
    for i in range(NM):
        prev = f"'Operating Budget'!{mlet(i-1)}{OB_['dpc']}" if i else "0"
        bk.fml(ws, L["cdpc"], MC0 + i, f"=-{prev}", S.FMT_CUR)
    bk.lbl(ws, L["cga"], "Facility + G&A + contingency (in month)")
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, L["cga"], MC0 + i,
               f"=-'Operating Budget'!{c}{OB_['fac']}-'Operating Budget'!{c}{OB_['ga']}"
               f"-'Operating Budget'!{c}{OB_['cont']}", S.FMT_CUR)
    bk.lbl(ws, L["ctax"], "Tax paid")
    for i in range(NM):
        bk.fml(ws, L["ctax"], MC0 + i, f"=-{mlet(i)}{L['tax']}", S.FMT_CUR)
    bk.lbl(ws, L["csell"], "Seller payments (incl. down M1 + payoff interest)")
    for i in range(NM):
        down = f"-{A['down']}" if i == 0 else ""
        bk.fml(ws, L["csell"], MC0 + i, f"=-{mlet(i)}{R['s_prin']}-{mlet(i)}{R['s_intpaid']}{down}", S.FMT_CUR)
    bk.lbl(ws, L["cfin"], "Financing proceeds - payments (balloon refi / full refi / SBA)")
    for i in range(NM):
        f = (f"=IF(AND({A['refi_on']}=0,{i+1}={A['balloon_mo']}),{bref},0)"
             f"+IF(AND({A['refi_on']}=1,{i+1}={A['refi_mo']}),{A['refi_amt']},0)"
             f"+IF(AND({A['sba_on']}=1,{i+1}={A['sba_mo']}),{A['sba_amt']},0)"
             f"-{mlet(i)}{R['bref_pmt']}-{mlet(i)}{R['refi_pmt']}-{mlet(i)}{R['sba_pmt']}")
        bk.fml(ws, L["cfin"], MC0 + i, f, S.FMT_CUR)
    bk.lbl(ws, L["rev_int"], "Revolver interest (on prior-month balance)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{L['rev_bal']}" if i else "0"
        bk.fml(ws, L["rev_int"], MC0 + i, f"={prev}*{A['rev_rate']}/12", S.FMT_CUR)
    bk.lbl(ws, L["cbefore"], "Cash before revolver")
    for i in range(NM):
        prev = f"{mlet(i-1)}{L['cash']}" if i else f"{A['cash0']}-{A['startup']}"
        c = mlet(i)
        bk.fml(ws, L["cbefore"], MC0 + i,
               f"={prev}+{c}{L['cin']}+{c}{L['cpay']}+{c}{L['cdpc']}+{c}{L['cga']}+{c}{L['ctax']}"
               f"+{c}{L['csell']}+{c}{L['cfin']}-{c}{L['rev_int']}", S.FMT_CUR)
    bk.lbl(ws, L["rev_draw"], "Revolver draw (to hold minimum cash)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{L['rev_bal']}" if i else "0"
        bk.fml(ws, L["rev_draw"], MC0 + i,
               f"=MIN(MAX(0,{A['floor']}-{mlet(i)}{L['cbefore']}),{A['rev_limit']}-{prev})", S.FMT_CUR)
    bk.lbl(ws, L["rev_repay"], "Revolver repayment (surplus above minimum cash)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{L['rev_bal']}" if i else "0"
        bk.fml(ws, L["rev_repay"], MC0 + i,
               f"=MIN({prev},MAX(0,{mlet(i)}{L['cbefore']}-{A['floor']}))", S.FMT_CUR)
    bk.lbl(ws, L["rev_bal"], "Revolver balance (end)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{L['rev_bal']}" if i else "0"
        c = mlet(i)
        bk.fml(ws, L["rev_bal"], MC0 + i, f"={prev}+{c}{L['rev_draw']}-{c}{L['rev_repay']}", S.FMT_CUR)
    bk.lbl(ws, L["cash"], "ENDING CASH", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, L["cash"], MC0 + i,
               f"={c}{L['cbefore']}+{c}{L['rev_draw']}-{c}{L['rev_repay']}", S.FMT_CUR, bold=True,
               fill=S.fill(S.LIGHTBLUE))
    ws.conditional_formatting.add(f"C{L['cash']}:AL{L['cash']}", CellIsRule(operator="lessThan",
        formula=[bk.addr["floor"]], fill=S.fill("F4CCCC")))
    bk.lbl(ws, L["dscr"], "DSCR (EBITDA / recurring debt service)")
    for i in range(NM):
        c = mlet(i)
        ds = (f"({c}{R['bref_pmt']}+{c}{R['refi_pmt']}+{c}{R['sba_pmt']}"
              f"+IF(AND({i+1}>={A['s_first']},{i+1}<={A['s_last']},{A['refi_on']}=0),{A['s_pmt']},0))")
        bk.fml(ws, L["dscr"], MC0 + i, f"=IF({ds}>0,'P&L'!{c}{P['ebitda']}/{ds},\"-\")", S.FMT_MULT)
    bk.lbl(ws, L["mincash"], "Minimum cash (36 months)", bold=True)
    bk.fml(ws, L["mincash"], 2, f"=MIN(C{L['cash']}:AL{L['cash']})", S.FMT_CUR, bold=True)
    return ws


# =====================================================================
def balance_sheet(bk: OB):
    ws = bk.sheet("Balance Sheet", tab="7F7F7F")
    bk.title(ws, "BALANCE SHEET - no plug cells; the check row must be zero in every month.")
    bk.widths(ws)
    A = bk.addr
    CF = bk.rows["Cash Flow & Runway"]; OB_ = bk.rows["Operating Budget"]; R = bk.rows["Balance Sheet"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    def row(label, f_tmpl, key, bold=False):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f_tmpl(i), S.FMT_CUR, bold=bold)
        R[key] = r; r += 1
    cf = "Cash Flow & Runway"
    row("Cash", lambda i: f"='{cf}'!{mlet(i)}{CF['cash']}", "cash")
    row("Accounts receivable", lambda i: f"='{cf}'!{mlet(i)}{CF['ar']}", "ar")
    row("License intangible - gross (1980)", lambda i: f"=IF({i+1}>={A['chow_mo']},{A['price']},0)", "lic")
    row("Less: accumulated amortization",
        lambda i: f"=-SUM('P&L'!$C${bk.rows['P&L']['amort']}:{mlet(i)}{bk.rows['P&L']['amort']})", "accam")
    row("TOTAL ASSETS", lambda i: f"={mlet(i)}{R['cash']}+{mlet(i)}{R['ar']}+{mlet(i)}{R['lic']}+{mlet(i)}{R['accam']}",
        "assets", bold=True)
    row("A/P - direct patient care (net-30)", lambda i: f"='Operating Budget'!{mlet(i)}{OB_['dpc']}", "ap")
    row("Seller note payable + accrued interest",
        lambda i: (f"='{cf}'!{mlet(i)}{CF['s_end']}"
                   f"+SUM('{cf}'!$C${CF['s_int']}:{mlet(i)}{CF['s_int']})"
                   f"-SUM('{cf}'!$C${CF['s_intpaid']}:{mlet(i)}{CF['s_intpaid']})"), "sn")
    row("Balloon refi payable", lambda i: f"='{cf}'!{mlet(i)}{CF['bref_bal']}", "bref")
    row("Full refi payable", lambda i: f"='{cf}'!{mlet(i)}{CF['refi_bal']}", "refi")
    row("SBA loan payable", lambda i: f"='{cf}'!{mlet(i)}{CF['sba_bal']}", "sba")
    row("Revolver payable", lambda i: f"='{cf}'!{mlet(i)}{CF['rev_bal']}", "rev")
    row("TOTAL LIABILITIES",
        lambda i: f"={mlet(i)}{R['ap']}+{mlet(i)}{R['sn']}+{mlet(i)}{R['bref']}+{mlet(i)}{R['refi']}+{mlet(i)}{R['sba']}+{mlet(i)}{R['rev']}",
        "liab", bold=True)
    row("Contributed capital", lambda i: f"={A['cash0']}", "equity")
    row("Retained earnings (cum. net income - startup spend)",
        lambda i: f"=SUM('{cf}'!$C${CF['ni']}:{mlet(i)}{CF['ni']})-{A['startup']}", "re")
    row("TOTAL LIABILITIES + EQUITY",
        lambda i: f"={mlet(i)}{R['liab']}+{mlet(i)}{R['equity']}+{mlet(i)}{R['re']}", "le", bold=True)
    row("CHECK (assets - L-E) = 0", lambda i: f"={mlet(i)}{R['assets']}-{mlet(i)}{R['le']}", "check", bold=True)
    return ws


# =====================================================================
def summary(bk: OB):
    ws = bk.sheet("3-Year Summary", tab="BF8F00")
    bk.title(ws, "3-YEAR SUMMARY - annual rollup; GLOBAL DSCR is the operative SBA test (per audit).")
    ws.column_dimensions["A"].width = 44
    for c in "BCDE":
        ws.column_dimensions[c].width = 15
    P = bk.rows["P&L"]; CF = bk.rows["Cash Flow & Runway"]; A = bk.addr
    r = 4
    for j, h in enumerate(["", "Year 1", "Year 2", "Year 3"]):
        bk._set(ws, r, 1 + j, h, S.f_label(bold=True), None, S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
    r += 1
    yrs = [("C", "N"), ("O", "Z"), ("AA", "AL")]
    def row(label, sheet, srow, key=None, bold=False, fmt=S.FMT_CUR):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for j, (c0, c1) in enumerate(yrs):
            bk.fml(ws, r, 2 + j, f"=SUM('{sheet}'!{c0}{srow}:{c1}{srow})", fmt, bold=bold)
        if key: bk.rows["3-Year Summary"][key] = r
        r += 1
    row("Net patient revenue", "P&L", P["npr"], "npr", bold=True)
    row("EBITDA", "P&L", P["ebitda"], "ebitda", bold=True)
    row("Net income", "Cash Flow & Runway", CF["ni"], "ni")
    bk.lbl(ws, r, "Debt service (P&I, all notes)")
    for j, (c0, c1) in enumerate(yrs):
        cf = "Cash Flow & Runway"
        f = (f"=SUM('{cf}'!{c0}{CF['bref_pmt']}:{c1}{CF['bref_pmt']})+SUM('{cf}'!{c0}{CF['refi_pmt']}:{c1}{CF['refi_pmt']})"
             f"+SUM('{cf}'!{c0}{CF['sba_pmt']}:{c1}{CF['sba_pmt']})"
             f"+SUM('{cf}'!{c0}{CF['s_prin']}:{c1}{CF['s_prin']})+SUM('{cf}'!{c0}{CF['s_intpaid']}:{c1}{CF['s_intpaid']})")
        bk.fml(ws, r, 2 + j, f, S.FMT_CUR)
    ds_row = r; r += 1
    bk.lbl(ws, r, "DSCR (per year)", bold=True)
    for j in range(3):
        col = gcl(2 + j)
        bk.fml(ws, r, 2 + j, f"=IF({col}{ds_row}=0,\"-\",{col}6/{col}{ds_row})", S.FMT_MULT, bold=True)
    r += 2
    bk.lbl(ws, r, "GLOBAL 3-YEAR DSCR (operative SBA test)", bold=True)
    bk.fml(ws, r, 2, f"=SUM(B6:D6)/SUM(B{ds_row}:D{ds_row})", S.FMT_MULT, bold=True, fill=S.fill(S.LIGHTBLUE))
    r += 1
    bk.lbl(ws, r, "Minimum cash across 36 months", bold=True)
    bk.fml(ws, r, 2, f"='Cash Flow & Runway'!B{CF['mincash']}", S.FMT_CUR, bold=True)
    r += 2
    bk.lbl(ws, r, "Note: Y1 DSCR is structurally thin for a hospice startup (license maturation + Medicare "
                  "payment cycle). Per the audit: lead with Global DSCR; do not force Y1 to 1.25x.", italic=True)
    return ws


# =====================================================================
def avb(bk: OB):
    ws = bk.sheet("Actuals vs Budget", tab="538135")
    bk.title(ws, "Enter ACTUALS monthly (yellow); variances compute. First 12 months.")
    bk.widths(ws, n=12, w=12)
    P = bk.rows["P&L"]; CF = bk.rows["Cash Flow & Runway"]
    r = 4
    for i in range(12):
        bk._set(ws, r, MC0 + i, MONTH_LABELS[i], S.f_label(bold=True), None, S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
    r += 1
    for label, sheet, srow in [("Net patient revenue", "P&L", P["npr"]),
                               ("Total clinical labor", "P&L", P["clin"]),
                               ("EBITDA", "P&L", P["ebitda"]),
                               ("Cash collections", "Cash Flow & Runway", CF["coll"]),
                               ("Ending cash", "Cash Flow & Runway", CF["cash"])]:
        bk.lbl(ws, r, f"{label} - BUDGET", bold=True)
        for i in range(12):
            bk.fml(ws, r, MC0 + i, f"='{sheet}'!{mlet(i)}{srow}", S.FMT_CUR, link=True)
        r += 1
        bk.lbl(ws, r, f"{label} - ACTUAL [enter]")
        for i in range(12):
            bk.inp(ws, r, MC0 + i, None, S.FMT_CUR)
        r += 1
        bk.lbl(ws, r, f"{label} - variance", italic=True)
        for i in range(12):
            bk.fml(ws, r, MC0 + i, f"=IF({mlet(i)}{r-1}=\"\",\"\",{mlet(i)}{r-1}-{mlet(i)}{r-2})", S.FMT_CUR)
        r += 2
    return ws


# =====================================================================
def checks(bk: OB):
    ws = bk.sheet("Checks", tab="FF0000")
    bk.title(ws, "CHECKS - every integrity test in one place. MASTER must read 0 / OK.")
    bk.widths(ws)
    B = bk.rows["Balance Sheet"]; CF = bk.rows["Cash Flow & Runway"]; A = bk.addr
    r = 4
    hr = bk.mheader(ws, r); r = hr
    def crow(label, f_tmpl, key):
        nonlocal r
        bk.lbl(ws, r, label)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f_tmpl(i), S.FMT_NUM2)
        bk.rows["Checks"][key] = r; r += 1
    crow("BS out of balance (abs > $0.01 -> 1)",
         lambda i: f"=IF(ABS('Balance Sheet'!{mlet(i)}{B['check']})>0.01,1,0)", "bs")
    crow("Cash below zero (-> 1)",
         lambda i: f"=IF('Cash Flow & Runway'!{mlet(i)}{CF['cash']}<-0.01,1,0)", "neg")
    crow("Revolver over commitment (-> 1)",
         lambda i: f"=IF('Cash Flow & Runway'!{mlet(i)}{CF['rev_bal']}>{A['rev_limit']}+0.01,1,0)", "revlim")
    crow("Revolver drawn while cash above floor (+$1 tol) (-> 1)",
         lambda i: (f"=IF(AND('Cash Flow & Runway'!{mlet(i)}{CF['rev_bal']}>0.01,"
                    f"'Cash Flow & Runway'!{mlet(i)}{CF['cash']}>{A['floor']}+1),1,0)"), "revcash")
    crow("Seller balance negative (-> 1)",
         lambda i: f"=IF('Cash Flow & Runway'!{mlet(i)}{CF['s_end']}<-0.01,1,0)", "sneg")
    r += 1
    bk.lbl(ws, r, "Collections conservation: cum collected + AR - cum NPR (should be 0)", bold=True)
    bk.fml(ws, r, 2, (f"=SUM('Cash Flow & Runway'!C{CF['coll']}:AL{CF['coll']})"
                      f"+'Cash Flow & Runway'!AL{CF['ar']}"
                      f"-SUM('Cash Flow & Runway'!C{CF['earned']}:AL{CF['earned']})"), S.FMT_CUR, bold=True)
    cons = r; r += 1
    bk.lbl(ws, r, "MASTER CHECK (0 = OK)", bold=True)
    rows = bk.rows["Checks"]
    parts = "+".join(f"SUM(C{rows[k]}:AL{rows[k]})" for k in ("bs", "neg", "revlim", "revcash", "sneg"))
    bk.fml(ws, r, 2, f"={parts}+IF(ABS(B{cons})>1,1,0)", S.FMT_NUM, bold=True, fill=S.fill(S.LIGHTBLUE))
    bk.rows["Checks"]["master"] = r
    return ws


def dashboard(bk: OB):
    ws = bk.sheet("Dashboard", tab="00B050")
    bk.title(ws, "KPI DASHBOARD - the rows a lender / investment committee reads first. All live links.")
    bk.widths(ws)
    A = bk.addr
    C = bk.rows["Census Waterfall"]; V = bk.rows["Revenue Model"]; P = bk.rows["P&L"]
    ST = bk.rows["Staffing"]; CF = bk.rows["Cash Flow & Runway"]; OB_ = bk.rows["Operating Budget"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    def row(label, f_tmpl, fmt=S.FMT_CUR, bold=False, key=None):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            bk.fml(ws, r, MC0 + i, f_tmpl(i), fmt, bold=bold)
        if key: bk.rows["Dashboard"][key] = r
        r += 1
    row("ADC", lambda i: f"='Census Waterfall'!{mlet(i)}{C['adc']}", S.FMT_NUM1, key="adc")
    row("Net patient revenue", lambda i: f"='P&L'!{mlet(i)}{P['npr']}", key="npr")
    row("NPR per patient-day", lambda i: f"=IF('Census Waterfall'!{mlet(i)}{C['pd']}=0,0,'P&L'!{mlet(i)}{P['npr']}/'Census Waterfall'!{mlet(i)}{C['pd']})", S.FMT_RATE)
    row("Direct cost per patient-day (COS/PD)", lambda i: f"=IF('Census Waterfall'!{mlet(i)}{C['pd']}=0,0,'P&L'!{mlet(i)}{P['cos']}/'Census Waterfall'!{mlet(i)}{C['pd']})", S.FMT_RATE)
    row("Contribution per patient-day", lambda i: f"=IF('Census Waterfall'!{mlet(i)}{C['pd']}=0,0,'P&L'!{mlet(i)}{P['gp']}/'Census Waterfall'!{mlet(i)}{C['pd']})", S.FMT_RATE)
    row("Total labor % of NPR", lambda i: f"=IF('P&L'!{mlet(i)}{P['npr']}=0,0,('Staffing'!{mlet(i)}{ST['clin_tot']}+'Staffing'!{mlet(i)}{ST['ind_tot']})/'P&L'!{mlet(i)}{P['npr']})", S.FMT_PCT)
    row("EBITDA", lambda i: f"='P&L'!{mlet(i)}{P['ebitda']}", bold=True, key="ebitda")
    row("EBITDA margin % (pre-debt-service)", lambda i: f"='P&L'!{mlet(i)}{P['em']}", S.FMT_PCT, bold=True, key="em")
    row("Debt service (P&I, all facilities)",
        lambda i: (f"='Cash Flow & Runway'!{mlet(i)}{CF['bref_pmt']}+'Cash Flow & Runway'!{mlet(i)}{CF['refi_pmt']}"
                   f"+'Cash Flow & Runway'!{mlet(i)}{CF['sba_pmt']}+'Cash Flow & Runway'!{mlet(i)}{CF['s_prin']}"
                   f"+'Cash Flow & Runway'!{mlet(i)}{CF['s_intpaid']}"), key="ds")
    row("DSCR (monthly)", lambda i: f"='Cash Flow & Runway'!{mlet(i)}{CF['dscr']}", S.FMT_MULT, key="dscr")
    row("DSCR TTM (from M12)",
        lambda i: ("=\"-\"" if i < 11 else
                   f"=SUM('P&L'!{mlet(i-11)}{P['ebitda']}:{mlet(i)}{P['ebitda']})/SUM({mlet(i-11)}{bk.rows['Dashboard']['ds']}:{mlet(i)}{bk.rows['Dashboard']['ds']})"),
        S.FMT_MULT)
    row("Covenant headroom vs 1.25x (TTM basis after M12)",
        lambda i: ("=\"-\"" if i < 11 else
                   f"=SUM('P&L'!{mlet(i-11)}{P['ebitda']}:{mlet(i)}{P['ebitda']})/SUM({mlet(i-11)}{bk.rows['Dashboard']['ds']}:{mlet(i)}{bk.rows['Dashboard']['ds']})-1.25"),
        S.FMT_MULT)
    row("DSO (days: AR / NPR x days-in-month)",
        lambda i: f"=IF('P&L'!{mlet(i)}{P['npr']}=0,0,'Cash Flow & Runway'!{mlet(i)}{CF['ar']}/'P&L'!{mlet(i)}{P['npr']}*{A['days_mo']})", S.FMT_NUM1)
    row("Ending cash", lambda i: f"='Cash Flow & Runway'!{mlet(i)}{CF['cash']}", key="cash")
    row("Revolver balance", lambda i: f"='Cash Flow & Runway'!{mlet(i)}{CF['rev_bal']}")
    row("TOTAL LIQUIDITY (cash + undrawn revolver)",
        lambda i: f"='Cash Flow & Runway'!{mlet(i)}{CF['cash']}+{A['rev_limit']}-'Cash Flow & Runway'!{mlet(i)}{CF['rev_bal']}", bold=True)
    row("Days cash on hand (vs monthly opex)",
        lambda i: (f"=IF(('Staffing'!{mlet(i)}{ST['payroll_cash']}+'Operating Budget'!{mlet(i)}{OB_['dpc']}"
                   f"+'Operating Budget'!{mlet(i)}{OB_['fac']}+'Operating Budget'!{mlet(i)}{OB_['ga']})=0,0,"
                   f"'Cash Flow & Runway'!{mlet(i)}{CF['cash']}/('Staffing'!{mlet(i)}{ST['payroll_cash']}"
                   f"+'Operating Budget'!{mlet(i)}{OB_['dpc']}+'Operating Budget'!{mlet(i)}{OB_['fac']}"
                   f"+'Operating Budget'!{mlet(i)}{OB_['ga']})*{A['days_mo']})"), S.FMT_NUM1)
    r += 1
    bk.lbl(ws, r, "Breakeven ADC (fixed costs / contribution per PD / days)", bold=True)
    bk.fml(ws, r, 2,
           (f"=('Staffing'!N{ST['ind_tot']}+'Operating Budget'!N{OB_['fac']}+'Operating Budget'!N{OB_['ga']})"
            f"/(IF('Census Waterfall'!N{C['pd']}=0,1,'P&L'!N{P['gp']}/'Census Waterfall'!N{C['pd']}))/{A['days_mo']}"),
           S.FMT_NUM1, bold=True)
    r += 1
    bk.lbl(ws, r, "MASTER CHECK (from Checks tab; 0 = OK)", bold=True)
    bk.fml(ws, r, 2, f"='Checks'!B{bk.rows['Checks']['master']}", S.FMT_NUM, bold=True, fill=S.fill(S.LIGHTBLUE))
    return ws


# =====================================================================
def build():
    bk = OB()
    control_tower(bk)
    census(bk)
    revenue(bk)
    staffing(bk)
    opbudget(bk)
    pl(bk)
    debt_and_cash(bk)
    balance_sheet(bk)
    checks(bk)
    dashboard(bk)
    summary(bk)
    avb(bk)
    out = "financial_models/output/Azalea_Hospice_Proforma_Rev3.00_DYNAMIC.xlsx"
    bk.wb.save(out)
    import json
    with open("financial_models/output/proforma_v3_rowmap.json", "w") as f:
        json.dump({"rows": bk.rows, "addr": {k: v for k, v in bk.addr.items() if isinstance(v, str)}}, f, indent=1)
    print("wrote", out)
    return bk


if __name__ == "__main__":
    build()
