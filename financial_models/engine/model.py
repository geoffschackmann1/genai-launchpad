"""Azalea Hospice - shared financial engine.

Single source of truth for assumptions, period structure, the Python
reference calculation (used for QA), and the Inputs-tab writer.

Framing: NEW Medicare hospice startup (NOT an acquisition). Paloma's
clinicians join Azalea and Paloma's patients migrate over M1-M2. All
assumptions trued up to Paloma's 2025 actuals (Tyler_PL_2025.xlsx) and
Apr-May 2025 payrolls.
"""

from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from . import styles as S


# =========================================================================
# PERIOD STRUCTURE  — Y1 monthly (M1-M12), Y2-Y3 quarterly (8 quarters)
# Columns: period i -> spreadsheet column C+i (col 3+i). Kept identical on
# every period-based tab so same-column cross-sheet links line up.
# =========================================================================
AVG_DAYS_MONTH = 30.4
DAYS_QUARTER = AVG_DAYS_MONTH * 3   # 91.2

def _build_periods():
    periods = []
    for m in range(1, 13):
        periods.append(dict(label=f"M{m}", kind="M", months=1, days=AVG_DAYS_MONTH,
                            month_index=m, year=1, escal=1.00))
    for q in range(1, 5):
        periods.append(dict(label=f"Y2 Q{q}", kind="Q", months=3, days=DAYS_QUARTER,
                            month_index=12 + q * 3, year=2, escal=1.03))
    for q in range(1, 5):
        periods.append(dict(label=f"Y3 Q{q}", kind="Q", months=3, days=DAYS_QUARTER,
                            month_index=24 + q * 3, year=3, escal=1.03 ** 2))
    return periods

PERIODS = _build_periods()
NP = len(PERIODS)                       # 20
FIRST_COL = 3                           # column C
def pcol(i):  return FIRST_COL + i      # 1-based column index for period i
def plet(i):  return get_column_letter(pcol(i))
LAST_LET = plet(NP - 1)                 # 'V'

Y1_IDX = [i for i, p in enumerate(PERIODS) if p["year"] == 1]
Y2_IDX = [i for i, p in enumerate(PERIODS) if p["year"] == 2]
Y3_IDX = [i for i, p in enumerate(PERIODS) if p["year"] == 3]


# =========================================================================
# ASSUMPTIONS  (scalars -> named ranges on the Inputs tab)
# Each: key -> (label, value, number_format, source/flag note)
# =========================================================================
SRC_PL = "Source: Paloma P&L 2025 (Tyler_PL_2025.xlsx)."
SRC_PR = "Source: Paloma payrolls Apr-May 2025."
FLAG = "FLAG / CONFIRM: go-forward assumption, not in Paloma actuals."

# -- Block A: Geography & RHC rates --
RATES = [
    ("rhc_tier1",   "RHC per-diem, days 1-60 ($/day)", 230.83, S.FMT_RATE, "CMS FY2026 RHC routine home care, Tier 1 (national, pre-wage-index)."),
    ("rhc_tier2",   "RHC per-diem, days 61+ ($/day)",   182.36, S.FMT_RATE, "CMS FY2026 RHC routine home care, Tier 2 (national, pre-wage-index)."),
    ("wage_index",  "Tyler CBSA 46340 wage index",       0.88,  S.FMT_NUM2, "Tyler, TX CBSA wage index; confirm against CMS WI addendum."),
    ("labor_share", "RHC labor share",                    0.68,  S.FMT_PCT,  "Hospice CHC labor portion subject to wage-index adjustment."),
    ("tier1_pct",   "% patient-days in Tier 1 (days 1-60)", 0.40, S.FMT_PCT, "Lower than a fresh-admit book: migrated patients arrive mid-episode (many past day 60). Tune to validate blended rate. " + SRC_PL),
    ("seq",         "Sequestration reduction",            0.02,  S.FMT_PCT,  "2.0% input; Paloma actual ran 2.1-2.7% of gross. " + SRC_PL),
    ("writeoff_pct","Write-offs (% of gross)",            0.0015, S.FMT_PCT2,"Net-due-zero / claim write-offs. " + SRC_PL),
    ("blended_actual","Blended NET rate per PD - ACTUAL ($/day)", 181.00, S.FMT_RATE, "Validated against Paloma actuals (~$118K net / ~22 ADC / 30.4 days). " + SRC_PL),
    ("medicare_cap", "Medicare aggregate cap / beneficiary ($/yr)", 35361.44, S.FMT_CUR, "CMS FY2026 hospice aggregate cap per beneficiary."),
    ("rhc_escalation", "Medicare RHC rate escalation (annual)", 0.025, S.FMT_PCT, "CMS historical 2-3%/yr. Applied to gross rates in Y2 (x1.025) and Y3 (x1.025^2)."),
]

# -- Block C: Census migration ramp (scenario scalars) --
CENSUS_SCALARS = [
    ("capture_rate",   "Patient-migration capture rate", 1.00, S.FMT_PCT, "Base 100% -> steady-state ADC 22. Downside 80% -> ADC ~18. Each Paloma patient must revoke and re-elect Azalea."),
    ("census_scenario","Census scenario (1=Base flat, 2=Upside growth)", 1, S.FMT_INT, "1 = hold migrated panel flat at steady state (SBA conservative). 2 = grow from new referrals toward ADC ~35-60 (investor upside)."),
]
# Base path (100% capture, flat after migration completes) and Upside path
# (referral growth). Effective ADC = chosen path x capture_rate.
# BASE = patient-migration ramp (M1–M2) then BACK-LOADED growth path. Y1 holds
# near 22-24 (stabilize migrated Paloma panel + build referral pipeline), Y2
# accelerates 24→50 as BD Director's pipeline matures, Y3 climbs to 56. The higher
# census funds a full, realistic direct-care team while holding ~21/25/27% margins.
ADC_BASE = [12.0, 19.8, 22.0, 22.2, 22.4, 22.7, 22.9, 23.1, 23.3, 23.6, 23.8, 24.0,
            30.5, 37.0, 43.5, 50.0,
            51.5, 53.0, 54.5, 56.0]
# UPSIDE = aggressive growth to ADC 60 by Y3Q4 (existing investor-upside path).
ADC_UPSIDE = [12.0, 19.8, 22.0, 23.0, 24.0, 25.0, 26.0, 27.0, 28.0, 29.0, 30.0, 31.0,
              33.0, 35.0, 37.0, 39.0,
              44.0, 50.0, 55.0, 60.0]

# -- Block D: Staggered FT roster (2B).  start_month is the model month. --
# role_group: 'direct' (clinical COGS) or 'indirect' (SG&A overhead)
ROSTER = [
    # name, role, annual_salary, start_month, role_group, open_req
    ("Brad Woodard",     "Administrator / Director",      120000, 1, "indirect", False),
    ("Angie Frost",      "RN Case Manager",                77504, 1, "direct",   False),
    ("Tanisha Robison",  "CNA / Hospice Aide",             46511, 1, "direct",   False),
    ("Dana Davenport",   "Director of Clinical Services",  110000, 2, "indirect", False),
    ("Silas Shelton",    "Administrator",                  130000, 3, "indirect", False),
    ("Jodi McCollum",    "RN Case Manager (2nd seat)",      75008, 4, "direct",   True),
    # Growth driver — needed Y1 H2 to build referral pipeline that fuels Y2-Y3 census
    ("Director of Business Development", "Sales / referrals (growth driver)", 95000, 6, "indirect", False),
    # Capacity-driven FT hires triggered by census growth (thresholds editable on Inputs)
    ("Capacity hire - 2nd CNA",  "CNA / Hospice Aide (capacity-driven, triggers ADC>=25)", 46511, 13, "direct", False),
    # Core IDG clinical team — convert PRN social work / chaplain / LVN to FT as census scales (direct labor)
    ("FT Social Worker (MSW)",   "Core IDG MSW (converts PRN SW to salaried)",             58000, 14, "direct", False),
    ("FT LVN - visits / on-call","Core IDG LVN (converts PRN LVN to salaried)",            52000, 16, "direct", False),
    ("FT Chaplain",              "Core IDG spiritual care (converts PRN chaplain to FT)",  55000, 18, "direct", False),
    # Capacity-driven RN/CNA hires (richer ratios: 1:12 RN, 1:10 CNA at scale)
    ("Capacity hire - 3rd RN",   "RN Case Manager (capacity-driven, ADC>=30)",             76256, 13, "direct", False),
    ("Capacity hire - 3rd CNA",  "CNA / Hospice Aide (capacity-driven, ADC>=32)",          46511, 14, "direct", False),
    ("Capacity hire - 4th RN",   "RN Case Manager (capacity-driven, ADC>=38)",             76256, 20, "direct", False),
    ("Capacity hire - 4th CNA",  "CNA / Hospice Aide (capacity-driven, ADC>=40)",          46511, 20, "direct", False),
    ("Capacity hire - 5th RN",   "RN Case Manager (capacity-driven, ADC>=48)",             76256, 27, "direct", False),
    ("Capacity hire - 5th CNA",  "CNA / Hospice Aide (capacity-driven, ADC>=50)",          46511, 27, "direct", False),
    ("Capacity hire - 6th CNA",  "CNA / Hospice Aide (capacity-driven, ADC>=54)",          46511, 31, "direct", False),
    ("Capacity hire - 6th RN",   "RN Case Manager (capacity-driven, ADC>=55)",             76256, 34, "direct", False),
    # Back-office build-out as the agency scales — legitimate CoP / scale roles (indirect).
    # NOTE: Billing/AR Specialist dropped — billing is outsourced at 1.5% of revenue.
    ("Quality / Compliance Manager", "QAPI program | CoP compliance | surveys",            80000, 16, "indirect", False),
    ("Intake / Admissions Coordinator", "Referral intake | IDG admit coordination",        55000, 20, "indirect", False),
    ("Volunteer Coordinator",        "Medicare CoP requirement | volunteer hours",         50000, 26, "indirect", False),
]

# -- PRN / per-visit roster (census-driven, all months) --
# key, label, visits_per_pt_per_month, rate_per_visit, role_group
PRN_VISIT = [
    ("prn_lvn",  "PRN On-Call LVN (Kendricks-Tuck, Foster)", 0.60, 100.0, "direct"),
    ("prn_chap", "PRN Chaplain (Ector, Johnson)",            0.75,  75.0, "direct"),
    ("prn_sw",   "PRN Social Worker (Sanders)",              1.00,  75.0, "direct"),
    ("prn_np",   "PRN Nurse Practitioner (Rogers)",          0.15, 200.0, "direct"),
    ("prn_cna",  "PRN CNA / per-visit (Henson, Mills)",      0.50,  21.0, "direct"),
]
# PRN flat office support (Rhonda Smith) - W-2, no health
RHONDA_ANNUAL = 34000

# -- Block E: ratios (Y2-Y3 capacity reference) --
RATIOS = [
    ("rn_caseload",  "RN case-manager caseload (1 : N ADC)", 12, S.FMT_INT, "Richer 1:12 caseload at scale (was 1:15). Capacity planning benchmark."),
    ("aide_caseload","CNA / aide caseload (1 : N ADC)",      10, S.FMT_INT, "Richer 1:10 caseload at scale (was 1:12). Capacity planning benchmark."),
    ("visits_rn_day","Target visits / RN / day",            5.5, S.FMT_NUM1,"Ops KPI."),
    ("visits_aide_day","Target visits / aide / day",          6, S.FMT_NUM1,"Ops KPI."),
    ("merit",        "Annual merit / COL increase (Y2, Y3)", 0.03, S.FMT_PCT,"Applied to salaried wages from Y2."),
]

# -- Block F: benefits & ER taxes --
BENEFITS = [
    ("benefits_load", "Benefits load on W-2 wages", 0.1827, S.FMT_PCT, "Blended employer burden (FICA 7.65% + FUTA/SUTA + WC + other). " + SRC_PR),
    ("health_pm",     "Health insurance ($/FT employee/mo)", 550, S.FMT_CUR, "Per full-time W-2 employee. " + FLAG),
]

# -- Block G: per-PD COGS (trued up to actuals) --
COGS_PD = [
    ("supplies_pd", "Medical supplies ($/patient-day)", 3.99, S.FMT_CUR2, SRC_PL),
    ("dme_pd",      "DME ($/patient-day)",              1.77, S.FMT_CUR2, SRC_PL),
    ("pharmacy_pd", "Pharmacy ($/patient-day)",         3.23, S.FMT_CUR2, SRC_PL),
]

# -- Block G2: fixed monthly G&A (trued up to actuals) --
GA_FIXED = [
    ("ga_utilities", "Utilities",                293,  S.FMT_CUR, SRC_PL),
    ("ga_internet",  "Internet",                 779,  S.FMT_CUR, SRC_PL),
    ("ga_telephone", "Telephone / fax",          403,  S.FMT_CUR, SRC_PL),
    ("ga_bvtm",      "After-hours messaging (BVTM)", 48, S.FMT_CUR, SRC_PL),
    ("ga_emr",       "EMR system",               1281, S.FMT_CUR, SRC_PL),
    ("ga_pcr",       "PCR (CPM)",                1001, S.FMT_CUR, SRC_PL),
    ("ga_cc",        "Credit-card fees",         1004, S.FMT_CUR, SRC_PL),
    ("ga_liability", "Liability / D&O insurance", 250, S.FMT_CUR, SRC_PL),
    ("ga_training",  "Training / CEUs",           100, S.FMT_CUR, SRC_PL),
    ("ga_other",     "Other (client)",           1050, S.FMT_CUR, SRC_PL),
    ("ga_bankfees",  "Bank / payroll fees",      1750, S.FMT_CUR, SRC_PL),
    ("ga_rent",      "Facility rent",            3000, S.FMT_CUR, "Paloma actual was $0 (embedded/rent-free). Azalea needs its own lease. " + FLAG),
    ("ga_marketing", "Marketing / patient acquisition", 3000, S.FMT_CUR, "Paloma actual was $0. " + FLAG),
]
GA_VAR = [
    ("billing_fee_pct", "Outsourced billing fee (% of gross)", 0.015, S.FMT_PCT2, "Outsourced billing service at 1.5% of gross revenue (replaces Paloma's flat BCBP fee)."),
    ("qr_fee_pct", "QR payment fee (% of gross)", 0.0075, S.FMT_PCT2, SRC_PL),
]
MED_DIRECTOR_PM = ("med_director", "Medical Director 1 (1099, $/mo)", 4000, S.FMT_CUR,
                   "1099 contract, flat, NO benefits. Paloma actual ran $1,000/mo; spec models $4,000/mo go-forward. " + FLAG)
MED_DIRECTOR2_PM = ("med_director2", "Medical Director 2 (1099, $/mo, conditional)", 5000, S.FMT_CUR,
                    "Second MD relationship needed at higher census volumes (e.g., palliative-care specialist or rotating consultant). 1099, no benefits.")
MED_DIRECTOR2_START_M = ("med_director2_start", "Medical Director 2 start month", 22, S.FMT_INT,
                          "Defaults to M22 (start of Y2Q4) when census approaches ADC ~38. Editable.")

# -- Block H: capital structure (NO acquisition) --
CAPITAL = [
    ("sba_principal", "SBA 7(a) loan principal ($)", 500000, S.FMT_CUR, "Sized to fund ramp-period operating need + startup costs. " + FLAG),
    ("sba_rate",      "SBA interest rate (APR)",      0.115, S.FMT_PCT, FLAG),
    ("sba_term_mo",   "SBA term (months)",            120,   S.FMT_INT, "10-year amortization. " + FLAG),
    ("equity",        "Owner equity injection ($)",   250000, S.FMT_CUR, FLAG),
    ("loc_limit",     "Working-capital line limit ($)", 100000, S.FMT_CUR, FLAG),
    ("loc_rate",      "Working-capital line rate (APR)", 0.105, S.FMT_PCT, FLAG),
    ("min_cash",      "Minimum cash floor ($)",        25000, S.FMT_CUR, "Operating cash buffer; LOC draws to hold this floor. " + FLAG),
    ("capex",         "Startup capex - equipment ($)", 15000, S.FMT_CUR, "Computers, office furniture; depreciated straight-line. " + FLAG),
    ("deprec_yrs",    "Depreciation / amortization life (yrs)", 5, S.FMT_INT, FLAG),
    ("tx_tax",        "TX franchise/margin tax (eff.)", 0.00375, S.FMT_PCT2, "Applied to revenue when pre-tax income positive."),
    ("license_cost",  "Hickory Medicare license - acquisition cost ($)", 300000, S.FMT_CUR, "CHOW: acquire Hickory Hospice's existing Medicare-certified provider number (San Antonio + Tyler alternative-delivery site). Eliminates 855A enrollment gap - Azalea bills from day 1."),
    ("license_rate",  "License note interest rate (APR)", 0.06, S.FMT_PCT, "Seller financing on the license acquisition."),
    ("license_term",  "License note term (months)", 36, S.FMT_INT, "Monthly P+I amortization."),
    ("license_amort_yrs", "License intangible amortization life (yrs)", 15, S.FMT_INT, "GAAP intangible amortization - non-cash, below EBITDA. No effect on DSCR."),
]

# -- Block H2: startup one-time uses (Sources & Uses) --
STARTUP = [
    ("su_enroll",   "CHOW filing / 855A change-of-ownership", 3000, S.FMT_CUR, "Preserves Hickory's existing Medicare number; no fresh enrollment delay."),
    ("su_license",  "TX state licensure",                 5000,  S.FMT_CUR, FLAG),
    ("su_emr",      "EMR implementation / setup",        10000,  S.FMT_CUR, FLAG),
    ("su_supplies", "Initial medical-supply stock",      10000,  S.FMT_CUR, FLAG),
    ("su_legal",    "Legal / entity formation",          15000,  S.FMT_CUR, FLAG),
    ("su_conting",  "Contingency",                       20000,  S.FMT_CUR, FLAG),
]

# -- Block I: working capital --
WORKING_CAP = [
    ("ar_days", "Accounts-receivable days (Medicare RAP-to-final lag)", 45, S.FMT_INT,
     "Cash receipts lag accrued revenue ~30-45 days. NOTE: as a new provider, 855A enrollment can delay first cash 3-6 months; modeled here on accrual per owner direction (toggle off)."),
    ("ap_days", "Accounts-payable days", 30, S.FMT_INT, FLAG),
]

# -- Paloma actuals (for the Actuals-vs-Model reconciliation tab) --
ACTUALS = {
    "months": ["Mar 2025", "Apr 2025", "May 2025", "Jun 2025"],
    "gross":   [138295.80, 116422.51, 127893.15, 119990.52],
    "writeoff":[179.62, 177.59, 172.62, 172.62],
    "seq":     [3073.97, 2549.02, 2739.17, 3248.08],
    "medadj":  [86.96, 0, 0, 0],
    "pharmacy":[2001.11, 2149.05, 2293.00, 2293.00],
    "supplies":[1846.79, 1671.97, 3650.00, 3650.00],
    "dme":     [1200.91, 1200.00, 1200.00, 1200.00],
    "days":    [31, 30, 31, 30],
    # semi-monthly payroll (1st half / 2nd half) from P&L
    "pay1":    [20727.47, 22921.54, 22117.27, 31000.00],
    "pay2":    [41512.90, 22765.10, 21499.29, 27000.00],
}


# =========================================================================
# PYTHON REFERENCE CALCULATION  (the "truth" used to QA the workbook)
# =========================================================================
def _val(rows, key):
    for r in rows:
        if r[0] == key:
            return r[2]
    raise KeyError(key)

def assumptions_dict():
    d = {}
    for grp in (RATES, CENSUS_SCALARS, RATIOS, BENEFITS, COGS_PD, GA_FIXED,
                GA_VAR, CAPITAL, STARTUP, WORKING_CAP):
        for r in grp:
            d[r[0]] = r[2]
    d["med_director"] = MED_DIRECTOR_PM[2]
    d["med_director2"] = MED_DIRECTOR2_PM[2]
    d["med_director2_start"] = MED_DIRECTOR2_START_M[2]
    return d


def compute(capture_rate=None, scenario=None):
    """Return per-period arrays mirroring the workbook formulas."""
    a = assumptions_dict()
    if capture_rate is not None:
        a["capture_rate"] = capture_rate
    if scenario is not None:
        a["census_scenario"] = scenario

    wiadj = a["labor_share"] * a["wage_index"] + (1 - a["labor_share"])
    rate_t1 = a["rhc_tier1"] * wiadj
    rate_t2 = a["rhc_tier2"] * wiadj
    blended_gross = a["tier1_pct"] * rate_t1 + (1 - a["tier1_pct"]) * rate_t2
    net_rate = blended_gross * (1 - a["seq"] - a["writeoff_pct"])

    path = ADC_BASE if a["census_scenario"] == 1 else ADC_UPSIDE
    adc = [path[i] * a["capture_rate"] for i in range(NP)]

    R = {k: [0.0] * NP for k in (
        "adc", "pd", "gross", "net", "ft_direct", "ft_indirect", "prn", "rhonda",
        "med_dir", "burden_d", "burden_i", "health_d", "health_i", "cogs_patient",
        "cogs", "gp", "sga_labor", "ga_fixed", "qr", "billing", "sga", "ebitda", "da",
        "int_sba", "int_loc", "pretax", "tax", "ni", "ds", "ar", "ap",
        "cfo", "principal", "loc_draw", "loc_repay", "end_cash", "loc_bal",
        "ppe_net", "intang_net", "license_net", "ft_head_d", "ft_head_i",
        "sba_bal", "lic_bal", "lic_int", "lic_prin", "re")}

    # ---- SBA amortization at PERIOD grain (mirrors workbook formulas) ----
    P = a["sba_principal"]; rm = a["sba_rate"] / 12; n = int(a["sba_term_mo"])
    pmt = P * rm / (1 - (1 + rm) ** -n) if rm else P / n  # monthly payment
    sba_bal = P

    # ---- Hickory license note (seller financing) ----
    lic_P = a["license_cost"]; lic_rm = a["license_rate"] / 12
    lic_n = int(a["license_term"])
    lic_pmt = lic_P * lic_rm / (1 - (1 + lic_rm) ** -lic_n) if lic_rm else lic_P / lic_n
    lic_bal = lic_P

    startup_total = sum(_val(STARTUP, k[0]) for k in STARTUP)
    capex = a["capex"]
    # Three asset pools, three lives
    da_capex_pm = capex / a["deprec_yrs"] / 12
    da_startup_pm = startup_total / a["deprec_yrs"] / 12
    da_license_pm = lic_P / a["license_amort_yrs"] / 12
    da_pm = da_capex_pm + da_startup_pm + da_license_pm

    # License is financed by the seller note → does not consume opening cash
    beg_cash = a["equity"] + a["sba_principal"] - startup_total - capex
    loc_bal = 0.0; re_cum = 0.0
    accum_capex = 0.0; accum_startup = 0.0; accum_license = 0.0
    prev_ar = 0.0; prev_ap = 0.0

    for i, p in enumerate(PERIODS):
        m = p["months"]; days = p["days"]; esc = p["escal"]
        rhc_esc = (1 + a["rhc_escalation"]) ** (p["year"] - 1)
        R["adc"][i] = adc[i]
        R["pd"][i] = adc[i] * days
        R["gross"][i] = R["pd"][i] * blended_gross * rhc_esc
        R["net"][i] = R["pd"][i] * net_rate * rhc_esc

        # FT salaried
        ftd = fti = 0.0; hd = hi = 0
        for (_, _, sal, start, grp, _o) in ROSTER:
            active = p["month_index"] >= start
            if active:
                cost = sal / 12 * m * esc
                if grp == "direct": ftd += cost; hd += 1
                else: fti += cost; hi += 1
        R["ft_direct"][i] = ftd; R["ft_indirect"][i] = fti
        R["ft_head_d"][i] = hd; R["ft_head_i"][i] = hi

        # PRN visit + Rhonda
        prn = sum(adc[i] * vpm * rate * m for (_, _, vpm, rate, _g) in PRN_VISIT)
        R["prn"][i] = prn
        R["rhonda"][i] = RHONDA_ANNUAL / 12 * m * esc
        R["med_dir"][i] = a["med_director"] * m
        if p["month_index"] >= a["med_director2_start"]:
            R["med_dir"][i] += a["med_director2"] * m

        # Benefits (W-2). direct W2 = ft_direct + prn ; indirect W2 = ft_indirect + rhonda
        w2_d = ftd + prn; w2_i = fti + R["rhonda"][i]
        R["burden_d"][i] = w2_d * a["benefits_load"]
        R["burden_i"][i] = w2_i * a["benefits_load"]
        R["health_d"][i] = a["health_pm"] * hd * m
        R["health_i"][i] = a["health_pm"] * hi * m

        # COGS
        R["cogs_patient"][i] = R["pd"][i] * (a["supplies_pd"] + a["dme_pd"] + a["pharmacy_pd"])
        cogs_labor = ftd + prn + R["med_dir"][i] + R["burden_d"][i] + R["health_d"][i]
        R["cogs"][i] = cogs_labor + R["cogs_patient"][i]
        R["gp"][i] = R["net"][i] - R["cogs"][i]

        # SG&A
        R["sga_labor"][i] = fti + R["rhonda"][i] + R["burden_i"][i] + R["health_i"][i]
        R["ga_fixed"][i] = sum(_val(GA_FIXED, k[0]) for k in GA_FIXED) * m
        R["qr"][i] = R["gross"][i] * a["qr_fee_pct"]
        R["billing"][i] = R["gross"][i] * a["billing_fee_pct"]
        R["sga"][i] = R["sga_labor"][i] + R["ga_fixed"][i] + R["qr"][i] + R["billing"][i]
        R["ebitda"][i] = R["gp"][i] - R["sga"][i]

        # Below EBITDA — SBA + license note, each on its own beginning balance
        R["da"][i] = da_pm * m
        ints = sba_bal * rm * m
        prins = pmt * m - ints
        sba_bal -= prins
        R["int_sba"][i] = ints
        R["sba_bal"][i] = sba_bal
        # License note — closed-form amortization to ensure clean payoff at term
        if lic_bal > 0 and lic_rm > 0:
            factor = (1 + lic_rm) ** m
            new_bal = lic_bal * factor - lic_pmt * (factor - 1) / lic_rm
            new_bal = max(0.0, new_bal)
        else:
            new_bal = max(0.0, lic_bal - lic_pmt * m)
        lic_prin = lic_bal - new_bal
        lic_int = max(0.0, lic_pmt * m - lic_prin) if lic_bal > 0 else 0.0
        # If we paid off mid-period, payment = principal + interest only
        if new_bal == 0 and lic_bal > 0:
            # interest accrued only on the months we actually owed
            lic_int = lic_pmt * m - lic_prin
            lic_int = max(0.0, lic_int)
        lic_bal = new_bal
        R["lic_int"][i] = lic_int
        R["lic_prin"][i] = lic_prin
        R["lic_bal"][i] = lic_bal
        R["int_loc"][i] = loc_bal * a["loc_rate"] * (m / 12)
        R["pretax"][i] = R["ebitda"][i] - R["da"][i] - R["int_sba"][i] - lic_int - R["int_loc"][i]
        R["tax"][i] = R["net"][i] * a["tx_tax"] if R["pretax"][i] > 0 else 0.0
        R["ni"][i] = R["pretax"][i] - R["tax"][i]
        R["ds"][i] = ints + prins + lic_int + lic_prin
        R["principal"][i] = prins + lic_prin

        # Working capital
        R["ar"][i] = R["net"][i] * (a["ar_days"] / days)
        opex_payable = R["ga_fixed"][i] + R["qr"][i] + R["cogs_patient"][i]
        R["ap"][i] = opex_payable * (a["ap_days"] / days)
        d_ar = R["ar"][i] - prev_ar; d_ap = R["ap"][i] - prev_ap
        prev_ar = R["ar"][i]; prev_ap = R["ap"][i]
        R["cfo"][i] = R["ni"][i] + R["da"][i] - d_ar + d_ap

        # Cash + LOC sweep (LOC interest on prior balance -> no circularity)
        end = beg_cash + R["cfo"][i] - prins - lic_prin
        draw = repay = 0.0
        if end < a["min_cash"]:
            draw = a["min_cash"] - end
        elif loc_bal > 0:
            repay = min(loc_bal, end - a["min_cash"])
        end = end + draw - repay
        loc_bal = loc_bal + draw - repay
        R["loc_draw"][i] = draw; R["loc_repay"][i] = repay
        R["end_cash"][i] = end; R["loc_bal"][i] = loc_bal
        beg_cash = end

        # Balance-sheet asset roll-forward (three pools, three lives)
        accum_capex += da_capex_pm * m
        accum_startup += da_startup_pm * m
        accum_license += da_license_pm * m
        R["ppe_net"][i] = max(0.0, capex - accum_capex)
        R["intang_net"][i] = max(0.0, startup_total - accum_startup)
        R["license_net"][i] = max(0.0, lic_P - accum_license)
        re_cum += R["ni"][i]
        R["re"][i] = re_cum

    R["_scalars"] = dict(net_rate=net_rate, blended_gross=blended_gross,
                         wiadj=wiadj, startup_total=startup_total, pmt=pmt,
                         rate_t1=rate_t1, rate_t2=rate_t2)
    return R, a
