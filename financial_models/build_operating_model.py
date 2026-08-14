"""Azalea / Refuge OPERATING MODEL - formula-driven workbook.

Tabs: Assumptions | Census | Collections | Payroll | Weekly Cash | Monthly x36 |
Actuals vs Budget | Investor. Every business number lives on Assumptions (blue
input cells); downstream cells are formulas, so the file is a living tool.

Grain: weekly x13 (Jul-Sep 2026) + monthly x36 (Jul-26 .. Jun-29).
Debt base case: accepted seller terms ($125K down + $25K Sep-Dec + Jan balloon @6%).
Toggles: September $500K/6%/36mo refinance; SBA $450K.
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as gcl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

from .engine import styles as S
from .engine.model import ROSTER, PRN_VISIT, RHONDA_ANNUAL

NM, NW = 36, 13
MC0 = 3                                   # first monthly data col (C)
WC0 = 3                                   # first weekly data col (C)
MONTH_LABELS = []
_names = ["Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun"]
for i in range(NM):
    MONTH_LABELS.append(f"{_names[i % 12]}-{26 + ((i + 6) // 12)}")
WEEK_LABELS = [f"W{i+1}" for i in range(NW)]
WEEK_MONTH = [1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3]      # calendar-ish map
PAY_WEEKS = {3: 1, 5: 1, 7: 2, 9: 2, 11: 3, 13: 3}         # week -> month whose half-payroll pays

# ---- calibrated census defaults (Geoff's 7/6 plan) ----
# Full panel migrates wks 1-2 (open ~24-25); net growth +1.7/mo M1-2 (end M2 27.4),
# +4..+6 across M3-4 (end M4 32.4 mid), +3..+5 across M5-6 (end M6 35.0 mid);
# then to 50 by M24 and 56 by M36 (the validated growth plan, from the higher base).
ALOS_D = 82.0
WK_ADMITS = [12.0, 8.02, 5.62, 4.26, 2.76, 2.61, 2.65, 2.68, 2.61, 3.04, 3.0, 3.05, 3.1]
MO_ADMITS = [0, 0, 0, 13.58, 13.31, 13.79, 13.78, 14.17, 14.41, 14.7, 15.1, 15.33,
             15.63, 16.03, 16.26, 16.56, 16.95, 17.19, 17.48, 17.88, 18.11, 18.41, 18.81, 19.04,
             19.04, 19.22, 19.41, 19.59, 19.78, 19.96, 20.15, 20.33, 20.52, 20.7, 20.89, 21.08]
# End-of-month census the defaults reproduce (used by QA):
CENSUS_TARGETS = [25.29, 27.4, 29.9, 32.4, 33.7, 35.0, 35.8, 36.7, 37.5, 38.3, 39.2, 40.0,
                  40.8, 41.7, 42.5, 43.3, 44.2, 45.0, 45.8, 46.7, 47.5, 48.3, 49.2, 50.0,
                  50.5, 51.0, 51.5, 52.0, 52.5, 53.0, 53.5, 54.0, 54.5, 55.0, 55.5, 56.0]
# Hire-month overrides: capacity hires re-timed to when the NEW census crosses their
# ADC triggers; IDG conversions/quality hires pulled forward to match (engine converted
# at census ~30-37, which this plan reaches in months 4-9).
START_OVERRIDES = {
    "Capacity hire - 2nd CNA": 2, "Capacity hire - 3rd RN": 4, "Capacity hire - 3rd CNA": 5,
    "Capacity hire - 4th RN": 11, "Capacity hire - 4th CNA": 13, "Capacity hire - 5th RN": 23,
    "Capacity hire - 5th CNA": 25, "Capacity hire - 6th CNA": 33, "Capacity hire - 6th RN": 35,
    "FT Social Worker (MSW)": 6, "FT LVN - visits / on-call": 8, "FT Chaplain": 10,
    "Quality / Compliance Manager": 8, "Intake / Admissions Coordinator": 10,
    "Volunteer Coordinator": 12,
}

GA_LINES = [   # label, monthly $, freq months, pay day-of-month
    ("Facility rent", 2000, 1, 1),
    ("Utilities", 293, 1, 15),
    ("Internet", 779, 1, 15),
    ("Telephone / fax", 403, 1, 15),
    ("After-hours messaging", 48, 1, 15),
    ("EMR system", 1281, 1, 1),
    ("PCR (CPM)", 1001, 1, 15),
    ("Training / CEUs", 100, 1, 15),
    ("Other (client)", 1050, 1, 15),
    ("Bank / payroll fees", 1750, 1, 15),
    ("Marketing / patient acquisition", 3000, 1, 1),
    ("PL/GL insurance (corrected)", 1200, 3, 1),
    ("On-call stipends", 1500, 1, 15),
    ("Bereavement program", 300, 1, 15),
]

def mlet(i):  return gcl(MC0 + i)
def wlet(i):  return gcl(WC0 + i)
MRANGE = lambda r: f"$C${r}:$AL${r}"


class OB:
    """Operating-model book: styled-cell helpers + Assumptions address registry."""
    def __init__(self):
        self.wb = Workbook(); self.wb.remove(self.wb.active)
        self.addr = {}
        self.rows = {}          # sheet -> {key: row}

    def sheet(self, name, tab=None):
        ws = self.wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        if tab: ws.sheet_properties.tabColor = tab
        self.rows[name] = {}
        return ws

    def _set(self, ws, r, c, v, font, fmt=None, fill=None, align=None, border=None):
        cell = ws.cell(r, c, v); cell.font = font
        if fmt: cell.number_format = fmt
        if fill: cell.fill = fill
        if align: cell.alignment = align
        if border: cell.border = border
        return cell

    def lbl(self, ws, r, text, c=1, bold=False, indent=0, italic=False):
        cell = ws.cell(r, c, text)
        cell.font = S.Font(name=S.BASE_FONT, size=10, bold=bold, italic=italic, color=S.BLACK)
        cell.alignment = S.Alignment(horizontal="left", vertical="center", indent=indent)
        return cell

    def inp(self, ws, r, c, v, fmt, key=None):
        cell = self._set(ws, r, c, v, S.f_input(), fmt, S.fill(S.INPUTFILL), S.RIGHT, S.BORDER_THIN)
        if key:
            self.addr[key] = f"'{ws.title}'!${gcl(c)}${r}"
        return cell

    def fml(self, ws, r, c, f, fmt, bold=False, fill=None, link=False):
        return self._set(ws, r, c, f, S.f_link(bold) if link else S.f_formula(bold), fmt, fill, S.RIGHT)

    def title(self, ws, subtitle, span=14):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
        self._set(ws, 1, 1, "AZALEA HOSPICE - REFUGE OPERATING MODEL", S.f_title(), fill=S.fill(S.NAVY), align=S.LEFT)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
        self._set(ws, 2, 1, subtitle, S.f_subtitle(), fill=S.fill(S.NAVY), align=S.LEFT)
        ws.row_dimensions[1].height = 22

    def section(self, ws, r, text, span=14):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
        self._set(ws, r, 1, text, S.f_section(), fill=S.fill(S.MIDBLUE), align=S.LEFT)

    def mheader(self, ws, r):
        self.lbl(ws, r, "Month #", bold=True)
        for i in range(NM):
            self._set(ws, r, MC0 + i, i + 1, S.f_label(bold=True), S.FMT_INT,
                      S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
            self._set(ws, r + 1, MC0 + i, MONTH_LABELS[i], S.f_label(bold=True), None,
                      S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
        self.lbl(ws, r + 1, "", bold=True)
        ws.freeze_panes = ws.cell(r + 2, MC0)
        return r + 2

    def wheader(self, ws, r):
        for i in range(NW):
            self._set(ws, r, WC0 + i, WEEK_LABELS[i], S.f_label(bold=True), None,
                      S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
        ws.freeze_panes = ws.cell(r + 1, WC0)
        return r + 1

    def widths(self, ws, label=40, n=NM, w=10.5):
        ws.column_dimensions["A"].width = label
        ws.column_dimensions["B"].width = 13
        for i in range(n):
            ws.column_dimensions[gcl(MC0 + i)].width = w


# =====================================================================
def assumptions(bk: OB):
    ws = bk.sheet("Assumptions", tab="1F3864")
    bk.title(ws, "Assumptions - ALL editable cells are yellow/blue. Everything downstream is formulas.")
    bk.widths(ws, label=46, n=16, w=11)
    r = 4
    def item(key, label, val, fmt, note=None):
        nonlocal r
        bk.lbl(ws, r, label); bk.inp(ws, r, 2, val, fmt, key=key)
        if note: bk.lbl(ws, r, note, c=4, italic=True)
        r += 1

    bk.section(ws, r, "A. RATES & PAYER MIX"); r += 1
    item("net_rate", "Medicare blended NET rate ($/patient-day)", 181.0, S.FMT_RATE)
    item("mcaid_rate", "Medicaid blended NET rate ($/patient-day)", 155.0, S.FMT_RATE)
    item("mix_mcare", "Medicare share of patient days", 0.95, S.FMT_PCT)
    item("rate_esc", "Rate escalation (annual, yrs 2-3)", 0.025, S.FMT_PCT)
    item("days_mo", "Average days per month", 30.4, S.FMT_NUM1)
    r += 1

    bk.section(ws, r, "B. CENSUS ENGINE (admissions are the operating lever - edit rows on Census tab)"); r += 1
    item("alos", "Average length of stay (days)", ALOS_D, S.FMT_NUM1)
    item("death_share", "Deaths as share of separations", 0.55, S.FMT_PCT)
    item("capture", "SCENARIO: census capture (1.00 = plan)", 1.00, S.FMT_PCT)
    r += 1

    bk.section(ws, r, "C. BILLING & COLLECTIONS"); r += 1
    item("noe_loss_days", "Days lost per late NOE (per affected admit)", 5.0, S.FMT_NUM1)
    item("noe_rate", "Late-NOE rate (share of admits)", 0.02, S.FMT_PCT)
    item("hold_mo", "CHOW payment hold (months; 0/1/2)", 1, S.FMT_INT,
         "MAC hold on first claims after ownership change")
    item("lag_mo", "Base collection lag (months)", 1, S.FMT_INT)
    r += 1

    bk.section(ws, r, "D. PAYROLL & BENEFITS (per-employee register on Payroll tab)"); r += 1
    item("burden", "Payroll burden (taxes/WC) on W-2 + PRN", 0.1827, S.FMT_PCT2)
    item("health_pm", "Health insurance ($/covered head/mo)", 550.0, S.FMT_CUR)
    item("health_wait", "Health waiting period (months)", 3, S.FMT_INT)
    item("pay_adj", "SCENARIO: payroll adjustment (+/- %)", 0.0, S.FMT_PCT)
    item("sal_esc", "Salary escalation (annual, yrs 2-3)", 0.03, S.FMT_PCT)
    item("defer_pct", "Owner salary deferral % (0 = off; 0.40 = lever on)", 0.0, S.FMT_PCT)
    item("defer_end", "Deferral window ends (month #)", 6, S.FMT_INT)
    item("med_dir", "Medical director ($/mo, contracted)", 4000.0, S.FMT_CUR)
    item("med_dir2", "2nd medical director ($/mo)", 5000.0, S.FMT_CUR)
    item("med_dir2_mo", "2nd med director start month", 18, S.FMT_INT)
    item("rhonda", "Bookkeeping/admin support (annual)", float(RHONDA_ANNUAL), S.FMT_CUR)
    r += 1

    bk.section(ws, r, "E. PRN PER-VISIT ROSTER (census-driven)"); r += 1
    bk.lbl(ws, r, "Role", bold=True); bk.lbl(ws, r, "Visits/pt/mo", c=2, bold=True)
    bk.lbl(ws, r, "$/visit", c=3, bold=True); r += 1
    prn_first = r
    for (key, label, vpm, rate, _g) in PRN_VISIT:
        bk.lbl(ws, r, label)
        bk.inp(ws, r, 2, vpm, S.FMT_NUM2); bk.inp(ws, r, 3, rate, S.FMT_RATE)
        r += 1
    bk.addr["prn_block"] = (prn_first, r - 1)   # rows of the PRN table
    r += 1

    bk.section(ws, r, "F. PATIENT COSTS & NEW (BLIND-SPOT) LINES"); r += 1
    item("cogs_pd", "Patient costs $/patient-day (pharmacy+supplies+DME)", 8.99, S.FMT_RATE,
         "paid net-30 (following month)")
    item("mileage_adc", "Mileage $/ADC/month", 136.36, S.FMT_RATE, "~$3,000/mo at 22 ADC; paid following month")
    item("fees_pct", "Billing + QR fees (% of net revenue)", 0.023, S.FMT_PCT2)
    item("cap_pct", "Cap-liability reserve (% of Medicare revenue, accrual)", 0.02, S.FMT_PCT)
    r += 1

    bk.section(ws, r, "G. G&A BILL-TIMING CALENDAR"); r += 1
    bk.lbl(ws, r, "Line", bold=True); bk.lbl(ws, r, "$/period", c=2, bold=True)
    bk.lbl(ws, r, "Freq (mo)", c=3, bold=True); bk.lbl(ws, r, "Pay day", c=4, bold=True); r += 1
    ga_first = r
    for (label, amt, freq, day) in GA_LINES:
        bk.lbl(ws, r, label)
        bk.inp(ws, r, 2, float(amt), S.FMT_CUR); bk.inp(ws, r, 3, freq, S.FMT_INT)
        bk.inp(ws, r, 4, day, S.FMT_INT)
        r += 1
    bk.addr["ga_block"] = (ga_first, r - 1)
    r += 1

    bk.section(ws, r, "H. DEAL & DEBT (base = accepted seller terms; toggles below)"); r += 1
    item("price", "Purchase price", 500000.0, S.FMT_CUR)
    item("down", "Down payment at closing (Jul)", 125000.0, S.FMT_CUR)
    item("s_pmt", "Seller monthly payment", 25000.0, S.FMT_CUR)
    item("s_first", "First seller payment (month #)", 3, S.FMT_INT, "3 = September")
    item("s_last", "Last seller payment (month #)", 6, S.FMT_INT, "6 = December")
    item("s_rate", "Seller note rate (simple, annual)", 0.06, S.FMT_PCT)
    item("balloon_mo", "Balloon month (7 = January, per 36-month rule)", 7, S.FMT_INT)
    item("bref_rate", "January balloon refinance rate (APR)", 0.09, S.FMT_PCT,
         "personal refi that funds the balloon when the $500K refi toggle is OFF")
    item("bref_term", "January balloon refinance term (months)", 120, S.FMT_INT)
    item("refi_on", "REFI TOGGLE (1 = refinance full $500K)", 0, S.FMT_INT)
    item("refi_mo", "Refi funding month (3 = September)", 3, S.FMT_INT)
    item("refi_amt", "Refi amount", 500000.0, S.FMT_CUR)
    item("refi_rate", "Refi rate (APR)", 0.06, S.FMT_PCT)
    item("refi_term", "Refi term (months)", 36, S.FMT_INT)
    item("sba_on", "SBA TOGGLE (1 = SBA loan funds)", 0, S.FMT_INT)
    item("sba_mo", "SBA funding month", 3, S.FMT_INT)
    item("sba_amt", "SBA amount", 450000.0, S.FMT_CUR)
    item("sba_rate", "SBA rate (APR)", 0.105, S.FMT_PCT)
    item("sba_term", "SBA term (months)", 180, S.FMT_INT)
    r += 1

    bk.section(ws, r, "I. OPENING & OTHER"); r += 1
    item("cash0", "Opening cash (equity on deposit)", 250000.0, S.FMT_CUR)
    item("floor", "Minimum-cash floor (flag threshold)", 25000.0, S.FMT_CUR)
    item("amort_yrs", "License intangible amortization (years)", 15, S.FMT_INT)
    item("tx_tax", "TX franchise/margin tax (% of revenue if profitable)", 0.00375, S.FMT_PCT2)
    dv = DataValidation(type="list", formula1='"0,1,2"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(bk.addr["hold_mo"].split("!")[1].replace("$", ""))
    return ws


# =====================================================================
def census(bk: OB):
    ws = bk.sheet("Census", tab="2E75B6")
    bk.title(ws, "Census engine - admissions are BLUE inputs; flows follow ALOS. Weekly (13) then monthly (36).")
    bk.widths(ws, n=NM)
    A = bk.addr; R = bk.rows["Census"]
    r = 4
    bk.section(ws, r, "WEEKLY - first 13 weeks (Jul-Sep 2026)"); r += 1
    hr = bk.wheader(ws, r); r = hr
    bk.lbl(ws, r, "Admissions (per week)  [INPUT]")
    for i in range(NW): bk.inp(ws, r, WC0 + i, WK_ADMITS[i], S.FMT_NUM1)
    R["w_admits"] = r; r += 1
    bk.lbl(ws, r, "Separations (census x 7/ALOS)")
    for i in range(NW):
        prev = f"{wlet(i-1)}{r+2}" if i else "0"
        bk.fml(ws, r, WC0 + i, f"={prev}*7/{A['alos']}", S.FMT_NUM1)
    R["w_sep"] = r; r += 1
    bk.lbl(ws, r, "  of which deaths")
    for i in range(NW):
        bk.fml(ws, r, WC0 + i, f"={wlet(i)}{R['w_sep']}*{A['death_share']}", S.FMT_NUM1)
    r += 1
    bk.lbl(ws, r, "End-of-week census", bold=True)
    for i in range(NW):
        prev = f"{wlet(i-1)}{r}" if i else "0"
        bk.fml(ws, r, WC0 + i, f"={prev}+{wlet(i)}{R['w_admits']}-{wlet(i)}{R['w_sep']}",
               S.FMT_NUM1, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["w_end"] = r; r += 1
    bk.lbl(ws, r, "Average census (week)")
    for i in range(NW):
        prev = f"{wlet(i-1)}{R['w_end']}" if i else "0"
        bk.fml(ws, r, WC0 + i, f"=({prev}+{wlet(i)}{R['w_end']})/2*{A['capture']}", S.FMT_NUM1)
    R["w_avg"] = r; r += 1
    bk.lbl(ws, r, "Patient days (week)")
    for i in range(NW):
        bk.fml(ws, r, WC0 + i, f"={wlet(i)}{R['w_avg']}*7", S.FMT_NUM1)
    R["w_pd"] = r; r += 2

    bk.section(ws, r, "MONTHLY - 36 months (M1-M3 tie to the weekly engine)"); r += 1
    hr = bk.mheader(ws, r); r = hr
    bk.lbl(ws, r, "Admissions (per month)  [INPUT M4+]")
    for i in range(NM):
        if i < 3:
            wk = [(0, 3), (4, 8), (9, 12)][i]
            bk.fml(ws, r, MC0 + i,
                   f"=SUM({wlet(wk[0])}{R['w_admits']}:{wlet(wk[1])}{R['w_admits']})", S.FMT_NUM1, link=True)
        else:
            bk.inp(ws, r, MC0 + i, MO_ADMITS[i], S.FMT_NUM1)
    R["m_admits"] = r; r += 1
    bk.lbl(ws, r, "Separations")
    for i in range(NM):
        if i < 3:
            wk = [(0, 3), (4, 8), (9, 12)][i]
            bk.fml(ws, r, MC0 + i, f"=SUM({wlet(wk[0])}{R['w_sep']}:{wlet(wk[1])}{R['w_sep']})", S.FMT_NUM1)
        else:
            bk.fml(ws, r, MC0 + i, f"={mlet(i-1)}{r+2}*{A['days_mo']}/{A['alos']}", S.FMT_NUM1)
    R["m_sep"] = r; r += 1
    bk.lbl(ws, r, "  of which deaths")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['m_sep']}*{A['death_share']}", S.FMT_NUM1)
    R["m_deaths"] = r; r += 1
    bk.lbl(ws, r, "End-of-month census", bold=True)
    for i in range(NM):
        if i < 3:
            wk_end = [3, 8, 12][i]
            bk.fml(ws, r, MC0 + i, f"={wlet(wk_end)}{R['w_end']}", S.FMT_NUM1, bold=True)
        else:
            bk.fml(ws, r, MC0 + i, f"={mlet(i-1)}{r}+{mlet(i)}{R['m_admits']}-{mlet(i)}{R['m_sep']}",
                   S.FMT_NUM1, bold=True)
    R["m_end"] = r; r += 1
    bk.lbl(ws, r, "ADC (avg daily census, x capture)", bold=True)
    for i in range(NM):
        if i < 3:
            wk = [(0, 3), (4, 8), (9, 12)][i]
            # end-of-week census basis (matches the calibrated proven path)
            bk.fml(ws, r, MC0 + i,
                   f"=AVERAGE({wlet(wk[0])}{R['w_end']}:{wlet(wk[1])}{R['w_end']})*{A['capture']}",
                   S.FMT_NUM1, bold=True, fill=S.fill(S.LIGHTBLUE))
        else:
            bk.fml(ws, r, MC0 + i, f"=({mlet(i-1)}{R['m_end']}+{mlet(i)}{R['m_end']})/2*{A['capture']}",
                   S.FMT_NUM1, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["m_adc"] = r; r += 1
    bk.lbl(ws, r, "Patient days (month)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['m_adc']}*{A['days_mo']}", S.FMT_NUM)
    R["m_pd"] = r; r += 1
    return ws


# =====================================================================
def collections(bk: OB):
    ws = bk.sheet("Collections", tab="2E75B6")
    bk.title(ws, "Revenue earned -> cash collected. NOE losses, sequential claims, CHOW payment hold on first claims.")
    bk.widths(ws, n=NM)
    A = bk.addr; C = bk.rows["Census"]; R = bk.rows["Collections"]
    r = 4
    hr = bk.mheader(ws, r); r = hr
    bk.lbl(ws, r, "Blended net rate ($/PD, escalated)")
    for i in range(NM):
        yr = i // 12
        bk.fml(ws, r, MC0 + i,
               f"=({A['mix_mcare']}*{A['net_rate']}+(1-{A['mix_mcare']})*{A['mcaid_rate']})"
               f"*(1+{A['rate_esc']})^{yr}", S.FMT_RATE)
    R["rate"] = r; r += 1
    bk.lbl(ws, r, "Revenue earned (PD x rate)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=Census!{mlet(i)}{C['m_pd']}*{mlet(i)}{R['rate']}", S.FMT_CUR)
    R["earned"] = r; r += 1
    bk.lbl(ws, r, "less: NOE losses (late notices)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=-Census!{mlet(i)}{C['m_admits']}*{A['noe_rate']}*{A['noe_loss_days']}*{mlet(i)}{R['rate']}",
               S.FMT_CUR)
    R["noe"] = r; r += 1
    bk.lbl(ws, r, "NET REVENUE (billable)", bold=True)
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['earned']}+{mlet(i)}{R['noe']}", S.FMT_CUR,
               bold=True, fill=S.fill(S.LIGHTBLUE))
    R["net"] = r; r += 1
    bk.lbl(ws, r, "Collection month (svc mo + lag + hold on first claims)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"={i+1}+{A['lag_mo']}+MAX(0,{A['hold_mo']}-{i})", S.FMT_INT)
    R["cmo"] = r; r += 1
    bk.lbl(ws, r, "CASH COLLECTED", bold=True)
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUMPRODUCT(({MRANGE(R['cmo'])}={i+1})*{MRANGE(R['net'])})", S.FMT_CUR,
               bold=True, fill=S.fill(S.GREENBAR if hasattr(S, 'GREENBAR') else S.LIGHTBLUE))
    R["cash"] = r; r += 1
    bk.lbl(ws, r, "AR balance (earned not yet collected)", italic=True)
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUM($C${R['net']}:{mlet(i)}{R['net']})-SUM($C${R['cash']}:{mlet(i)}{R['cash']})", S.FMT_CUR)
    R["ar"] = r; r += 2

    bk.section(ws, r, "WEEKLY RECEIPTS (13 weeks; month receipts spread across its weeks)"); r += 1
    hr = bk.wheader(ws, r); r = hr
    bk.lbl(ws, r, "Cash received (week)")
    for i in range(NW):
        mo = WEEK_MONTH[i]
        nweeks = WEEK_MONTH.count(mo)
        bk.fml(ws, r, WC0 + i, f"={mlet(mo-1)}{R['cash']}/{nweeks}", S.FMT_CUR)
    R["w_cash"] = r; r += 1
    return ws


# =====================================================================
def payroll(bk: OB):
    ws = bk.sheet("Payroll", tab="548235")
    bk.title(ws, "Per-employee register (blue inputs) -> monthly cost matrix -> summary rows. Semi-monthly cash timing on Weekly Cash tab.")
    ws.column_dimensions["A"].width = 30
    for c in "BCDEFG": ws.column_dimensions[c].width = 11
    for i in range(NM): ws.column_dimensions[gcl(9 + i)].width = 10.5
    A = bk.addr; R = bk.rows["Payroll"]
    r = 4
    bk.section(ws, r, "REGISTER  (I..AT = monthly gross cost incl. scenario adj, deferral, escalation)", span=8); r += 1
    heads = ["Employee / role", "Salary", "Hire mo", "Direct?", "Hlth wait", "Defer %", "Defer end", ""]
    for j, h in enumerate(heads):
        bk.lbl(ws, r, h, c=1 + j, bold=True)
    for i in range(NM):
        bk._set(ws, r, 9 + i, f"M{i+1}", S.f_label(bold=True), None, S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
    r += 1
    first = r
    OWNERS = ("Brad Woodard", "Dana Davenport", "Silas Shelton")
    for (name, role, sal, start, grp, _o) in ROSTER:
        # Capacity hires are census-triggered (ADC 25-55). Default hire months match
        # the 7/6 census plan; if admissions lag, push these months out.
        capacity = name.startswith("Capacity hire")
        start_eff = START_OVERRIDES.get(name, int(start))
        bk.lbl(ws, r, f"{name} - {role[:30]}" + ("  [census-triggered - slip if ramp lags]" if capacity else ""))
        bk.inp(ws, r, 2, float(sal), S.FMT_CUR)
        bk.inp(ws, r, 3, start_eff, S.FMT_INT)
        bk.inp(ws, r, 4, "D" if grp == "direct" else "I", None)
        bk.inp(ws, r, 5, int(3), S.FMT_INT)
        is_owner = any(o in name for o in OWNERS)
        bk.fml(ws, r, 6, f"=IF({1 if is_owner else 0},{A['defer_pct']},0)", S.FMT_PCT)
        bk.fml(ws, r, 7, f"={A['defer_end']}", S.FMT_INT)
        for i in range(NM):
            yr = i // 12
            f = (f"=IF({i+1}>=$C{r},$B{r}/12*(1+{A['pay_adj']})*(1+{A['sal_esc']})^{yr}"
                 f"*(1-IF({i+1}<=$G{r},$F{r},0)),0)")
            bk.fml(ws, r, 9 + i, f, S.FMT_CUR)
        r += 1
    last = r - 1
    R["reg"] = (first, last); r += 1

    def sumrow(label, crit, key, bold=False):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            col = gcl(9 + i)
            if crit:
                f = f"=SUMPRODUCT(($D${first}:$D${last}=\"{crit}\")*{col}{first}:{col}{last})"
            else:
                f = f"=SUM({col}{first}:{col}{last})"
            bk.fml(ws, r, 9 + i, f, S.FMT_CUR, bold=bold)
        R[key] = r; r += 1

    sumrow("W-2 salaries - DIRECT (clinical)", "D", "w2_d")
    sumrow("W-2 salaries - INDIRECT (admin/sales)", "I", "w2_i")
    sumrow("W-2 total", None, "w2", bold=True)

    C = bk.rows["Census"]
    pf, pl = A["prn_block"]
    bk.lbl(ws, r, "PRN per-visit staff (census-driven)")
    for i in range(NM):
        bk.fml(ws, r, 9 + i,
               f"=Census!{mlet(i)}{C['m_adc']}*SUMPRODUCT(Assumptions!$B${pf}:$B${pl},Assumptions!$C${pf}:$C${pl})",
               S.FMT_CUR)
    R["prn"] = r; r += 1
    bk.lbl(ws, r, "Medical director(s) (contracted)")
    for i in range(NM):
        bk.fml(ws, r, 9 + i, f"={A['med_dir']}+IF({i+1}>={A['med_dir2_mo']},{A['med_dir2']},0)", S.FMT_CUR)
    R["medd"] = r; r += 1
    bk.lbl(ws, r, "Bookkeeping/admin support")
    for i in range(NM):
        bk.fml(ws, r, 9 + i, f"={A['rhonda']}/12", S.FMT_CUR)
    R["rhonda"] = r; r += 1
    bk.lbl(ws, r, "Burden (taxes/WC)")
    for i in range(NM):
        col = gcl(9 + i)
        bk.fml(ws, r, 9 + i, f"=({col}{R['w2']}+{col}{R['prn']}+{col}{R['rhonda']})*{A['burden']}", S.FMT_CUR)
    R["burden"] = r; r += 1
    bk.lbl(ws, r, "Health insurance (after waiting period)")
    for i in range(NM):
        bk.fml(ws, r, 9 + i,
               f"=SUMPRODUCT((($C${first}:$C${last}+$E${first}:$E${last})<={i+1})*1)*{A['health_pm']}",
               S.FMT_CUR)
    R["health"] = r; r += 1
    bk.lbl(ws, r, "TOTAL LABOR COST", bold=True)
    for i in range(NM):
        col = gcl(9 + i)
        bk.fml(ws, r, 9 + i,
               f"={col}{R['w2']}+{col}{R['prn']}+{col}{R['medd']}+{col}{R['rhonda']}+{col}{R['burden']}+{col}{R['health']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["total"] = r; r += 1
    return ws


# =====================================================================
def monthly(bk: OB):
    ws = bk.sheet("Monthly x36", tab="7F7F7F")
    bk.title(ws, "36-month pro forma: P&L, debt schedules (seller / refi toggle / SBA toggle), direct-method cash, DSCR & floor flags.")
    bk.widths(ws, n=NM)
    A = bk.addr; C = bk.rows["Census"]; L = bk.rows["Collections"]; P = bk.rows["Payroll"]
    R = bk.rows["Monthly x36"]
    gaf, gal = A["ga_block"]
    r = 4
    R["_hdr_num"] = r
    hr = bk.mheader(ws, r); r = hr

    def link(label, sheet, srow, key, srccol0=MC0, bold=False, fill=None, neg=False):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for i in range(NM):
            src = f"'{sheet}'!{gcl(srccol0 + i)}{srow}"
            bk.fml(ws, r, MC0 + i, f"={'-' if neg else ''}{src}", S.FMT_CUR, bold=bold, fill=fill, link=True)
        R[key] = r; r += 1

    bk.section(ws, r, "P&L"); r += 1
    bk.lbl(ws, r, "ADC")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=Census!{mlet(i)}{C['m_adc']}", S.FMT_NUM1, link=True)
    R["adc"] = r; r += 1
    link("NET REVENUE (after NOE losses)", "Collections", bk.rows["Collections"]["net"], "net", bold=True,
         fill=S.fill(S.LIGHTBLUE))
    link("Total labor cost", "Payroll", P["total"], "labor", srccol0=9)
    bk.lbl(ws, r, "Patient costs (pharmacy/supplies/DME)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=Census!{mlet(i)}{C['m_pd']}*{A['cogs_pd']}", S.FMT_CUR)
    R["pcogs"] = r; r += 1
    bk.lbl(ws, r, "Mileage")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=Census!{mlet(i)}{C['m_adc']}*{A['mileage_adc']}", S.FMT_CUR)
    R["mileage"] = r; r += 1
    bk.lbl(ws, r, "G&A (bill calendar, monthly-equivalent)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUMPRODUCT(Assumptions!$B${gaf}:$B${gal}/Assumptions!$C${gaf}:$C${gal})", S.FMT_CUR)
    R["ga"] = r; r += 1
    bk.lbl(ws, r, "Billing + QR fees (% of net)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['net']}*{A['fees_pct']}", S.FMT_CUR)
    R["fees"] = r; r += 1
    bk.lbl(ws, r, "EBITDA", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={c}{R['net']}-{c}{R['labor']}-{c}{R['pcogs']}-{c}{R['mileage']}-{c}{R['ga']}-{c}{R['fees']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["ebitda"] = r; r += 1
    bk.lbl(ws, r, "Cap-liability reserve (accrual, non-cash)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-{mlet(i)}{R['net']}*{A['mix_mcare']}*{A['cap_pct']}", S.FMT_CUR)
    R["cap"] = r; r += 1
    bk.lbl(ws, r, "License amortization (non-cash)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-{A['price']}/{A['amort_yrs']}/12", S.FMT_CUR)
    R["amort"] = r; r += 1

    bk.section(ws, r, "DEBT - SELLER NOTE (base terms; auto-paid off if refi toggled)"); r += 1
    bk.lbl(ws, r, "Seller beg balance")
    for i in range(NM):
        prev = f"{mlet(i-1)}{r+4}" if i else f"={A['price']}-{A['down']}"
        bk.fml(ws, r, MC0 + i, prev if i == 0 else f"={prev}", S.FMT_CUR)
    R["s_beg"] = r; r += 1
    bk.lbl(ws, r, "Seller interest accrued (6% simple)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['s_beg']}*{A['s_rate']}/12", S.FMT_CUR)
    R["s_int"] = r; r += 1
    bk.lbl(ws, r, "Seller principal payments")
    for i in range(NM):
        f = (f"=IF({mlet(i)}{R['s_beg']}<=0,0,"
             f"IF(AND({A['refi_on']}=1,{i+1}={A['refi_mo']}),{mlet(i)}{R['s_beg']},"
             f"IF(AND({i+1}>={A['s_first']},{i+1}<={A['s_last']}),{A['s_pmt']},"
             f"IF({i+1}={A['balloon_mo']},{mlet(i)}{R['s_beg']},0))))")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["s_prin"] = r; r += 1
    bk.lbl(ws, r, "Seller interest PAID (at payoff/balloon)")
    for i in range(NM):
        f = (f"=IF(OR(AND({A['refi_on']}=1,{i+1}={A['refi_mo']}),"
             f"AND({A['refi_on']}=0,{i+1}={A['balloon_mo']})),"
             f"SUM($C${R['s_int']}:{mlet(i)}{R['s_int']}),0)")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["s_intpaid"] = r; r += 1
    bk.lbl(ws, r, "Seller end balance")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['s_beg']}-{mlet(i)}{R['s_prin']}", S.FMT_CUR)
    R["s_end"] = r; r += 1

    bk.section(ws, r, "DEBT - REFI (toggle) & SBA (toggle)"); r += 1
    bk.lbl(ws, r, "Refi payment (P&I, from month after funding)")
    for i in range(NM):
        f = (f"=IF(AND({A['refi_on']}=1,{i+1}>{A['refi_mo']},{i+1}<={A['refi_mo']}+{A['refi_term']}),"
             f"-PMT({A['refi_rate']}/12,{A['refi_term']},{A['refi_amt']}),0)")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["refi_pmt"] = r; r += 1
    bk.lbl(ws, r, "Refi balance (end)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{r}" if i else "0"
        f = (f"=IF({A['refi_on']}=0,0,IF({i+1}<{A['refi_mo']},0,"
             f"IF({i+1}={A['refi_mo']},{A['refi_amt']},"
             f"MAX(0,{prev}*(1+{A['refi_rate']}/12)-{mlet(i)}{R['refi_pmt']}))))")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["refi_bal"] = r; r += 1
    bk.lbl(ws, r, "Refi interest")
    for i in range(NM):
        prev = f"{mlet(i-1)}{R['refi_bal']}" if i else "0"
        bk.fml(ws, r, MC0 + i, f"={prev}*{A['refi_rate']}/12*({mlet(i)}{R['refi_pmt']}>0)", S.FMT_CUR)
    R["refi_int"] = r; r += 1
    bk.lbl(ws, r, "SBA payment (P&I) / interest")
    for i in range(NM):
        f = (f"=IF(AND({A['sba_on']}=1,{i+1}>{A['sba_mo']}),"
             f"-PMT({A['sba_rate']}/12,{A['sba_term']},{A['sba_amt']}),0)")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["sba_pmt"] = r; r += 1
    bk.lbl(ws, r, "Balloon refinance amount (funds Jan balloon when refi OFF)")
    hdr_row = R.get("_hdr_num")
    bk.fml(ws, r, 2,
           f"=IF({A['refi_on']}=1,0,SUMPRODUCT(($C${hdr_row}:$AL${hdr_row}={A['balloon_mo']})"
           f"*($C${R['s_prin']}:$AL${R['s_prin']}+$C${R['s_intpaid']}:$AL${R['s_intpaid']})))", S.FMT_CUR)
    R["bref_amt"] = r
    bref_amt = f"$B${r}"; r += 1
    bk.lbl(ws, r, "Balloon-refi payment (P&I, from month after balloon)")
    for i in range(NM):
        f = (f"=IF(AND({A['refi_on']}=0,{i+1}>{A['balloon_mo']},{i+1}<={A['balloon_mo']}+{A['bref_term']}),"
             f"-PMT({A['bref_rate']}/12,{A['bref_term']},{bref_amt}),0)")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["bref_pmt"] = r; r += 1
    bk.lbl(ws, r, "Balloon-refi balance (end)")
    for i in range(NM):
        prev = f"{mlet(i-1)}{r}" if i else "0"
        f = (f"=IF({A['refi_on']}=1,0,IF({i+1}<{A['balloon_mo']},0,"
             f"IF({i+1}={A['balloon_mo']},{bref_amt},"
             f"MAX(0,{prev}*(1+{A['bref_rate']}/12)-{mlet(i)}{R['bref_pmt']}))))")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["bref_bal"] = r; r += 1
    bk.lbl(ws, r, "Balloon-refi interest")
    for i in range(NM):
        prev = f"{mlet(i-1)}{R['bref_bal']}" if i else "0"
        bk.fml(ws, r, MC0 + i, f"={prev}*{A['bref_rate']}/12*({mlet(i)}{R['bref_pmt']}>0)", S.FMT_CUR)
    R["bref_int"] = r; r += 1

    bk.section(ws, r, "NET INCOME"); r += 1
    bk.lbl(ws, r, "Pre-tax income")
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={c}{R['ebitda']}+{c}{R['cap']}+{c}{R['amort']}-{c}{R['s_int']}-{c}{R['refi_int']}"
               f"-{c}{R['bref_int']}"
               f"-IF({A['sba_on']}=1,{A['sba_amt']}*{A['sba_rate']}/12,0)*({c}{R['sba_pmt']}>0)",
               S.FMT_CUR)
    R["pretax"] = r; r += 1
    bk.lbl(ws, r, "TX margin tax")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=IF({mlet(i)}{R['pretax']}>0,{mlet(i)}{R['net']}*{A['tx_tax']},0)", S.FMT_CUR)
    R["tax"] = r; r += 1
    bk.lbl(ws, r, "NET INCOME", bold=True)
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['pretax']}-{mlet(i)}{R['tax']}", S.FMT_CUR, bold=True)
    R["ni"] = r; r += 1

    bk.section(ws, r, "CASH (direct method)"); r += 1
    link("Cash collected", "Collections", L["cash"], "cash_in", bold=False)
    bk.lbl(ws, r, "Payroll & labor paid (in month)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-'Payroll'!{gcl(9+i)}{P['total']}", S.FMT_CUR)
    R["cash_pay"] = r; r += 1
    bk.lbl(ws, r, "Vendors: G&A (in month) + patient costs & mileage (net-30)")
    for i in range(NM):
        prevm = f"({mlet(i-1)}{R['pcogs']}+{mlet(i-1)}{R['mileage']})" if i else "0"
        bk.fml(ws, r, MC0 + i, f"=-{mlet(i)}{R['ga']}-{mlet(i)}{R['fees']}-{prevm}", S.FMT_CUR)
    R["cash_vend"] = r; r += 1
    bk.lbl(ws, r, "Seller payments (incl. down M1, interest at payoff)")
    for i in range(NM):
        down = f"-{A['down']}" if i == 0 else ""
        bk.fml(ws, r, MC0 + i, f"=-{mlet(i)}{R['s_prin']}-{mlet(i)}{R['s_intpaid']}{down}", S.FMT_CUR)
    R["cash_seller"] = r; r += 1
    bk.lbl(ws, r, "Financing: refi/SBA/balloon-refi proceeds and payments")
    for i in range(NM):
        f = (f"=IF(AND({A['refi_on']}=1,{i+1}={A['refi_mo']}),{A['refi_amt']},0)"
             f"-{mlet(i)}{R['refi_pmt']}"
             f"+IF(AND({A['refi_on']}=0,{i+1}={A['balloon_mo']}),$B${R['bref_amt']},0)"
             f"-{mlet(i)}{R['bref_pmt']}"
             f"+IF(AND({A['sba_on']}=1,{i+1}={A['sba_mo']}),{A['sba_amt']},0)-{mlet(i)}{R['sba_pmt']}")
        bk.fml(ws, r, MC0 + i, f, S.FMT_CUR)
    R["cash_fin"] = r; r += 1
    bk.lbl(ws, r, "ENDING CASH", bold=True)
    for i in range(NM):
        prev = f"{mlet(i-1)}{r}" if i else f"{A['cash0']}"
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={prev}+{c}{R['cash_in']}+{c}{R['cash_pay']}+{c}{R['cash_vend']}+{c}{R['cash_seller']}+{c}{R['cash_fin']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["cash_end"] = r
    rng = f"C{r}:AL{r}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan",
        formula=[bk.addr["floor"]], fill=S.fill("F4CCCC")))
    r += 1
    bk.lbl(ws, r, "DSCR (EBITDA / recurring debt service)")
    for i in range(NM):
        c = mlet(i)
        ds = f"({c}{R['refi_pmt']}+{c}{R['bref_pmt']}+{c}{R['sba_pmt']}+IF(AND({i+1}>={A['s_first']},{i+1}<={A['s_last']},{A['refi_on']}=0),{A['s_pmt']},0))"
        bk.fml(ws, r, MC0 + i, f"=IF({ds}>0,{c}{R['ebitda']}/{ds},\"-\")", S.FMT_MULT)
    R["dscr"] = r; r += 1
    return ws


# =====================================================================
def weekly(bk: OB):
    ws = bk.sheet("Weekly Cash 13wk", tab="C00000")
    bk.title(ws, "First 13 weeks (Jul-Sep 2026), the danger zone. Payroll on the 15th/EOM; vendor bills by calendar day; receipts from Collections.")
    bk.widths(ws, n=NW, w=11.5)
    A = bk.addr; L = bk.rows["Collections"]; P = bk.rows["Payroll"]; M = bk.rows["Monthly x36"]
    R = bk.rows["Weekly Cash 13wk"]
    gaf, gal = A["ga_block"]
    r = 4
    hr = bk.wheader(ws, r); r = hr
    bk.lbl(ws, r, "Cash received")
    for i in range(NW):
        bk.fml(ws, r, WC0 + i, f"=Collections!{wlet(i)}{L['w_cash']}", S.FMT_CUR, link=True)
    R["rcpt"] = r; r += 1
    bk.lbl(ws, r, "Payroll (semi-monthly: 15th & EOM)")
    for i in range(NW):
        wk = i + 1
        if wk in PAY_WEEKS:
            mo = PAY_WEEKS[wk]
            bk.fml(ws, r, WC0 + i, f"=-'Payroll'!{gcl(9+mo-1)}{P['total']}/2", S.FMT_CUR)
        else:
            bk.fml(ws, r, WC0 + i, "=0", S.FMT_CUR)
    R["pay"] = r; r += 1
    bk.lbl(ws, r, "Vendor bills (G&A calendar day)")
    for i in range(NW):
        wk = i + 1; mo = WEEK_MONTH[i]
        # week covers days: position within month
        wk_in_mo = [w for w in range(NW) if WEEK_MONTH[w] == mo].index(i)
        nwks = WEEK_MONTH.count(mo)
        d0 = int(round(wk_in_mo * 30.4 / nwks)) + 1
        d1 = int(round((wk_in_mo + 1) * 30.4 / nwks))
        f = (f"=-SUMPRODUCT((Assumptions!$D${gaf}:$D${gal}>={d0})*(Assumptions!$D${gaf}:$D${gal}<={d1})"
             f"*(MOD({mo}-1,Assumptions!$C${gaf}:$C${gal})=0)"
             f"*Assumptions!$B${gaf}:$B${gal})")
        bk.fml(ws, r, WC0 + i, f, S.FMT_CUR)
    R["vend"] = r; r += 1
    bk.lbl(ws, r, "Patient costs + mileage (net-30: paid 15th of following month)")
    for i in range(NW):
        wk = i + 1
        if wk == 7:      # Aug 15 pays July
            bk.fml(ws, r, WC0 + i, f"=-('Monthly x36'!C{M['pcogs']}+'Monthly x36'!C{M['mileage']})", S.FMT_CUR)
        elif wk == 11:   # Sep 15 pays August
            bk.fml(ws, r, WC0 + i, f"=-('Monthly x36'!D{M['pcogs']}+'Monthly x36'!D{M['mileage']})", S.FMT_CUR)
        else:
            bk.fml(ws, r, WC0 + i, "=0", S.FMT_CUR)
    R["pc"] = r; r += 1
    bk.lbl(ws, r, "Fees (billing/QR, with receipts)")
    for i in range(NW):
        mo = WEEK_MONTH[i]; nwks = WEEK_MONTH.count(mo)
        bk.fml(ws, r, WC0 + i, f"=-'Monthly x36'!{mlet(mo-1)}{M['fees']}/{nwks}", S.FMT_CUR)
    R["fees"] = r; r += 1
    bk.lbl(ws, r, "Seller: down payment (wk 1) / Sept payment (wk 11) / refi if ON (wk 13)")
    for i in range(NW):
        wk = i + 1
        if wk == 1:
            bk.fml(ws, r, WC0 + i, f"=-{A['down']}", S.FMT_CUR)
        elif wk == 11:
            bk.fml(ws, r, WC0 + i,
                   f"=-IF(AND({A['refi_on']}=0,{A['s_first']}<=3),{A['s_pmt']},0)", S.FMT_CUR)
        elif wk == 13:
            bk.fml(ws, r, WC0 + i,
                   f"=IF({A['refi_on']}=1,{A['refi_amt']}"
                   f"-('Monthly x36'!E{M['s_beg']}+SUM('Monthly x36'!$C${M['s_int']}:$E${M['s_int']})),0)", S.FMT_CUR)
        else:
            bk.fml(ws, r, WC0 + i, "=0", S.FMT_CUR)
    R["seller"] = r; r += 1
    bk.lbl(ws, r, "SBA proceeds (if toggled, lands 1st week of funding month)")
    MONTH_FIRST_WEEK = {1: 1, 2: 5, 3: 10}
    for i in range(NW):
        wk = i + 1
        conds = [f"AND({A['sba_on']}=1,{A['sba_mo']}={mo},{wk}={fw})" for mo, fw in MONTH_FIRST_WEEK.items()]
        f = f"=IF(OR({','.join(conds)}),{A['sba_amt']},0)"
        bk.fml(ws, r, WC0 + i, f, S.FMT_CUR)
    R["sba"] = r; r += 1
    bk.lbl(ws, r, "ENDING CASH", bold=True)
    for i in range(NW):
        prev = f"{wlet(i-1)}{r}" if i else f"{A['cash0']}"
        c = wlet(i)
        bk.fml(ws, r, WC0 + i,
               f"={prev}+{c}{R['rcpt']}+{c}{R['pay']}+{c}{R['vend']}+{c}{R['pc']}+{c}{R['fees']}+{c}{R['seller']}+{c}{R['sba']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    R["end"] = r
    ws.conditional_formatting.add(f"C{r}:O{r}", CellIsRule(operator="lessThan",
        formula=[bk.addr["floor"]], fill=S.fill("F4CCCC")))
    r += 2
    bk.lbl(ws, r, "Note: weekly receipts spread each month's collections evenly across its weeks; payroll halves on the "
                  "15th/EOM; vendor hits per the Assumptions bill calendar. Precision is operational, not accounting-grade.",
           italic=True)
    return ws


# =====================================================================
def avb(bk: OB):
    ws = bk.sheet("Actuals vs Budget", tab="BF8F00")
    bk.title(ws, "Enter ACTUALS monthly (blue cells); variance computes automatically. First 12 months.")
    bk.widths(ws, n=12, w=12)
    M = bk.rows["Monthly x36"]
    r = 4
    bk.lbl(ws, r, "", bold=True)
    for i in range(12):
        bk._set(ws, r, MC0 + i, MONTH_LABELS[i], S.f_label(bold=True), None, S.fill(S.GREYHDR), S.CENTER, S.BORDER_THIN)
    r += 1
    lines = [("Net revenue", "net"), ("Total labor", "labor"), ("EBITDA", "ebitda"), ("Ending cash", "cash_end")]
    for label, key in lines:
        bk.lbl(ws, r, f"{label} - BUDGET", bold=True)
        for i in range(12):
            bk.fml(ws, r, MC0 + i, f"='Monthly x36'!{mlet(i)}{M[key]}", S.FMT_CUR, link=True)
        r += 1
        bk.lbl(ws, r, f"{label} - ACTUAL  [enter]")
        for i in range(12):
            bk.inp(ws, r, MC0 + i, None, S.FMT_CUR)
        r += 1
        bk.lbl(ws, r, f"{label} - variance", italic=True)
        for i in range(12):
            bk.fml(ws, r, MC0 + i, f"=IF({mlet(i)}{r-1}=\"\",\"\",{mlet(i)}{r-1}-{mlet(i)}{r-2})", S.FMT_CUR)
        r += 2
    return ws


# =====================================================================
def investor(bk: OB):
    ws = bk.sheet("Investor Summary", tab="538135")
    bk.title(ws, "Jim Bullard - position, debt stack, distributable view. Links to the live model.")
    ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 50
    A = bk.addr; M = bk.rows["Monthly x36"]
    r = 4
    bk.section(ws, r, "CAP TABLE", span=3); r += 1
    for name, pct, note in [("Geoff Schackmann (Manager, direct)", 0.399, "sole 20%+ owner; personal guaranty"),
                            ("James E. Bullard", 0.195, "$195,000 cash; passive; no SBA PG"),
                            ("Silas Shelton / Dana Davenport / Brad Woodard", 0.399, "13.3% each, restricted/vesting"),
                            ("Unissued pool", 0.007, "reserved")]:
        bk.lbl(ws, r, name); bk.fml(ws, r, 2, f"={pct}", S.FMT_PCT); bk.lbl(ws, r, note, c=3, italic=True); r += 1
    r += 1
    bk.section(ws, r, "DEBT STACK (live)", span=3); r += 1
    bk.lbl(ws, r, "Seller note balance - end of M6 (Dec-26)")
    bk.fml(ws, r, 2, f"='Monthly x36'!H{M['s_end']}", S.FMT_CUR, link=True); r += 1
    bk.lbl(ws, r, "Seller fully paid by")
    bk.fml(ws, r, 2, f"=IF({A['refi_on']}=1,\"Sep-26 (refi)\",\"Jan-27 (balloon)\")", None); r += 1
    bk.lbl(ws, r, "Refi balance - end of Yr 1 (if toggled)")
    bk.fml(ws, r, 2, f"='Monthly x36'!N{M['refi_bal']}", S.FMT_CUR, link=True); r += 1
    r += 1
    bk.section(ws, r, "YEAR-1 RESULTS (live)", span=3); r += 1
    bk.lbl(ws, r, "FY1 net revenue")
    bk.fml(ws, r, 2, f"=SUM('Monthly x36'!C{M['net']}:N{M['net']})", S.FMT_CUR, link=True); r += 1
    bk.lbl(ws, r, "FY1 EBITDA")
    bk.fml(ws, r, 2, f"=SUM('Monthly x36'!C{M['ebitda']}:N{M['ebitda']})", S.FMT_CUR, link=True); r += 1
    bk.lbl(ws, r, "FY1 net income")
    bk.fml(ws, r, 2, f"=SUM('Monthly x36'!C{M['ni']}:N{M['ni']})", S.FMT_CUR, link=True); r += 1
    bk.lbl(ws, r, "Cash at end of FY1")
    bk.fml(ws, r, 2, f"='Monthly x36'!N{M['cash_end']}", S.FMT_CUR, link=True); r += 1
    r += 1
    bk.section(ws, r, "DISTRIBUTABLE VIEW (S-corp: strictly pro-rata)", span=3); r += 1
    bk.lbl(ws, r, "Estimated tax distribution rate"); bk.inp(ws, r, 2, 0.30, S.FMT_PCT); td = r; r += 1
    bk.lbl(ws, r, "FY1 tax distributions (all members)")
    bk.fml(ws, r, 2, f"=MAX(0,SUM('Monthly x36'!C{M['ni']}:N{M['ni']}))*$B${td}", S.FMT_CUR); r += 1
    bk.lbl(ws, r, "  Jim's 19.5% share")
    bk.fml(ws, r, 2, f"=$B${r-1}*0.195", S.FMT_CUR); r += 1
    bk.lbl(ws, r, "Cash above floor at FY1 end (discretionary capacity)")
    bk.fml(ws, r, 2, f"=MAX(0,'Monthly x36'!N{M['cash_end']}-{A['floor']})", S.FMT_CUR); r += 1
    bk.lbl(ws, r, "Note: distributions are discretionary, subject to solvency, lender covenants, and the "
                  "S-corp pro-rata requirement; the OA mandates annual tax distributions.", italic=True)
    return ws


# =====================================================================
def build():
    bk = OB()
    assumptions(bk)
    census(bk)
    collections(bk)
    payroll(bk)
    monthly(bk)
    weekly(bk)
    avb(bk)
    investor(bk)
    bk.wb["Assumptions"].sheet_view.tabSelected = True
    out = "financial_models/output/Azalea_Refuge_Operating_Model.xlsx"
    bk.wb.save(out)
    import json, os
    rowmap = {sh: rows for sh, rows in bk.rows.items()}
    with open("financial_models/output/operating_model_rowmap.json", "w") as f:
        json.dump({"rows": rowmap, "addr": {k: v for k, v in bk.addr.items() if isinstance(v, str)}}, f, indent=1)
    print("wrote", out)
    return bk


if __name__ == "__main__":
    build()
