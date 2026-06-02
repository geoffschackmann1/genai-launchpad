"""Excel writers for the shared Azalea engine and the three workbooks.

Every period-based tab uses identical period columns (C..V) so same-column
cross-sheet links line up. Scalars live on the Inputs tab and are referenced
(no hardcoded business numbers in downstream formulas).
"""

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from . import styles as S
from .model import (
    PERIODS, NP, pcol, plet, FIRST_COL, LAST_LET, Y1_IDX, Y2_IDX, Y3_IDX,
    RATES, CENSUS_SCALARS, ADC_BASE, ADC_UPSIDE, ROSTER, PRN_VISIT, RHONDA_ANNUAL,
    RATIOS, BENEFITS, COGS_PD, GA_FIXED, GA_VAR, MED_DIRECTOR_PM,
    MED_DIRECTOR2_PM, MED_DIRECTOR2_START_M, CAPITAL,
    STARTUP, WORKING_CAP, ACTUALS, AVG_DAYS_MONTH,
)

DASH = '"-"'


# =========================================================================
# Book: a thin wrapper holding the workbook, the Inputs address registry,
# and styled-cell helpers.
# =========================================================================
class Book:
    def __init__(self, title):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.title = title
        self.addr = {}            # key -> "Inputs!$B$row"
        self.adc_row = None       # effective-ADC row on Inputs
        self.roster = []          # dicts: name, role, sal, start, grp, open
        self.prn = []             # dicts: label, vpm, rate, grp
        self.rows = {}            # sheet -> {rowkey: rownum} for cross-refs

    # ---- defined names ----
    def name(self, key, ref):
        try:
            self.wb.defined_names.add(DefinedName(key, attr_text=ref))
        except Exception:
            self.wb.defined_names[key] = DefinedName(key, attr_text=ref)

    # ---- styled cells ----
    def _set(self, ws, r, c, value, font, fmt=None, fill=None, align=None,
             border=None, comment=None):
        cell = ws.cell(r, c, value)
        cell.font = font
        if fmt: cell.number_format = fmt
        if fill: cell.fill = fill
        if align: cell.alignment = align
        if border: cell.border = border
        if comment: cell.comment = comment
        return cell

    def lbl(self, ws, r, text, c=1, bold=False, indent=0, italic=False):
        cell = ws.cell(r, c, text)
        cell.font = S.Font(name=S.BASE_FONT, size=10, bold=bold, italic=italic,
                           color=S.BLACK)
        cell.alignment = S.Alignment(horizontal="left", vertical="center", indent=indent)
        return cell

    def inp(self, ws, r, c, value, fmt, note=None):
        return self._set(ws, r, c, value, S.f_input(), fmt, S.fill(S.INPUTFILL),
                         S.RIGHT, S.BORDER_THIN, S.note(note) if note else None)

    def fml(self, ws, r, c, formula, fmt, link=False, bold=False, fill=None):
        return self._set(ws, r, c, formula, S.f_link(bold) if link else S.f_formula(bold),
                         fmt, fill, S.RIGHT)

    def title_block(self, ws, subtitle, span):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
        self._set(ws, 1, 1, "AZALEA HOSPICE & PALLIATIVE CARE — Tyler, TX",
                  S.f_title(), fill=S.fill(S.NAVY), align=S.LEFT)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
        self._set(ws, 2, 1, subtitle, S.f_subtitle(), fill=S.fill(S.NAVY), align=S.LEFT)
        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 14

    def section(self, ws, r, text, span):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
        self._set(ws, r, 1, text, S.f_section(), fill=S.fill(S.MIDBLUE), align=S.LEFT)

    def section_range(self, ws, r, text, c0, c1):
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
        self._set(ws, r, c0, text, S.f_section(), fill=S.fill(S.MIDBLUE), align=S.LEFT)

    def period_header(self, ws, r, label_text="($ per period)"):
        """Year band on row r, period labels on row r+1 (cols C..V)."""
        self._set(ws, r, 1, label_text, S.f_sub(), align=S.LEFT)
        # year bands
        bands = [("YEAR 1 — monthly", Y1_IDX), ("YEAR 2 — quarterly", Y2_IDX),
                 ("YEAR 3 — quarterly", Y3_IDX)]
        for name, idx in bands:
            c0, c1 = pcol(idx[0]), pcol(idx[-1])
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
            self._set(ws, r, c0, name, S.f_sub(), fill=S.fill(S.LIGHTBLUE), align=S.CENTER)
        for i, p in enumerate(PERIODS):
            self._set(ws, r + 1, pcol(i), p["label"], S.f_label(bold=True),
                      fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
        self._set(ws, r + 1, 1, "", S.f_label())
        ws.freeze_panes = ws.cell(r + 2, FIRST_COL)
        return r + 2   # first data row

    def widths(self, ws, label_w=42, b_w=14, period_w=11):
        ws.column_dimensions["A"].width = label_w
        ws.column_dimensions["B"].width = b_w
        for i in range(NP):
            ws.column_dimensions[get_column_letter(pcol(i))].width = period_w

    def months_mult(self, i):
        return 1 if PERIODS[i]["kind"] == "M" else 3


# =========================================================================
# INPUTS TAB
# =========================================================================
def inputs_tab(bk: Book):
    ws = bk.wb.create_sheet("Inputs")
    bk.widths(ws, label_w=46, b_w=15, period_w=11)
    bk.title_block(ws, "Assumptions Engine — CHOW relaunch of Hickory Hospice (San Antonio + Tyler alternative-delivery) · "
                       "blue = input · black = formula · green = cross-sheet link",
                   span=22)
    r = 4

    def scalar_block(title, rows_def):
        nonlocal r
        bk.section(ws, r, title, 22); r += 1
        for key, label, value, fmt, note in rows_def:
            bk.lbl(ws, r, label, indent=1)
            bk.inp(ws, r, 2, value, fmt, note)
            bk.addr[key] = f"Inputs!$B${r}"
            bk.name(key, f"Inputs!$B${r}")
            r += 1
        r += 1

    scalar_block("A.  GEOGRAPHY & MEDICARE RHC RATES (FY2026, 100% Medicare)", RATES)

    # derived rate build (formulas)
    bk.section(ws, r, "A2.  DERIVED RATE BUILD (formulas)", 22); r += 1
    bk.lbl(ws, r, "Wage-index adjustment factor", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['labor_share']}*{bk.addr['wage_index']}+(1-{bk.addr['labor_share']})", S.FMT_NUM4)
    bk.addr["wiadj"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Wage-adjusted rate, days 1-60", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['rhc_tier1']}*{bk.addr['wiadj']}", S.FMT_RATE)
    bk.addr["rate_t1"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Wage-adjusted rate, days 61+", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['rhc_tier2']}*{bk.addr['wiadj']}", S.FMT_RATE)
    bk.addr["rate_t2"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Blended GROSS rate ($/PD)", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['tier1_pct']}*{bk.addr['rate_t1']}+(1-{bk.addr['tier1_pct']})*{bk.addr['rate_t2']}", S.FMT_RATE)
    bk.addr["blended_gross"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Blended NET rate ($/PD)  [model]", indent=1, bold=True)
    bk.fml(ws, r, 2, f"={bk.addr['blended_gross']}*(1-{bk.addr['seq']}-{bk.addr['writeoff_pct']})", S.FMT_RATE, bold=True)
    bk.addr["net_rate"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Validation: model net rate vs actual $181", indent=1, italic=True)
    bk.fml(ws, r, 2, f"={bk.addr['net_rate']}-{bk.addr['blended_actual']}", S.FMT_RATE)
    ws.cell(r, 3, "should be near $0").font = S.f_note(); r += 2

    # timing scalar
    bk.section(ws, r, "B.  TIMING", 22); r += 1
    bk.lbl(ws, r, "Average days per month", indent=1)
    bk.inp(ws, r, 2, AVG_DAYS_MONTH, S.FMT_NUM1, "Used for patient-days; quarters = x3.")
    bk.addr["avg_days_month"] = f"Inputs!$B${r}"; bk.name("avg_days_month", f"Inputs!$B${r}"); r += 2

    # Census scenario scalars + paths
    bk.section(ws, r, "C.  CENSUS — PATIENT-MIGRATION RAMP (front-loaded M1-M2)", 22); r += 1
    for key, label, value, fmt, note in CENSUS_SCALARS:
        bk.lbl(ws, r, label, indent=1)
        bk.inp(ws, r, 2, value, fmt, note)
        bk.addr[key] = f"Inputs!$B${r}"; bk.name(key, f"Inputs!$B${r}")
        r += 1
    r += 1
    base_row = r
    bk.lbl(ws, r, "ADC path — BASE (100% capture, flat steady-state)", indent=1)
    for i in range(NP):
        bk.inp(ws, r, pcol(i), ADC_BASE[i], S.FMT_NUM1,
               "Migrated panel. M1=12 (~75% migrated), M2=19.8, M3+=22 steady. " + ("Source: Paloma P&L 2025." if i>=2 else ""))
    r += 1
    up_row = r
    bk.lbl(ws, r, "ADC path — UPSIDE (referral growth)", indent=1)
    for i in range(NP):
        bk.inp(ws, r, pcol(i), ADC_UPSIDE[i], S.FMT_NUM1, "Growth from Azalea's own referral pipeline.")
    r += 1
    bk.lbl(ws, r, "EFFECTIVE ADC  = chosen path x capture", indent=1, bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i),
               f"=IF({bk.addr['census_scenario']}=2,{c}{up_row},{c}{base_row})*{bk.addr['capture_rate']}",
               S.FMT_NUM1, bold=True, fill=S.fill(S.PALEBLUE))
    bk.adc_row = r; r += 2

    # Roster (Block D)
    bk.section(ws, r, "D.  STAGGERED FT ROSTER (named hires)", 22); r += 1
    hdr = ["Name", "Role", "Annual salary", "Start month", "Cost group", "Status"]
    for j, h in enumerate(hdr):
        bk._set(ws, r, 1 + j, h, S.f_sub(), fill=S.fill(S.LIGHTBLUE), align=S.LEFT, border=S.BORDER_THIN)
    r += 1
    for (name, role, sal, start, grp, openreq) in ROSTER:
        bk.lbl(ws, r, name, indent=1)
        bk.lbl(ws, r, role, c=2)
        bk.inp(ws, r, 3, sal, S.FMT_CUR, S.note("Source: Paloma payrolls Apr-May 2025.") and "Source: Paloma payrolls Apr-May 2025.")
        bk.inp(ws, r, 4, start, S.FMT_INT, "Model month this hire starts full-time.")
        bk.lbl(ws, r, "Direct (COGS)" if grp == "direct" else "Indirect (SG&A)", c=5)
        bk.lbl(ws, r, "OPEN REQ" if openreq else "Filled", c=6,
               italic=openreq)
        bk.roster.append(dict(name=name, role=role, sal=f"Inputs!$C${r}",
                              start=f"Inputs!$D${r}", grp=grp, open=openreq))
        r += 1
    bk.lbl(ws, r, "Rhonda Smith — Office Mgr (PRN, W-2)", indent=1)
    bk.lbl(ws, r, "Per-visit / part-time office support", c=2)
    bk.inp(ws, r, 3, RHONDA_ANNUAL, S.FMT_CUR, "PRN run-rate ($/yr). Source: Paloma payrolls.")
    bk.addr["rhonda_annual"] = f"Inputs!$C${r}"; r += 2

    # PRN visit roster
    bk.section(ws, r, "D2.  PRN / PER-VISIT ROSTER (census-driven)", 22); r += 1
    for j, h in enumerate(["Role", "", "Visits / pt / mo", "Rate / visit"]):
        bk._set(ws, r, 1 + j, h, S.f_sub(), fill=S.fill(S.LIGHTBLUE), align=S.LEFT, border=S.BORDER_THIN)
    r += 1
    for (key, label, vpm, rate, grp) in PRN_VISIT:
        bk.lbl(ws, r, label, indent=1)
        bk.inp(ws, r, 3, vpm, S.FMT_NUM1, "Visits per patient per month.")
        bk.inp(ws, r, 4, rate, S.FMT_CUR, "Per-visit pay rate. Source: Paloma payrolls.")
        bk.prn.append(dict(label=label, vpm=f"Inputs!$C${r}", rate=f"Inputs!$D${r}", grp=grp))
        r += 1
    r += 1

    scalar_block("E.  STAFFING RATIOS (Y2-Y3 capacity reference)", RATIOS)
    scalar_block("F.  BENEFITS & EMPLOYER BURDEN", BENEFITS)
    # med director scalars (1099 contracts — no benefits)
    bk.section(ws, r, "F2.  MEDICAL DIRECTORS (1099 — no benefits)", 22); r += 1
    for scalar_def in (MED_DIRECTOR_PM, MED_DIRECTOR2_PM, MED_DIRECTOR2_START_M):
        key, label, value, fmt, note = scalar_def
        bk.lbl(ws, r, label, indent=1); bk.inp(ws, r, 2, value, fmt, note)
        bk.addr[key] = f"Inputs!$B${r}"; bk.name(key, f"Inputs!$B${r}"); r += 1
    r += 1

    scalar_block("G.  PATIENT-RELATED COGS (per patient-day, trued to actuals)", COGS_PD)
    scalar_block("G2.  FIXED MONTHLY G&A (trued to actuals)", GA_FIXED)
    scalar_block("G3.  VARIABLE G&A", GA_VAR)
    scalar_block("H.  CAPITAL STRUCTURE (startup — NO acquisition)", CAPITAL)
    scalar_block("H2.  STARTUP ONE-TIME USES (Sources & Uses)", STARTUP)
    scalar_block("I.  WORKING CAPITAL", WORKING_CAP)

    # startup total + opening cash helpers
    bk.section(ws, r, "DERIVED CAPITAL HELPERS", 22); r += 1
    su_first = bk.addr[STARTUP[0][0]].split('$')[-1]
    su_last = bk.addr[STARTUP[-1][0]].split('$')[-1]
    bk.lbl(ws, r, "Total startup one-time uses", indent=1, bold=True)
    bk.fml(ws, r, 2, f"=SUM(B{su_first}:B{su_last})", S.FMT_CUR, bold=True)
    bk.addr["startup_total"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Opening cash = equity + loan - startup - capex", indent=1, bold=True)
    bk.fml(ws, r, 2, f"={bk.addr['equity']}+{bk.addr['sba_principal']}-{bk.addr['startup_total']}-{bk.addr['capex']}",
           S.FMT_CUR, bold=True)
    bk.addr["opening_cash"] = f"Inputs!$B${r}"; r += 1
    bk.lbl(ws, r, "Monthly D&A (capex+startup over 5 yr, license over 15 yr)", indent=1)
    bk.fml(ws, r, 2,
           f"=({bk.addr['capex']}+{bk.addr['startup_total']})/{bk.addr['deprec_yrs']}/12"
           f"+{bk.addr['license_cost']}/{bk.addr['license_amort_yrs']}/12", S.FMT_CUR)
    bk.addr["da_pm"] = f"Inputs!$B${r}"; r += 1

    ws.sheet_view.showGridLines = False
    return ws


# ---- small helpers for period tabs ----
def days_formula(bk, i):
    """Days-in-period referencing avg_days_month input."""
    mult = bk.months_mult(i)
    return f"={bk.addr['avg_days_month']}*{mult}"

def adc_ref(bk, i):
    return f"Inputs!{plet(i)}${bk.adc_row.__index__()}" if False else f"Inputs!{plet(i)}${bk.adc_row}"


# =========================================================================
# Period-tab helpers
# =========================================================================
def _rowmap(bk, sheet):
    bk.rows.setdefault(sheet, {})
    return bk.rows[sheet]

def _esc(bk, i):
    y = PERIODS[i]["year"]
    return "1" if y == 1 else f"(1+{bk.addr['merit']})^{y-1}"

def _link(sheet, col, row):
    sref = f"'{sheet}'" if " " in sheet or "&" in sheet else sheet
    return f"{sref}!{col}{row}"

def _ga_range(bk, mult):
    r1 = bk.addr[GA_FIXED[0][0]].rsplit('$', 1)[1]
    r2 = bk.addr[GA_FIXED[-1][0]].rsplit('$', 1)[1]
    return f"=SUM(Inputs!$B${r1}:$B${r2})*{mult}"


# =========================================================================
# REVENUE MODEL
# =========================================================================
def revenue_tab(bk: Book):
    ws = bk.wb.create_sheet("Revenue Model")
    bk.widths(ws); bk.title_block(ws, "Revenue Model — 100% Medicare RHC · accrual basis", 22)
    bk.section(ws, 4, "RHC RATE (per patient-day) — linked from Inputs", 22)
    bk.lbl(ws, 5, "Blended GROSS rate ($/PD)", indent=1)
    bk.fml(ws, 5, 2, f"={bk.addr['blended_gross']}", S.FMT_RATE, link=True)
    bk.lbl(ws, 6, "Blended NET rate ($/PD)", indent=1)
    bk.fml(ws, 6, 2, f"={bk.addr['net_rate']}", S.FMT_RATE, link=True)
    dr = bk.period_header(ws, 8)
    rm = _rowmap(bk, "Revenue Model")
    r = dr
    bk.lbl(ws, r, "Average daily census (ADC)", bold=False)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"=Inputs!{plet(i)}${bk.adc_row}", S.FMT_NUM1, link=True)
    rm["adc"] = r; r += 1
    bk.lbl(ws, r, "Days in period")
    for i in range(NP):
        bk.fml(ws, r, pcol(i), days_formula(bk, i), S.FMT_NUM1, link=True)
    rm["days"] = r; r += 1
    bk.lbl(ws, r, "Patient-days", bold=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['adc']}*{plet(i)}{rm['days']}", S.FMT_NUM, bold=True)
    rm["pd"] = r; r += 1
    bk.lbl(ws, r, "Rate escalation factor (Y1 = 1.000, Y2+ compounded)", italic=True)
    for i in range(NP):
        yr = PERIODS[i]["year"]
        f = "1" if yr == 1 else f"(1+{bk.addr['rhc_escalation']})^{yr-1}"
        bk.fml(ws, r, pcol(i), f"={f}", S.FMT_NUM4)
    rm["rate_factor"] = r; r += 1
    bk.lbl(ws, r, "Gross Medicare revenue")
    for i in range(NP):
        bk.fml(ws, r, pcol(i),
               f"={plet(i)}{rm['pd']}*{bk.addr['blended_gross']}*{plet(i)}{rm['rate_factor']}",
               S.FMT_CUR)
    rm["gross"] = r; r += 1
    bk.lbl(ws, r, "Less: sequestration", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"=-{plet(i)}{rm['gross']}*{bk.addr['seq']}", S.FMT_CUR)
    r += 1
    bk.lbl(ws, r, "Less: write-offs", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"=-{plet(i)}{rm['gross']}*{bk.addr['writeoff_pct']}", S.FMT_CUR)
    r += 1
    bk.lbl(ws, r, "NET REVENUE", bold=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['gross']}*(1-{bk.addr['seq']}-{bk.addr['writeoff_pct']})",
               S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    rm["net"] = r; r += 2
    bk.lbl(ws, r, "Memo: net revenue per patient-day", italic=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"=IF({plet(i)}{rm['pd']}=0,0,{plet(i)}{rm['net']}/{plet(i)}{rm['pd']})", S.FMT_RATE)
    r += 1
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# STAFFING & PAYROLL
# =========================================================================
def staffing_tab(bk: Book):
    ws = bk.wb.create_sheet("Staffing & Payroll")
    bk.widths(ws, label_w=44); bk.title_block(ws,
        "Staffing & Payroll — staggered hires + census-driven PRN · burden on W-2 only · Med Director 1099 (no benefits)", 22)
    dr = bk.period_header(ws, 4)
    rm = _rowmap(bk, "Staffing & Payroll")
    r = dr
    bk.lbl(ws, r, "ADC (driver)", italic=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Revenue Model", plet(i), bk.rows["Revenue Model"]["adc"]), S.FMT_NUM1, link=True)
    rm["adc"] = r; r += 1

    # FT salaried
    bk.section(ws, r, "DIRECT-CARE FT SALARIES (Cost of Care)", 22); r += 1
    direct_ft_rows = []; indirect_ft_rows = []
    for h in bk.roster:
        flag = " [OPEN REQ]" if h["open"] else ""
        bk.lbl(ws, r, f"{h['name']} — {h['role']}{flag}", indent=1, italic=h["open"])
        for i in range(NP):
            mi = PERIODS[i]["month_index"]; mult = bk.months_mult(i)
            bk.fml(ws, r, pcol(i),
                   f"=IF({mi}>={h['start']},{h['sal']}/12*{mult}*{_esc(bk,i)},0)",
                   S.FMT_CUR)
        if h["grp"] == "direct":
            direct_ft_rows.append(r)
        else:
            indirect_ft_rows.append(r)
        h["row"] = r
        r += 1
    bk.lbl(ws, r, "Subtotal — direct-care FT salaries", bold=True)
    for i in range(NP):
        terms = "+".join(f"{plet(i)}{rr}" for rr in direct_ft_rows)
        bk.fml(ws, r, pcol(i), f"={terms}", S.FMT_CUR, bold=True)
    rm["ft_direct"] = r; r += 2

    bk.section(ws, r, "PRN / PER-VISIT DIRECT CARE (census-driven)", 22); r += 1
    prn_rows = []
    for p in bk.prn:
        bk.lbl(ws, r, p["label"], indent=1)
        for i in range(NP):
            mult = bk.months_mult(i)
            bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['adc']}*{p['vpm']}*{p['rate']}*{mult}", S.FMT_CUR)
        prn_rows.append(r); r += 1
    bk.lbl(ws, r, "Subtotal — PRN visit wages", bold=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + "+".join(f"{plet(i)}{rr}" for rr in prn_rows), S.FMT_CUR, bold=True)
    rm["prn"] = r; r += 1
    bk.lbl(ws, r, "Medical Director 1 + 2 (1099, MD2 conditional)", indent=1)
    for i in range(NP):
        mi = PERIODS[i]["month_index"]; mm = bk.months_mult(i)
        f = (f"={bk.addr['med_director']}*{mm}"
             f"+IF({mi}>={bk.addr['med_director2_start']},{bk.addr['med_director2']}*{mm},0)")
        bk.fml(ws, r, pcol(i), f, S.FMT_CUR)
    rm["med_dir"] = r; r += 2

    bk.section(ws, r, "INDIRECT LABOR (SG&A overhead)", 22); r += 1
    bk.lbl(ws, r, "Subtotal — indirect FT salaries", bold=True)
    for i in range(NP):
        terms = "+".join(f"{plet(i)}{rr}" for rr in indirect_ft_rows)
        bk.fml(ws, r, pcol(i), f"={terms}", S.FMT_CUR, bold=True)
    rm["ft_indirect"] = r; r += 1
    bk.lbl(ws, r, "Rhonda Smith — Office Mgr (PRN, W-2)", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={bk.addr['rhonda_annual']}/12*{bk.months_mult(i)}*{_esc(bk,i)}", S.FMT_CUR)
    rm["rhonda"] = r; r += 2

    # benefits
    bk.section(ws, r, "EMPLOYER BURDEN & BENEFITS (W-2 only)", 22); r += 1
    bk.lbl(ws, r, "Direct W-2 wage base (FT direct + PRN)", italic=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['ft_direct']}+{plet(i)}{rm['prn']}", S.FMT_CUR)
    rm["w2_d"] = r; r += 1
    bk.lbl(ws, r, "Indirect W-2 wage base (FT indirect + Rhonda)", italic=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['ft_indirect']}+{plet(i)}{rm['rhonda']}", S.FMT_CUR)
    rm["w2_i"] = r; r += 1
    bk.lbl(ws, r, "Burden — direct (load %)", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['w2_d']}*{bk.addr['benefits_load']}", S.FMT_CUR)
    rm["burden_d"] = r; r += 1
    bk.lbl(ws, r, "Burden — indirect (load %)", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['w2_i']}*{bk.addr['benefits_load']}", S.FMT_CUR)
    rm["burden_i"] = r; r += 1
    # health: count active FT by group
    d_starts = [h["start"] for h in bk.roster if h["grp"] == "direct"]
    i_starts = [h["start"] for h in bk.roster if h["grp"] == "indirect"]
    bk.lbl(ws, r, "Health insurance — direct FT", indent=1)
    for i in range(NP):
        mi = PERIODS[i]["month_index"]; mult = bk.months_mult(i)
        cnt = "+".join(f"IF({mi}>={s},1,0)" for s in d_starts)
        bk.fml(ws, r, pcol(i), f"={bk.addr['health_pm']}*({cnt})*{mult}", S.FMT_CUR)
    rm["health_d"] = r; r += 1
    bk.lbl(ws, r, "Health insurance — indirect FT", indent=1)
    for i in range(NP):
        mi = PERIODS[i]["month_index"]; mult = bk.months_mult(i)
        cnt = "+".join(f"IF({mi}>={s},1,0)" for s in i_starts)
        bk.fml(ws, r, pcol(i), f"={bk.addr['health_pm']}*({cnt})*{mult}", S.FMT_CUR)
    rm["health_i"] = r; r += 2

    # totals
    bk.lbl(ws, r, "TOTAL DIRECT-CARE LABOR (COGS)", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i),
               f"={c}{rm['ft_direct']}+{c}{rm['prn']}+{c}{rm['med_dir']}+{c}{rm['burden_d']}+{c}{rm['health_d']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    rm["cogs_labor"] = r; r += 1
    bk.lbl(ws, r, "TOTAL INDIRECT LABOR (SG&A)", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i),
               f"={c}{rm['ft_indirect']}+{c}{rm['rhonda']}+{c}{rm['burden_i']}+{c}{rm['health_i']}",
               S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    rm["sga_labor"] = r; r += 1
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# OPERATING BUDGET  (monthly P&L down to Net Income)
# =========================================================================
def opbudget_tab(bk: Book):
    ws = bk.wb.create_sheet("Operating Budget")
    bk.widths(ws, label_w=44); bk.title_block(ws,
        "Operating Budget — P&L · Net Rev − COGS = Gross Profit − SG&A = EBITDA − D&A − Interest = Net Income", 22)
    dr = bk.period_header(ws, 4)
    rm = _rowmap(bk, "Operating Budget")
    rev = bk.rows["Revenue Model"]; stf = bk.rows["Staffing & Payroll"]
    r = dr
    bk.lbl(ws, r, "NET REVENUE", bold=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Revenue Model", plet(i), rev["net"]), S.FMT_CUR, link=True, bold=True)
    rm["net"] = r; r += 2
    bk.section(ws, r, "COST OF SERVICES", 22); r += 1
    bk.lbl(ws, r, "Direct-care labor", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Staffing & Payroll", plet(i), stf["cogs_labor"]), S.FMT_CUR, link=True)
    rm["cogs_labor"] = r; r += 1
    for key, lab in (("supplies_pd", "Medical supplies"), ("dme_pd", "DME"), ("pharmacy_pd", "Pharmacy")):
        bk.lbl(ws, r, lab, indent=1)
        for i in range(NP):
            bk.fml(ws, r, pcol(i), f"={_link('Revenue Model', plet(i), rev['pd'])}*{bk.addr[key]}", S.FMT_CUR, link=True)
        rm[key] = r; r += 1
    bk.lbl(ws, r, "Patient-related COGS subtotal", italic=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['supplies_pd']}+{c}{rm['dme_pd']}+{c}{rm['pharmacy_pd']}", S.FMT_CUR)
    rm["cogs_patient"] = r; r += 1
    bk.lbl(ws, r, "TOTAL COST OF SERVICES", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['cogs_labor']}+{c}{rm['cogs_patient']}", S.FMT_CUR, bold=True)
    rm["cogs"] = r; r += 1
    bk.lbl(ws, r, "GROSS PROFIT", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['net']}-{c}{rm['cogs']}", S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    rm["gp"] = r; r += 1
    bk.lbl(ws, r, "Gross margin %", italic=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"=IF({c}{rm['net']}=0,0,{c}{rm['gp']}/{c}{rm['net']})", S.FMT_PCT)
    r += 2
    bk.section(ws, r, "SELLING, GENERAL & ADMINISTRATIVE", 22); r += 1
    bk.lbl(ws, r, "Indirect labor", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Staffing & Payroll", plet(i), stf["sga_labor"]), S.FMT_CUR, link=True)
    rm["sga_labor"] = r; r += 1
    bk.lbl(ws, r, "Fixed G&A (utilities, EMR, billing, rent, mktg…)", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), _ga_range(bk, bk.months_mult(i)), S.FMT_CUR, link=True)
    rm["ga_fixed"] = r; r += 1
    bk.lbl(ws, r, "Outsourced billing fee (1.5% of gross)", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={_link('Revenue Model', plet(i), rev['gross'])}*{bk.addr['billing_fee_pct']}", S.FMT_CUR, link=True)
    rm["billing"] = r; r += 1
    bk.lbl(ws, r, "QR payment fee (% of gross)", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={_link('Revenue Model', plet(i), rev['gross'])}*{bk.addr['qr_fee_pct']}", S.FMT_CUR, link=True)
    rm["qr"] = r; r += 1
    bk.lbl(ws, r, "TOTAL SG&A", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['sga_labor']}+{c}{rm['ga_fixed']}+{c}{rm['billing']}+{c}{rm['qr']}", S.FMT_CUR, bold=True)
    rm["sga"] = r; r += 1
    bk.lbl(ws, r, "EBITDA", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['gp']}-{c}{rm['sga']}", S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    rm["ebitda"] = r; r += 1
    bk.lbl(ws, r, "EBITDA margin %", italic=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"=IF({c}{rm['net']}=0,0,{c}{rm['ebitda']}/{c}{rm['net']})", S.FMT_PCT)
    rm["ebitda_margin"] = r; r += 2
    bk.section(ws, r, "BELOW EBITDA", 22); r += 1
    bk.lbl(ws, r, "Depreciation & amortization", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={bk.addr['da_pm']}*{bk.months_mult(i)}", S.FMT_CUR, link=True)
    rm["da"] = r; r += 1
    bk.lbl(ws, r, "Interest — SBA loan", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=0", S.FMT_CUR, link=True)   # patched after Debt tab
    rm["int_sba"] = r; r += 1
    bk.lbl(ws, r, "Interest — Hickory license note", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=0", S.FMT_CUR, link=True)   # patched after Debt tab
    rm["int_license"] = r; r += 1
    bk.lbl(ws, r, "Interest — working-capital line", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=0", S.FMT_CUR, link=True)   # patched after Cash Flow tab
    rm["int_loc"] = r; r += 1
    bk.lbl(ws, r, "PRE-TAX INCOME", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i),
               f"={c}{rm['ebitda']}-{c}{rm['da']}-{c}{rm['int_sba']}-{c}{rm['int_license']}-{c}{rm['int_loc']}",
               S.FMT_CUR, bold=True)
    rm["pretax"] = r; r += 1
    bk.lbl(ws, r, "TX franchise / margin tax", indent=1)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"=IF({c}{rm['pretax']}>0,{c}{rm['net']}*{bk.addr['tx_tax']},0)", S.FMT_CUR)
    rm["tax"] = r; r += 1
    bk.lbl(ws, r, "NET INCOME", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['pretax']}-{c}{rm['tax']}", S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    rm["ni"] = r; r += 1
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# DEBT SCHEDULE (SBA 7(a) at period grain) + DSCR
# =========================================================================
def debt_tab(bk: Book):
    ws = bk.wb.create_sheet("Debt Schedule")
    bk.widths(ws, label_w=40); bk.title_block(ws,
        "Debt Schedule — SBA 7(a) + Hickory license seller note · Combined DSCR shown (target ≥ 1.25x)", 22)
    bk.section(ws, 4, "LOAN TERMS (from Inputs)", 22)
    bk.lbl(ws, 5, "SBA monthly payment (PMT)", indent=1)
    bk.fml(ws, 5, 2, f"=PMT({bk.addr['sba_rate']}/12,{bk.addr['sba_term_mo']},-{bk.addr['sba_principal']})", S.FMT_CUR)
    bk.lbl(ws, 6, "License note monthly payment (PMT)", indent=1)
    bk.fml(ws, 6, 2, f"=PMT({bk.addr['license_rate']}/12,{bk.addr['license_term']},-{bk.addr['license_cost']})", S.FMT_CUR)
    SBA_PMT = "'Debt Schedule'!$B$5"
    LIC_PMT = "'Debt Schedule'!$B$6"
    dr = bk.period_header(ws, 8)
    rm = _rowmap(bk, "Debt Schedule"); ob = bk.rows["Operating Budget"]
    r = dr

    # ---------------- SBA 7(a) ----------------
    bk.section(ws, r, "SBA 7(a) LOAN", 22); r += 1
    bk.lbl(ws, r, "Beginning balance")
    for i in range(NP):
        if i == 0:
            bk.fml(ws, r, pcol(i), f"={bk.addr['sba_principal']}", S.FMT_CUR, link=True)
        else:
            bk.fml(ws, r, pcol(i), f"={plet(i-1)}{r+4}", S.FMT_CUR)
    rm["beg"] = r; r += 1
    bk.lbl(ws, r, "Payment", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={SBA_PMT}*{bk.months_mult(i)}", S.FMT_CUR)
    rm["pmt"] = r; r += 1
    bk.lbl(ws, r, "Interest", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['beg']}*{bk.addr['sba_rate']}/12*{bk.months_mult(i)}", S.FMT_CUR)
    rm["int"] = r; r += 1
    bk.lbl(ws, r, "Principal", indent=1)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['pmt']}-{c}{rm['int']}", S.FMT_CUR)
    rm["prin"] = r; r += 1
    bk.lbl(ws, r, "Ending balance", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['beg']}-{c}{rm['prin']}", S.FMT_CUR, bold=True)
    rm["end"] = r; r += 2

    # ---------------- Hickory license seller note (closed-form amort) -----
    bk.section(ws, r, "HICKORY LICENSE SELLER NOTE  ($300K, 6%, 36 mo)", 22); r += 1
    bk.lbl(ws, r, "Beginning balance")
    for i in range(NP):
        if i == 0:
            bk.fml(ws, r, pcol(i), f"={bk.addr['license_cost']}", S.FMT_CUR, link=True)
        else:
            # prior ending balance — Ending row is r+4 below
            bk.fml(ws, r, pcol(i), f"={plet(i-1)}{r+4}", S.FMT_CUR)
    rm["lic_beg"] = r; r += 1
    bk.lbl(ws, r, "Ending balance (closed-form)", indent=1)
    # B*(1+r)^m − pmt*((1+r)^m − 1)/r, floored at 0
    for i in range(NP):
        c = plet(i); m = bk.months_mult(i)
        f = (f"=MAX(0,{c}{r-1}*(1+{bk.addr['license_rate']}/12)^{m}"
             f"-{LIC_PMT}*((1+{bk.addr['license_rate']}/12)^{m}-1)/({bk.addr['license_rate']}/12))")
        bk.fml(ws, r, pcol(i), f, S.FMT_CUR)
    rm["lic_end_calc"] = r; r += 1
    bk.lbl(ws, r, "Principal", indent=1)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['lic_beg']}-{c}{rm['lic_end_calc']}", S.FMT_CUR)
    rm["lic_prin"] = r; r += 1
    bk.lbl(ws, r, "Interest", indent=1)
    # interest = pmt*m − principal (clamped, never negative; 0 once balance = 0)
    for i in range(NP):
        c = plet(i)
        f = f"=MAX(0,IF({c}{rm['lic_beg']}=0,0,{LIC_PMT}*{bk.months_mult(i)}-{c}{rm['lic_prin']}))"
        bk.fml(ws, r, pcol(i), f, S.FMT_CUR)
    rm["lic_int"] = r; r += 1
    bk.lbl(ws, r, "Ending balance", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['lic_end_calc']}", S.FMT_CUR, bold=True)
    rm["lic_end"] = r; r += 2

    # ---------------- Combined coverage ----------------
    bk.section(ws, r, "COMBINED DEBT SERVICE & COVERAGE", 22); r += 1
    bk.lbl(ws, r, "Total interest", italic=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['int']}+{c}{rm['lic_int']}", S.FMT_CUR)
    rm["total_int"] = r; r += 1
    bk.lbl(ws, r, "Total principal", italic=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['prin']}+{c}{rm['lic_prin']}", S.FMT_CUR)
    rm["total_prin"] = r; r += 1
    bk.lbl(ws, r, "TOTAL DEBT SERVICE (P+I, combined)", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['total_int']}+{c}{rm['total_prin']}", S.FMT_CUR, bold=True)
    rm["ds"] = r; r += 1
    bk.lbl(ws, r, "EBITDA (link)", italic=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Operating Budget", plet(i), ob["ebitda"]), S.FMT_CUR, link=True)
    rm["ebitda"] = r; r += 1
    bk.lbl(ws, r, "COMBINED DSCR (period)", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"=IF({c}{rm['ds']}=0,0,{c}{rm['ebitda']}/{c}{rm['ds']})", S.FMT_MULT, bold=True,
               fill=S.fill(S.LIGHTBLUE))
    rm["dscr"] = r; r += 1
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# CASH FLOW & BALANCE SHEET  (col B = opening; C..V = periods)
# LOC interest uses PRIOR-period balance -> no circular reference.
# =========================================================================
def cashflow_tab(bk: Book):
    ws = bk.wb.create_sheet("Cash Flow & BS")
    bk.widths(ws, label_w=42, b_w=13); bk.title_block(ws,
        "Cash Flow (indirect) & Balance Sheet — BS balances through real AR/AP/LOC, never a plug", 23)
    ob = bk.rows["Operating Budget"]; rev = bk.rows["Revenue Model"]; dbt = bk.rows["Debt Schedule"]
    # header with opening column
    self_r = 4
    bk._set(ws, self_r, 2, "Opening", S.f_sub(), fill=S.fill(S.LIGHTBLUE), align=S.CENTER)
    dr = bk.period_header(ws, self_r)
    rm = _rowmap(bk, "Cash Flow & BS")
    OPEN = "B"
    r = dr
    bk.section(ws, r, "CASH FLOW (indirect method)", 23); r += 1
    bk.lbl(ws, r, "Net income", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Operating Budget", plet(i), ob["ni"]), S.FMT_CUR, link=True)
    rm["ni"] = r; r += 1
    bk.lbl(ws, r, "+ Depreciation & amortization", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), "=" + _link("Operating Budget", plet(i), ob["da"]), S.FMT_CUR, link=True)
    rm["da"] = r; r += 1
    # AR / AP balances (helper rows, used for deltas + BS)
    bk.lbl(ws, r, "Accounts receivable (balance)", indent=1, italic=True)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={_link('Revenue Model', plet(i), rev['net'])}*{bk.addr['ar_days']}/{_link('Revenue Model', plet(i), rev['days'])}", S.FMT_CUR, link=True)
    rm["ar"] = r; r += 1
    bk.lbl(ws, r, "Accounts payable (balance)", indent=1, italic=True)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        c = plet(i)
        opex = f"({_link('Operating Budget', c, ob['ga_fixed'])}+{_link('Operating Budget', c, ob['qr'])}+{_link('Operating Budget', c, ob['cogs_patient'])})"
        bk.fml(ws, r, pcol(i), f"={opex}*{bk.addr['ap_days']}/{_link('Revenue Model', c, rev['days'])}", S.FMT_CUR, link=True)
    rm["ap"] = r; r += 1
    bk.lbl(ws, r, "− Δ Accounts receivable", indent=1)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i), f"=-({plet(i)}{rm['ar']}-{prev}{rm['ar']})", S.FMT_CUR)
    rm["dar"] = r; r += 1
    bk.lbl(ws, r, "+ Δ Accounts payable", indent=1)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['ap']}-{prev}{rm['ap']}", S.FMT_CUR)
    rm["dap"] = r; r += 1
    bk.lbl(ws, r, "Cash flow from operations", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['ni']}+{c}{rm['da']}+{c}{rm['dar']}+{c}{rm['dap']}", S.FMT_CUR, bold=True)
    rm["cfo"] = r; r += 1
    bk.lbl(ws, r, "− SBA principal repayment", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"=-{_link('Debt Schedule', plet(i), dbt['prin'])}", S.FMT_CUR, link=True)
    rm["prin"] = r; r += 1
    bk.lbl(ws, r, "− Hickory license note principal", indent=1)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"=-{_link('Debt Schedule', plet(i), dbt['lic_prin'])}", S.FMT_CUR, link=True)
    rm["lic_prin"] = r; r += 1

    # LOC interest (prior balance) -- referenced by Operating Budget
    bk.lbl(ws, r, "Working-capital line interest", indent=1, italic=True)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i), f"={prev}{r+4}*{bk.addr['loc_rate']}*{bk.months_mult(i)}/12", S.FMT_CUR)  # prior LOC balance (locbal = locint + 4)
    rm["locint"] = r; LOCINT_ROW = r; r += 1

    # cash before LOC, draw/repay, ending cash, LOC balance
    bk.lbl(ws, r, "Beginning cash", indent=1)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        prev = f"={bk.addr['opening_cash']}" if i == 0 else f"={plet(i-1)}{r+4}"  # prior endcash (begcash + 4)
        bk.fml(ws, r, pcol(i), prev, S.FMT_CUR, link=(i == 0))
    rm["begcash"] = r; r += 1
    bk.lbl(ws, r, "Cash before LOC sweep", italic=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['begcash']}+{c}{rm['cfo']}+{c}{rm['prin']}+{c}{rm['lic_prin']}", S.FMT_CUR)
    rm["precash"] = r; r += 1
    bk.lbl(ws, r, "LOC draw / (repay)", indent=1)
    for i in range(NP):
        c = plet(i); prevloc = OPEN if i == 0 else plet(i-1)
        # draw if below floor; else repay down to floor up to outstanding prior balance
        f = (f"=IF({c}{rm['precash']}<{bk.addr['min_cash']},{bk.addr['min_cash']}-{c}{rm['precash']},"
             f"-MIN({prevloc}{r+1},MAX(0,{c}{rm['precash']}-{bk.addr['min_cash']})))")
        bk.fml(ws, r, pcol(i), f, S.FMT_CUR)
    rm["locflow"] = r; r += 1
    bk.lbl(ws, r, "Working-capital line balance", bold=True)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        c = plet(i); prevloc = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i), f"={prevloc}{r}+{c}{rm['locflow']}", S.FMT_CUR, bold=True)
    rm["locbal"] = r; r += 1
    bk.lbl(ws, r, "ENDING CASH", bold=True)
    for i in range(NP):
        c = plet(i)
        bk.fml(ws, r, pcol(i), f"={c}{rm['precash']}+{c}{rm['locflow']}", S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    rm["endcash"] = r; r += 2

    # ---- Balance Sheet ----
    bk.section(ws, r, "BALANCE SHEET", 23); r += 1
    bk.lbl(ws, r, "Cash", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['opening_cash']}", S.FMT_CUR, link=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['endcash']}", S.FMT_CUR)
    rm["bs_cash"] = r; r += 1
    bk.lbl(ws, r, "Accounts receivable", indent=1)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['ar']}", S.FMT_CUR)
    rm["bs_ar"] = r; r += 1
    # Three asset pools, each amortized over its own life
    bk.lbl(ws, r, "PP&E, net", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['capex']}", S.FMT_CUR, link=True)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i),
               f"=MAX(0,{prev}{r}-{bk.addr['capex']}/{bk.addr['deprec_yrs']}/12*{bk.months_mult(i)})",
               S.FMT_CUR)
    rm["bs_ppe"] = r; r += 1
    bk.lbl(ws, r, "Startup & organizational costs, net", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['startup_total']}", S.FMT_CUR, link=True)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i),
               f"=MAX(0,{prev}{r}-{bk.addr['startup_total']}/{bk.addr['deprec_yrs']}/12*{bk.months_mult(i)})",
               S.FMT_CUR)
    rm["bs_intang"] = r; r += 1
    bk.lbl(ws, r, "Hickory Medicare license intangible, net", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['license_cost']}", S.FMT_CUR, link=True)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i),
               f"=MAX(0,{prev}{r}-{bk.addr['license_cost']}/{bk.addr['license_amort_yrs']}/12*{bk.months_mult(i)})",
               S.FMT_CUR)
    rm["bs_license"] = r; r += 1
    bk.lbl(ws, r, "TOTAL ASSETS", bold=True)
    for col in [OPEN] + [plet(i) for i in range(NP)]:
        ws[f"{col}{r}"] = (f"={col}{rm['bs_cash']}+{col}{rm['bs_ar']}+{col}{rm['bs_ppe']}"
                          f"+{col}{rm['bs_intang']}+{col}{rm['bs_license']}")
        ws[f"{col}{r}"].number_format = S.FMT_CUR; ws[f"{col}{r}"].font = S.f_total()
    rm["bs_assets"] = r; r += 2

    bk.lbl(ws, r, "Accounts payable", indent=1)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['ap']}", S.FMT_CUR)
    rm["bs_ap"] = r; r += 1
    bk.lbl(ws, r, "Working-capital line", indent=1)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={plet(i)}{rm['locbal']}", S.FMT_CUR)
    rm["bs_loc"] = r; r += 1
    bk.lbl(ws, r, "SBA loan balance", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['sba_principal']}", S.FMT_CUR, link=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={_link('Debt Schedule', plet(i), dbt['end'])}", S.FMT_CUR, link=True)
    rm["bs_sba"] = r; r += 1
    bk.lbl(ws, r, "Hickory license note balance", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['license_cost']}", S.FMT_CUR, link=True)
    for i in range(NP):
        bk.fml(ws, r, pcol(i), f"={_link('Debt Schedule', plet(i), dbt['lic_end'])}", S.FMT_CUR, link=True)
    rm["bs_lic"] = r; r += 1
    bk.lbl(ws, r, "Paid-in equity", indent=1)
    for col in [OPEN] + [plet(i) for i in range(NP)]:
        ws[f"{col}{r}"] = f"={bk.addr['equity']}"
        ws[f"{col}{r}"].number_format = S.FMT_CUR; ws[f"{col}{r}"].font = S.f_link()
    rm["bs_equity"] = r; r += 1
    bk.lbl(ws, r, "Retained earnings", indent=1)
    bk.fml(ws, r, 2, "=0", S.FMT_CUR)
    for i in range(NP):
        prev = OPEN if i == 0 else plet(i-1)
        bk.fml(ws, r, pcol(i), f"={prev}{r}+{plet(i)}{rm['ni']}", S.FMT_CUR)
    rm["bs_re"] = r; r += 1
    bk.lbl(ws, r, "TOTAL LIABILITIES & EQUITY", bold=True)
    for col in [OPEN] + [plet(i) for i in range(NP)]:
        ws[f"{col}{r}"] = (f"={col}{rm['bs_ap']}+{col}{rm['bs_loc']}+{col}{rm['bs_sba']}"
                          f"+{col}{rm['bs_lic']}+{col}{rm['bs_equity']}+{col}{rm['bs_re']}")
        ws[f"{col}{r}"].number_format = S.FMT_CUR; ws[f"{col}{r}"].font = S.f_total()
    rm["bs_le"] = r; r += 1
    bk.lbl(ws, r, "CHECK: Assets − (L+E)  → 0", bold=True)
    for col in [OPEN] + [plet(i) for i in range(NP)]:
        cell = ws[f"{col}{r}"]
        cell.value = f"={col}{rm['bs_assets']}-{col}{rm['bs_le']}"
        cell.number_format = S.FMT_CUR; cell.font = S.f_total(); cell.fill = S.fill(S.GREENFILL)
    rm["bs_check"] = r; r += 1
    ws.sheet_view.showGridLines = False
    return ws


def patch_pl_interest(bk: Book):
    """Fill Operating Budget interest rows now that Debt & Cash Flow exist."""
    ws = bk.wb["Operating Budget"]
    ob = bk.rows["Operating Budget"]; dbt = bk.rows["Debt Schedule"]; cf = bk.rows["Cash Flow & BS"]
    for i in range(NP):
        ws.cell(ob["int_sba"], pcol(i)).value = "=" + _link("Debt Schedule", plet(i), dbt["int"])
        ws.cell(ob["int_license"], pcol(i)).value = "=" + _link("Debt Schedule", plet(i), dbt["lic_int"])
        ws.cell(ob["int_loc"], pcol(i)).value = "=" + _link("Cash Flow & BS", plet(i), cf["locint"])


# =========================================================================
# ACTUALS vs MODEL  (Paloma Mar-Jun 2025 reconciliation)
# =========================================================================
def actuals_tab(bk: Book):
    ws = bk.wb.create_sheet("Actuals vs Model")
    ws.column_dimensions["A"].width = 38
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 14
    bk.title_block(ws, "Paloma Tyler actuals (Mar–Jun 2025) vs Azalea model steady-state — proof the migrated panel is a proven, EBITDA-positive book", 8)
    A = ACTUALS
    hdr_row = 4
    heads = ["Metric"] + A["months"] + ["Apr–Jun avg", "Model (M3)", "Var %"]
    for j, h in enumerate(heads):
        bk._set(ws, hdr_row, 1 + j, h, S.f_label(bold=True), fill=S.fill(S.GREYHDR),
                align=S.CENTER if j else S.LEFT, border=S.BORDER_THIN)
    ws.freeze_panes = "B5"
    r = hdr_row + 1
    ob = bk.rows["Operating Budget"]; rev = bk.rows["Revenue Model"]
    M3 = plet(2)  # model month 3 = steady state

    def actual_row(label, key, fmt=S.FMT_CUR, note=None):
        nonlocal r
        bk.lbl(ws, r, label, indent=1)
        for j, v in enumerate(A[key]):
            bk.inp(ws, r, 2 + j, v, fmt, note)
        # Apr-Jun avg = Apr(C), May(D), Jun(E)
        bk.fml(ws, r, 6, f"=AVERAGE(C{r}:E{r})", fmt)
        return r

    # Gross
    gr = actual_row("Gross revenue", "gross", note="Source: Paloma P&L 2025.")
    bk.fml(ws, gr, 7, f"={_link('Revenue Model', M3, rev['gross'])}", S.FMT_CUR, link=True)
    bk.fml(ws, gr, 8, f"=IF(F{gr}=0,0,(G{gr}-F{gr})/F{gr})", S.FMT_PCT); r += 1
    wo = actual_row("Less: write-offs", "writeoff", note="Source: Paloma P&L 2025."); r += 1
    sq = actual_row("Less: sequestration", "seq", note="Source: Paloma P&L 2025."); r += 1
    # net revenue (formula from actual rows)
    bk.lbl(ws, r, "NET REVENUE", bold=True)
    for j in range(4):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"={cl}{gr}-{cl}{wo}-{cl}{sq}", S.FMT_CUR, bold=True)
    bk.fml(ws, r, 6, f"=AVERAGE(C{r}:E{r})", S.FMT_CUR, bold=True)
    bk.fml(ws, r, 7, f"={_link('Revenue Model', M3, rev['net'])}", S.FMT_CUR, link=True, bold=True)
    bk.fml(ws, r, 8, f"=IF(F{r}=0,0,(G{r}-F{r})/F{r})", S.FMT_PCT, bold=True)
    net_row = r; r += 1
    # implied ADC
    bk.lbl(ws, r, "Implied ADC (net ÷ $181 ÷ days)", italic=True)
    for j in range(4):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"={cl}{net_row}/{bk.addr['blended_actual']}/{A['days'][j]}", S.FMT_NUM1)
    bk.fml(ws, r, 6, f"=AVERAGE(C{r}:E{r})", S.FMT_NUM1)
    bk.fml(ws, r, 7, f"={_link('Revenue Model', M3, rev['adc'])}", S.FMT_NUM1, link=True)
    bk.fml(ws, r, 8, f"=IF(F{r}=0,0,(G{r}-F{r})/F{r})", S.FMT_PCT); r += 2

    bk.section(ws, r, "PATIENT-RELATED COGS (monthly $)", 8); r += 1
    for lbl, key, mkey in (("Pharmacy", "pharmacy", "pharmacy_pd"),
                           ("Medical supplies", "supplies", "supplies_pd"),
                           ("DME", "dme", "dme_pd")):
        rr = actual_row(lbl, key, note="Source: Paloma P&L 2025.")
        bk.fml(ws, rr, 7, f"={_link('Revenue Model', M3, rev['pd'])}*{bk.addr[mkey]}", S.FMT_CUR, link=True)
        bk.fml(ws, rr, 8, f"=IF(F{rr}=0,0,(G{rr}-F{rr})/F{rr})", S.FMT_PCT); r += 1
    r += 1

    bk.section(ws, r, "PAYROLL (semi-monthly actual)", 8); r += 1
    p1 = actual_row("Payroll — 1st half", "pay1", note="Source: Paloma payrolls."); r += 1
    p2 = actual_row("Payroll — 2nd half", "pay2", note="Source: Paloma payrolls."); r += 1
    bk.lbl(ws, r, "Total monthly payroll (actual)", bold=True)
    for j in range(4):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"={cl}{p1}+{cl}{p2}", S.FMT_CUR, bold=True)
    bk.fml(ws, r, 6, f"=AVERAGE(C{r}:E{r})", S.FMT_CUR, bold=True)
    bk.fml(ws, r, 7, f"={_link('Staffing & Payroll', M3, bk.rows['Staffing & Payroll']['cogs_labor'])}"
                     f"+{_link('Staffing & Payroll', M3, bk.rows['Staffing & Payroll']['sga_labor'])}",
           S.FMT_CUR, link=True, bold=True)
    r += 2
    note = ("Model steady-state (M3) reconciles to Paloma's proven Apr–Jun book: ~$118K net / ~22 ADC. "
            "Payroll differs by design — Azalea's staggered W-2 roster + benefits vs Paloma's blended run.")
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=8)
    bk._set(ws, r, 1, note, S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# ASSEMBLY
# =========================================================================
def build_engine(bk: Book):
    """Create the shared engine tabs (identical in all three workbooks)."""
    inputs_tab(bk)
    revenue_tab(bk)
    staffing_tab(bk)
    opbudget_tab(bk)
    debt_tab(bk)
    cashflow_tab(bk)
    patch_pl_interest(bk)
    pl_detail_tab(bk)
    actuals_tab(bk)
    return bk


# =========================================================================
# DETAILED P&L  (QuickBooks-style, cash-basis layout — matches client template
# row-for-row; cash-basis "Net Income" = operating result, ties to EBITDA)
# Columns: B..M = M1..M12, N = Year 2, O = Year 3, P = TOTAL (3-yr)
# =========================================================================
FMT_ACCT = '#,##0.00;(#,##0.00);""'
FMT_TOT = '#,##0.00;(#,##0.00)'

def pl_detail_tab(bk: Book):
    ws = bk.wb.create_sheet("Detailed P&L")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    for c in range(2, 17):
        ws.column_dimensions[get_column_letter(c)].width = 12.5
    MCOL = lambda i: 2 + i          # B..M
    Y2C, Y3C, TOT = 14, 15, 16      # N, O, P
    rev, stf = bk.rows["Revenue Model"], bk.rows["Staffing & Payroll"]
    RG, PD = rev["gross"], rev["pd"]

    # title
    ws["A1"] = "Azalea Hospice & Palliative Care — Tyler, TX"
    ws["A1"].font = S.Font(name=S.BASE_FONT, size=13, bold=True, color=S.NAVY)
    ws["A2"] = "Profit and Loss"
    ws["A2"].font = S.Font(name=S.BASE_FONT, size=11, bold=True, color=S.BLACK)
    ws["A3"] = "Year 1 (monthly)  ·  Years 2–3 (annual)  ·  model data, cash basis"
    ws["A3"].font = S.f_note()

    # column header (row 5)
    hdr = 5
    labels = [PERIODS[i]["label"] for i in range(12)] + ["Year 2", "Year 3", "TOTAL"]
    for k, lab in enumerate(labels):
        c = ws.cell(hdr, 2 + k, lab)
        c.font = S.f_label(bold=True); c.alignment = S.CENTER
        c.border = S.Border(bottom=S.Side(style="thin", color=S.BLACK))
    ws.freeze_panes = "B6"

    # ---- engine reference builders (bare refs; '=' added by writers) ----
    def g_m(i):  return f"'Revenue Model'!{plet(i)}{RG}"
    def g_y(a, b): return f"SUM('Revenue Model'!{plet(a)}{RG}:{plet(b)}{RG})"
    def pd_m(i): return f"'Revenue Model'!{plet(i)}{PD}"
    def pd_y(a, b): return f"SUM('Revenue Model'!{plet(a)}{PD}:{plet(b)}{PD})"
    def s_m(key, i): return f"'Staffing & Payroll'!{plet(i)}{stf[key]}"
    def s_y(key, a, b): return f"SUM('Staffing & Payroll'!{plet(a)}{stf[key]}:{plet(b)}{stf[key]})"

    def label(R, text, indent, bold=False):
        cc = ws.cell(R, 1, text)
        cc.font = S.Font(name=S.BASE_FONT, size=10, bold=bold, color=S.BLACK)
        cc.alignment = S.Alignment(horizontal="left", vertical="center", indent=indent)

    def account(R, text, indent, m_fn, y_fn):
        """m_fn(i)->expr ; y_fn(a,b)->expr ; expr bare (no '=')."""
        label(R, text, indent)
        for i in range(12):
            c = ws.cell(R, MCOL(i), "=" + m_fn(i)); c.number_format = FMT_ACCT
            c.font = S.f_formula(); c.alignment = S.RIGHT
        for col, (a, b) in ((Y2C, (12, 15)), (Y3C, (16, 19))):
            c = ws.cell(R, col, "=" + y_fn(a, b)); c.number_format = FMT_ACCT
            c.font = S.f_link(); c.alignment = S.RIGHT
        c = ws.cell(R, TOT, f"=SUM(B{R}:M{R})+N{R}+O{R}"); c.number_format = FMT_ACCT
        c.font = S.f_total(); c.alignment = S.RIGHT

    def blank_account(R, text, indent):
        account(R, text, indent, lambda i: "0", lambda a, b: "0")

    def total(R, text, indent, sf, top=True, dbl=False, pct=False):
        label(R, text, indent, bold=True)
        for col in range(2, 17):
            cl = get_column_letter(col)
            c = ws.cell(R, col, "=" + sf(cl))
            c.number_format = S.FMT_PCT if pct else FMT_TOT
            c.font = S.f_total(); c.alignment = S.RIGHT
            bd = {}
            if top: bd["top"] = S.Side(style="thin", color=S.BLACK)
            if dbl: bd["bottom"] = S.Side(style="double", color=S.BLACK)
            if bd: c.border = S.Border(**bd)

    def header(R, text):
        label(R, text, 0, bold=True)

    # ===== INCOME =====
    header(6, "Income")
    account(7, "Gross Billed Revenue", 1, lambda i: g_m(i), lambda a, b: g_y(a, b))
    blank_account(8, "Less: CDAs Adjustment", 1)
    account(9, "Less: Net Due Zero (Write-offs)", 1,
            lambda i: f"{g_m(i)}*{bk.addr['writeoff_pct']}", lambda a, b: f"{g_y(a,b)}*{bk.addr['writeoff_pct']}")
    account(10, "Less: Sequestration", 1,
            lambda i: f"{g_m(i)}*{bk.addr['seq']}", lambda a, b: f"{g_y(a,b)}*{bk.addr['seq']}")
    blank_account(11, "Less: Medication Adjustments", 1)
    total(12, "Total Income", 1, lambda c: f"{c}7-{c}8-{c}9-{c}10-{c}11")

    # ===== COST OF SERVICES =====
    header(14, "Cost of Services")
    account(15, "Pharmacy", 1, lambda i: f"{pd_m(i)}*{bk.addr['pharmacy_pd']}", lambda a, b: f"{pd_y(a,b)}*{bk.addr['pharmacy_pd']}")
    account(16, "Medical Supplies", 1, lambda i: f"{pd_m(i)}*{bk.addr['supplies_pd']}", lambda a, b: f"{pd_y(a,b)}*{bk.addr['supplies_pd']}")
    account(17, "Durable Medical Equipment (DME)", 1, lambda i: f"{pd_m(i)}*{bk.addr['dme_pd']}", lambda a, b: f"{pd_y(a,b)}*{bk.addr['dme_pd']}")
    blank_account(18, "Respite Care", 1)
    blank_account(19, "Transportation", 1)
    blank_account(20, "Infusion / Labs", 1)
    blank_account(21, "PT / OT / ST", 1)
    total(22, "Total Cost of Services", 1, lambda c: f"SUM({c}15:{c}21)")

    # ===== GROSS PROFIT =====
    total(24, "Gross Profit", 0, lambda c: f"{c}12-{c}22")

    # ===== OPERATING EXPENSES =====
    header(26, "Operating Expenses")
    label(27, "Payroll & Related", 1, bold=True)
    account(28, "Payroll — 1st Half", 2,
            lambda i: f"({s_m('ft_direct',i)}+{s_m('ft_indirect',i)}+{s_m('prn',i)}+{s_m('rhonda',i)})/2",
            lambda a, b: f"({s_y('ft_direct',a,b)}+{s_y('ft_indirect',a,b)}+{s_y('prn',a,b)}+{s_y('rhonda',a,b)})/2")
    account(29, "Payroll — 2nd Half", 2,
            lambda i: f"({s_m('ft_direct',i)}+{s_m('ft_indirect',i)}+{s_m('prn',i)}+{s_m('rhonda',i)})/2",
            lambda a, b: f"({s_y('ft_direct',a,b)}+{s_y('ft_indirect',a,b)}+{s_y('prn',a,b)}+{s_y('rhonda',a,b)})/2")
    blank_account(30, "Payroll YTD Adjustments", 2)
    account(31, "Employer Taxes (FUTA/SUI/FICA/WC)", 2,
            lambda i: f"{s_m('burden_d',i)}+{s_m('burden_i',i)}", lambda a, b: f"{s_y('burden_d',a,b)}+{s_y('burden_i',a,b)}")
    blank_account(32, "Contract Labor", 2)
    account(33, "Medical Director", 2, lambda i: s_m("med_dir", i), lambda a, b: s_y("med_dir", a, b))
    blank_account(34, "Medical Director 2", 2)
    blank_account(35, "Intercompany Transfer", 2)
    account(36, "Staff Other (health insurance)", 2,
            lambda i: f"{s_m('health_d',i)}+{s_m('health_i',i)}", lambda a, b: f"{s_y('health_d',a,b)}+{s_y('health_i',a,b)}")
    total(37, "Total Payroll & Related", 1, lambda c: f"SUM({c}28:{c}36)")

    # G&A
    label(39, "General & Administrative", 1, bold=True)
    def ga(R, text, key):
        account(R, text, 2, lambda i: bk.addr[key], lambda a, b: f"{bk.addr[key]}*{(b-a+1)*3}")
    ga(40, "Bank / Payroll Fees", "ga_bankfees")
    ga(41, "Triple Net Rent", "ga_rent")
    ga(42, "Utilities", "ga_utilities")
    ga(43, "Internet", "ga_internet")
    ga(44, "Telephone / Fax", "ga_telephone")
    ga(45, "After Hours Messaging (BVTM)", "ga_bvtm")
    blank_account(46, "CFO / AP Support", 2)
    blank_account(47, "Support Services", 2)
    blank_account(48, "Messaging / CRG Signer", 2)
    blank_account(49, "NM Room & Board", 2)
    blank_account(50, "Corporate Labor", 2)
    ga(51, "Marketing", "ga_marketing")
    account(52, "Outsourced Billing Fee (1.5%)", 2,
            lambda i: f"{g_m(i)}*{bk.addr['billing_fee_pct']}", lambda a, b: f"{g_y(a,b)}*{bk.addr['billing_fee_pct']}")
    ga(53, "EMR System", "ga_emr")
    blank_account(54, "Other EMR", 2)
    ga(55, "PCR (CPM)", "ga_pcr")
    blank_account(56, "Ancillary", 2)
    ga(57, "Liability Insurance (D&O)", "ga_liability")
    account(58, "QR Payment Fee (0.75%)", 2,
            lambda i: f"{g_m(i)}*{bk.addr['qr_fee_pct']}", lambda a, b: f"{g_y(a,b)}*{bk.addr['qr_fee_pct']}")
    ga(59, "Training / CEUs", "ga_training")
    ga(60, "Credit Card Fees", "ga_cc")
    blank_account(61, "Other (QPf)", 2)
    ga(62, "Other (Client)", "ga_other")
    total(63, "Total General & Administrative", 1, lambda c: f"SUM({c}40:{c}62)")

    total(65, "Total Operating Expenses", 0, lambda c: f"{c}37+{c}63")
    total(67, "Net Income", 0, lambda c: f"{c}24-{c}65", dbl=True)
    total(69, "Net Margin %", 1, lambda c: f"IF({c}12=0,0,{c}67/{c}12)", top=False, pct=True)

    ws.merge_cells(start_row=71, start_column=1, end_row=72, end_column=8)
    ws.cell(71, 1, "Cash-basis operating P&L (QuickBooks layout): all payroll is in Operating Expenses, "
            "Cost of Services holds patient supplies only. \"Net Income\" here = operating result and ties to "
            "the engine's EBITDA; interest, D&A and tax are below EBITDA on the Operating Budget / 3-statement tabs.").font = S.f_note()
    ws.cell(71, 1).alignment = S.LEFT_WRAP
    return ws
