"""Audience-specific tabs layered on top of the shared engine:
 A = SBA loan package, B = investor/equity model, C = operations dashboard.
All numbers link back to the engine tabs (green) or Inputs.
"""
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

from . import styles as S
from .model import (PERIODS, NP, pcol, plet, Y1_IDX, Y2_IDX, Y3_IDX, ROSTER,
                    PRN_VISIT, GA_FIXED, STARTUP, ACTUALS, LICENSE_PAID_AT_CLOSE)
from .build import Book, _link

YR = {1: (0, 11), 2: (12, 15), 3: (16, 19)}


def ysum(sheet, row, year):
    a, b = YR[year]
    s = f"'{sheet}'" if (" " in sheet or "&" in sheet) else sheet
    return f"SUM({s}!{plet(a)}{row}:{plet(b)}{row})"


def yend(sheet, row, year):
    """Value at the last period of a year (for balances)."""
    b = YR[year][1]
    s = f"'{sheet}'" if (" " in sheet or "&" in sheet) else sheet
    return f"{s}!{plet(b)}{row}"


def _year_header(bk, ws, r, cols_from=3):
    for j, y in enumerate(("Year 1", "Year 2", "Year 3")):
        bk._set(ws, r, cols_from + j, y, S.f_label(bold=True), fill=S.fill(S.GREYHDR),
                align=S.CENTER, border=S.BORDER_THIN)


# =========================================================================
# A — SBA: SOURCES & USES
# =========================================================================
def sources_uses_tab(bk: Book):
    ws = bk.wb.create_sheet("Sources & Uses")
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 44
    ws.column_dimensions["D"].width = 16
    bk.title_block(ws, ("Sources & Uses of Financing - Hickory CHOW (seller paid at close) + working capital | SBA + equity"
                        if LICENSE_PAID_AT_CLOSE else
                        "Sources & Uses of Financing - Hickory CHOW acquisition + working capital | SBA + seller note + equity"), 4)
    bk.section_range(ws, 4, "SOURCES", 1, 2)
    bk.section_range(ws, 4, "USES", 3, 4)
    # Sources
    r = 5
    bk.lbl(ws, r, "SBA 7(a) loan", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['sba_principal']}", S.FMT_CUR, link=True); r += 1
    if not LICENSE_PAID_AT_CLOSE:
        bk.lbl(ws, r, "Hickory license seller note", indent=1)
        bk.fml(ws, r, 2, f"={bk.addr['license_cost']}", S.FMT_CUR, link=True); r += 1
    bk.lbl(ws, r, "Owner equity injection", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['equity']}", S.FMT_CUR, link=True); r += 1
    src_total = r
    bk.lbl(ws, r, "TOTAL SOURCES", bold=True)
    bk.fml(ws, r, 2, f"=SUM(B5:B{r-1})", S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    # Uses (column C/D)
    ur = 5
    bk.lbl(ws, ur, "Hickory Medicare license (CHOW)", c=3, indent=1)
    bk.fml(ws, ur, 4, f"={bk.addr['license_cost']}", S.FMT_CUR, link=True); ur += 1
    for key, label, *_ in STARTUP:
        bk.lbl(ws, ur, label, c=3, indent=1)
        bk.fml(ws, ur, 4, f"={bk.addr[key]}", S.FMT_CUR, link=True); ur += 1
    bk.lbl(ws, ur, "Startup equipment (capex)", c=3, indent=1)
    bk.fml(ws, ur, 4, f"={bk.addr['capex']}", S.FMT_CUR, link=True); ur += 1
    bk.lbl(ws, ur, "Working-capital reserve (opening cash)", c=3, indent=1)
    bk.fml(ws, ur, 4, f"={bk.addr['opening_cash']}", S.FMT_CUR, link=True); ur += 1
    bk.lbl(ws, ur, "TOTAL USES", c=3, bold=True)
    bk.fml(ws, ur, 4, f"=SUM(D5:D{ur-1})", S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    chk = max(src_total, ur) + 2
    bk.lbl(ws, chk, "CHECK: Sources - Uses to 0", bold=True)
    bk.fml(ws, chk, 2, f"=B{src_total}-D{ur}", S.FMT_CUR, bold=True, fill=S.fill(S.GREENFILL))
    ws.merge_cells(start_row=chk + 2, start_column=1, end_row=chk + 4, end_column=4)
    bk._set(ws, chk + 2, 1,
            ("CHOW acquisition: Azalea acquires Hickory Hospice's existing Medicare-certified provider number "
             "($300K, paid in full at close from SBA loan + equity proceeds) - gives Azalea immediate billing "
             "capability in San Antonio and an alternative-delivery site serving the East Texas / Tyler market. "
             "CHAP/ACHC accreditation transfers with the CHOW; 855A change-of-ownership preserves the provider "
             "number with no fresh enrollment delay. With no seller note, the SBA loan is the only debt the "
             "business carries."
             if LICENSE_PAID_AT_CLOSE else
             "CHOW acquisition: Azalea acquires Hickory Hospice's existing Medicare-certified provider number "
             "($300K, 36 months at 6% seller-financed) - gives Azalea immediate billing capability in San Antonio "
             "and an alternative-delivery site serving the East Texas / Tyler market. CHAP/ACHC accreditation "
             "transfers with the CHOW; 855A change-of-ownership preserves the provider number with no fresh "
             "enrollment delay. The seller note self-finances the license - it shows on both sides above."),
            S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# A — SBA: LENDER SUMMARY (annual DSCR, min cash, runway)
# =========================================================================
def lender_summary_tab(bk: Book):
    ws = bk.wb.create_sheet("Lender Summary")
    ws.column_dimensions["A"].width = 40
    for c in "BCDE":
        ws.column_dimensions[c].width = 16
    bk.title_block(ws, ("Lender Summary - SBA 7(a) debt-service coverage (seller paid at close) vs SBA floor 1.25x"
                        if LICENSE_PAID_AT_CLOSE else
                        "Lender Summary - COMBINED debt-service coverage (SBA + Hickory seller note) vs SBA floor 1.25x"), 5)
    ob, dbt, cf = "Operating Budget", "Debt Schedule", "Cash Flow & BS"
    R = bk.rows
    r = 4
    _year_header(bk, ws, r); r += 1
    def line(label, fn, fmt=S.FMT_CUR, bold=False, fill=None):
        nonlocal r
        bk.lbl(ws, r, label, bold=bold)
        for j, y in enumerate((1, 2, 3)):
            bk.fml(ws, r, 3 + j, "=" + fn(y), fmt, link=True, bold=bold, fill=fill)
        rr = r; r += 1; return rr
    line("Net revenue", lambda y: ysum(ob, R[ob]["net"], y))
    line("EBITDA", lambda y: ysum(ob, R[ob]["ebitda"], y), bold=True)
    line(" EBITDA margin %", lambda y: f"{ysum(ob, R[ob]['ebitda'], y)}/{ysum(ob, R[ob]['net'], y)}", S.FMT_PCT)
    line("Debt service (P+I)", lambda y: ysum(dbt, R[dbt]["ds"], y))
    dscr = line("DSCR", lambda y: f"{ysum(ob, R[ob]['ebitda'], y)}/{ysum(dbt, R[dbt]['ds'], y)}", S.FMT_MULT, bold=True, fill=S.fill(S.LIGHTBLUE))
    line("Net income", lambda y: ysum(ob, R[ob]["ni"], y))
    line("Ending cash", lambda y: yend(cf, R[cf]["endcash"], y))
    r += 1
    bk.section(ws, r, "KEY LENDER METRICS", 5); r += 1
    bk.lbl(ws, r, "Global 3-yr DSCR (target >= 1.25x)", bold=True)
    bk.fml(ws, r, 2,
           f"=({ysum(ob,R[ob]['ebitda'],1)}+{ysum(ob,R[ob]['ebitda'],2)}+{ysum(ob,R[ob]['ebitda'],3)})/"
           f"({ysum(dbt,R[dbt]['ds'],1)}+{ysum(dbt,R[dbt]['ds'],2)}+{ysum(dbt,R[dbt]['ds'],3)})",
           S.FMT_MULT, bold=True, fill=S.fill(S.GREENFILL)); r += 1
    bk.lbl(ws, r, "Minimum monthly DSCR (Y1)", bold=True)
    c0, c1 = plet(0), plet(11)
    bk.fml(ws, r, 2, f"=MIN('Debt Schedule'!{c0}{R[dbt]['dscr']}:{c1}{R[dbt]['dscr']})", S.FMT_MULT); r += 1
    bk.lbl(ws, r, "Minimum cash balance (any period)", bold=True)
    bk.fml(ws, r, 2, f"=MIN('Cash Flow & BS'!{c0}{R[cf]['endcash']}:{plet(NP-1)}{R[cf]['endcash']})", S.FMT_CUR); r += 1
    bk.lbl(ws, r, "Peak working-capital line draw", bold=True)
    bk.fml(ws, r, 2, f"=MAX('Cash Flow & BS'!{c0}{R[cf]['locbal']}:{plet(NP-1)}{R[cf]['locbal']})", S.FMT_CUR); r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 3, end_column=5)
    bk._set(ws, r, 1,
            ("Azalea acquires Hickory Hospice's already-certified Medicare provider number via a $300K CHOW "
             "(paid in full at close), eliminating the 855A enrollment cash gap a fresh startup would face - "
             "Azalea bills from day 1 in San Antonio and operates an alternative-delivery site for Tyler/East "
             "Texas. The migrated Paloma clinical team brings a proven ~$118K/mo, ~22 ADC book of business. "
             "With the seller paid at close, the SBA loan is the only debt the business carries, so the coverage "
             "shown here is SBA-only - strong from Year 1."
             if LICENSE_PAID_AT_CLOSE else
             "Azalea acquires Hickory Hospice's already-certified Medicare provider number via a $300K CHOW "
             "(seller-financed at 6% over 36 months), eliminating the 855A enrollment cash gap a fresh startup "
             "would face - Azalea bills from day 1 in San Antonio and operates an alternative-delivery site for "
             "Tyler/East Texas. The migrated Paloma clinical team brings a proven ~$118K/mo, ~22 ADC book of "
             "business. Coverage shown here is COMBINED (SBA + seller note); the seller note retires at month 36, "
             "after which DSCR jumps as only the SBA service remains."),
            S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# A — SBA: STRESS TESTS (steady-state annual sensitivity via shocks)
# =========================================================================
def stress_tab(bk: Book):
    ws = bk.wb.create_sheet("Stress Tests")
    ws.column_dimensions["A"].width = 34
    for c in "BCDEF":
        ws.column_dimensions[c].width = 15
    bk.title_block(ws, "Stress Tests - Year-2 steady-state | does the loan stay serviced under shocks?", 6)
    ob, dbt = "Operating Budget", "Debt Schedule"
    R = bk.rows
    # base annual (Y2) anchors
    base_rev = ysum(ob, R[ob]["net"], 2)
    base_ebitda = ysum(ob, R[ob]["ebitda"], 2)
    base_ds = ysum(dbt, R[dbt]["ds"], 2)
    # variable cost ratio (patient COGS + PRN + QR) approximated from Y2
    var_y2 = (f"({ysum(ob,R[ob]['cogs_patient'],2)}+{ysum(ob,R[ob]['qr'],2)}+"
              f"SUM('Staffing & Payroll'!{plet(12)}{R['Staffing & Payroll']['prn']}:{plet(15)}{R['Staffing & Payroll']['prn']}))")
    r = 4
    bk.section(ws, r, "SHOCK INPUTS (blue = adjust)", 6); r += 1
    bk.lbl(ws, r, "Revenue / census shock %", indent=1)
    bk.inp(ws, r, 2, -0.20, S.FMT_PCT, "e.g. census miss or capture-rate shortfall."); rev_shock = f"$B${r}"; r += 1
    bk.lbl(ws, r, "Wage / fixed-cost shock %", indent=1)
    bk.inp(ws, r, 2, 0.10, S.FMT_PCT, "labor inflation on fixed costs."); wage_shock = f"$B${r}"; r += 2

    bk.section(ws, r, "SCENARIO RESULTS (Year-2 annualized)", 6); r += 1
    heads = ["Metric", "Base", "Census -20%", "Census -35%", "Wage +10%", "Combined"]
    for j, h in enumerate(heads):
        bk._set(ws, r, 1 + j, h, S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    r += 1
    # multipliers per scenario: (rev_mult, fixed_mult)
    scen = [("=1", "=1"), ("=1-0.20", "=1"), ("=1-0.35", "=1"), ("=1", "=1+0.10"),
            (f"=1+{rev_shock}", f"=1+{wage_shock}")]
    mrow = r
    bk.lbl(ws, r, "Revenue multiplier", italic=True)
    for j, (rv, fx) in enumerate(scen):
        bk.fml(ws, r, 2 + j, rv, S.FMT_NUM1)
    r += 1
    bk.lbl(ws, r, "Fixed-cost multiplier", italic=True)
    for j, (rv, fx) in enumerate(scen):
        bk.fml(ws, r, 2 + j, fx, S.FMT_NUM1)
    fxrow = r; r += 1
    # fixed cost base = EBITDA = rev - var - fixed  => fixed = rev - var - ebitda
    bk.lbl(ws, r, "Net revenue", )
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"={base_rev}*{cl}{mrow}", S.FMT_CUR)
    rrev = r; r += 1
    bk.lbl(ws, r, "Variable costs")
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"=-{var_y2}*{cl}{mrow}", S.FMT_CUR)
    rvar = r; r += 1
    bk.lbl(ws, r, "Fixed costs")
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"=-({base_rev}-{var_y2}-{base_ebitda})*{cl}{fxrow}", S.FMT_CUR)
    rfix = r; r += 1
    bk.lbl(ws, r, "EBITDA", bold=True)
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"={cl}{rrev}+{cl}{rvar}+{cl}{rfix}", S.FMT_CUR, bold=True, fill=S.fill(S.LIGHTBLUE))
    reb = r; r += 1
    bk.lbl(ws, r, "Debt service")
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"={base_ds}", S.FMT_CUR)
    rds = r; r += 1
    bk.lbl(ws, r, "DSCR", bold=True)
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f"=IF({cl}{rds}=0,0,{cl}{reb}/{cl}{rds})", S.FMT_MULT, bold=True)
    rdscr = r; r += 1
    bk.lbl(ws, r, "Covers SBA floor 1.25x?", bold=True)
    for j in range(5):
        cl = get_column_letter(2 + j)
        bk.fml(ws, r, 2 + j, f'=IF({cl}{rdscr}>=1.25,"YES","NO")', S.FMT_INT, bold=True)
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=6)
    bk._set(ws, r, 1,
            "Simplified Year-2 steady-state sensitivity: revenue and variable costs scale with the census "
            "shock; fixed costs scale with the wage shock. For a full dynamic re-run, change Capture rate / "
            "Census scenario on the Inputs tab and the entire model recalculates.",
            S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    bk.rows["Stress Tests"] = {"dscr": rdscr}
    return ws


# =========================================================================
# B — INVESTOR: helpers
# =========================================================================
def _annual_fixed(bk):
    """Steady-state annual fixed cost (labor + benefits + fixed G&A), from Inputs."""
    sal = "+".join(h["sal"] for h in bk.roster)
    ga_r1 = bk.addr[GA_FIXED[0][0]].rsplit('$', 1)[1]
    ga_r2 = bk.addr[GA_FIXED[-1][0]].rsplit('$', 1)[1]
    nft = len(bk.roster)
    return (f"(({sal}+{bk.addr['rhonda_annual']})*(1+{bk.addr['benefits_load']})"
            f"+{bk.addr['health_pm']}*{nft}*12+{bk.addr['med_director']}*12"
            f"+SUM(Inputs!$B${ga_r1}:$B${ga_r2})*12)")


def _var_per_pd(bk):
    """Variable cost per patient-day (patient COGS + PRN/visit), from Inputs."""
    prn = "+".join(f"{p['vpm']}*{p['rate']}" for p in bk.prn)
    return (f"({bk.addr['supplies_pd']}+{bk.addr['dme_pd']}+{bk.addr['pharmacy_pd']}"
            f"+({prn})/{bk.addr['avg_days_month']})")


# =========================================================================
# B — INVESTOR: RETURNS
# =========================================================================
def returns_tab(bk: Book):
    ws = bk.wb.create_sheet("Returns")
    ws.column_dimensions["A"].width = 40
    for c in "BCDE":
        ws.column_dimensions[c].width = 16
    bk.title_block(ws, "Equity Returns - IRR | MOIC | cash-on-cash | 3-yr hold, exit at EBITDA multiple", 5)
    ob, dbt, cf = "Operating Budget", "Debt Schedule", "Cash Flow & BS"
    R = bk.rows
    r = 4
    bk.section(ws, r, "ASSUMPTIONS", 5); r += 1
    bk.lbl(ws, r, "Exit EBITDA multiple", indent=1)
    bk.inp(ws, r, 2, 5.0, S.FMT_MULT, "Exit enterprise value = multiple x Year-3 EBITDA."); exit_mult = f"$B${r}"; r += 1
    bk.lbl(ws, r, "Equity invested", indent=1)
    bk.fml(ws, r, 2, f"={bk.addr['equity']}", S.FMT_CUR, link=True); equity = f"$B${r}"; r += 2

    bk.section(ws, r, "EQUITY CASH FLOWS", 5); r += 1
    bk._set(ws, r, 2, "Close", S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    _year_header(bk, ws, r, cols_from=3); r += 1
    # FCFE per year = CFO + principal(negative) ; distributions assumed paid out
    bk.lbl(ws, r, "Free cash flow to equity", )
    bk.fml(ws, r, 2, f"=-{equity}", S.FMT_CUR)
    for j, y in enumerate((1, 2, 3)):
        bk.fml(ws, r, 3 + j, f"={ysum(cf, R[cf]['cfo'], y)}+{ysum(cf, R[cf]['prin'], y)}", S.FMT_CUR, link=True)
    fcfe = r; r += 1
    bk.lbl(ws, r, "Terminal equity value (exit)", )
    bk.fml(ws, r, 5, f"={exit_mult}*{ysum(ob, R[ob]['ebitda'], 3)}-{yend(dbt, R[dbt]['end'], 3)}-{yend(cf, R[cf]['locbal'], 3)}",
           S.FMT_CUR, link=True)
    term = r; r += 1
    bk.lbl(ws, r, "Net equity cash flow", bold=True)
    bk.fml(ws, r, 2, f"=B{fcfe}", S.FMT_CUR, bold=True)
    for j in range(3):
        cl = get_column_letter(3 + j)
        extra = f"+{cl}{term}" if j == 2 else ""
        bk.fml(ws, r, 3 + j, f"={cl}{fcfe}{extra}", S.FMT_CUR, bold=True, fill=S.fill(S.PALEBLUE))
    necf = r; r += 2

    bk.section(ws, r, "RETURNS", 5); r += 1
    bk.lbl(ws, r, "Equity IRR (3-yr)", bold=True)
    bk.fml(ws, r, 2, f"=IRR(B{necf}:E{necf})", S.FMT_PCT, bold=True, fill=S.fill(S.GREENFILL)); r += 1
    bk.lbl(ws, r, "MOIC (gross multiple)", bold=True)
    bk.fml(ws, r, 2, f"=SUM(C{necf}:E{necf})/{equity}", S.FMT_MULT, bold=True, fill=S.fill(S.GREENFILL)); r += 1
    bk.lbl(ws, r, "Avg cash-on-cash (Y1-Y3)", bold=True)
    bk.fml(ws, r, 2, f"=AVERAGE(C{fcfe}:E{fcfe})/{equity}", S.FMT_PCT, bold=True); r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=5)
    bk._set(ws, r, 1,
            "Base case (ADC 22 flat) is already profitable; switch Census scenario to 2 (Upside) on the "
            "Inputs tab to see growth-driven returns. Distribution policy assumes free cash flow to equity "
            "is distributed annually; terminal value exits at the EBITDA multiple net of remaining debt.",
            S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# B — INVESTOR: UNIT ECONOMICS
# =========================================================================
def unit_econ_tab(bk: Book):
    ws = bk.wb.create_sheet("Unit Economics")
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    bk.title_block(ws, "Unit Economics - contribution margin per patient-day", 4)
    r = 4
    bk.lbl(ws, r, "Net revenue per patient-day", bold=True)
    bk.fml(ws, r, 2, f"={bk.addr['net_rate']}", S.FMT_CUR2, link=True); nr = r; r += 1
    for key, lab in (("supplies_pd", "Medical supplies / PD"), ("dme_pd", "DME / PD"),
                     ("pharmacy_pd", "Pharmacy / PD")):
        bk.lbl(ws, r, lab, indent=1)
        bk.fml(ws, r, 2, f"=-{bk.addr[key]}", S.FMT_CUR2, link=True); r += 1
    bk.lbl(ws, r, "PRN / per-visit labor / PD", indent=1)
    prn = "+".join(f"{p['vpm']}*{p['rate']}" for p in bk.prn)
    bk.fml(ws, r, 2, f"=-({prn})/{bk.addr['avg_days_month']}", S.FMT_CUR2, link=True); r += 1
    bk.lbl(ws, r, "QR payment fee / PD", indent=1)
    bk.fml(ws, r, 2, f"=-{bk.addr['blended_gross']}*{bk.addr['qr_fee_pct']}", S.FMT_CUR2, link=True); r += 1
    bk.lbl(ws, r, "CONTRIBUTION MARGIN / PD", bold=True)
    bk.fml(ws, r, 2, f"=B{nr}-SUM(B{nr+1}:B{r-1})", S.FMT_CUR2, bold=True, fill=S.fill(S.LIGHTBLUE))
    cm = r; r += 1
    bk.lbl(ws, r, "Contribution margin %", italic=True)
    bk.fml(ws, r, 2, f"=B{cm}/B{nr}", S.FMT_PCT); r += 2
    bk.lbl(ws, r, "Daily fixed-cost burden (annual fixed ÷ 365)", indent=1)
    bk.fml(ws, r, 2, f"={_annual_fixed(bk)}/({bk.addr['avg_days_month']}*12)", S.FMT_CUR2, link=True); fx = r; r += 1
    bk.lbl(ws, r, "Break-even ADC (fixed ÷ contribution/PD)", bold=True)
    bk.fml(ws, r, 2, f"=B{fx}/B{cm}", S.FMT_NUM1, bold=True, fill=S.fill(S.GREENFILL)); r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=4)
    bk._set(ws, r, 1, "Low per-PD COGS (~$9, trued to Paloma actuals) drives a high contribution margin; "
            "operating leverage improves as census grows beyond the migrated panel.", S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# B — INVESTOR: SENSITIVITY GRID (ADC x net rate -> annual EBITDA)
# =========================================================================
def sensitivity_tab(bk: Book):
    ws = bk.wb.create_sheet("Sensitivity")
    ws.column_dimensions["A"].width = 22
    for j in range(6):
        ws.column_dimensions[get_column_letter(2 + j)].width = 13
    bk.title_block(ws, "Sensitivity - steady-state annual EBITDA by ADC x net rate/PD", 7)
    adcs = [16, 18, 20, 22, 26, 30, 35]
    rates = [170, 176, 181, 186, 192]
    daysyr = f"({bk.addr['avg_days_month']}*12)"
    fixed = _annual_fixed(bk)
    varpd = _var_per_pd(bk)
    r = 5
    bk._set(ws, r - 1, 1, "EBITDA ($/yr)", S.f_sub())
    bk._set(ws, r, 1, "ADC (rows) / net rate (cols)", S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    for j, rt in enumerate(rates):
        bk._set(ws, r, 2 + j, rt, S.f_input(), S.FMT_RATE, S.fill(S.INPUTFILL), S.CENTER, S.BORDER_THIN)
    rate_row = r; r += 1
    for adc in adcs:
        bk._set(ws, r, 1, adc, S.f_input(), S.FMT_NUM1, S.fill(S.INPUTFILL), S.CENTER, S.BORDER_THIN)
        for j in range(len(rates)):
            rc = get_column_letter(2 + j)
            # EBITDA = ADC*days*(rate - varpd) - fixed   (rate here already net; QR handled via varpd? keep simple)
            f = f"=$A{r}*{daysyr}*({rc}${rate_row}-{varpd})-{fixed}"
            bk.fml(ws, r, 2 + j, f, S.FMT_CUR)
        r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=7)
    bk._set(ws, r, 1, "Steady-state annual approximation: contribution = ADC x 365 x (net rate - variable/PD) "
            "less annual fixed cost. Base case ≈ ADC 22 @ $181.", S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# C — OPERATIONS: DASHBOARD
# =========================================================================
def dashboard_tab(bk: Book):
    ws = bk.wb.create_sheet("Dashboard")
    ws.column_dimensions["A"].width = 30
    for c in "BCDEFG":
        ws.column_dimensions[c].width = 15
    bk.title_block(ws, "Operations Dashboard - live KPIs (all formula-driven from the engine)", 7)
    ob, dbt, cf, rev = "Operating Budget", "Debt Schedule", "Cash Flow & BS", "Revenue Model"
    R = bk.rows
    cards = [
        ("Y3 run-rate net revenue", ysum(ob, R[ob]["net"], 3), S.FMT_CUR),
        ("Y3 EBITDA margin", f"{ysum(ob,R[ob]['ebitda'],3)}/{ysum(ob,R[ob]['net'],3)}", S.FMT_PCT),
        ("Min cash balance", f"MIN('Cash Flow & BS'!{plet(0)}{R[cf]['endcash']}:{plet(NP-1)}{R[cf]['endcash']})", S.FMT_CUR),
        ("Min DSCR (Y1 monthly)", f"MIN('Debt Schedule'!{plet(0)}{R[dbt]['dscr']}:{plet(11)}{R[dbt]['dscr']})", S.FMT_MULT),
        ("Peak working-capital draw", f"MAX('Cash Flow & BS'!{plet(0)}{R[cf]['locbal']}:{plet(NP-1)}{R[cf]['locbal']})", S.FMT_CUR),
        ("Blended net rate / PD", f"{bk.addr['net_rate']}", S.FMT_RATE),
    ]
    r = 4
    bk.section(ws, r, "KEY PERFORMANCE INDICATORS", 7); r += 1
    for k, (label, fml, fmt) in enumerate(cards):
        col = 1 + (k % 3) * 2
        row = r + (k // 3) * 3
        bk._set(ws, row, col, label, S.f_label(bold=True), fill=S.fill(S.LIGHTBLUE), align=S.LEFT, border=S.BORDER_THIN)
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        c = bk.fml(ws, row + 1, col, "=" + fml, fmt, link=True, bold=True)
        c.font = S.Font(name=S.BASE_FONT, size=14, bold=True, color=S.NAVY)
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
        ws[get_column_letter(col) + str(row + 1)].alignment = S.LEFT
    r = r + 6
    # ADC ramp line chart (data on Revenue Model)
    chart = LineChart(); chart.title = "ADC ramp (M1 - Y3 Q4)"; chart.height = 6.5; chart.width = 18
    chart.y_axis.title = "ADC"; chart.legend = None
    data = Reference(bk.wb[rev], min_col=pcol(0), max_col=pcol(NP - 1), min_row=R[rev]["adc"])
    chart.add_data(data, from_rows=True, titles_from_data=False)
    ws.add_chart(chart, f"A{r}")
    # Ending-cash line chart (data on Cash Flow & BS)
    cash = LineChart(); cash.title = "Ending cash by period"; cash.height = 6.5; cash.width = 18
    cash.y_axis.title = "$"; cash.legend = None
    cdata = Reference(bk.wb[cf], min_col=pcol(0), max_col=pcol(NP - 1), min_row=R[cf]["endcash"])
    cash.add_data(cdata, from_rows=True, titles_from_data=False)
    ws.add_chart(cash, f"A{r + 14}")
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# C — OPERATIONS: ACTUALS-vs-BUDGET VARIANCE (monthly, auto-flag >10%)
# =========================================================================
def variance_tab(bk: Book):
    ws = bk.wb.create_sheet("Actuals vs Budget")
    ws.column_dimensions["A"].width = 30
    for c in "BCDE":
        ws.column_dimensions[c].width = 16
    bk.title_block(ws, "Monthly Actuals vs Budget - enter month actuals (blue); variance auto-flags > 10%", 5)
    ob = "Operating Budget"; R = bk.rows
    bk.lbl(ws, 4, "Reporting month (1-12)", bold=True)
    bk.inp(ws, 4, 2, 3, S.FMT_INT, "Which model month to compare against (uses INDEX into Operating Budget).")
    mcell = "$B$4"
    hdr = 6
    for j, h in enumerate(["Line", "Budget", "Actual", "Variance", "Flag"]):
        bk._set(ws, hdr, 1 + j, h, S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    lines = [("Net revenue", "net"), ("Total COGS", "cogs"), ("Gross profit", "gp"),
             ("Total SG&A", "sga"), ("EBITDA", "ebitda"), ("Net income", "ni")]
    r = hdr + 1
    c0, c1 = plet(0), plet(11)
    for label, key in lines:
        bk.lbl(ws, r, label, indent=1)
        # budget via INDEX of the month columns
        bk.fml(ws, r, 2, f"=INDEX('Operating Budget'!{c0}{R[ob][key]}:{c1}{R[ob][key]},{mcell})", S.FMT_CUR, link=True)
        bk.inp(ws, r, 3, 0, S.FMT_CUR, "Enter actual for the reporting month.")
        bk.fml(ws, r, 4, f"=IF(B{r}=0,0,(C{r}-B{r})/B{r})", S.FMT_PCT)
        bk.fml(ws, r, 5, f'=IF(ABS(D{r})>0.1,"FLAG >10%","ok")', S.FMT_INT)
        r += 1
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 2, end_column=5)
    bk._set(ws, r + 1, 1, "Budget pulls the selected month from the engine; enter actuals from the close. "
            "Flag fires when |variance| > 10%.", S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# C — OPERATIONS: STAFFING TRACKER (hire calendar, FTE vs demand, open reqs)
# =========================================================================
def staffing_tracker_tab(bk: Book):
    ws = bk.wb.create_sheet("Staffing Tracker")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 26
    for c in "CDEF":
        ws.column_dimensions[c].width = 13
    bk.title_block(ws, "Staffing Tracker - staggered-hire calendar | capacity vs demand | open requisitions", 7)
    R = bk.rows; stf = "Staffing & Payroll"
    r = 4
    bk.section(ws, r, "HIRE CALENDAR", 7); r += 1
    for j, h in enumerate(["Name", "Role", "Annual salary", "Start month", "Status"]):
        bk._set(ws, r, 1 + j, h, S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.LEFT, border=S.BORDER_THIN)
    r += 1
    for h in bk.roster:
        bk.lbl(ws, r, h["name"], indent=1)
        bk.lbl(ws, r, h["role"], c=2)
        bk.fml(ws, r, 3, f"={h['sal']}", S.FMT_CUR, link=True)
        bk.fml(ws, r, 4, f"={h['start']}", S.FMT_INT, link=True)
        bk.lbl(ws, r, "OPEN REQ" if h["open"] else "Filled", c=5, italic=h["open"], bold=h["open"])
        if h["open"]:
            ws[f"E{r}"].fill = S.fill(S.REDFLAG)
        r += 1
    r += 1
    bk.section(ws, r, "CAPACITY vs DEMAND (steady-state ADC 22)", 7); r += 1
    bk.lbl(ws, r, "RN case managers needed = ADC ÷ caseload", indent=1)
    bk.fml(ws, r, 3, f"=ROUNDUP('Revenue Model'!{plet(2)}{R['Revenue Model']['adc']}/{bk.addr['rn_caseload']},0)", S.FMT_NUM1, link=True); r += 1
    bk.lbl(ws, r, "Aides needed = ADC ÷ caseload", indent=1)
    bk.fml(ws, r, 3, f"=ROUNDUP('Revenue Model'!{plet(2)}{R['Revenue Model']['adc']}/{bk.addr['aide_caseload']},0)", S.FMT_NUM1, link=True); r += 1
    bk.lbl(ws, r, "Target visits / RN / day", indent=1)
    bk.fml(ws, r, 3, f"={bk.addr['visits_rn_day']}", S.FMT_NUM1, link=True); r += 1
    bk.lbl(ws, r, "Target visits / aide / day", indent=1)
    bk.fml(ws, r, 3, f"={bk.addr['visits_aide_day']}", S.FMT_NUM1, link=True); r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=7)
    bk._set(ws, r, 1, "Jodi McCollum's 2nd RN-CM seat is funded but vacant - modeled to start M4 as census "
            "coverage requires. Track real churn beyond this one open req.", S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# C — OPERATIONS: MEDICARE CAP MONITOR + PAYROLL RECONCILIATION
# =========================================================================
def cap_payroll_tab(bk: Book):
    ws = bk.wb.create_sheet("Cap & Payroll")
    ws.column_dimensions["A"].width = 38
    for c in "BCDEF":
        ws.column_dimensions[c].width = 15
    bk.title_block(ws, "Medicare Cap Monitor (inherited day-counts) + Payroll Reconciliation", 6)
    R = bk.rows; ob = "Operating Budget"; rev = "Revenue Model"
    r = 4
    bk.section(ws, r, "MEDICARE AGGREGATE CAP - by year", 6); r += 1
    _year_header(bk, ws, r); r += 1
    bk.lbl(ws, r, "Unique beneficiaries (est.)", indent=1)
    bk.inp(ws, r, 3, 40, S.FMT_INT, "Migrated panel arrives mid-episode; estimate beneficiaries served / cap year.")
    bk.inp(ws, r, 4, 55, S.FMT_INT, ""); bk.inp(ws, r, 5, 60, S.FMT_INT, ""); ben = r; r += 1
    bk.lbl(ws, r, "Cap allowable = beneficiaries x cap", indent=1)
    for j in range(3):
        cl = get_column_letter(3 + j)
        bk.fml(ws, r, 3 + j, f"={cl}{ben}*{bk.addr['medicare_cap']}", S.FMT_CUR, link=True)
    capr = r; r += 1
    bk.lbl(ws, r, "Gross Medicare revenue", indent=1)
    for j, y in enumerate((1, 2, 3)):
        bk.fml(ws, r, 3 + j, "=" + ysum(rev, R[rev]["gross"], y), S.FMT_CUR, link=True)
    grossr = r; r += 1
    bk.lbl(ws, r, "Cap headroom (allowable - gross)", bold=True)
    for j in range(3):
        cl = get_column_letter(3 + j)
        bk.fml(ws, r, 3 + j, f"={cl}{capr}-{cl}{grossr}", S.FMT_CUR, bold=True, fill=S.fill(S.GREENFILL)); 
    r += 2
    bk.section(ws, r, "PAYROLL RECONCILIATION (actual semi-monthly vs model)", 6); r += 1
    for j, h in enumerate(["Run", "1st half", "2nd half", "Month total"]):
        bk._set(ws, r, 1 + j, h, S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    r += 1
    for k, mo in enumerate(["Apr 2025", "May 2025", "Jun 2025"]):
        idx = k + 1  # ACTUALS arrays index (skip Mar)
        bk.lbl(ws, r, mo, indent=1)
        bk.inp(ws, r, 2, ACTUALS["pay1"][idx], S.FMT_CUR, "Source: Paloma payrolls.")
        bk.inp(ws, r, 3, ACTUALS["pay2"][idx], S.FMT_CUR, "Source: Paloma payrolls.")
        bk.fml(ws, r, 4, f"=B{r}+C{r}", S.FMT_CUR, bold=True)
        r += 1
    bk.lbl(ws, r, "Model total labor (M3 steady)", bold=True)
    m3 = plet(2)
    bk.fml(ws, r, 4, f"='Staffing & Payroll'!{m3}{R['Staffing & Payroll']['cogs_labor']}"
                     f"+'Staffing & Payroll'!{m3}{R['Staffing & Payroll']['sga_labor']}", S.FMT_CUR, link=True, bold=True)
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=6)
    bk._set(ws, r, 1, "Cap monitor uses inherited mid-episode day-counts (estimate beneficiaries), not fresh admits. "
            "Payroll ties the 3 actual semi-monthly runs to the model's loaded labor.", S.f_note(), align=S.LEFT_WRAP)
    ws.sheet_view.showGridLines = False
    return ws


# =========================================================================
# ASSEMBLY — three workbooks
# =========================================================================
def assemble_sba(bk: Book):
    from .build import build_engine
    build_engine(bk)
    sources_uses_tab(bk)
    lender_summary_tab(bk)
    stress_tab(bk)
    # order: put summary tabs right after Inputs
    _reorder(bk, ["Inputs", "Sources & Uses", "Lender Summary", "Stress Tests",
                  "Revenue Model", "Staffing & Payroll", "Operating Budget", "Detailed P&L",
                  "Debt Schedule", "Cash Flow & BS", "Actuals vs Model"])
    return bk


def assemble_investor(bk: Book):
    from .build import build_engine
    build_engine(bk)
    returns_tab(bk)
    investment_options_tab(bk)
    unit_econ_tab(bk)
    sensitivity_tab(bk)
    _reorder(bk, ["Inputs", "Returns", "Investment Options", "Unit Economics", "Sensitivity",
                  "Revenue Model", "Staffing & Payroll", "Operating Budget", "Detailed P&L",
                  "Debt Schedule", "Cash Flow & BS", "Actuals vs Model"])
    return bk


def assemble_ops(bk: Book):
    from .build import build_engine
    build_engine(bk)
    dashboard_tab(bk)
    variance_tab(bk)
    staffing_tracker_tab(bk)
    cap_payroll_tab(bk)
    _reorder(bk, ["Inputs", "Dashboard", "Actuals vs Budget", "Staffing Tracker",
                  "Cap & Payroll", "Revenue Model", "Staffing & Payroll",
                  "Operating Budget", "Detailed P&L", "Debt Schedule", "Cash Flow & BS", "Actuals vs Model"])
    return bk


def _reorder(bk: Book, order):
    sheets = bk.wb._sheets
    by_name = {ws.title: ws for ws in sheets}
    new = [by_name[n] for n in order if n in by_name]
    new += [ws for ws in sheets if ws.title not in order]
    bk.wb._sheets = new


# =========================================================================
# B — INVESTOR: INVESTMENT OPTIONS  (debt / hybrid / preferred equity menu)
# Five-year hold. Investor inputs amount + term choices; tab models
# year-by-year cash flow, IRR, MOIC, total return for each instrument
# and presents a side-by-side comparison.
# =========================================================================
def investment_options_tab(bk: Book):
    ws = bk.wb.create_sheet("Investment Options")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 16
    bk.title_block(ws, "Investor Capital - Three Instruments | debt | convertible | preferred equity | choose by yield vs. upside", 8)

    ob = "Operating Budget"
    R = bk.rows

    # ===================== SHARED INVESTOR INPUTS =====================
    bk.section(ws, 4, "SHARED INPUTS (edit blue cells)", 8)
    bk.lbl(ws, 5, "Investment amount ($)", indent=1)
    bk.inp(ws, 5, 2, 50000, S.FMT_CUR, "Per-investor check size. Tiers $25K / $50K / $100K are common.")
    INV = "$B$5"
    bk.lbl(ws, 6, "Hold period (years)", indent=1)
    bk.inp(ws, 6, 2, 5, S.FMT_INT, "Years to maturity / exit (model assumes redemption/exit at end).")
    HOLD = "$B$6"
    bk.lbl(ws, 7, "Exit EBITDA multiple", indent=1)
    bk.inp(ws, 7, 2, 5.0, S.FMT_MULT, "Drives convertible and preferred equity exit value.")
    EXM = "$B$7"
    bk.lbl(ws, 8, "Y3 EBITDA (link)", indent=1, italic=True)
    bk.fml(ws, 8, 2, "=" + ysum(ob, R[ob]["ebitda"], 3), S.FMT_CUR, link=True)
    YEB = "$B$8"
    bk.lbl(ws, 9, "Equity exit value (Y3 EBITDA x multiple - net debt)", indent=1)
    bk.fml(ws, 9, 2, f"={EXM}*{YEB}-{yend('Debt Schedule', R['Debt Schedule']['end'], 3)}-{yend('Cash Flow & BS', R['Cash Flow & BS']['locbal'], 3)}",
           S.FMT_CUR, link=True)
    EQV = "$B$9"

    # column heads (B-D are the three options)
    hdr_row = 11
    headers = [("OPTION A - Promissory Note", "Pure DEBT"),
               ("OPTION B - Convertible Note", "HYBRID"),
               ("OPTION C - Preferred Equity", "EQUITY w/ floor")]
    for k, (nm, tag) in enumerate(headers):
        col = 2 + k * 2  # B, D, F
        ws.merge_cells(start_row=hdr_row, start_column=col, end_row=hdr_row, end_column=col + 1)
        bk._set(ws, hdr_row, col, nm, S.f_section(), fill=S.fill(S.NAVY), align=S.CENTER)
        ws.merge_cells(start_row=hdr_row + 1, start_column=col, end_row=hdr_row + 1, end_column=col + 1)
        bk._set(ws, hdr_row + 1, col, tag, S.f_sub(), fill=S.fill(S.LIGHTBLUE), align=S.CENTER)
    ws.row_dimensions[hdr_row].height = 22

    # ===================== TERMS ROW =====================
    r = hdr_row + 3
    bk.section_range(ws, r, "TERMS (blue = adjust per offer)", 1, 7); r += 1

    # --- Option A (Promissory Note) ---
    bk.lbl(ws, r, "Interest rate (APR)", indent=1)
    bk.inp(ws, r, 2, 0.10, S.FMT_PCT, "Senior subordinated debt: 9-12% typical given subordination to SBA.")
    A_RATE = f"$B${r}"
    bk.inp(ws, r, 4, 0.07, S.FMT_PCT, "Convertible coupon (lower than straight debt; equity upside compensates).")
    B_CPN = f"$D${r}"
    bk.inp(ws, r, 6, 0.08, S.FMT_PCT, "Preferred dividend rate, cumulative.")
    C_DIV = f"$F${r}"; r += 1

    bk.lbl(ws, r, "Term / amortization", indent=1)
    bk.inp(ws, r, 2, 5, S.FMT_INT, "Years; monthly P+I amortization in this model.")
    A_TERM = f"$B${r}"
    bk.inp(ws, r, 4, 5, S.FMT_INT, "Years to maturity / conversion trigger.")
    B_TERM = f"$D${r}"
    bk.inp(ws, r, 6, 5, S.FMT_INT, "Years to redemption / exit.")
    C_TERM = f"$F${r}"; r += 1

    bk.lbl(ws, r, "Payment style / conversion / preference", indent=1)
    bk.lbl(ws, r, "Monthly P+I", c=2, italic=True)
    bk.lbl(ws, r, "Discount + valuation cap", c=4, italic=True)
    bk.lbl(ws, r, "1x liq preference, then participate", c=6, italic=True); r += 1

    # convertible-only: discount + cap
    bk.lbl(ws, r, "(Convertible) Discount on equity round", indent=1)
    bk.inp(ws, r, 4, 0.20, S.FMT_PCT, "Discount applied at conversion vs. exit/round price.")
    B_DISC = f"$D${r}"; r += 1
    bk.lbl(ws, r, "(Convertible) Valuation cap ($)", indent=1)
    bk.inp(ws, r, 4, 3000000, S.FMT_CUR, "Implied pre-money cap; conversion uses MIN(cap, exit_val x (1-discount)).")
    B_CAP = f"$D${r}"; r += 1

    # preferred-only: participation
    bk.lbl(ws, r, "(Preferred) Participate after preference", indent=1)
    bk.inp(ws, r, 6, 1, S.FMT_INT, "1 = participating (1x back, then pro-rata). 0 = non-participating (greater of preference or as-converted).")
    C_PART = f"$F${r}"; r += 1

    # ===================== CASH FLOW SCHEDULE (5 yrs) =====================
    r += 1
    bk.section_range(ws, r, "INVESTOR CASH FLOW SCHEDULE (annual, Year 0 = close)", 1, 7); r += 1
    # row of year labels
    yhdr = r
    bk._set(ws, r, 1, "Year", S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.LEFT, border=S.BORDER_THIN)
    for y in range(0, 6):
        bk._set(ws, r, 2 + y, f"Y{y}", S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    r += 1

    # --- OPTION A: Promissory note. Monthly P+I; annual cash flow = pmt*12 within term, 0 after.
    bk.lbl(ws, r, "OPTION A - Promissory Note", indent=1, bold=True)
    # Annual payment = PMT × 12 if year ≤ term else 0; Y0 = −investment
    PMT_A = f"PMT({A_RATE}/12,{A_TERM}*12,-{INV})"
    bk.fml(ws, r, 2, f"=-{INV}", S.FMT_CUR, bold=True)
    for y in range(1, 6):
        cl = get_column_letter(2 + y)
        bk.fml(ws, r, 2 + y, f"=IF({y}<={A_TERM},{PMT_A}*12,0)", S.FMT_CUR)
    A_ROW = r; r += 1

    # --- OPTION B: Convertible note. Annual coupon paid Y1..term; at maturity/exit pick GREATER of (P+remaining accrual) or equity-converted.
    bk.lbl(ws, r, "OPTION B - Convertible Note", indent=1, bold=True)
    # Conversion ownership %: investment / MIN(cap, exit_val × (1−discount))
    OWN_B = f"({INV}/MIN({B_CAP},{EQV}*(1-{B_DISC})))"
    # Equity-converted proceeds at exit
    EQ_PROC_B = f"({OWN_B}*{EQV})"
    # Face value at maturity = principal back
    FACE_B = f"{INV}"
    # Greater of two at exit year
    bk.fml(ws, r, 2, f"=-{INV}", S.FMT_CUR, bold=True)
    for y in range(1, 6):
        cl = get_column_letter(2 + y)
        # interim years: coupon if y < term, else 0
        # exit year (y == term): coupon + MAX(face, eq_proceeds) − face_payoff
        # Simpler: years 1..term-1 = coupon; year term = MAX(face + coupon, eq_proceeds)
        expr = (f"=IF({y}<{B_TERM},{INV}*{B_CPN},"
                f"IF({y}={B_TERM},MAX({INV}+{INV}*{B_CPN},{EQ_PROC_B}),0))")
        bk.fml(ws, r, 2 + y, expr, S.FMT_CUR)
    B_ROW = r; r += 1

    # --- OPTION C: Preferred equity. Annual dividend; at exit year: 1× preference back + (if participating) pro-rata share of residual.
    bk.lbl(ws, r, "OPTION C - Preferred Equity", indent=1, bold=True)
    OWN_C = f"({INV}/{EQV})"   # as-converted ownership % of exit value
    # Residual after preference returned to all preferreds (just this investor for simplicity)
    RESIDUAL = f"MAX(0,{EQV}-{INV})"
    PARTICIPATION = f"IF({C_PART}=1,{OWN_C}*{RESIDUAL},MAX({INV},{OWN_C}*{EQV})-{INV})"
    bk.fml(ws, r, 2, f"=-{INV}", S.FMT_CUR, bold=True)
    for y in range(1, 6):
        cl = get_column_letter(2 + y)
        # dividend every year up to and including exit; at exit year add preference + participation
        expr = (f"=IF({y}<{C_TERM},{INV}*{C_DIV},"
                f"IF({y}={C_TERM},{INV}*{C_DIV}+{INV}+{PARTICIPATION},0))")
        bk.fml(ws, r, 2 + y, expr, S.FMT_CUR)
    C_ROW = r; r += 2

    # ===================== RETURNS COMPARISON =====================
    bk.section_range(ws, r, "RETURNS COMPARISON (Y0 = -investment; subsequent years per schedule above)", 1, 8); r += 1
    # metric rows
    # Per-option headline rate (stated APR/coupon/dividend) for direct comparison.
    HEADLINE = {A_ROW: A_RATE, B_ROW: B_CPN, C_ROW: C_DIV}
    # IRR formula per option. Option A pays MONTHLY P+I, so its true investor
    # yield is the effective annual rate (1 + APR/12)^12 − 1, NOT the annual
    # IRR of the lumpy annual series. Options B and C pay annually; annual IRR
    # is exact.
    IRR_FORM = {A_ROW: lambda row: f"(1+{A_RATE}/12)^12-1",
                B_ROW: lambda row: f"IRR(B{row}:G{row})",
                C_ROW: lambda row: f"IRR(B{row}:G{row})"}
    metrics = [
        ("Headline rate (APR / coupon / pref div)", lambda row: HEADLINE[row], S.FMT_PCT, False),
        ("Total cash to investor (Y1-Y5)", lambda row: f"SUM(C{row}:G{row})", S.FMT_CUR, False),
        ("Total profit (cash - investment)", lambda row: f"SUM(C{row}:G{row})-{INV}", S.FMT_CUR, False),
        ("MOIC (gross multiple)", lambda row: f"SUM(C{row}:G{row})/{INV}", S.FMT_MULT, True),
        ("IRR / effective annual yield", lambda row: IRR_FORM[row](row), S.FMT_PCT, True),
        ("Average annual cash yield (Y1-Y5)", lambda row: f"AVERAGE(C{row}:G{row})/{INV}", S.FMT_PCT, False),
    ]
    metric_total_row = None
    metric_irr_row = None
    for label, fn, fmt, highlight in metrics:
        bk.lbl(ws, r, label, indent=1, bold=True)
        for k, src_row in enumerate([A_ROW, B_ROW, C_ROW]):
            col = 2 + k * 2
            bk.fml(ws, r, col, "=" + fn(src_row), fmt, bold=True,
                   fill=S.fill(S.GREENFILL) if highlight else None)
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        if label.startswith("Total cash to investor"):
            metric_total_row = r
        if label.startswith("IRR / effective"):
            metric_irr_row = r
        r += 1
    r += 1

    # ===================== RISK / WHO IT'S FOR =====================
    bk.section_range(ws, r, "RISK & FIT", 1, 7); r += 1
    bk._set(ws, r, 1, "Profile", S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.LEFT, border=S.BORDER_THIN)
    for k, txt in enumerate(("Lowest risk | fixed yield", "Medium risk | upside option", "Highest risk | highest upside")):
        col = 2 + k * 2
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        bk._set(ws, r, col, txt, S.f_label(bold=True), fill=S.fill(S.LIGHTBLUE), align=S.CENTER, border=S.BORDER_THIN)
    r += 1
    bk._set(ws, r, 1, "Security / position", S.f_label(), align=S.LEFT)
    for k, txt in enumerate(("Senior subordinated debt; subordinate to SBA",
                             "Debt at maturity OR equity at exit (greater)",
                             "Equity; subordinate to all debt")):
        col = 2 + k * 2
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        bk._set(ws, r, col, txt, S.f_note(), align=S.LEFT_WRAP)
    r += 1
    bk._set(ws, r, 1, "Cash flow to investor", S.f_label(), align=S.LEFT)
    for k, txt in enumerate(("Monthly P+I, predictable",
                             "Annual coupon, lump at exit",
                             "Annual dividend, lump at exit")):
        col = 2 + k * 2
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        bk._set(ws, r, col, txt, S.f_note(), align=S.LEFT_WRAP)
    r += 1
    bk._set(ws, r, 1, "Tax treatment", S.f_label(), align=S.LEFT)
    for k, txt in enumerate(("Interest = ordinary income",
                             "Interest = ordinary; conversion = capital gain",
                             "Dividends + capital gain on exit")):
        col = 2 + k * 2
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        bk._set(ws, r, col, txt, S.f_note(), align=S.LEFT_WRAP)
    r += 1
    bk._set(ws, r, 1, "Best for", S.f_label(), align=S.LEFT)
    for k, txt in enumerate(("Yield-focused, risk-averse passive lenders",
                             "Investors who want yield AND a shot at equity upside",
                             "Believers in the upside who can take equity risk")):
        col = 2 + k * 2
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        bk._set(ws, r, col, txt, S.f_note(), align=S.LEFT_WRAP)
    r += 2

    # ===================== TIER MENU (sample check sizes) =====================
    bk.section_range(ws, r, "ILLUSTRATIVE TIERS (recompute by changing 'Investment amount' above)", 1, 7); r += 1
    for j, h in enumerate(["Check size", "Option A: 5-yr IRR", "Option A: total cash", "Option B: IRR @5x exit", "Option B: total cash", "Option C: IRR @5x exit", "Option C: total cash"]):
        bk._set(ws, r, 1 + j, h, S.f_label(bold=True), fill=S.fill(S.GREYHDR), align=S.CENTER, border=S.BORDER_THIN)
    r += 1
    # These cells use the same option formulas with substituted amount.
    # Simplest: have each row reference its own scaled cash flows by ratio
    # IRR is independent of scale -> use the live IRR cells; cash scales linearly with amount.
    irr_row = A_ROW  # not used; we reference the metrics rows below
    # Recompute total cash by ratio: amount/INV * (live total)
    # IRR is amount-independent in all three options *given the same terms*.
    A_LIVE_IRR = "B" + str(A_ROW + 5 - 1)   # 4th metric row (IRR) for option A column B
    # Simpler: take IRR cells we just built
    # metrics order: total, profit, moic, IRR, yield -> IRR is at metrics row index 3
    # A_ROW + (3 rows up to metrics start) ... we tracked metrics with sequential r so:
    # find IRR rows via their addresses:
    # IRR row for opt A in col B = r_metrics_start + 3
    # we don't track r_metrics_start here -- instead embed inline below using PMT/IRR per-row
    # IRR and total-cash for Options B and C scale linearly with check size
    # (terms and EQV fixed) — reference the live metrics rows above.
    for amt in (25000, 50000, 100000, 250000):
        bk._set(ws, r, 1, amt, S.f_input(), S.FMT_CUR, S.fill(S.INPUTFILL), S.CENTER, S.BORDER_THIN)
        # Option A: monthly P+I → IRR ≈ EFFECT(A_RATE,12); cash = PMT × 12 × term
        bk.fml(ws, r, 2, f"=(1+{A_RATE}/12)^12-1", S.FMT_PCT)
        bk.fml(ws, r, 3, f"=PMT({A_RATE}/12,{A_TERM}*12,-A{r})*12*{A_TERM}", S.FMT_CUR)
        # Option B: IRR same regardless of amount (terms fixed); cash scales A{r}/INV
        bk.fml(ws, r, 4, f"=$D${metric_irr_row}", S.FMT_PCT)
        bk.fml(ws, r, 5, f"=A{r}/{INV}*$D${metric_total_row}", S.FMT_CUR)
        # Option C: same logic
        bk.fml(ws, r, 6, f"=$F${metric_irr_row}", S.FMT_PCT)
        bk.fml(ws, r, 7, f"=A{r}/{INV}*$F${metric_total_row}", S.FMT_CUR)
        r += 1

    r += 1
    note = ("Conventions: Year-0 cash flow is the investor's check (negative). Subsequent years are cash "
            "received. Option A IRR ≈ rate (slight compounding bump from monthly payments). Options B and C "
            "IRR/cash scale linearly with check size when terms and exit value are fixed - use the tier menu "
            "to size offers, then change 'Investment amount' to see the live cash-flow schedule update. "
            "Conversion math: Option B holder receives MAX(principal + final coupon, equity-as-converted) at "
            "year = term, where conversion ownership = investment ÷ MIN(cap, exit x (1 - discount)).")
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 4, end_column=8)
    bk._set(ws, r, 1, note, S.f_note(), align=S.LEFT_WRAP)
    return ws
