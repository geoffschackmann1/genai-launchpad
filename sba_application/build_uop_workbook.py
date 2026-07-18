"""UOP workbook for John Hart / Mary Brownmiller - Path A structure (Rev 4.10 basis).

Story: equity $250K funds the $125K license down payment + ramp working capital;
Jim Bullard's bank note ($500K/6%/36) funds September and pays the sellers $375K;
the SBA 7(a) $500K funds ~1/15/2027 (at the 51% transfer) and REFINANCES the bank
note (~$461,676 payoff), dropping debt service from $15,211/mo to ~$6,747/mo.
The $250K equity is the SBA injection (33% of the $750K project), traced
dollar-for-dollar on the Equity Injection Trace tab.
Output: sba_application/13_checklist_response_2026-07-16/UOP Azalea Refuge Rev2.00.xlsx
"""
import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side

OUT = "sba_application/13_checklist_response_2026-07-16/UOP Azalea Refuge Rev2.00.xlsx"

H = Font(bold=True, size=12, color="1F3B2D")
B = Font(bold=True)
IT = Font(italic=True, size=9)
MONEY = "#,##0"
PCT = "0.0%"
FILL = PatternFill("solid", fgColor="EAF1EC")


def sty(ws, cell, value, font=None, fmt=None, fill=None):
    c = ws[cell]
    c.value = value
    if font: c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    return c


def tab_structure(wb):
    ws = wb.active
    ws.title = "Loan Structure"
    ws.column_dimensions["A"].width = 52
    for col in "BCDEF": ws.column_dimensions[col].width = 15

    sty(ws, "A1", "PROPOSED SBA 7(a) LOAN STRUCTURE - AZALEA HOSPICE & PALLIATIVE CARE", H)
    sty(ws, "A2", "Tyler Hospice Hold, LLC (EIN 41-4966640) | Refuge Hospice acquisition | Rev 2.00 basis: Rev 4.10 proforma | 7/16/2026", IT)

    sty(ws, "A4", "OVERALL PROJECT (all phases)", B, fill=FILL)
    sty(ws, "A5", "Refuge Hospice, LLC license purchase price")
    sty(ws, "B5", 500000, fmt=MONEY)
    sty(ws, "A6", "Working capital (ramp payroll ahead of Medicare collections)")
    sty(ws, "B6", 235000, fmt=MONEY)
    sty(ws, "A7", "Soft costs (SBA packaging, guaranty fee, legal/closing)")
    sty(ws, "B7", 15000, fmt=MONEY)
    sty(ws, "A8", "TOTAL PROJECT", B); sty(ws, "B8", "=B5+B6+B7", fmt=MONEY).font = B
    sty(ws, "A9", "Equity injection (Bullard $195K + Schackmann $55K - see Injection Trace tab)")
    sty(ws, "B9", 250000, fmt=MONEY)
    sty(ws, "C9", "=B9/B8", fmt=PCT)
    sty(ws, "A10", "SBA 7(a) loan request", B)
    sty(ws, "B10", 500000, fmt=MONEY).font = B
    sty(ws, "C10", "=B10/B8", fmt=PCT)
    sty(ws, "A11", "Check: sources - uses (0 = tied)")
    sty(ws, "B11", "=B9+B10-B8", fmt=MONEY)

    sty(ws, "A13", "FINANCING SEQUENCE (how the project is funded through SBA closing)", B, fill=FILL)
    for r, (a, b) in enumerate([
        ("~8/1/2026 - Equity pays license down payment (49% interest transfers)", "$125,000"),
        ("Aug-Dec - Equity balance funds ramp working capital", "$125,000"),
        ("Sept 2026 - Interim bank note (Bullard-relationship TX bank, $500K, 6%, 36-mo) pays sellers in full", "$375,000 + interest"),
        ("Sept 2026 - Bank note surplus to working capital", "~$123,000"),
        ("Oct-Dec - Bank note service (interim)", "$15,211/mo"),
        ("1/15/2027 - SBA 7(a) funds at the 51% transfer; REFINANCES the bank note", "~$461,676 payoff"),
        ("1/15/2027 - SBA surplus: soft costs + additional working capital", "~$38,324"),
        ("From Feb 2027 - single SBA debt service", "$6,747/mo"),
    ], start=14):
        sty(ws, f"A{r}", a); sty(ws, f"B{r}", b)

    sty(ws, "A23", "USE OF SBA PROCEEDS (at January closing)", B, fill=FILL)
    sty(ws, "A24", "Refinance interim bank acquisition note (payoff incl. accrued interest)")
    sty(ws, "B24", 461676, fmt=MONEY)
    sty(ws, "A25", "SBA soft costs (packaging $2,500 + guaranty fee est. + legal)")
    sty(ws, "B25", 15000, fmt=MONEY)
    sty(ws, "A26", "Additional working capital")
    sty(ws, "B26", "=B28-B24-B25", fmt=MONEY)
    sty(ws, "A27", "Guaranty fee est.: 75% x $500K x 1.7% =")
    sty(ws, "B27", "=ROUND(500000*0.75*0.017,0)", fmt=MONEY)
    sty(ws, "A28", "TOTAL SBA LOAN", B); sty(ws, "B28", 500000, fmt=MONEY).font = B
    sty(ws, "A29", "Why this refinance is SBA-clean: the bank note funded the business acquisition;", IT)
    sty(ws, "A30", "SBA funds at the 100% ownership transfer (42 CFR 424.550(b) date), so no seller guaranties apply.", IT)

    sty(ws, "A32", "PROPOSED SBA TERMS & COVERAGE (EBITDA per Rev 4.10 proforma, pre-debt-service)", B, fill=FILL)
    sty(ws, "A33", "Rate (est. Prime + spread; lender to confirm)"); sty(ws, "B33", 0.105, fmt="0.00%")
    sty(ws, "A34", "Term (years)"); sty(ws, "B34", 10)
    sty(ws, "A35", "Monthly payment"); sty(ws, "B35", "=-PMT(B33/12,B34*12,B28)", fmt=MONEY)
    sty(ws, "A36", "Annual debt service"); sty(ws, "B36", "=B35*12", fmt=MONEY)
    sty(ws, "A37", "Year"); sty(ws, "B37", "EBITDA", B); sty(ws, "C37", "Debt service", B); sty(ws, "D37", "DSCR", B)
    for i, (yrl, e) in enumerate([("Year 1", 299195), ("Year 2", 695981), ("Year 3", 985429)]):
        r = 38 + i
        sty(ws, f"A{r}", yrl); sty(ws, f"B{r}", e, fmt=MONEY)
        sty(ws, f"C{r}", "=$B$36", fmt=MONEY)
        sty(ws, f"D{r}", f"=B{r}/C{r}", fmt="0.00x")
    sty(ws, "A42", "Interim-period note: Oct-Dec debt service is the bank note at $15,211/mo, fully covered", IT)
    sty(ws, "A43", "by operations + working capital per the monthly proforma (cash never goes negative in base).", IT)


def tab_wc(wb):
    ws = wb.create_sheet("Working Capital Detail")
    ws.column_dimensions["A"].width = 42
    for i in range(2, 15):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 11

    sty(ws, "A1", "WORKING CAPITAL - COMPONENTS & MONTHLY DRAW TIMELINE", H)
    sty(ws, "A2", "Responds to SourceFunding note of 6/3 (break WC into components with a draw timeline). Borrower "
                  "consents to lender-controlled disbursement tied to census milestones.", IT)
    sty(ws, "A3", "Funding sequence: equity ($125K after the down payment) covers Aug-Sep; the bank-note surplus "
                  "(~$123K) covers Oct-Dec; the SBA surplus (~$23K net of soft costs) tops up at January closing. "
                  "Base case cash never goes negative (min month ~$3.4K - tight, disclosed).", IT)

    labels = ["Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26", "Jan-27",
              "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27", "Jul-27"]
    sty(ws, "A5", "Component / month", B, fill=FILL)
    for i, m in enumerate(labels):
        c = ws.cell(row=5, column=2 + i, value=m); c.font = B; c.fill = FILL
    sty(ws, "N5", "Total", B, fill=FILL)
    rows = [
        ("Clinical payroll & benefits (net of collections)", [24000, 21000, 17000, 13000, 10000, 8000, 6000, 4000, 2000, 0, 0, 0]),
        ("Administrative payroll (net of collections)",      [10000, 9000, 7000, 6000, 5000, 4000, 2000, 2000, 0, 0, 0, 0]),
        ("Patient care costs (pharmacy, DME, supplies)",     [6000, 5500, 5000, 4000, 3500, 2500, 2000, 1500, 0, 0, 0, 0]),
        ("Rent + CAM (13387 Hwy 69 N per executed lease)",   [2143, 2143, 2143, 2143, 2143, 2143, 2143, 2149, 0, 0, 0, 0]),
        ("Insurance (GL/PL, workers comp)",                  [2000, 1750, 1750, 1500, 1500, 1250, 1250, 1000, 0, 0, 0, 0]),
        ("Billing, EMR, IT & communications",                [1600, 1500, 1400, 1300, 1200, 1200, 1200, 1450, 0, 0, 0, 0]),
        ("Contingency (unallocated)",                        [0, 0, 2000, 2000, 2500, 2500, 2000, 2000, 1000, 1000, 0, 0]),
    ]
    r = 6
    for name, vals in rows:
        sty(ws, f"A{r}", name)
        for i, v in enumerate(vals):
            ws.cell(row=r, column=2 + i, value=v).number_format = MONEY
        ws[f"N{r}"] = f"=SUM(B{r}:M{r})"; ws[f"N{r}"].number_format = MONEY
        r += 1
    last = r - 1
    sty(ws, f"A{r}", "Monthly draw", B)
    for i in range(12):
        col = openpyxl.utils.get_column_letter(2 + i)
        ws.cell(row=r, column=2 + i, value=f"=SUM({col}6:{col}{last})").number_format = MONEY
        ws.cell(row=r, column=2 + i).font = B
    ws[f"N{r}"] = f"=SUM(N6:N{last})"; ws[f"N{r}"].number_format = MONEY; ws[f"N{r}"].font = B
    draw_row = r; r += 1
    sty(ws, f"A{r}", "Cumulative draw", B)
    ws.cell(row=r, column=2, value=f"=B{draw_row}").number_format = MONEY
    for i in range(1, 12):
        col = openpyxl.utils.get_column_letter(2 + i)
        prev = openpyxl.utils.get_column_letter(1 + i)
        ws.cell(row=r, column=2 + i, value=f"={prev}{r}+{col}{draw_row}").number_format = MONEY
    r += 2
    sty(ws, f"A{r}", "Notes:", B)
    for note in [
        "Amounts are the projected monthly cash-flow SHORTFALL (operating costs less collections), not gross costs;",
        "gross costs and collections are in the attached Rev 4.10 proforma (36 monthly periods, formula-driven).",
        "Draws taper to zero by month 9-10 as Medicare collections catch and pass payroll.",
        "Owner-deferral lever: the three owner-operators defer salary until break-even census (~18, crossed month 2-3);",
        "$42,500 accrues and is repaid in January - modeled explicitly on the proforma's Staffing tab.",
        "Prior structure carried a $672,000 reserve; this plan needs $235,000 because the license bills day one,",
        "equity covers the down payment, and the bank note (then SBA) covers the acquisition.",
    ]:
        sty(ws, f"A{r+1}", note, Font(size=9)); r += 1
    sty(ws, "A24", "WORKING CAPITAL TOTAL", B, fill=FILL)
    sty(ws, "B24", f"=N{draw_row}", fmt=MONEY).font = B


def tab_timeline(wb):
    ws = wb.create_sheet("Financing Timeline")
    ws.column_dimensions["A"].width = 40
    for col in "BCDEF": ws.column_dimensions[col].width = 17

    sty(ws, "A1", "SELLER PAYOFF & DEBT SEQUENCE (per MIPA 7/14/2026 + Path A refinance)", H)
    sty(ws, "A2", "Seller note prepayable without penalty | payments to DHJR, LP as sellers' agent", IT)
    sty(ws, "A4", "Event", B, fill=FILL); sty(ws, "B4", "Date", B, fill=FILL)
    sty(ws, "C4", "Amount", B, fill=FILL); sty(ws, "D4", "Funded by", B, fill=FILL)
    events = [
        ("License down payment (49% transfers)", "~8/1/2026", 125000, "Equity injection"),
        ("Seller balance paid IN FULL (principal)", "Sept 2026", 375000, "Interim bank note ($500K/6%/36)"),
        ("Seller accrued interest at payoff", "Sept 2026", "=ROUND(375000*0.06/12*2,0)", "Interim bank note"),
        ("Bank note service (interim)", "Oct-Dec 2026", "15,211/mo", "Operations"),
        ("Bank note REFINANCED by SBA 7(a) (payoff)", "1/15/2027", 461676, "SBA loan proceeds"),
        ("Remaining 51% of membership interests transfer", "1/15/2027", "-", "36 months after CCN date 1/8/2024"),
        ("Single SBA debt service thereafter", "Feb 2027 on", "6,747/mo", "Operations"),
    ]
    for i, (ev, dt, amt, src) in enumerate(events):
        r = 5 + i
        sty(ws, f"A{r}", ev); sty(ws, f"B{r}", dt)
        sty(ws, f"C{r}", amt, fmt=MONEY if isinstance(amt, (int, float)) or str(amt).startswith("=") else None)
        sty(ws, f"D{r}", src)
    sty(ws, "A13", "Sellers are paid in full in September - four months earlier than the MIPA requires. The SBA")
    sty(ws, "A14", "loan then refinances BANK debt (not seller debt) at the 100% ownership transfer, which avoids")
    sty(ws, "A15", "the seller-guaranty rules on partial changes of ownership and simplifies lender placement.")


def tab_injection(wb):
    ws = wb.create_sheet("Equity Injection Trace")
    ws.column_dimensions["A"].width = 52
    for col in "BCDE": ws.column_dimensions[col].width = 16

    sty(ws, "A1", "SBA EQUITY INJECTION - DOLLAR-FOR-DOLLAR TRACE", H)
    sty(ws, "A2", "The $250,000 equity is the SBA injection AND the cash that launches the project. SBA counts "
                  "documented prior expenditure on project costs as injection; this tab is the documentation map.", IT)

    sty(ws, "A4", "SOURCES", B, fill=FILL)
    sty(ws, "A5", "James Bullard - wire 1 (CONFIRMED 5/7/2026, to Mercury ****1275)"); sty(ws, "B5", 100000, fmt=MONEY)
    sty(ws, "A6", "James Bullard - wire 2 (committed; date TBD - must land & be documented)"); sty(ws, "B6", 95000, fmt=MONEY)
    sty(ws, "A7", "Geoff Schackmann - owner cash (date TBD - must land & be documented)"); sty(ws, "B7", 55000, fmt=MONEY)
    sty(ws, "A8", "TOTAL EQUITY SOURCES", B); sty(ws, "B8", "=SUM(B5:B7)", fmt=MONEY).font = B

    sty(ws, "A10", "USES (traced to the Rev 4.10 monthly cash flow)", B, fill=FILL)
    sty(ws, "A11", "License down payment at closing (~8/1/2026; MIPA receipt = documentation)"); sty(ws, "B11", 125000, fmt=MONEY)
    sty(ws, "A12", "Ramp working capital Aug-Dec (payroll, rent, patient care ahead of collections)"); sty(ws, "B12", "=B8-B11", fmt=MONEY)
    sty(ws, "A13", "TOTAL USES (all INSIDE the project)", B); sty(ws, "B13", "=B11+B12", fmt=MONEY).font = B

    sty(ws, "A15", "SBA INJECTION TEST (at January closing)", B, fill=FILL)
    sty(ws, "A16", "Total project cost"); sty(ws, "B16", 750000, fmt=MONEY)
    sty(ws, "A17", "Equity injected (all documented, all spent in-project)"); sty(ws, "B17", "=B8", fmt=MONEY)
    sty(ws, "A18", "Injection %", B); sty(ws, "B18", "=B17/B16", fmt=PCT).font = B
    sty(ws, "A19", "SOP 50 10 minimum for startups/changes of ownership"); sty(ws, "B19", 0.10, fmt=PCT)
    sty(ws, "A20", "Headroom vs minimum", B); sty(ws, "B20", "=B18-B19", fmt=PCT).font = B

    sty(ws, "A22", "DOCUMENTATION CHECKLIST (each item required before the injection counts)", B, fill=FILL)
    for i, x in enumerate([
        "1. Bullard wire 1: wire confirmation 5/7/2026 (ON FILE) + Bullard bank statement covering 30+ days prior",
        "2. Bullard wire 2: wire confirmation + 30-day-prior source statement (PENDING - must complete before SBA close)",
        "3. Schackmann $55K: deposit evidence + source statement (PENDING)",
        "4. Mercury account statements showing funds on deposit and the $125K down-payment disbursement",
        "5. MIPA + settlement receipt for the $125,000 down payment",
        "6. Monthly operating statements tying the working-capital spend to the project",
    ], start=23):
        sty(ws, f"A{i}", x, Font(size=9))
    sty(ws, "A30", "RULES THAT BITE:", B)
    sty(ws, "A31", "Injection must be TRUE EQUITY - proceeds of the investor NOTE raise (debt) do NOT count.", Font(size=9, bold=True))
    sty(ws, "A32", "Money must be spent inside the project (license, WC, soft costs) - it all is, per the trace above.", Font(size=9))


if __name__ == "__main__":
    wb = openpyxl.Workbook()
    tab_structure(wb)
    tab_wc(wb)
    tab_timeline(wb)
    tab_injection(wb)
    wb.save(OUT)
    print("wrote", OUT)
