"""UOP workbook for John Hart / Mary Brownmiller - Avant structure (Rev 5.00 basis).

Story: equity $250K funds the operating ramp (no down payment - fully deferred seller note);
Avant seller (Kimberly Carlisle) carries the full $300K at 6% simple with all payments deferred;
the SBA 7(a) $500K funds month 2 (license 100% in borrower name at CHOW) and retires the seller
SBA funds month 2 and retires the seller note (~$303,000 incl. accrued interest); single $6,747/mo payment after.
The $250K equity is the SBA injection (33% of the $750K project), traced
dollar-for-dollar on the Equity Injection Trace tab.
Output: sba_application/13_checklist_response_2026-07-16/UOP Azalea Avant Rev3.00.xlsx
"""
import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side

OUT = "sba_application/13_checklist_response_2026-07-16/UOP Azalea Avant Rev3.00.xlsx"

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
    sty(ws, "A2", "Tyler Hospice Hold, LLC (EIN 41-4966640) | Avant Hospice acquisition | Rev 3.00 basis: Rev 5.00 AVANT proforma | 8/14/2026", IT)

    sty(ws, "A4", "OVERALL PROJECT (all phases)", B, fill=FILL)
    sty(ws, "A5", "Avant Hospice, LLC license purchase price")
    sty(ws, "B5", 300000, fmt=MONEY)
    sty(ws, "A6", "Working capital (ramp payroll ahead of Medicare collections + prepayment review)")
    sty(ws, "B6", 435000, fmt=MONEY)
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
        ("CHOW close - 100% of membership interests transfer; NO down payment", "$0 at close"),
        ("From close - Equity ($250,000) funds the operating ramp", "$250,000"),
        ("Months 1-2 - Seller carries the full balance at 6% simple; nothing due", "$300,000 carried"),
        ("Month 2 - SBA 7(a) funds (license already 100% in borrower name)", "$500,000"),
        ("Month 2 - SBA retires the seller note (principal + accrued interest)", "~$303,000 payoff"),
        ("Month 2 - SBA surplus: soft costs + additional working capital", "~$197,000"),
        ("From month 3 - single SBA debt service", "$6,747/mo"),
        ("Fallback if SBA slips - seller deferral continues until prepayment review clears", "no payments due"),
    ], start=14):
        sty(ws, f"A{r}", a); sty(ws, f"B{r}", b)

    sty(ws, "A23", "USE OF SBA PROCEEDS (at month-2 funding)", B, fill=FILL)
    sty(ws, "A24", "Retire seller acquisition note (payoff incl. accrued interest)")
    sty(ws, "B24", 303000, fmt=MONEY)
    sty(ws, "A25", "SBA soft costs (packaging $2,500 + guaranty fee est. + legal)")
    sty(ws, "B25", 15000, fmt=MONEY)
    sty(ws, "A26", "Additional working capital")
    sty(ws, "B26", "=B28-B24-B25", fmt=MONEY)
    sty(ws, "A27", "Guaranty fee est.: 75% x $500K x 1.7% =")
    sty(ws, "B27", "=ROUND(500000*0.75*0.017,0)", fmt=MONEY)
    sty(ws, "A28", "TOTAL SBA LOAN", B); sty(ws, "B28", 500000, fmt=MONEY).font = B
    sty(ws, "A29", "Why this is SBA-clean: 100% of the membership interests transfer at CHOW close, so the", IT)
    sty(ws, "A30", "license is fully in the borrower name BEFORE SBA funds - no split closing, no waiting period.", IT)

    sty(ws, "A32", "PROPOSED SBA TERMS & COVERAGE (EBITDA per Rev 5.00 AVANT proforma, pre-debt-service)", B, fill=FILL)
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
    sty(ws, "A42", "Interim-period note: no debt service at all until SBA funds (seller payments fully deferred);", IT)
    sty(ws, "A43", "base-case cash never goes negative (minimum month ~$192K per the monthly proforma).", IT)


def tab_wc(wb):
    ws = wb.create_sheet("Working Capital Detail")
    ws.column_dimensions["A"].width = 42
    for i in range(2, 15):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 11

    sty(ws, "A1", "WORKING CAPITAL - COMPONENTS & MONTHLY DRAW TIMELINE", H)
    sty(ws, "A2", "Responds to SourceFunding note of 6/3 (break WC into components with a draw timeline). Borrower "
                  "consents to lender-controlled disbursement tied to census milestones.", IT)
    sty(ws, "A3", "Funding sequence: equity ($250K, no down payment) covers the early ramp; the SBA surplus "
                  "(~$182K net of soft costs) lands at month 2. "
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
        "gross costs and collections are in the attached Rev 5.00 AVANT proforma (36 monthly periods, formula-driven).",
        "Draws taper to zero by month 9-10 as Medicare collections catch and pass payroll.",
        "Owner-deferral lever: the three owner-operators defer salary until break-even census (~18, crossed month 2-3);",
        "$42,500 accrues and is repaid in January - modeled explicitly on the proforma's Staffing tab.",
        "The $435,000 reserve is deliberately large (58% of project): the target has never billed Medicare, so",
        "first claims enter prepayment review - the reserve carries payroll through that window (base 2 mo / downside 4).",
    ]:
        sty(ws, f"A{r+1}", note, Font(size=9)); r += 1
    sty(ws, "A24", "WORKING CAPITAL TOTAL", B, fill=FILL)
    sty(ws, "B24", f"=N{draw_row}", fmt=MONEY).font = B


def tab_timeline(wb):
    ws = wb.create_sheet("Financing Timeline")
    ws.column_dimensions["A"].width = 40
    for col in "BCDEF": ws.column_dimensions[col].width = 17

    sty(ws, "A1", "SELLER PAYOFF & DEBT SEQUENCE (Avant - deferred seller note, SBA takeout)", H)
    sty(ws, "A2", "Seller note prepayable without penalty | seller: Kimberly Carlisle, sole member", IT)
    sty(ws, "A4", "Event", B, fill=FILL); sty(ws, "B4", "Date", B, fill=FILL)
    sty(ws, "C4", "Amount", B, fill=FILL); sty(ws, "D4", "Funded by", B, fill=FILL)
    events = [
        ("CHOW close - 100% of membership interests transfer", "At CHOW approval", "-", "No cash due at close"),
        ("Seller carries full balance (6% simple, all payments deferred)", "Months 1-2", 300000, "Seller financing"),
        ("Seller note RETIRED by SBA 7(a) (principal + accrued interest)", "Month 2", "=ROUND(300000*(1+0.06/12*2),0)", "SBA loan proceeds"),
        ("SBA surplus to soft costs + working capital", "Month 2", "=500000-303000", "SBA loan proceeds"),
        ("Single SBA debt service thereafter", "Month 3 on", "6,747/mo", "Operations"),
        ("Fallback if SBA slips: deferral continues until prepayment review clears", "-", "-", "Seller terms"),
    ]
    for i, (ev, dt, amt, src) in enumerate(events):
        r = 5 + i
        sty(ws, f"A{r}", ev); sty(ws, f"B{r}", dt)
        sty(ws, f"C{r}", amt, fmt=MONEY if isinstance(amt, (int, float)) or str(amt).startswith("=") else None)
        sty(ws, f"D{r}", src)
    sty(ws, "A13", "The seller is paid in full at SBA funding (month 2) - and owes nothing before that. Because")
    sty(ws, "A14", "100% of ownership transfers at CHOW close, SBA funds AFTER the license is fully in the")
    sty(ws, "A15", "borrower name - no partial-change-of-ownership seller-guaranty complications, simple placement.")


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

    sty(ws, "A10", "USES (traced to the Rev 5.00 AVANT monthly cash flow)", B, fill=FILL)
    sty(ws, "A11", "License down payment (none - seller note fully deferred)"); sty(ws, "B11", 0, fmt=MONEY)
    sty(ws, "A12", "Ramp working capital (payroll, rent, patient care ahead of collections + prepayment review)"); sty(ws, "B12", "=B8-B11", fmt=MONEY)
    sty(ws, "A13", "TOTAL USES (all INSIDE the project)", B); sty(ws, "B13", "=B11+B12", fmt=MONEY).font = B

    sty(ws, "A15", "SBA INJECTION TEST (at month-2 funding)", B, fill=FILL)
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
        "4. Mercury account statements showing funds on deposit and working-capital disbursements",
        "5. Executed Avant MIPA / payment-confirmation schedule (no down payment to document)",
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
