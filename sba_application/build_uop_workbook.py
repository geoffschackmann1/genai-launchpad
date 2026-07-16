"""UOP follow-up workbook for John Hart / Mary Brownmiller (Refuge structure).

Responds to John's 6/3 preliminary structure sheet and his core ask: break the working
capital into components with an anticipated draw timeline. Old structure had $672K WC;
current structure carries $235K, fully detailed here. Layout mirrors John's template.
All calculation cells are formulas. Output:
sba_application/13_checklist_response_2026-07-16/UOP Azalea Refuge Rev1.00.xlsx
"""
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = "sba_application/13_checklist_response_2026-07-16/UOP Azalea Refuge Rev1.00.xlsx"

H = Font(bold=True, size=12, color="1F3B2D")
B = Font(bold=True)
MONEY = "#,##0"
PCT = "0.0%"
FILL = PatternFill("solid", fgColor="EAF1EC")
THIN = Border(bottom=Side(style="thin", color="999999"))


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
    ws.column_dimensions["A"].width = 46
    for col in "BCDEF": ws.column_dimensions[col].width = 16

    sty(ws, "A1", "PROPOSED SBA 7(a) LOAN STRUCTURE - AZALEA HOSPICE & PALLIATIVE CARE", H)
    sty(ws, "A2", "Tyler Hospice Hold, LLC (EIN 41-4966640) | Refuge Hospice acquisition | 7/16/2026", Font(italic=True, size=9))

    sty(ws, "A4", "USE OF PROCEEDS", B, fill=FILL)
    sty(ws, "A5", "Medicare/Medicaid hospice license - Refuge Hospice, LLC (purchase price)")
    sty(ws, "B5", 500000, fmt=MONEY)
    sty(ws, "A6", "   of which: equity down payment at close (~8/1/2026, 49% interest)")
    sty(ws, "B6", 125000, fmt=MONEY)
    sty(ws, "A7", "   of which: seller-financed balance at 6% (SBA takeout at the 51% transfer, 1/15/2027)")
    sty(ws, "B7", "=B5-B6", fmt=MONEY)
    sty(ws, "A8", "Working capital reserve (detail + draw timeline on next tab)")
    sty(ws, "B8", "='Working Capital Detail'!B24", fmt=MONEY)
    sty(ws, "A9", "Soft costs (packaging, SBA guaranty fee, legal/closing)")
    sty(ws, "B9", "=B13+B14+B15+B16", fmt=MONEY)
    sty(ws, "A10", "TOTAL PROJECT COST", B)
    sty(ws, "B10", "=B5+B8+B9", fmt=MONEY).font = B

    sty(ws, "A12", "Soft-cost detail:", Font(italic=True, size=9))
    sty(ws, "A13", "   7(a) loan packaging"); sty(ws, "B13", 2500, fmt=MONEY)
    sty(ws, "A14", "   SBA guaranty fee (estimate - see computation at F13; lender to confirm)")
    sty(ws, "B14", "=F15", fmt=MONEY)
    sty(ws, "A15", "   Legal & closing (rounded)"); sty(ws, "B15", 2400, fmt=MONEY)
    sty(ws, "A16", "   Contingency / rounding"); sty(ws, "B16", "=15000-B13-B14-B15", fmt=MONEY)

    sty(ws, "D12", "Guaranty fee computation", B)
    sty(ws, "D13", "Guaranty %"); sty(ws, "F13", 0.75, fmt=PCT)
    sty(ws, "D14", "Guaranteed amount"); sty(ws, "F14", "=B20*F13", fmt=MONEY)
    sty(ws, "D15", "Fee (est. 1.7% of guaranteed portion)"); sty(ws, "F15", "=ROUND(F14*0.017,0)", fmt=MONEY)

    sty(ws, "A18", "FINANCE STRUCTURE", B, fill=FILL)
    sty(ws, "A19", "Source"); sty(ws, "B19", "Amount", B); sty(ws, "C19", "% of project", B)
    sty(ws, "A20", "SBA 7(a) loan (funds Jan 2027: seller balloon takeout + working capital)")
    sty(ws, "B20", 500000, fmt=MONEY); sty(ws, "C20", "=B20/$B$10", fmt=PCT)
    sty(ws, "A21", "Borrower equity injection - James Bullard ($195,000; $100K wired 5/7/2026)")
    sty(ws, "B21", 195000, fmt=MONEY); sty(ws, "C21", "=B21/$B$10", fmt=PCT)
    sty(ws, "A22", "Borrower equity injection - Geoff Schackmann ($55,000)")
    sty(ws, "B22", 55000, fmt=MONEY); sty(ws, "C22", "=B22/$B$10", fmt=PCT)
    sty(ws, "A23", "TOTAL SOURCES", B)
    sty(ws, "B23", "=SUM(B20:B22)", fmt=MONEY).font = B
    sty(ws, "C23", "=B23/$B$10", fmt=PCT)
    sty(ws, "A24", "Check: sources = uses (0 = tied)")
    sty(ws, "B24", "=B23-B10", fmt=MONEY)

    sty(ws, "A26", "PROPOSED SBA LOAN TERMS & DEBT SERVICE", B, fill=FILL)
    sty(ws, "A27", "Loan amount"); sty(ws, "B27", "=B20", fmt=MONEY)
    sty(ws, "A28", "Interest rate (est. Prime + spread; lender to confirm)"); sty(ws, "B28", 0.105, fmt="0.00%")
    sty(ws, "A27", "Loan amount (funds ~1/15/2027, concurrent with the 51% transfer)")
    sty(ws, "A29", "Term (years)"); sty(ws, "B29", 10)
    sty(ws, "A30", "Monthly payment"); sty(ws, "B30", "=-PMT(B28/12,B29*12,B27)", fmt=MONEY)
    sty(ws, "A31", "Annual debt service"); sty(ws, "B31", "=B30*12", fmt=MONEY)

    sty(ws, "A33", "COVERAGE (EBITDA per Rev 3.00 proforma, before debt service)", B, fill=FILL)
    sty(ws, "A34", "Year"); sty(ws, "B34", "EBITDA", B); sty(ws, "C34", "Debt service", B); sty(ws, "D34", "DSCR", B)
    for i, (yr, e) in enumerate([("Year 1", 263000), ("Year 2", 680000), ("Year 3", 976000)]):
        r = 35 + i
        sty(ws, f"A{r}", yr); sty(ws, f"B{r}", e, fmt=MONEY)
        sty(ws, f"C{r}", "=$B$31", fmt=MONEY)
        sty(ws, f"D{r}", f"=B{r}/C{r}", fmt="0.00x")
    sty(ws, "A39", "BASE CASE: SBA funds January 2027 at the 51% transfer (a clean 100% change of ownership;")
    sty(ws, "A40", "no seller guaranties required). September funding is an UPSIDE if a lender can paper the")
    sty(ws, "A41", "partial-CHOW guaranty rules. The Rev 3.10 proforma conservatively models the takeout at a")
    sty(ws, "A42", "36-month amortization ($7,957/mo from Feb); actual 7(a) terms above are the cheaper case.")


def tab_wc(wb):
    ws = wb.create_sheet("Working Capital Detail")
    ws.column_dimensions["A"].width = 40
    for i in range(2, 15):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 11

    sty(ws, "A1", "WORKING CAPITAL RESERVE - COMPONENTS & ANTICIPATED DRAW TIMELINE", H)
    sty(ws, "A2", "Responds to SourceFunding note of 6/3: components + draw schedule. Borrower consents to "
                  "lender-controlled disbursement tied to census milestones.", Font(italic=True, size=9))
    sty(ws, "A3", "Purpose: fund operations through the Medicare payment lag (NOE-to-cash ~30-60 days) while census "
                  "ramps to break-even (~17-18 patients, crossed month 2-3 per the Rev 3.10 proforma). BASE CASE: the "
                  "SBA loan funds ~1/15/2027; amounts shown before then are bridged on an interim LOC (peak need "
                  "~$183K base / ~$375K downside) and repaid from the reserve at SBA funding.", Font(italic=True, size=9))

    # months across: M1..M12 (Aug 2026 - Jul 2027)
    labels = ["Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26", "Jan-27",
              "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27", "Jul-27"]
    sty(ws, "A5", "Component / month", B, fill=FILL)
    for i, m in enumerate(labels):
        c = ws.cell(row=5, column=2 + i, value=m); c.font = B; c.fill = FILL
    sty(ws, "N5", "Total", B, fill=FILL)

    # draw rows: (name, monthly amounts M1..M12)
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
    draw_row = r
    r += 1
    sty(ws, f"A{r}", "Cumulative draw", B)
    ws.cell(row=r, column=2, value=f"=B{draw_row}").number_format = MONEY
    for i in range(1, 12):
        col = openpyxl.utils.get_column_letter(2 + i)
        prev = openpyxl.utils.get_column_letter(1 + i)
        ws.cell(row=r, column=2 + i, value=f"={prev}{r}+{col}{draw_row}").number_format = MONEY
    cum_row = r
    r += 1
    sty(ws, f"A{r}", "Remaining reserve", B)
    for i in range(12):
        col = openpyxl.utils.get_column_letter(2 + i)
        ws.cell(row=r, column=2 + i, value=f"=$B$24-{col}{cum_row}").number_format = MONEY

    r += 2
    sty(ws, f"A{r}", "Notes:", B)
    for note in [
        "Amounts are the projected cash-flow SHORTFALL each month (operating costs less collections), not gross costs;",
        "gross monthly operating costs and collections are in the attached Rev 3.10 proforma (36 monthly periods).",
        "Draws taper to zero by month 9-10 as Medicare collections catch and pass payroll; reserve is a bridge, not a cushion.",
        "Prior structure carried a $672,000 reserve; the current structure needs $235,000 because the license is",
        "billing-ready day one and equity covers the down payment plus the seller installments through December.",
        "Base case peak interim-LOC need before SBA funding is ~$183K; the documented slow-census downside needs a",
        "~$375K facility - disclosed, with mitigants (census recovery levers, owner deferral, September takeout upside).",
    ]:
        sty(ws, f"A{r+1}", note, Font(size=9)); r += 1

    sty(ws, "A24", "WORKING CAPITAL RESERVE TOTAL", B, fill=FILL)
    sty(ws, "B24", f"=N{draw_row}", fmt=MONEY).font = B


def tab_seller(wb):
    ws = wb.create_sheet("Seller Note & SBA Takeout")
    ws.column_dimensions["A"].width = 34
    for col in "BCDEF": ws.column_dimensions[col].width = 17

    sty(ws, "A1", "SELLER NOTE & SBA TAKEOUT TIMELINE (per MIPA dated 7/14/2026)", H)
    sty(ws, "A2", "Refuge Hospice, LLC | $500,000 price | prepayable without penalty | payments to DHJR, LP as sellers' agent",
        Font(italic=True, size=9))

    sty(ws, "A4", "Event", B, fill=FILL); sty(ws, "B4", "Date", B, fill=FILL)
    sty(ws, "C4", "Payment", B, fill=FILL); sty(ws, "D4", "Note balance after", B, fill=FILL)
    sty(ws, "E4", "Funded by", B, fill=FILL)
    events = [
        ("Down payment (49% interest transfers)", "~8/1/2026", 125000, "=500000-C5", "Equity injection"),
        ("Monthly installment ($31,250 incl. 6% interest)", "9/1/2026", 31250, "", "Equity injection"),
        ("Monthly installment", "10/1/2026", 31250, "", "Equity injection"),
        ("Monthly installment", "11/1/2026", 31250, "", "Equity injection"),
        ("Monthly installment", "12/1/2026", 31250, "", "Equity injection"),
        ("Final payment per MIPA amortization (Exhibit C)", "1/15/2027", 258489.46, 0, "SBA 7(a) proceeds"),
    ]
    for i, (ev, dt, pay, bal, src) in enumerate(events):
        r = 5 + i
        sty(ws, f"A{r}", ev); sty(ws, f"B{r}", dt)
        sty(ws, f"C{r}", pay, fmt="#,##0.00" if pay == 258489.46 else MONEY)
        if bal != "": sty(ws, f"D{r}", bal, fmt=MONEY)
        sty(ws, f"E{r}", src)
    sty(ws, "A12", "BASE CASE: the SBA 7(a) loan funds ~1/15/2027, concurrent with the transfer of the remaining 51%")
    sty(ws, "A13", "of membership interests - the first date past 36 months from the company's CMS certification")
    sty(ws, "A14", "effective date (1/8/2024) per 42 CFR 424.550(b). Funding at the 100% transfer avoids the seller-")
    sty(ws, "A15", "guaranty requirements that apply to SBA loans made during a partial change of ownership.")
    sty(ws, "A17", "UPSIDE - September takeout (note is prepayable without penalty):", B)
    sty(ws, "A18", "If a lender can fund September 2026, payoff is $375,000 + ~$1,875 accrued interest, saving four")
    sty(ws, "A19", "installments and ~$8,500 of seller interest. Requires the lender to paper seller guaranties under")
    sty(ws, "A20", "the partial-change-of-ownership rules, so it is presented as upside, not base.")


if __name__ == "__main__":
    wb = openpyxl.Workbook()
    tab_structure(wb)
    tab_wc(wb)
    tab_seller(wb)
    wb.save(OUT)
    print("wrote", OUT)
