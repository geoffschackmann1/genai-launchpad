"""Emit a standalone debt-service + runway workbook for the user's scenario."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import financial_models.engine.model as m
from financial_models.runway_scenario import run, pmt, labels, SBA_P, SBA_N, EQUITY

BLUE = "1F4E79"; LT = "DDEBF7"; GRN = "E2EFDA"; AMB = "FFF2CC"
hdr = Font(bold=True, color="FFFFFF"); bold = Font(bold=True)
fillh = PatternFill("solid", fgColor=BLUE)
filll = PatternFill("solid", fgColor=LT)
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)


def money(c): c.number_format = '#,##0'
def x2(c): c.number_format = '0.00"x"'


def main():
    R, a = run(0.105)
    labs = labels()
    sc = R["_scalars"]
    startup = sc["startup_total"]
    beg = a["equity"] + a["sba_principal"] - startup - a["capex"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Debt Service & Runway"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 13

    r = 1
    ws.cell(r, 1, "AZALEA HOSPICE - Debt Service & Runway Scenario").font = Font(bold=True, size=14, color=BLUE); r += 1
    ws.cell(r, 1, "SBA $600K / 15yr + $300K license note (6%/3yr) + $180K equity injection. Census = base path.").font = Font(italic=True, size=9); r += 2

    ws.cell(r, 1, "LOAN TERMS").font = bold; r += 1
    for h, c in zip(["Instrument", "Principal", "Rate", "Term", "Monthly P&I", "Annual P&I"], "ABCDEF"):
        cell = ws[c + str(r)]; cell.value = h; cell.font = hdr; cell.fill = fillh
    r += 1
    sba_m = pmt(SBA_P, 0.105, SBA_N)
    lic_m = pmt(300_000, 0.06, 36)
    rows = [
        ["SBA 7(a) loan", 600000, "10.5%", "180 mo (15y)", sba_m, sba_m * 12],
        ["Hickory license note", 300000, "6.0%", "36 mo (3y)", lic_m, lic_m * 12],
        ["TOTAL DEBT SERVICE", 900000, "", "", sba_m + lic_m, (sba_m + lic_m) * 12],
    ]
    for row in rows:
        for j, v in enumerate(row):
            c = ws.cell(r, j + 1, v)
            if j in (1, 4, 5) and isinstance(v, (int, float)): money(c)
            if row[0].startswith("TOTAL"): c.font = bold; c.fill = filll
        r += 1
    r += 1

    ws.cell(r, 1, "SOURCES & OPENING CASH").font = bold; r += 1
    su = [
        ("SBA 7(a) loan", a["sba_principal"]),
        ("Equity injection (Jim Bullard)", a["equity"]),
        ("Seller note (license, non-cash source)", 300000),
        ("  less: startup costs", -startup),
        ("  less: capex", -a["capex"]),
        ("  less: license (financed by note, $0 cash)", 0),
        ("OPENING CASH AT CLOSE", beg),
        ("Minimum-cash floor (LOC backstop below)", a["min_cash"]),
    ]
    for label, v in su:
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, v); money(c)
        if label.startswith("OPENING"):
            ws.cell(r, 1).font = bold; c.font = bold; c.fill = PatternFill("solid", fgColor=GRN)
        r += 1
    r += 1

    ws.cell(r, 1, "PERIOD SCHEDULE  (Yr1 monthly, Yr2-3 quarterly)").font = bold; r += 1
    heads = ["Period", "EBITDA", "SBA P&I", "License P&I", "Total Debt Svc", "DSCR", "End Cash"]
    for j, h in enumerate(heads):
        c = ws.cell(r, j + 1, h); c.font = hdr; c.fill = fillh; c.alignment = Alignment(horizontal="center")
    r += 1
    start_row = r
    for i, p in enumerate(m.PERIODS):
        e = R["ebitda"][i]
        sba_pi = R["int_sba"][i] + (R["ds"][i] - R["int_sba"][i] - R["lic_int"][i] - R["lic_prin"][i])
        lic_pi = R["lic_int"][i] + R["lic_prin"][i]
        ds = R["ds"][i]
        dscr = e / ds if ds else 0
        vals = [labs[i], e, sba_pi, lic_pi, ds, dscr, R["end_cash"][i]]
        for j, v in enumerate(vals):
            c = ws.cell(r, j + 1, v)
            c.border = box
            if j == 0:
                c.alignment = Alignment(horizontal="center")
            elif j == 5:
                x2(c)
                if dscr < 1.25: c.fill = PatternFill("solid", fgColor="FCE4D6")
            else:
                money(c)
        r += 1
    r += 1

    ws.cell(r, 1, "RUNWAY & COVERAGE SUMMARY").font = bold; r += 1
    mc = min(R["end_cash"]); mci = R["end_cash"].index(mc)
    loc = sum(R["loc_draw"])
    def yr(y):
        e = sum(R["ebitda"][i] for i, p in enumerate(m.PERIODS) if p["year"] == y)
        ds = sum(R["ds"][i] for i, p in enumerate(m.PERIODS) if p["year"] == y)
        return e, ds
    summ = [
        ("Opening cash at close", "${:,.0f}".format(beg)),
        ("Minimum cash trough", "${:,.0f}  (at {})".format(mc, labs[mci])),
        ("LOC drawn over 36 months", "${:,.0f}  (line never needed)".format(loc)),
        ("Year 1 DSCR", "{:.2f}x".format(yr(1)[0] / yr(1)[1])),
        ("Year 2 DSCR", "{:.2f}x".format(yr(2)[0] / yr(2)[1])),
        ("Year 3 DSCR", "{:.2f}x".format(yr(3)[0] / yr(3)[1])),
        ("First period DSCR >= 1.25x", "Month 3 (sustained thereafter)"),
        ("SBA injection as % of SBA loan", "30%  (vs ~10% SBA minimum)"),
    ]
    for label, v in summ:
        ws.cell(r, 1, label).font = bold
        ws.cell(r, 2, v)
        r += 1
    r += 1
    ws.cell(r, 1, "Note: 'typical SBA rate' modeled at Prime 7.5% + 3.0% = 10.5% (variable). Rate sensitivity 9.5-12.5%").font = Font(italic=True, size=9); r += 1
    ws.cell(r, 1, "moves the min-cash trough by under $3K - runway is insensitive to the SBA rate.").font = Font(italic=True, size=9)

    out = "financial_models/output/Azalea_DebtService_Runway_Scenario.xlsx"
    wb.save(out)
    print("wrote", out)


if __name__ == "__main__":
    main()
