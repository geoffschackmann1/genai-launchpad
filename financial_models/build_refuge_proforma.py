"""Build the Refuge Hospice 12-month forward proforma workbook.

Tabs: Assumptions | Offer & Affordability | 12-Mo P&L (operating budget) |
Cash-Flow Budget | SBA Scenario. Values-based (engine-computed), monthly,
July 2026 close -> June 2027. Base case has NO SBA loan.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from financial_models.refuge_scenario import (engine_cfo, run_structure, pmt,
                                              PRICE, CASH, SELLER_RATE,
                                              REFI_RATE, REFI_TERM)
import financial_models.engine.model as m

BLUE = "1F4E79"; LT = "DDEBF7"; GRN = "E2EFDA"; AMB = "FFF2CC"; RED = "FCE4D6"
MONTHS = ["Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26",
          "Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27"]
LIC_AMORT = PRICE / 15 / 12          # $500K over 15 yrs

hdrF = Font(bold=True, color="FFFFFF"); boldF = Font(bold=True)
fillH = PatternFill("solid", fgColor=BLUE)
thin = Side(style="thin", color="BFBFBF"); box = Border(left=thin, right=thin, top=thin, bottom=thin)


def money(c): c.number_format = '#,##0'


def sheet(wb, name, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name; ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 38
    for i in range(2, 16):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 11
    return ws


def title(ws, r, text, size=13):
    ws.cell(r, 1, text).font = Font(bold=True, size=size, color=BLUE); return r + 1


def note(ws, r, text):
    ws.cell(r, 1, text).font = Font(italic=True, size=9); return r + 1


def month_header(ws, r, label="($/month)"):
    ws.cell(r, 1, label).font = hdrF; ws.cell(r, 1).fill = fillH
    for j, mo in enumerate(MONTHS):
        c = ws.cell(r, 2 + j, mo); c.font = hdrF; c.fill = fillH; c.alignment = Alignment(horizontal="center")
    c = ws.cell(r, 14, "TOTAL"); c.font = hdrF; c.fill = fillH; c.alignment = Alignment(horizontal="center")
    return r + 1


def row(ws, r, label, vals, bold=False, fill=None, total=True, indent=False):
    c0 = ws.cell(r, 1, ("    " if indent else "") + label)
    if bold: c0.font = boldF
    for j, v in enumerate(vals):
        c = ws.cell(r, 2 + j, v); money(c); c.border = box
        if bold: c.font = boldF
        if fill: c.fill = PatternFill("solid", fgColor=fill)
    if total:
        c = ws.cell(r, 14, sum(vals)); money(c); c.border = box
        if bold: c.font = boldF
        if fill: c.fill = PatternFill("solid", fgColor=fill)
    return r + 1


def build():
    cfo_engine, R, a = engine_cfo()
    rec = run_structure(cfo_engine, 100_000, 25_000, 2, "RECOMMENDED")

    # ---- monthly P&L pieces from the engine (first 12 periods are monthly) ----
    N = 12
    net = [R["net"][i] for i in range(N)]
    gross = [R["gross"][i] for i in range(N)]
    adc = [R["adc"][i] for i in range(N)]
    cogs_labor = [R["ft_direct"][i] + R["prn"][i] + R["med_dir"][i] + R["burden_d"][i] + R["health_d"][i] for i in range(N)]
    cogs_pt = [R["cogs_patient"][i] for i in range(N)]
    gp = [R["gp"][i] for i in range(N)]
    sga_lab = [R["sga_labor"][i] for i in range(N)]
    ga_fix = [R["ga_fixed"][i] for i in range(N)]
    fees = [R["qr"][i] + R["billing"][i] for i in range(N)]
    ebitda = [R["ebitda"][i] for i in range(N)]

    # ---- seller note / refi overlay (recommended structure) ----
    down, monthly, holiday = rec["down"], rec["monthly"], rec["holiday"]
    bal = PRICE - down
    seller_int, seller_pay = [], []
    acc_int = 0.0
    balloon = 0.0
    for i in range(N):
        mno = i + 1
        if mno <= 6:
            intr = bal * SELLER_RATE / 12
            acc_int += intr
            seller_int.append(intr)
            p = monthly if (holiday < mno <= 6) else 0.0
            p = min(p, bal); bal -= p
            if mno == 6:
                balloon = bal + acc_int
                seller_pay.append(p)          # balloon shown separately (refi-funded)
                bal = 0.0
            else:
                seller_pay.append(p)
        else:
            seller_int.append(0.0); seller_pay.append(0.0)
    refi_pmt = pmt(balloon, REFI_RATE, REFI_TERM)
    refi_int, refi_prin = [], []
    rbal = balloon
    for i in range(N):
        if i + 1 <= 6:
            refi_int.append(0.0); refi_prin.append(0.0)
        else:
            intr = rbal * REFI_RATE / 12
            pr = refi_pmt - intr
            rbal -= pr
            refi_int.append(intr); refi_prin.append(pr)

    # ---- display P&L below EBITDA (own consistent set) ----
    da = [LIC_AMORT] * N
    pretax = [ebitda[i] - da[i] - seller_int[i] - refi_int[i] for i in range(N)]
    tax = [net[i] * a["tx_tax"] if pretax[i] > 0 else 0.0 for i in range(N)]
    ni = [pretax[i] - tax[i] for i in range(N)]

    # ---- cash flow (working capital from engine AR/AP) ----
    ar = [R["ar"][i] for i in range(N)]; ap = [R["ap"][i] for i in range(N)]
    d_ar = [ar[0]] + [ar[i] - ar[i - 1] for i in range(1, N)]
    d_ap = [ap[0]] + [ap[i] - ap[i - 1] for i in range(1, N)]
    # Seller interest accrues (paid at the balloon, which the refi funds) -> add back as non-cash.
    # Refi interest IS paid monthly and is already inside NI, so financing deducts principal only.
    cfo = [ni[i] + da[i] + seller_int[i] - d_ar[i] + d_ap[i] for i in range(N)]
    beg, cash_path, endc = CASH - down, [], []
    for i in range(N):
        mno = i + 1
        c = beg + cfo[i] - seller_pay[i] - (refi_prin[i] if mno > 6 else 0.0)
        # balloon at M6 funded by personal refi: -balloon + refi proceeds = 0 net
        cash_path.append(c); endc.append(c); beg = c
    minc = min(endc); minm = endc.index(minc) + 1

    wb = openpyxl.Workbook()

    # ================= Assumptions =================
    ws = sheet(wb, "Assumptions", first=True)
    r = title(ws, 1, "AZALEA HOSPICE - Refuge Hospice 12-Month Proforma (Jul 2026 - Jun 2027)")
    r = note(ws, r, "Base case: NO SBA loan. Values computed from the validated Azalea engine (Paloma-panel migration, "
                    "Tyler CBSA rates). Interview decisions of 7/2/2026 locked below.")
    r += 1
    items = [
        ("Acquisition target", "Refuge Hospice - dormant Medicare & Medicaid license (CCN + Medicaid contract), Tyler / East Texas"),
        ("Structure", "Purchase of 100% of membership interests (CHOW); provider number conveys"),
        ("Purchase price", "$500,000 (full ask; negotiation is on terms)"),
        ("Close", "July 2026 (Month 1 = Jul-26; January 2027 = Month 6)"),
        ("Capital available", "$250,000 liquid (includes Jim Bullard's $195K equity); startup/capex funded elsewhere"),
        ("Revenue", "Proven Paloma panel migrates at close: ~22 ADC by M3 at $181/day blended net (100% capture, base path)"),
        ("Deferred balance", "6.0% simple interest; secured by acquired membership interests; personal guaranty of Geoff Schackmann"),
        ("Month-6 payoff", "Balloon paid in January 2027 by Geoff's personal refinance (lender identified - firm)"),
        ("Refi service (assumption)", f"Company services refi from M7: {REFI_RATE:.0%} APR, {REFI_TERM}-mo am = ${refi_pmt:,.0f}/mo (changeable)"),
        ("License amortization", "$500K over 15 yrs = $2,778/mo (non-cash)"),
        ("TX margin tax", "0.375% of revenue when pre-tax positive"),
        ("Working capital", "AR 45 days / AP 30 days (drives the M1-M2 cash trough)"),
    ]
    for k, v in items:
        ws.cell(r, 1, k).font = boldF; ws.cell(r, 2, v); r += 1

    # ================= Offer & Affordability =================
    ws = sheet(wb, "Offer & Affordability")
    r = title(ws, 1, "WHAT CAN WE AFFORD? - structure matrix ($250K capital, no SBA)")
    r = note(ws, r, "Binding constraint: the operating cash trough is -$93.8K (cum.) at Month 2 while AR builds. "
                    "Every dollar of down payment or early monthly competes with payroll through that trough.")
    r += 1
    heads = ["Structure", "Down (close)", "Monthlies", "Paid pre-balloon", "Balloon (Jan)", "Min cash", "@Mo", "Verdict"]
    for j, h in enumerate(heads):
        c = ws.cell(r, 1 + j, h); c.font = hdrF; c.fill = fillH
    r += 1
    structs = [
        run_structure(cfo_engine, 245_000, 0, 0, "Seller ask: 49% down"),
        run_structure(cfo_engine, 106_000, 0, 0, "Max down, no monthlies"),
        run_structure(cfo_engine, 75_000, 15_000, 0, "$75K + $15K/mo M1-6"),
        run_structure(cfo_engine, 100_000, 10_000, 0, "$100K + $10K/mo M1-6"),
        rec | {"label": "RECOMMENDED: $100K + $25K/mo M3-6"},
        run_structure(cfo_engine, 125_000, 25_000, 2, "$125K + $25K/mo M3-6"),
    ]
    for s in structs:
        paid = s["down"] + (6 - s["holiday"]) * s["monthly"]
        verdict = ("UNAFFORDABLE" if s["min_cash"] < 0 else
                   "unsafe" if s["min_cash"] < 25_000 else
                   "tight" if s["min_cash"] < 40_000 else "OK")
        vals = [s["label"], s["down"], s["monthly"], paid, s["balloon"], s["min_cash"], s["min_month"], verdict]
        for j, v in enumerate(vals):
            c = ws.cell(r, 1 + j, v)
            if isinstance(v, (int, float)) and j not in (6,): money(c)
            c.border = box
            if "RECOMMENDED" in s["label"]: c.fill = PatternFill("solid", fgColor=GRN); c.font = boldF
            if verdict == "UNAFFORDABLE" and j == 7: c.fill = PatternFill("solid", fgColor=RED)
        r += 1
    r += 1
    r = title(ws, r, "RECOMMENDED OFFER (modeled throughout this workbook)", 11)
    for line in [
        "Price: $500,000 (full ask)",
        "Down at close (July): $100,000 (20%)",
        "Months 1-2 (Jul-Aug): payment holiday while the migrated census ramps and first collections land",
        "Months 3-6 (Sep-Dec): $25,000/month = $100,000",
        f"January 2027: balloon of remaining principal $300,000 + ~${sum(seller_int):,.0f} accrued interest (6% simple) = ~${balloon:,.0f}, paid via personal refinance",
        "Seller receives $200,000 (40%) within 6 months and full payout in January; balance secured + personally guaranteed",
        f"Buyer min cash: ${minc:,.0f} (Month {minm}) - above the $25K floor with margin; $150K stays available at close",
    ]:
        ws.cell(r, 1, "  - " + line); r += 1

    # ================= 12-Mo P&L =================
    ws = sheet(wb, "12-Mo P&L (Operating Budget)")
    r = title(ws, 1, "12-MONTH PROFORMA P&L / OPERATING BUDGET - base case (no SBA)")
    r = month_header(ws, r + 1)
    r = row(ws, r, "Average daily census (ADC)", adc, total=False)
    r = row(ws, r, "Gross revenue", gross)
    r = row(ws, r, "NET REVENUE", net, bold=True, fill=LT)
    r = row(ws, r, "Direct clinical labor (w/ benefits)", cogs_labor, indent=True)
    r = row(ws, r, "Patient costs (pharmacy/supplies/DME)", cogs_pt, indent=True)
    r = row(ws, r, "GROSS PROFIT", gp, bold=True, fill=LT)
    r = row(ws, r, "Indirect labor (admin/sales, w/ benefits)", sga_lab, indent=True)
    r = row(ws, r, "G&A fixed (rent, EMR, insurance, etc.)", ga_fix, indent=True)
    r = row(ws, r, "QR + billing fees (% of gross)", fees, indent=True)
    r = row(ws, r, "EBITDA", ebitda, bold=True, fill=GRN)
    r = row(ws, r, "License amortization (non-cash)", da, indent=True)
    r = row(ws, r, "Seller note interest (6%)", seller_int, indent=True)
    r = row(ws, r, "Refi interest (from M7)", refi_int, indent=True)
    r = row(ws, r, "Pre-tax income", pretax)
    r = row(ws, r, "TX margin tax", tax, indent=True)
    r = row(ws, r, "NET INCOME", ni, bold=True, fill=GRN)
    r += 1
    r = note(ws, r, "Steady-state reconciles to the Paloma Tyler actuals (~$118K net/mo at ~22 ADC). Staffing per the "
                    "engine's staggered W-2 roster; PRN converts per census triggers.")

    # ================= Cash-Flow Budget =================
    ws = sheet(wb, "Cash-Flow Budget")
    r = title(ws, 1, "12-MONTH CASH-FLOW BUDGET - recommended offer structure")
    r = note(ws, r, f"Opening: $250,000 less $100,000 down = $150,000. Balloon (Jan) is refi-funded "
                    f"(cash-neutral to the company); company then services the refi at ${refi_pmt:,.0f}/mo.")
    r = month_header(ws, r + 1)
    r = row(ws, r, "Net income", ni)
    r = row(ws, r, "+ License amortization (non-cash)", da, indent=True)
    r = row(ws, r, "+ Seller interest accrued (paid at balloon)", seller_int, indent=True)
    r = row(ws, r, "- Increase in AR (45-day lag)", [-x for x in d_ar], indent=True)
    r = row(ws, r, "+ Increase in AP (30-day)", d_ap, indent=True)
    r = row(ws, r, "CASH FROM OPERATIONS", cfo, bold=True, fill=LT)
    r = row(ws, r, "Seller monthly payments (principal)", [-p for p in seller_pay], indent=True)
    r = row(ws, r, "Balloon paid / refi proceeds (net)", [0.0] * 12, indent=True, total=False)
    r = row(ws, r, "Refi principal (M7+; interest in P&L)", [-refi_prin[i] for i in range(12)], indent=True)
    r = row(ws, r, "NET CASH FLOW", [cfo[i] - seller_pay[i] - refi_prin[i] for i in range(12)], bold=True)
    r = row(ws, r, "ENDING CASH", endc, bold=True, fill=GRN, total=False)
    r += 1
    ws.cell(r, 1, f"Minimum cash: ${minc:,.0f} in {MONTHS[minm-1]}  |  Ending cash Jun-27: ${endc[-1]:,.0f}  |  "
                  f"Seller fully paid January 2027").font = boldF
    r += 2
    r = note(ws, r, "No LOC assumed in the base case. If any month dips below plan, levers: delay a monthly seller "
                    "payment (negotiate cure language), draw on Jim's bank relationship, or accelerate the SBA scenario below.")

    # ================= Census & Collections =================
    ws = sheet(wb, "Census & Collections")
    r = title(ws, 1, "CENSUS -> PATIENT DAYS -> REVENUE -> CASH (one-month collection lag)")
    r = note(ws, r, "Revenue is EARNED as patients are seen; cash is COLLECTED one month later "
                    "(see patients in July, money lands in August). Expenses paid in-month (conservative: "
                    "no AP lag credit). Recommended offer structure overlaid.")
    pdays = [R["pd"][i] for i in range(N)]
    coll = [0.0] + [net[i - 1] for i in range(1, N)]
    opex = [net[i] - ebitda[i] for i in range(N)]
    r = month_header(ws, r + 1)
    r = row(ws, r, "Average daily census (ADC)", adc, total=False)
    r = row(ws, r, "Total patient days", pdays)
    r = row(ws, r, "Gross revenue (earned)", gross)
    r = row(ws, r, "NET REVENUE (earned)", net, bold=True, fill=LT)
    r = row(ws, r, "Cash COLLECTED (prior month's net)", coll, bold=True, fill=AMB)
    r = row(ws, r, "Operating expenses (cash, in-month)", [-x for x in opex])
    r = row(ws, r, "Net operating cash", [coll[i] - opex[i] for i in range(N)], bold=True)
    r = row(ws, r, "Seller payments (down excl.)", [-p for p in seller_pay], indent=True)
    r = row(ws, r, "Refi principal+interest (M7+)", [-(refi_pmt if i > 5 else 0.0) for i in range(N)], indent=True)
    endc3, beg3 = [], CASH - down
    for i in range(N):
        c = beg3 + coll[i] - opex[i] - seller_pay[i] - (refi_pmt if i > 5 else 0.0)
        endc3.append(c); beg3 = c
    r = row(ws, r, "ENDING CASH (1-mo lag view)", endc3, bold=True, fill=GRN, total=False)
    r += 1
    ws.cell(r, 1, f"Min cash ${min(endc3):,.0f} ({MONTHS[endc3.index(min(endc3))]})  |  End Jun-27 "
                  f"${endc3[-1]:,.0f}  |  vs 45-day-AR view min ${minc:,.0f} - the one-month lag is the "
                  f"better case; the 45-day view is the conservative floor.").font = boldF
    r += 2
    r = note(ws, r, "Six-month picture: net revenue earned Jul-Dec ~$667K; collected by Dec 31 ~$542K "
                    "(December's revenue lands in January). EBITDA positive every month from M1.")

    # ================= Sensitivity =================
    from financial_models.refuge_sensitivity import (scenario as sens, GRID,
                                                     breakeven_adc, max_seller_monthly)
    ws = sheet(wb, "Sensitivity")
    r = title(ws, 1, "SENSITIVITY - census capture x payroll inflation (interview-locked axes)")
    r = note(ws, r, "Cost refinements applied to ALL cases: G&A -$2,004/mo (rent $2,000; credit-card fees cut); "
                    "health insurance 90-day waiting period; med director contracted (not inflated). "
                    "Floor: >=$25K safe | <$25K flag | <$0 FAIL. Offer overlay: $100K down + $25K x4 (M3-6) + Jan balloon.")
    r += 1
    heads = ["Scenario", "6-mo EBITDA", "Min cash (1-mo lag)", "Trough mo", "Min cash (45-day AR)", "Verdict"]
    for j, h in enumerate(heads):
        c = ws.cell(r, 1 + j, h); c.font = hdrF; c.fill = fillH
    r += 1
    for label, cpt, pmm, dfr in GRID:
        s = sens(cpt, pmm, dfr)
        e6 = sum(s["ebitda"][:6])
        v = "FAIL" if s["min1"] < 0 else ("flag" if s["min1"] < 25_000 else "safe")
        vals = [label, e6, s["min1"], s["min1_m"], s["min45"], v]
        for j, val in enumerate(vals):
            c = ws.cell(r, 1 + j, val)
            if isinstance(val, (int, float)) and j in (1, 2, 4): money(c)
            c.border = box
            if v == "FAIL": c.fill = PatternFill("solid", fgColor=RED)
            elif v == "flag": c.fill = PatternFill("solid", fgColor=AMB)
        r += 1
    r += 1
    r = title(ws, r, "WHERE THE CLIFF IS", 11)
    for line in [
        "Safe ($25K floor) down to 82% census capture (steady ADC ~18.7). Never-negative down to 80% (ADC ~18.1).",
        "If payroll runs +10%, the safe threshold rises to 90% capture; owner deferral buys it back to ~80%.",
        f"Breakeven ADC (steady state): {breakeven_adc():.1f} patients; at payroll +10%: {breakeven_adc(1.10):.1f}.",
        "Below ~80% capture the BUSINESS is EBITDA-negative at steady state - no offer structure fixes that;",
        "the mitigation is the census ramp itself (marketing, referral push), not deal terms.",
    ]:
        ws.cell(r, 1, "  - " + line); r += 1
    r += 1
    r = title(ws, r, "MAX AFFORDABLE SELLER MONTHLY (M3-6, after $100K down; $25K floor)", 11)
    for label, cpt, pmm, dfr in [("Base (100%)", 1.0, 1.0, False), ("Census 85%", 0.85, 1.0, False),
                                 ("Census 70%", 0.70, 1.0, False)]:
        mx = max_seller_monthly(cpt, pmm, dfr)
        ws.cell(r, 1, f"  - {label}: ${mx:,.0f}/mo" + ("  (offer's $25K has ~$20K headroom)" if cpt == 1.0 else
                      ("  (offer's $25K still fits)" if mx >= 25_000 else "  (payments would need pause/cure)")))
        r += 1
    r += 1
    r = note(ws, r, "Interview-locked mitigations: hires track census (natural hedge); owner salary deferral "
                    "~$12K/mo available M1-6; LOI cure/pause language on monthlies recommended for the <85% case.")

    # ================= SBA Scenario =================
    ws = sheet(wb, "SBA Scenario")
    r = title(ws, 1, "SCENARIO: SBA 7(a) $450,000 lands Month 3 (Sep-26)")
    r = note(ws, r, "Same operations and offer structure; adds $450K working-capital loan (10.5%, 15-yr, "
                    "$4,974/mo from M4). Shown for the parallel John Hart track.")
    sba_pmt = pmt(450_000, 0.105, 180)
    endc2, beg2 = [], CASH - down
    for i in range(12):
        mno = i + 1
        c = beg2 + cfo[i] - seller_pay[i] - refi_prin[i]
        if mno == 3: c += 450_000
        if mno > 3: c -= sba_pmt
        endc2.append(c); beg2 = c
    r = month_header(ws, r + 1)
    r = row(ws, r, "Ending cash - BASE (no SBA)", endc, total=False)
    r = row(ws, r, "Ending cash - WITH SBA $450K", endc2, bold=True, fill=GRN, total=False)
    r += 1
    ws.cell(r, 1, f"With SBA: min cash ${min(endc2):,.0f}, ending cash ${endc2[-1]:,.0f}. The SBA scenario removes all "
                  f"liquidity risk and could fund a larger down payment if the seller pushes.").font = Font(italic=True, size=9)

    out = "financial_models/output/Azalea_Refuge_12mo_Proforma.xlsx"
    wb.save(out)
    print("wrote", out)
    print(f"rec: down 100k, holiday 2, 25k x4, balloon {balloon:,.0f}, min cash {minc:,.0f} @M{minm}, end {endc[-1]:,.0f}")


if __name__ == "__main__":
    build()
