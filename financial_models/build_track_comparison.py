"""Four-track cash-flow comparison workbook (Rev 1.00, 2026-07-28).

One operating chassis (census ramp, unit economics from the Rev 4.10 proforma),
four license/billing wrappers:

  T1  Jason ADS      - census bills through Jason's established number (normal lag,
                       no hold); flat participation fee to Jason; no license capital.
  T2  Refuge         - Rev 4.10 deal terms (seller schedule, Sept bank note, Jan SBA
                       refi) with a PPEO hold on first claims.
  T3  Avant          - idle burn until the Palmetto tie-in/CHOW clears (input month),
                       Medicare-only census, PPEO hold on first claims, May-2026
                       seller schedule time-shifted to start at PPEO clear ($50K down,
                       month-6 balloon $51,478.15, then $11,646.35 x 18 at 6%).
  T4a Refuge+Jason   - 5-patient PPEO seed on the owned license while growth census
  T4b Avant+Jason      runs on Jason's number; admissions shift at PPEO clear.

Funding basis: tracks are modeled WITHOUT equity inflows; the headline metric is
peak cumulative funding need, compared against the ~$215K actually available.
No revolver, no plugs - same real-cash discipline as the Rev 4.10 model.

Run from repo root: python3 -m financial_models.build_track_comparison
Output: financial_models/output/Azalea_Track_Comparison_Rev1.00.xlsx
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as gcl

OUT = "financial_models/output/Azalea_Track_Comparison_Rev1.00.xlsx"
ROWMAP_OUT = "financial_models/output/track_comparison_rowmap.json"
NM = 24
MC0 = 2  # month 1 = column B
MONTHS = ["Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27",
          "Mar-27", "Apr-27", "May-27", "Jun-27", "Jul-27", "Aug-27", "Sep-27",
          "Oct-27", "Nov-27", "Dec-27", "Jan-28", "Feb-28", "Mar-28", "Apr-28",
          "May-28", "Jun-28", "Jul-28"]

NAVY = "1F3864"
GREY = "D9D9D9"
BLUE = "DDEBF7"
YELL = "FFF2CC"
FMT_CUR = '#,##0;(#,##0)'
FMT_NUM = '0.0'
FMT_INT = '0'

THIN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


def mlet(i):
    return gcl(MC0 + i)


def MRANGE(row):
    return f"$B${row}:$Y${row}"


class WB:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.addr = {}
        self.rows = {}

    def sheet(self, name, tab):
        ws = self.wb.create_sheet(name)
        ws.sheet_properties.tabColor = tab
        ws.column_dimensions["A"].width = 52
        for i in range(NM):
            ws.column_dimensions[mlet(i)].width = 11
        self.rows[name] = {}
        return ws

    def title(self, ws, text):
        c = ws.cell(row=1, column=1, value=text)
        c.font = Font(bold=True, size=12, color=NAVY)

    def hdr(self, ws, r):
        for i in range(NM):
            c = ws.cell(row=r, column=MC0 + i, value=MONTHS[i])
            c.font = Font(bold=True, size=9)
            c.fill = PatternFill("solid", fgColor=GREY)
            c.alignment = Alignment(horizontal="center")
            c.border = THIN
        self.rows[ws.title]["_hdr"] = r
        return r + 1

    def section(self, ws, r, text):
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        return r + 1

    def lbl(self, ws, r, text, bold=False, italic=False):
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(bold=bold, italic=italic, size=9.5)

    def inp(self, ws, r, c, val, fmt=FMT_CUR):
        cell = ws.cell(row=r, column=c, value=val)
        cell.number_format = fmt
        cell.fill = PatternFill("solid", fgColor=YELL)
        cell.border = THIN
        cell.font = Font(size=9.5)
        return cell

    def fml(self, ws, r, c, formula, fmt=FMT_CUR, bold=False, fill=None):
        cell = ws.cell(row=r, column=c, value=formula)
        cell.number_format = fmt
        cell.font = Font(size=9.5, bold=bold)
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
        cell.border = THIN
        return cell


# ADC ramp (chassis default; matches the Rev 4.10 proven path through M12,
# then grows ~+0.9/mo toward 50 by M24)
ADC = [9.0, 18.0, 25.0, 28.0, 30.0, 32.0, 34.0, 35.5, 36.5, 37.5, 38.5, 39.5,
       40.4, 41.3, 42.2, 43.1, 44.0, 45.0, 46.0, 47.0, 48.0, 49.0, 49.5, 50.0]


def assumptions(bk):
    ws = bk.sheet("Assumptions", "C00000")
    bk.title(ws, "ASSUMPTIONS - single source of truth. Yellow cells are inputs; every track tab references here. "
                 "Unit economics per the Rev 4.10 proforma Control Tower (net/day, cost rates, ADC path).")
    A = bk.addr
    r = 3

    def item(key, label, val, fmt=FMT_CUR, note=""):
        nonlocal r
        bk.lbl(ws, r, label)
        bk.inp(ws, r, 2, val, fmt)
        if note:
            ws.cell(row=r, column=4, value=note).font = Font(size=8.5, italic=True)
        A[key] = f"'Assumptions'!$B${r}"
        r += 1

    r = bk.section(ws, r, "SHARED CHASSIS (Rev 4.10 unit economics)")
    item("netday", "Net revenue per patient-day (blended, after sequestration/fees)", 172.0, FMT_CUR,
         "Rev 4.10 Control Tower; Medicaid r&b treated as cash-neutral pass-through")
    item("dpm", "Days per month (billing average)", 30.4, FMT_NUM)
    item("varpd", "Variable cost per patient-day (clinical labor + benefits + DPC)", 79.0, FMT_CUR,
         "Rev 4.10 Y1: clinical ~$63/pd + direct patient care ~$16/pd")
    item("fixmo", "Fixed overhead per month (admin labor, facility, G&A)", 62000.0, FMT_CUR,
         "Rev 4.10 Y1: indirect $49.7K/mo + facility/G&A $12.3K/mo; no owner deferral assumed (conservative)")
    item("col_lag_pct", "% collected month +1 (normal, no hold)", 1.0, '0%')
    item("avail", "Capital actually available (Jim's remaining $65K + $150K)", 215000.0, FMT_CUR)

    r += 1
    r = bk.section(ws, r, "ADC RAMP (chassis; editable)")
    bk.lbl(ws, r, "Month")
    for i in range(NM):
        c = ws.cell(row=r, column=MC0 + i, value=MONTHS[i])
        c.font = Font(bold=True, size=8.5)
    r += 1
    bk.lbl(ws, r, "Average daily census (total book)")
    for i in range(NM):
        bk.inp(ws, r, MC0 + i, ADC[i], FMT_NUM)
    bk.rows["Assumptions"]["adc"] = r
    A["adc_row"] = r
    r += 2

    r = bk.section(ws, r, "T1 - JASON ADS")
    item("j_fee", "Jason participation fee per month (flat; terms not yet agreed)", 15000.0, FMT_CUR,
         "PLACEHOLDER - AKS-compliant flat FMV fee or employment-mode equivalent")

    r += 1
    r = bk.section(ws, r, "T2 - REFUGE (Rev 4.10 deal terms)")
    item("r_hold", "PPEO hold on first claims (months)", 3, FMT_INT, "sensitivity: 4")
    item("r_down", "Down payment at closing (M1)", 125000.0)
    item("r_inst", "Seller installment (M2-M5, replaced by bank payoff M2)", 31250.0)
    item("r_bank", "Bank note funds (M2)", 500000.0)
    item("r_spay", "Seller payoff at bank funding (M2, principal + accrued)", 376875.0, FMT_CUR,
         "$375K balance + ~1 month interest at 6%")
    item("r_bpmt", "Bank note payment (M3-M5)", 15211.0)
    item("r_sba", "SBA funds (M6)", 500000.0)
    item("r_bpay", "Bank note payoff at SBA funding (M6)", 461676.0)
    item("r_spmt", "SBA payment (M7 on)", 6747.0)

    r += 1
    r = bk.section(ws, r, "T3 - AVANT ($300K, seller-financed; payments start at PPEO clear)")
    item("a_clear", "Palmetto tie-in / CHOW clearance month (census can start)", 2, FMT_INT,
         "UNRESOLVED GATE - multi-year Palmetto saga; sensitivity: 4, 6")
    item("a_hold", "PPEO hold on first claims (months after census start)", 4, FMT_INT, "sensitivity: 3, 6")
    item("a_burn", "Pre-clearance monthly burn (office, skeleton team)", 10000.0)
    item("a_down", "Down payment (paid at PPEO clear, per re-negotiated terms)", 50000.0)
    item("a_ball", "Month-6-after-clear balloon (per May schedule)", 51478.15)
    item("a_inst", "Monthly installment (18 months, starts month 7 after clear)", 11646.35)

    r += 1
    r = bk.section(ws, r, "T4 - COMBO (owned-license seed + Jason growth census)")
    item("c_seed", "Seed census on owned license (ADC)", 5.0, FMT_NUM)
    item("c_scost", "Seed carry cost per month (0.5 RN + 0.5 CNA + PRN allocation)", 12000.0)

    bk.wb["Assumptions"].freeze_panes = "B3"
    return ws


def chassis_rows(bk, ws, r, adc_formula):
    """Common P&L-to-cash rows given an ADC row formula per month.
    Returns dict of row indices: adc, npr(earned), varc, fixc."""
    name = ws.title
    A = bk.addr
    R = bk.rows[name]
    bk.lbl(ws, r, "ADC on this track")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, adc_formula(i), FMT_NUM)
    R["adc"] = r
    r += 1
    bk.lbl(ws, r, "Revenue earned (NPR)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['adc']}*{A['netday']}*{A['dpm']}")
    R["earned"] = r
    r += 1
    bk.lbl(ws, r, "Variable costs (clinical + patient care, paid in month)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-{mlet(i)}{R['adc']}*{A['varpd']}*{A['dpm']}")
    R["varc"] = r
    r += 1
    return r


def cum_and_need(bk, ws, r, net_row):
    name = ws.title
    R = bk.rows[name]
    bk.lbl(ws, r, "NET CASH FLOW (month)", bold=True)
    R["net"] = net_row
    bk.lbl(ws, r, "Cumulative cash (no equity assumed)", bold=True)
    for i in range(NM):
        prev = f"{mlet(i-1)}{r}+" if i > 0 else ""
        bk.fml(ws, r, MC0 + i, f"={prev}{mlet(i)}{net_row}", bold=True, fill=BLUE)
    R["cum"] = r
    r += 1
    bk.lbl(ws, r, "PEAK FUNDING NEED (max cumulative deficit)", bold=True)
    bk.fml(ws, r, 2, f"=MAX(0,-MIN({MRANGE(R['cum'])}))", bold=True, fill=YELL)
    R["peak"] = r
    r += 1
    return r


def track1(bk):
    ws = bk.sheet("T1 Jason ADS", "2E7D32")
    bk.title(ws, "T1 - JASON ALTERNATE DELIVERY SITE: census bills through Jason's established number; "
                 "collections at normal +1-month lag, NO hold; flat fee to Jason; no license capital.")
    A = bk.addr
    R = bk.rows["T1 Jason ADS"]
    r = bk.hdr(ws, 3)
    adcrow = A["adc_row"]
    r = chassis_rows(bk, ws, r, lambda i: f"=Assumptions!{mlet(i)}{adcrow}")
    bk.lbl(ws, r, "Collections (+1 month lag)")
    for i in range(NM):
        src = "0" if i == 0 else f"{mlet(i-1)}{R['earned']}*{A['col_lag_pct']}"
        bk.fml(ws, r, MC0 + i, f"={src}")
    R["coll"] = r
    r += 1
    bk.lbl(ws, r, "Fixed overhead")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-{A['fixmo']}")
    R["fix"] = r
    r += 1
    bk.lbl(ws, r, "Jason participation fee")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-{A['j_fee']}")
    R["fee"] = r
    r += 1
    bk.lbl(ws, r, "NET CASH FLOW", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['coll']}+{c}{R['varc']}+{c}{R['fix']}+{c}{R['fee']}", bold=True)
    net = r
    r += 1
    r = cum_and_need(bk, ws, r, net)
    ws.freeze_panes = "B4"
    return ws


def track2(bk):
    ws = bk.sheet("T2 Refuge", "C00000")
    bk.title(ws, "T2 - REFUGE (Rev 4.10 deal): own license from M1; PPEO hold on first claims; "
                 "seller schedule + Sept bank note + Jan SBA refi. No equity inflow modeled.")
    A = bk.addr
    R = bk.rows["T2 Refuge"]
    r = bk.hdr(ws, 3)
    adcrow = A["adc_row"]
    r = chassis_rows(bk, ws, r, lambda i: f"=Assumptions!{mlet(i)}{adcrow}")
    bk.lbl(ws, r, "Collection month for service month (lag + PPEO hold)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={i+1}+1+MAX(0,{A['r_hold']}-{i})", FMT_INT)
    R["cmo"] = r
    r += 1
    bk.lbl(ws, r, "Collections (held claims release together)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUMPRODUCT(({MRANGE(R['cmo'])}={i+1})*{MRANGE(R['earned'])})")
    R["coll"] = r
    r += 1
    bk.lbl(ws, r, "Fixed overhead")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-{A['fixmo']}")
    R["fix"] = r
    r += 1
    bk.lbl(ws, r, "License & financing flows (Rev 4.10 Path A)")
    for i in range(NM):
        m = i + 1
        f = (f"=-IF({m}=1,{A['r_down']},0)"
             f"+IF({m}=2,{A['r_bank']}-{A['r_spay']},0)"
             f"-IF(AND({m}>=3,{m}<=5),{A['r_bpmt']},0)"
             f"+IF({m}=6,{A['r_sba']}-{A['r_bpay']},0)"
             f"-IF({m}>=7,{A['r_spmt']},0)")
        bk.fml(ws, r, MC0 + i, f)
    R["deal"] = r
    r += 1
    bk.lbl(ws, r, "NET CASH FLOW", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['coll']}+{c}{R['varc']}+{c}{R['fix']}+{c}{R['deal']}", bold=True)
    net = r
    r += 1
    r = cum_and_need(bk, ws, r, net)
    ws.freeze_panes = "B4"
    return ws


def track3(bk):
    ws = bk.sheet("T3 Avant", "7030A0")
    bk.title(ws, "T3 - AVANT ($300K seller-financed, Medicare-only): idle until tie-in/CHOW clears, then census "
                 "starts; PPEO hold on first claims; payments start at PPEO clear (May schedule, time-shifted).")
    A = bk.addr
    R = bk.rows["T3 Avant"]
    r = bk.hdr(ws, 3)
    adcrow = A["adc_row"]
    # census starts at clearance month: ramp position = month - clear + 1
    r = chassis_rows(
        bk, ws, r,
        lambda i: (f"=IF({i+1}<{A['a_clear']},0,"
                   f"INDEX(Assumptions!{MRANGE(adcrow)},1,{i+1}-{A['a_clear']}+1))"))
    bk.lbl(ws, r, "Collection month for service month (lag + PPEO hold from census start)")
    for i in range(NM):
        # hold counts down from the first billing month (a_clear)
        bk.fml(ws, r, MC0 + i,
               f"={i+1}+1+MAX(0,{A['a_hold']}-({i+1}-{A['a_clear']}))", FMT_INT)
    R["cmo"] = r
    r += 1
    bk.lbl(ws, r, "Collections (held claims release together)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUMPRODUCT(({MRANGE(R['cmo'])}={i+1})*{MRANGE(R['earned'])})")
    R["coll"] = r
    r += 1
    bk.lbl(ws, r, "Overhead (pre-clearance burn, then full fixed)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=-IF({i+1}<{A['a_clear']},{A['a_burn']},{A['fixmo']})")
    R["fix"] = r
    r += 1
    bk.lbl(ws, r, "Seller payments (anchor = PPEO clear = clearance + hold)")
    for i in range(NM):
        m = i + 1
        anchor = f"({A['a_clear']}+{A['a_hold']})"
        f = (f"=-IF({m}={anchor},{A['a_down']},0)"
             f"-IF({m}={anchor}+6,{A['a_ball']},0)"
             f"-IF(AND({m}>={anchor}+7,{m}<={anchor}+24),{A['a_inst']},0)")
        bk.fml(ws, r, MC0 + i, f)
    R["deal"] = r
    r += 1
    bk.lbl(ws, r, "NET CASH FLOW", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i, f"={c}{R['coll']}+{c}{R['varc']}+{c}{R['fix']}+{c}{R['deal']}", bold=True)
    net = r
    r += 1
    r = cum_and_need(bk, ws, r, net)
    ws.freeze_panes = "B4"
    return ws


def track4(bk, name, tab, deal_kind):
    """Combo: seed ADC on owned license until its PPEO clears, growth census on Jason.
    deal_kind: 'refuge' or 'avant' - which owned-license payment schedule applies."""
    ws = bk.sheet(name, tab)
    bk.title(ws, f"{name} - COMBO: 5-patient PPEO seed on the owned license, growth census on Jason's number; "
                 "admissions shift to the owned license at PPEO clear; Jason fee ends after shift.")
    A = bk.addr
    R = bk.rows[name]
    r = bk.hdr(ws, 3)
    adcrow = A["adc_row"]

    if deal_kind == "refuge":
        start = "1"                       # owned license billable from M1
        clear = f"(1+{A['r_hold']})"       # seed claims pay after hold
    else:
        start = A["a_clear"]
        clear = f"({A['a_clear']}+{A['a_hold']})"

    bk.lbl(ws, r, "Owned-license ADC (seed until clear, then full book)")
    for i in range(NM):
        m = i + 1
        bk.fml(ws, r, MC0 + i,
               f"=IF({m}<{start},0,IF({m}<={clear},{A['c_seed']},"
               f"Assumptions!{mlet(i)}{adcrow}))", FMT_NUM)
    R["adc_own"] = r
    r += 1
    bk.lbl(ws, r, "Jason ADC (total book minus owned)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=MAX(0,Assumptions!{mlet(i)}{adcrow}-{mlet(i)}{R['adc_own']})", FMT_NUM)
    R["adc_j"] = r
    r += 1
    bk.lbl(ws, r, "Revenue earned - owned license")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"={mlet(i)}{R['adc_own']}*{A['netday']}*{A['dpm']}")
    R["earn_own"] = r
    r += 1
    bk.lbl(ws, r, "Collection month - owned (lag + hold from start)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=MAX({i+1}+1,{clear}+1)", FMT_INT)
    R["cmo"] = r
    r += 1
    bk.lbl(ws, r, "Collections - owned license")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=SUMPRODUCT(({MRANGE(R['cmo'])}={i+1})*{MRANGE(R['earn_own'])})")
    R["coll_own"] = r
    r += 1
    bk.lbl(ws, r, "Collections - Jason census (+1 lag, no hold)")
    for i in range(NM):
        src = "0" if i == 0 else f"{mlet(i-1)}{R['adc_j']}*{A['netday']}*{A['dpm']}*{A['col_lag_pct']}"
        bk.fml(ws, r, MC0 + i, f"={src}")
    R["coll_j"] = r
    r += 1
    bk.lbl(ws, r, "Variable costs (whole book)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i,
               f"=-(({mlet(i)}{R['adc_own']}+{mlet(i)}{R['adc_j']})*{A['varpd']}*{A['dpm']})")
    R["varc"] = r
    r += 1
    bk.lbl(ws, r, "Fixed overhead + seed carry")
    for i in range(NM):
        m = i + 1
        bk.fml(ws, r, MC0 + i,
               f"=-({A['fixmo']}+IF(AND({m}>={start},{m}<={clear}),{A['c_scost']},0))")
    R["fix"] = r
    r += 1
    bk.lbl(ws, r, "Jason fee (while Jason census active)")
    for i in range(NM):
        bk.fml(ws, r, MC0 + i, f"=-IF({mlet(i)}{R['adc_j']}>0,{A['j_fee']},0)")
    R["fee"] = r
    r += 1
    bk.lbl(ws, r, "Owned-license payments")
    for i in range(NM):
        m = i + 1
        if deal_kind == "refuge":
            f = (f"=-IF({m}=1,{A['r_down']},0)"
                 f"+IF({m}=2,{A['r_bank']}-{A['r_spay']},0)"
                 f"-IF(AND({m}>=3,{m}<=5),{A['r_bpmt']},0)"
                 f"+IF({m}=6,{A['r_sba']}-{A['r_bpay']},0)"
                 f"-IF({m}>=7,{A['r_spmt']},0)")
        else:
            anchor = f"({A['a_clear']}+{A['a_hold']})"
            f = (f"=-IF({m}={anchor},{A['a_down']},0)"
                 f"-IF({m}={anchor}+6,{A['a_ball']},0)"
                 f"-IF(AND({m}>={anchor}+7,{m}<={anchor}+24),{A['a_inst']},0)")
        bk.fml(ws, r, MC0 + i, f)
    R["deal"] = r
    r += 1
    bk.lbl(ws, r, "NET CASH FLOW", bold=True)
    for i in range(NM):
        c = mlet(i)
        bk.fml(ws, r, MC0 + i,
               f"={c}{R['coll_own']}+{c}{R['coll_j']}+{c}{R['varc']}+{c}{R['fix']}+{c}{R['fee']}+{c}{R['deal']}",
               bold=True)
    net = r
    r += 1
    r = cum_and_need(bk, ws, r, net)
    ws.freeze_panes = "B4"
    return ws


def comparison(bk):
    ws = bk.wb.create_sheet("Comparison", 0)
    ws.sheet_properties.tabColor = "BF8F00"
    ws.column_dimensions["A"].width = 46
    for col in "BCDEF":
        ws.column_dimensions[col].width = 17
    ws.column_dimensions["G"].width = 60
    bk.title(ws, "FOUR-TRACK COMPARISON - peak funding need vs capital actually available. "
                 "All tracks: same census ramp and unit economics; no equity inflows; no revolver.")

    tracks = [("T1 Jason ADS", "T1 Jason ADS"),
              ("T2 Refuge", "T2 Refuge"),
              ("T3 Avant", "T3 Avant"),
              ("T4a Refuge + Jason", "T4a Refuge+Jason"),
              ("T4b Avant + Jason", "T4b Avant+Jason")]
    gates = {
        "T1 Jason ADS": "Gates: Jason terms unagreed (fee is placeholder); AKS-compliant structure + county add + 855A location filing; no license owned at end.",
        "T2 Refuge": "Gates: 36-month window to 1/8/2027; Sept reactivation expiry; PPEO on first claim (letter); bank + SBA must both fund.",
        "T3 Avant": "Gates: Palmetto tie-in/CHOW UNRESOLVED (multi-year); Medicare-only; own 36-month/moratorium analysis unverified; license expires 7/8/2027.",
        "T4a Refuge+Jason": "Gates: T2 gates + Jason gates; two arrangements run at once.",
        "T4b Avant+Jason": "Gates: T3 gates + Jason gates; cheapest capital if Avant clears.",
    }
    hdrs = ["Track", "Peak funding need", "vs available", "Month-12 cum. cash", "Month-24 cum. cash", "Regulatory gates"]
    r = 3
    for j, h in enumerate(hdrs):
        c = ws.cell(row=r, column=1 + j, value=h)
        c.font = Font(bold=True, size=9.5, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.border = THIN
    r += 1
    A = bk.addr
    for label, sheet in tracks:
        R = bk.rows[sheet]
        ws.cell(row=r, column=1, value=label).font = Font(bold=True, size=9.5)
        bk_f = bk.fml
        bk_f(ws, r, 2, f"='{sheet}'!B{R['peak']}")
        bk_f(ws, r, 3, f"='{sheet}'!B{R['peak']}-{A['avail']}")
        bk_f(ws, r, 4, f"='{sheet}'!M{R['cum']}")
        bk_f(ws, r, 5, f"='{sheet}'!Y{R['cum']}")
        c = ws.cell(row=r, column=6, value=gates[sheet])
        c.font = Font(size=8.5)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    r += 1
    note = ws.cell(row=r, column=1, value=(
        "Notes: 'vs available' compares to the capital input on Assumptions (~$215K remaining of Jim's $250K). "
        "Positive = additional capital required beyond what is committed. Jason fee, Avant clearance month, and "
        "PPEO hold lengths are inputs - flex them on the Assumptions tab. Rev 1.00, prepared 7/28/2026."))
    note.font = Font(size=8.5, italic=True)
    ws.freeze_panes = "A4"
    return ws


def main():
    bk = WB()
    assumptions(bk)
    track1(bk)
    track2(bk)
    track3(bk)
    track4(bk, "T4a Refuge+Jason", "C55A11", "refuge")
    track4(bk, "T4b Avant+Jason", "7030A0", "avant")
    comparison(bk)
    bk.wb.active = 0
    bk.wb.save(OUT)
    with open(ROWMAP_OUT, "w") as f:
        json.dump({"rows": bk.rows, "addr": bk.addr}, f, indent=1)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
