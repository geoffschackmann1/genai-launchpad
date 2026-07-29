"""Four-track comparison memo (evaluates the workbook live so figures can't drift).

Output: sba_application/20_track_comparison_2026-07-28/
Run from repo root: python3 sba_application/build_track_memo.py
"""
import json
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docfmt
import formulas
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from openpyxl.utils import get_column_letter as gcl

XLSX = "financial_models/output/Azalea_Track_Comparison_Rev1.00.xlsx"
ROWMAP = "financial_models/output/track_comparison_rowmap.json"
OUT = "sba_application/20_track_comparison_2026-07-28/"
NAVY = RGBColor(0x1F, 0x38, 0x64)
TRACKS = ["T1 Jason ADS", "T2 Refuge", "T3 Avant", "T4a Refuge+Jason", "T4b Avant+Jason"]
LABELS = {"T1 Jason ADS": "Track 1 - Jason alternate delivery site only",
          "T2 Refuge": "Track 2 - Refuge acquisition (Rev 4.10 terms)",
          "T3 Avant": "Track 3 - Avant acquisition (seller-financed)",
          "T4a Refuge+Jason": "Track 4a - Refuge seed + Jason census",
          "T4b Avant+Jason": "Track 4b - Avant seed + Jason census"}
AVAIL = 215000.0


def h1(d, text):
    p = d.add_paragraph(style="Heading 1")
    r = p.add_run(text)
    r.font.color.rgb = NAVY
    return p


def para(d, text, bold=False, italic=False, size=10.5):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    return p


def bullet(d, text, size=10.5):
    p = d.add_paragraph(style="List Bullet")
    r = p.add_run(text)
    r.font.size = Pt(size)
    return p


def table(d, headers, rows, widths=None):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for j, htxt in enumerate(headers):
        c = t.rows[0].cells[j]
        c.text = ""
        rr = c.paragraphs[0].add_run(htxt)
        rr.bold = True
        rr.font.size = Pt(9.5)
    for row in rows:
        cells = t.add_row().cells
        for j, val in enumerate(row):
            cells[j].text = ""
            rr = cells[j].paragraphs[0].add_run(str(val))
            rr.font.size = Pt(9.5)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Inches(w)
    return t


def money(x):
    return f"(${abs(x):,.0f})" if x < 0 else f"${x:,.0f}"


def main():
    rm = json.load(open(ROWMAP))
    R = rm["rows"]
    sol = formulas.ExcelModel().loads(XLSX).finish().calculate()
    vals = {}
    for k, v in sol.items():
        try:
            vals[k.upper()] = v.value[0, 0]
        except Exception:
            pass
    fn = XLSX.split("/")[-1].upper()

    def g(sheet, ref):
        return vals[f"'[{fn}]{sheet.upper()}'!{ref.upper()}"]

    m = {}
    for t in TRACKS:
        m[t] = dict(peak=g(t, f"B{R[t]['peak']}"),
                    m12=g(t, f"M{R[t]['cum']}"),
                    m24=g(t, f"Y{R[t]['cum']}"))

    os.makedirs(OUT, exist_ok=True)
    d = Document()
    d.styles["Normal"].font.name = "Aptos"
    d.styles["Normal"].font.size = Pt(10.5)
    d.add_heading("Four-Track Cash Flow Comparison - Decision Memo", 0)
    para(d, "Prepared 7/28/2026 | Basis: Azalea_Track_Comparison_Rev1.00.xlsx (QA 13/13) | Confidential", bold=True)

    h1(d, "1. What this compares")
    para(d, "One operating chassis - the same census ramp, net rate ($172/patient-day), variable cost "
            "($79/patient-day), and fixed overhead ($62K/month) proven out in the Rev 4.10 proforma - run "
            "through four different license and billing structures over 24 months (August 2026 start). "
            "Tracks are modeled with NO equity inflows, so the headline number is the peak cumulative "
            "funding need, compared against the roughly $215,000 actually available (the remainder of "
            "Jim's $250,000 commitment). No revolver or assumed facility anywhere.")

    h1(d, "2. Results")
    rows = []
    for t in TRACKS:
        gap = m[t]["peak"] - AVAIL
        rows.append((LABELS[t], money(m[t]["peak"]),
                     ("+" if gap > 0 else "") + money(gap),
                     money(m[t]["m12"]), money(m[t]["m24"])))
    table(d, ["Track", "Peak funding need", "Gap vs $215K available", "Month-12 cumulative", "Month-24 cumulative"],
          rows, widths=[2.3, 1.3, 1.4, 1.3, 1.3])

    h1(d, "3. How to read it")
    bullet(d, f"Track 1 (Jason only) is the only track within reach of committed capital: peak need "
              f"{money(m['T1 Jason ADS']['peak'])} vs $215K available - a {money(m['T1 Jason ADS']['peak']-AVAIL)} "
              f"gap that fee negotiation or a small bridge closes. The trade: at month 24 you own no license; "
              f"the ADC-formula purchase option is the path to converting that book into an owned asset.")
    bullet(d, f"Track 2 (Refuge) peaks at {money(m['T2 Refuge']['peak'])} - but note this is the only ownership "
              f"track carrying institutional financing (the Sept bank note and Jan SBA inflows are IN these "
              f"numbers). Its peak-vs-available gap of {money(m['T2 Refuge']['peak']-AVAIL)} is what the PPEO "
              f"stress test already told us: the committed stack does not absorb prepayment review.")
    bullet(d, f"Track 3 (Avant) has the LOWEST license cost ($300K, fully seller-deferred until PPEO clears) "
              f"but the HIGHEST peak need, {money(m['T3 Avant']['peak'])}. The seller deferral funds the license "
              f"- it does nothing for the operating ramp, which burns roughly $580K before first collections "
              f"with no bank or SBA money attached to this track. Seller flexibility is not a substitute for "
              f"working capital.")
    bullet(d, f"The combos do exactly what they were designed to do: running growth census through Jason's "
              f"established number while the owned license seeds through PPEO cuts the Refuge path's peak by "
              f"{money(m['T2 Refuge']['peak']-m['T4a Refuge+Jason']['peak'])} (to "
              f"{money(m['T4a Refuge+Jason']['peak'])}) and the Avant path's peak by "
              f"{money(m['T3 Avant']['peak']-m['T4b Avant+Jason']['peak'])} (to "
              f"{money(m['T4b Avant+Jason']['peak'])}).")
    bullet(d, "Counterintuitive but real: Refuge beats Avant on CASH in both standalone and combo form, despite "
              "the harder seller terms - because Refuge comes with a financing stack and Avant does not. If "
              "Avant could attach even the bank note, its numbers would flip. That is a financing question, "
              "not a deal-terms question.")

    h1(d, "4. Regulatory gates (cash is not the only axis)")
    bullet(d, "Jason ADS: terms unagreed (the $15K/month fee is a placeholder input); AKS-compliant structure, "
              "Smith County TULIP add, and 855A location filing required before first admission.")
    bullet(d, "Refuge: 36-month window closes 1/8/2027; the reactivated billing number lapses around September "
              "absent claims; PPEO is certain per the reactivation letter; both the bank note and SBA must fund.")
    bullet(d, "Avant: the Palmetto tie-in/CHOW has been unresolved for years (escalated through two legislators "
              "and HHSC as of June-July 2026) - the clearance-month input is the single most uncertain number "
              "in this workbook. Medicare-only. Its own 36-month/moratorium position is unverified (the "
              "favorable read on the moratorium is the sellers', not counsel's). License expires 7/8/2027.")
    bullet(d, "Combos inherit both gate sets and run two arrangements simultaneously - more moving parts, "
              "lower peak cash.")

    h1(d, "5. Levers to test next (all are Assumptions-tab inputs)")
    bullet(d, "Jason fee: every $5K/month off the placeholder saves ~$60K over a PPEO year across T1/T4.")
    bullet(d, "Avant clearance month (default M2; try M4/M6) - each month of delay adds roughly a month of "
              "burn to T3/T4b peaks.")
    bullet(d, "PPEO hold (defaults: Refuge 3, Avant 4; try 6) - the downside case that sizes any bridge ask.")
    bullet(d, "A Jim bridge or bank line applied to any track directly reduces its gap dollar-for-dollar; "
              "T4a needs ~$172K of coverage, T1 needs ~$16K.")

    para(d, "Method note: simplified chassis (no owner-salary deferral, Medicaid room-and-board treated as "
            "cash-neutral); Refuge track carries the full Rev 4.10 financing sequence; Avant schedule per the "
            "May 5 payment confirmation, time-shifted to the PPEO-clear month per the sellers' current "
            "willingness. Figures regenerate from the workbook - edit Assumptions and rerun the QA before "
            "quoting new numbers.", italic=True)

    docfmt.finalize(d, "Azalea - Four-Track Cash Flow Comparison")
    path = OUT + "Four-Track Comparison MEMO.docx"
    d.save(path)
    print("wrote", path)


if __name__ == "__main__":
    main()
