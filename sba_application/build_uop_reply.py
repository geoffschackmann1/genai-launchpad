"""Reply email draft to John Hart + Mary Brownmiller and lease summary with flags.

Output: sba_application/13_checklist_response_2026-07-16/{8_Reply_Email_John_Mary_DRAFT,9_Lease_Summary_and_Flags}.docx
"""
from docx import Document
from docx.shared import Pt, RGBColor

OUT = "sba_application/13_checklist_response_2026-07-16/"


def _p(doc, text, bold=False, size=11):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    return p


def _h(doc, text):
    h = doc.add_heading(text, level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x3B, 0x2D)


def reply_email():
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading("Draft reply - SBA Follow up & Questions thread", 0)
    _p(doc, "To: John Hart; Mary Brownmiller | From: gs@h-care.us | Re: SBA Follow up & Questions", bold=True)
    for para in [
        "John and Mary,",
        "Structure update, and I think you'll both like it. The deal has moved to a stronger asset and a simpler ask.",
        "We are acquiring Refuge Hospice, a Texas license certified for both Medicare and Medicaid, ready to bill day "
        "one. Purchase agreement is at $500,000. The seller is carrying the balance short term: $125K down at closing "
        "for 49%, $31,250 a month from September at 6%, balance in January when the 36 month CMS rule lets the other "
        "51% transfer. The note is prepayable without penalty.",
        "The ask: one SBA 7(a) loan of $500,000. It takes out the seller note in September and funds working capital "
        "for the ramp. Equity injection is $250,000 cash, a full third of the $750,000 project. Jim Bullard $195K "
        "(first $100K wired in May), $55K from me. No piggyback bank loan anymore.",
        "John, your note on the $672K working capital reserve landed. The new structure needs $235,000, not $672,000, "
        "because the license bills immediately and the seller carries the acquisition until the SBA takeout. The "
        "attached workbook breaks the $235K into components with a month by month draw schedule. Draws taper to zero "
        "around month nine as Medicare collections catch payroll. I have no issue with the bank controlling "
        "disbursements against census milestones and said so in the schedule.",
        "Mary, on your priority list: my 2022, 2023 and 2024 personal returns are ready to send (2025 is with the "
        "preparer). Personal financial statement, cash flow and resume are in progress this week. I'll pull the credit "
        "report on Credit Karma like you suggested. The seller note is on the debt schedule. Also attached is the "
        "fully executed office lease for our Tyler location: 24 months plus two 12 month options, signed June 29.",
        "One item I want to be straight about: the lease tenant reads Hickory Hospice, LLC dba Azalea Hospice, the "
        "entity from the earlier deal that we set up. Since the operating company will now be Refuge Hospice, LLC "
        "under Tyler Hospice Hold, we will assign or amend the lease to the right entity before closing, and the "
        "landlord relationship is friendly. Flag it for the lender file however you prefer.",
        "Is your lender's acquisition promotion still live, and would this profile qualify for an exception under the "
        "$1.5M threshold? Happy to jump on a call this week.",
        "Thanks,",
        "Geoff",
    ]:
        _p(doc, para)
    path = OUT + "8_Reply_Email_John_Mary_DRAFT.docx"
    doc.save(path)
    print("wrote", path)


def lease_summary():
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)
    doc.add_heading("Executed Office Lease - Summary & Flags", 0)
    _p(doc, "Gary House / fully executed 6/29/2026 (DocuSign) | reviewed 7/16/2026", bold=True)

    _h(doc, "Key terms")
    t = doc.add_table(rows=0, cols=2)
    t.style = "Light Grid Accent 1"
    for k, v in [
        ("Premises", "13387 Hwy 69 N, Tyler, TX 75706 (Smith County) - the historic 'Gary House', ~1,607 sq ft, single-tenant, City of Tyler ETJ"),
        ("Landlord", "Fair Investments, Ltd., by Fair Management, LC (John R. Garrett, President). Notices: PO Box 689, Tyler TX 75710; cassie.hollenshead@fairoil.com"),
        ("Tenant", "Hickory Hospice, LLC d/b/a Azalea Hospice & Palliative Care - signed by Geoff Schackmann, Managing Member, 6/26/2026"),
        ("Term", "24 months: 7/1/2026 - 6/30/2028"),
        ("Options", "2 additional 12-month terms (tenant option, 120-day notice, CPI-adjusted rent) - total potential term 4 years"),
        ("Base rent", "$1,607.00/mo year 1 ($12.00/rsf/yr); $1,647.18/mo year 2 ($12.30/rsf/yr)"),
        ("Additional rent", "Net CAM reimbursement, 100% pro-rata share, projected $4.00/rsf/yr (~$535.67/mo). Landlord pays taxes, insurance, structural/roof"),
        ("All-in occupancy", "~$2,143/mo year 1 (base + projected CAM) - matches the working-capital detail rent line"),
        ("Security deposit", "$2,142.67 (paid at execution)"),
        ("Utilities", "Landlord: water, sewer, electric, gas, alarm. Tenant: phone, internet, cable, trash"),
        ("Early termination", "Tenant may terminate on 6 months notice + fee of 50% of remaining base rent (capped at $50,000)"),
        ("Guaranty", "Commercial Lease Guaranty addendum (TXR-2109) attached - confirm guarantor identity"),
        ("Special conditions", "Historic-structure restrictions (no structural/aesthetic changes; historic designation default trigger); signage on existing pole only, landlord approval required"),
    ]:
        row = t.add_row().cells
        row[0].text = k
        row[1].text = v
        for par in row[0].paragraphs:
            for run in par.runs: run.bold = True
        for cidx in (0, 1):
            for par in row[cidx].paragraphs:
                for run in par.runs: run.font.size = Pt(9)

    _h(doc, "Flags for the SBA file (and for counsel)")
    for x in [
        "1. TENANT ENTITY: the tenant is Hickory Hospice, LLC, the operating entity from the terminated Hickory deal. "
        "The operating company under the current structure is Refuge Hospice, LLC (under Tyler Hospice Hold, LLC). "
        "Assign or amend the lease to the correct entity before SBA closing. Also resolve whether Hickory Hospice, LLC "
        "(Geoff signing as Managing Member) is a Geoff-controlled entity today - if so it is an AFFILIATE that must be "
        "disclosed on the application (the Company Profile currently answers 'none'), even if dormant. Consistency here "
        "matters on a federal application.",
        "2. TERM vs LOAN TERM: 24 months + two 12-month options = 4 years maximum. The lender checklist wants the lease "
        "(with options) to run the term of the loan (10 years). Expect the lender to ask for additional renewal options "
        "or a landlord letter; the early-termination right cuts both ways and may actually help negotiate this.",
        "3. LANDLORD AGREEMENT: John's checklist notes SBA will want a lease assignment or landlord's "
        "subordination/agreement. Landlord contact is Fair Investments (Bob Garrett) - the relationship is warm.",
        "4. HISTORIC-PROPERTY DEFAULT TRIGGER: any action costing the property its Texas Historic Designation is a "
        "lease default. Operationally fine for an office, but flag to insurance and staff.",
        "5. RENT MATCHES THE MODEL: ~$2,143/mo all-in is consistent with the rent line in the Rev 3.00 proforma and the "
        "working-capital draw schedule - no reconciliation needed.",
    ]:
        _p(doc, x, size=10)

    path = OUT + "9_Lease_Summary_and_Flags.docx"
    doc.save(path)
    print("wrote", path)


if __name__ == "__main__":
    reply_email()
    lease_summary()
