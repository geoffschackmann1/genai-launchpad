"""Preliminary information package for William at healthfmv.com (Refuge Hospice FMV).

Compiled from the Gmail thread history with the sellers (Jorge Resendiz / Dennis Hendrix,
DHJR Limited Partnership dba the Refuge Hospice ownership group) and Jim Bullard, plus
the executed MIPA. Two things flagged prominently for accuracy: (1) the CMS PTAN went
through a deactivation/reactivation cycle (approval dated 3/20/2026) rather than
continuous billing since the 1/8/2024 CCN date; (2) no financials or tax returns have
been provided by the sellers as of this package - the three years that exist are
expected to show minimal revenue, consistent with a license-only operation.

Output: sba_application/15_valuation_2026-07-20/
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docfmt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor

OUT = "sba_application/15_valuation_2026-07-20/"
NAVY = RGBColor(0x1F, 0x38, 0x64)


def h1(d, text):
    p = d.add_paragraph(style="Heading 1")
    r = p.add_run(text)
    r.font.color.rgb = NAVY
    return p


def h2(d, text):
    p = d.add_paragraph(style="Heading 2")
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


def table(d, headers, rows, widths=None):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for j, htxt in enumerate(headers):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(htxt)
        r.bold = True
        r.font.size = Pt(9.5)
    for row in rows:
        cells = t.add_row().cells
        for j, val in enumerate(row):
            cells[j].text = ""
            r = cells[j].paragraphs[0].add_run(str(val))
            r.font.size = Pt(9.5)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Inches(w)
    return t


def info_package():
    d = Document()
    d.styles["Normal"].font.name = "Aptos"
    d.styles["Normal"].font.size = Pt(10.5)
    title = d.add_heading("Refuge Hospice, LLC - Preliminary Information for FMV Evaluation", 0)
    para(d, "Prepared for William, healthfmv.com | Compiled from correspondence with the sellers "
            "(July 2-16, 2026) | 7/20/2026", bold=True)
    para(d, "This is a preliminary information package to start the valuation process. It is compiled "
            "directly from email correspondence and the purchase agreement in negotiation, not from "
            "audited records. Two items are flagged below because they are directly relevant to a "
            "license valuation and should be confirmed before the valuation is finalized.", italic=True)

    h1(d, "1. Entity and transaction summary")
    table(d, ["Item", "Detail"], [
        ("Target entity", "Refuge Hospice, LLC, a Texas limited liability company"),
        ("Sellers", "Angel J. Resendiz and Dennis M. Hendrix, through DHJR Limited Partnership "
                    "(4831 Whirlwind Drive, San Antonio, TX 78217)"),
        ("Buyer", "Tyler Hospice Hold, LLC, a Wyoming LLC (Geoff Schackmann, Managing Member)"),
        ("Broker / intermediary", "Jorge (jorge@journeyshospicetx.com / jorge@dhjrlp.com), introduced "
                                  "via a mutual contact (Ann)"),
        ("Structure", "Two-step transfer under the CMS 36-month post-certification ownership rule: "
                      "49% of membership interests at closing, remaining 51% on the first permissible "
                      "date after the 36-month window"),
        ("Asking / agreed price", "$500,000 for 100% of the membership interests"),
    ], widths=[1.7, 4.6])

    h1(d, "2. License and certification detail (as represented by the sellers, 7/2/2026)")
    table(d, ["Item", "Detail"], [
        ("License status", "Active"),
        ("Service area", "Texas HHSC Region 8 - Atascosa, Bandera, Bexar (San Antonio), Calhoun, "
                          "Comal, DeWitt, Dimmit, Edwards, Frio, Gillespie, Goliad, Gonzales, "
                          "Guadalupe, Jackson, Karnes, Kendall, Kerr, Kinney, La Salle, Lavaca, "
                          "Maverick, Medina, Real, Uvalde, Val Verde, Victoria, Wilson, Zavala"),
        ("CMS Certification Number (CCN/PTAN)", "A91679"),
        ("PECOS enrollment status", "Active"),
        ("CHAP accreditation", "Effective 1/8/2024 - expires 1/8/2027"),
        ("CLIA number", "45D2282700, effective 5/23/2025 - expires 5/23/2027"),
        ("CMS 36-month rule", "Currently within the 36-month window from the 1/8/2024 certification "
                              "date; window expires 1/8/2027"),
        ("Medicaid (HHSC) contract", "Contract No. HHS000004700408; Provider No. 001035622; "
                                     "Cross Ref. No. 3342; Service Area Region 8; Service Codes "
                                     "1, 1A, 4, 24, 30, 31, 32, 33; Effective 6/26/2026, expires 12/23/2028"),
    ], widths=[2.0, 4.3])

    h1(d, "3. Flag - PTAN deactivation and reactivation")
    para(d, "On 7/3/2026 the sellers provided a CMS reactivation approval letter and confirmation "
            "email, both dated 3/20/2026, described as \"official confirmation from CMS ... for the "
            "reactivation of our PTAN and billing privileges.\" This indicates the PTAN was not "
            "continuously active for Medicare billing between the 1/8/2024 certification date and the "
            "March 2026 reactivation - it was reactivated only a few months before this sale process "
            "began. The reason for the deactivation is not stated in the correspondence on file and "
            "should be confirmed directly with the sellers (common causes include a lapsed "
            "revalidation cycle or a period of non-billing; either should be verifiable from the CMS "
            "letter itself). This is relevant to valuation because it bears on how long the license has "
            "actually been billing-capable versus simply certified.", bold=True)

    h1(d, "4. Flag - current operations and financial history")
    para(d, "The sellers described the agency directly, in their initial outreach (7/2/2026): "
            "\"the agency is clean, zero pts no staff and it is ready to take pts and start billing as "
            "early as today.\" No patients, no staff, and no material billing activity are represented "
            "as of the outreach date. Three years of entity tax returns and financial statements are "
            "expected to exist given the 1/8/2024 certification date, but based on the sellers' own "
            "description they will show minimal to no revenue. As of this package, the sellers have not "
            "yet provided financial statements or tax returns through email; those have been requested "
            "separately and will be forwarded when received. The valuation should be understood as "
            "primarily a license/certification valuation rather than a going-concern business "
            "valuation - there is no material operating history or patient census being acquired.", bold=True)

    h1(d, "5. Negotiated purchase terms")
    para(d, "The purchase price of $500,000 was agreed without negotiation. The payment structure was "
            "negotiated over several exchanges (7/3-7/6/2026) and is reflected in the Membership "
            "Interest Purchase Agreement now in redline:")
    table(d, ["Term", "Detail"], [
        ("Purchase price", "$500,000 for 100% of the membership interests"),
        ("Down payment", "$125,000 at closing (~8/1/2026), concurrent with the 49% transfer"),
        ("Installments", "$31,250/month, September 2026 through December 2026"),
        ("Interest rate", "6% per annum on the unpaid balance from day one"),
        ("Final payment", "Approximately $258,489.46 due 1/15/2027, the first date past the 36-month rule, "
                          "concurrent with the remaining 51% transfer"),
        ("Security", "Sellers hold a security interest in the company and a personal guaranty from the "
                     "buyer until paid in full; prepayable at any time without penalty"),
    ], widths=[1.7, 4.6])
    para(d, "Deal documents currently in redline (as of 7/16/2026): Membership Interest Sale and "
            "Purchase Agreement, Promissory Note, Guaranty, and Security Agreements (LLC membership "
            "interest as collateral) for each seller. These are available on request.")

    h1(d, "6. Market context")
    para(d, "Comparable Texas hospice license transactions reviewed during this process: Medicare-only "
            "license shells have traded at approximately $225,000-$350,000; Medicare and Medicaid "
            "dual-certified licenses (Refuge's category) at approximately $400,000-$500,000+; and "
            "licenses with material existing operations at $1,200,000-$1,700,000. Texas has an "
            "effective freeze on new Medicaid hospice enrollment, and CMS has a nationwide hospice "
            "enrollment moratorium in effect naming Texas, both of which constrain new supply and "
            "support values for existing certified licenses. A separate comparison memo with sourcing "
            "is available on request.")

    h1(d, "7. What we can provide on request")
    for item in [
        "The Membership Interest Purchase Agreement (current redline) and related transaction documents",
        "The sellers' CCN/PTAN reactivation letter and CMS confirmation email (3/20/2026)",
        "The CHAP accreditation certificate",
        "The Medicaid HHSC contract documentation",
        "Sellers' financial statements and tax returns, once received",
        "The buyer entity's formation documents and the operating team's background",
    ]:
        p = d.add_paragraph(style="List Bullet")
        r = p.add_run(item)
        r.font.size = Pt(10.5)

    docfmt.finalize(d, "Refuge Hospice - Preliminary Info for FMV Evaluation")
    path = OUT + "Refuge Hospice - Preliminary Information for FMV Evaluation.docx"
    d.save(path)
    print("wrote", path)


def email_to_william():
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading("Draft email to William (healthfmv.com)", 0)
    p = doc.add_paragraph()
    r = p.add_run("To: William | From: geoff@azaleahospice.com | Re: Preliminary information - Refuge Hospice license valuation")
    r.bold = True
    r.font.size = Pt(11)
    for text in [
        "William,",
        "I'm working on an SBA 7(a) loan for the acquisition of a Texas hospice license, Refuge Hospice, "
        "LLC, and my lender has confirmed a third party valuation will be required. I'd like to get the "
        "process started with you.",
        "I've put together everything I have so far in the attached summary: the license and "
        "certification details (CMS CCN, CHAP accreditation, Medicaid HHSC contract), the negotiated "
        "purchase terms, and market comps I pulled together on comparable Texas hospice license sales.",
        "Two things I want to flag directly rather than let you find them later. First, the seller's "
        "CMS PTAN went through a deactivation and reactivation - the reactivation approval is dated "
        "March 2026, so it wasn't continuously billing since the original 2024 certification date. I "
        "don't yet know the reason and I'm asking the sellers directly. Second, the agency has no "
        "patients and no staff right now - the sellers describe it as clean and ready to bill, not an "
        "active operation. So this is really a license and certification valuation, not a going concern "
        "valuation. Three years of tax returns technically exist but I expect them to show close to "
        "nothing, and I'll send them over once the sellers provide them.",
        "Let me know what else you need from me to get going, and what your timeline and fee structure "
        "look like. I'd like to move on this soon.",
        "Thanks,",
        "Geoff",
    ]:
        para(doc, text, size=11)
    docfmt.finalize(doc, "Azalea Hospice - Email to William (healthfmv.com)")
    path = OUT + "Email to William - Preliminary Valuation Info DRAFT.docx"
    doc.save(path)
    print("wrote", path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    info_package()
    email_to_william()
