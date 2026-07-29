"""Proposal to Jason: alternate delivery site in Tyler under his hospice's license.

Jason-facing document. Embeds the PPEO stress-test numbers (financial_models/
ppeo_stress.py, Rev 4.10 basis) as the honest "why now" and presents three
structure options ordered for a risk-averse partner, each inside a recognized
AKS safe harbor, with Jason's protections and a defined ADC-priced exit.

Placeholders: [JASON'S LAST NAME], [JASON'S HOSPICE, LLC] - fill before sending.
Output: sba_application/19_jason_proposal_2026-07-28/
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docfmt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor

OUT = "sba_application/19_jason_proposal_2026-07-28/"
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


def proposal():
    d = Document()
    d.styles["Normal"].font.name = "Aptos"
    d.styles["Normal"].font.size = Pt(10.5)

    d.add_heading("Tyler Alternate Delivery Site - Partnership Proposal", 0)
    para(d, "Prepared for Jason [JASON'S LAST NAME] | From Geoff Schackmann, Azalea Hospice | "
            "July 28, 2026 | Confidential - discussion draft, not a binding offer", bold=True)
    para(d, "Everything in this document is a starting point for a conversation, not a demand. "
            "The three structures in Section 3 are ordered by how much protection they give you, "
            "and I'm genuinely open to any of them. I've done the regulatory homework up front "
            "(Section 5) so your counsel can verify rather than start from scratch.", italic=True)

    h1(d, "1. What I'm asking for, in one paragraph")
    para(d, "Azalea has a complete hospice operation ready in Tyler: a leased and equipped office at "
            "13387 Highway 69 North, a clinical team (RN case management, CNAs, social work, chaplain "
            "coverage, a contracted medical director), insurance binding August 1, and patients ready "
            "to admit today. What we don't have is a billable Medicare number: our license acquisition "
            "is in process, but its billing number comes with a mandatory prepayment-review period that "
            "creates a payment gap I'd rather not fund with patients waiting. I'm proposing that your "
            "hospice add Tyler as an additional delivery site under your existing license and CCN, with "
            "my team operating it for a defined interim period, until our own license is fully in place. "
            "You'd add Smith County to your HHSC service area (a TULIP amendment) and report the Tyler "
            "location on your CMS-855A. We would run it; you would own it, control it, and be protected "
            "as set out below.")

    h1(d, "2. Why now - the honest numbers")
    para(d, "I want you to see the actual reason for the timing rather than a sales pitch. The license "
            "we're acquiring was reactivated by its Medicare contractor in March 2026, and under 42 CFR "
            "424.527 it enters a provisional period of enhanced oversight, with prepayment medical "
            "review, the moment it submits its first claim. Prepayment review commonly delays payment "
            "60 to 120+ days. We stress-tested our financial model against those timelines. The results, "
            "on top of our full committed capital stack:")
    table(d, ["Scenario", "Additional capital needed", "Cash goes negative", "Peak receivables held in review"], [
        ("Normal 30-day processing", "$0", "Never", "$306K (ordinary A/R)"),
        ("Prepayment review ~90 days", "$155,489", "Month 3", "$444,001"),
        ("Prepayment review ~120+ days", "$314,640", "Month 3", "$602,004"),
        ("~120+ days plus a slow census ramp", "$312,163", "Month 3 (14 months underwater)", "$511,703"),
    ], widths=[2.1, 1.6, 1.5, 1.6])
    para(d, "Two things worth noting. First, in every scenario the money is not lost - it is sitting in "
            "the review queue and releases when claims clear; our 36-month cash position ends identically. "
            "This is a timing problem, not a viability problem. Second, that timing problem is exactly what "
            "billing through your established, review-free number for an interim period solves. That's the "
            "entire economic logic of this proposal, and it's why the arrangement is genuinely temporary: "
            "the moment our own number clears review and our acquisition completes, the reason for this "
            "arrangement ends.")

    h1(d, "3. Three ways to structure it - your pick")
    para(d, "Each option keeps your hospice as the sole licensed, certified, enrolled operator of the "
            "Tyler site. The differences are in how my team plugs in and how I'm compensated. Each sits "
            "inside a recognized federal safe harbor, detailed in Section 5.")

    h2(d, "Option 1 - Employment (simplest answer to the ownership question: there isn't one)")
    para(d, "My Tyler clinical team and I become W-2 employees of your hospice for the interim period. "
            "You run payroll; we run the site day to day under your policies and your compliance program. "
            "Compensation can include a performance bonus tied to the Tyler site's results, because "
            "payments to bona fide employees are statutorily excepted from the federal Anti-Kickback "
            "Statute - this is the one structure where performance-based pay is expressly protected. "
            "Your exposure is ordinary employer exposure, covered by your existing workers' comp and "
            "our professional liability arrangements. Ownership of your company does not change by a "
            "single percentage point.")

    h2(d, "Option 2 - Management services agreement (flat fee, arm's length)")
    para(d, "My management company operates the Tyler site under a written management agreement with "
            "your hospice - the same structure we're using today to operate under our seller's license, "
            "so the documents are already built and lawyered. You'd hold reserved approval rights over "
            "admissions commencement, clinical leadership hires, contracts over a threshold, and anything "
            "touching your license or enrollment. Compensation is a flat, fair-market-value monthly fee "
            "set in advance for the term - deliberately NOT a percentage of revenue, because the "
            "personal-services safe harbor requires compensation that doesn't vary with the volume or "
            "value of business generated. Again, no ownership change.")

    h2(d, "Option 3 - Ring-fenced joint venture (only if you'd rather I have skin in the game)")
    para(d, "A new subsidiary of your hospice holds the Tyler site; I buy a bona fide minority stake "
            "(under 40%, at fair value, on the same terms any investor would get) and my team operates "
            "it under a management agreement per Option 2. Your existing company stays completely walled "
            "off. This is the most paperwork and the most counsel time, and I list it only because some "
            "partners prefer an operator who owns a piece of the outcome over a vendor or employee. If "
            "that's not you, Options 1 and 2 are cleaner.")

    h1(d, "4. Your protections, regardless of option")
    bullet(d, "You are the licensee. Nothing transfers your license, CCN, provider agreement, or any "
              "ownership interest. The agreement will say so expressly.")
    bullet(d, "Reserved approval rights: first admissions, clinical leadership changes, contracts and "
              "spending over an agreed threshold, and any filing touching CMS or HHSC require your "
              "written sign-off.")
    bullet(d, "Banking: Tyler-site revenue lands in your account. You keep an independent signatory at "
              "all times and can revoke any access we're granted, at will, on notice.")
    bullet(d, "Indemnification: we indemnify you for anything arising from the Tyler site's operations "
              "during the term; you're named additional insured on our professional and general "
              "liability coverage (binding August 1) and the site carries workers' comp from day one.")
    bullet(d, "Termination: you can terminate on short notice if your regulatory counsel concludes the "
              "arrangement threatens your enrollment or licensure - same clause our current interim "
              "agreement carries.")
    bullet(d, "A written, dated exit (below), so this never becomes an open-ended entanglement.")

    h1(d, "5. The regulatory homework (for your counsel)")
    bullet(d, "Mechanics: Smith County is added to your HHSC service area via a TULIP amendment; the "
              "Tyler office is reported as an additional practice location on your CMS-855A via PECOS. "
              "Multiple locations under one CCN are a single hospice for Medicare purposes - no new "
              "certification, no new survey cycle beyond your normal one, and your accreditor is "
              "notified per its reporting rules.")
    bullet(d, "Option 1 rests on the bona fide employment exception to the Anti-Kickback Statute "
              "(42 U.S.C. 1320a-7b(b)(3)(B); 42 CFR 1001.952(i)) - performance pay to W-2 employees "
              "is expressly protected.")
    bullet(d, "Option 2 is built to the personal services and management contracts safe harbor "
              "(42 CFR 1001.952(d)): written agreement of at least one year, services and methodology "
              "specified in advance, compensation at fair market value and not determined by volume or "
              "value of business generated.")
    bullet(d, "Option 3 is built to the small entity investment safe harbor (42 CFR 1001.952(a)(2)), "
              "including its caps on investor composition and revenue and its same-terms requirement.")
    bullet(d, "I'm not asking you to take my word on any of this - both sides should have healthcare "
              "regulatory counsel bless the final structure before a single patient is admitted. Ours "
              "is already engaged on the parallel license acquisition.")

    h1(d, "6. The exit - defined, dated, and priced in advance")
    para(d, "This arrangement ends when our own license is billable, which we currently expect in the "
            "first quarter of 2027. At that point, and only with your consent at signing rather than a "
            "negotiation later, we exercise a pre-agreed option to transition the Tyler site's patients, "
            "staff, and referral relationships to our license, priced by a formula we set now: a "
            "per-patient-of-average-daily-census amount anchored to current market comparables, applied "
            "to the census actually transferred on the transition date. You get paid for exactly the "
            "value that walks; there's nothing to argue about later because the formula, not the number, "
            "is what we sign. If we never stand up our own license, you keep the Tyler site, its census, "
            "and its staff - which is to say, your downside case is that you acquired a functioning "
            "Tyler operation at no acquisition cost.")

    h1(d, "7. Suggested next steps")
    bullet(d, "Pick the option (or tell me what's wrong with all three).")
    bullet(d, "We each have healthcare counsel review a one-page term sheet before any drafting.")
    bullet(d, "TULIP county amendment and 855A location filing prepared for your signature - my team "
              "does the paperwork, you review and file.")
    bullet(d, "Target: first Tyler admission under your license within 2-3 weeks of the term sheet.")

    docfmt.finalize(d, "Tyler Alternate Delivery Site - Proposal for Jason")
    path = OUT + "Tyler Alternate Site Proposal - Jason DRAFT.docx"
    d.save(path)
    print("wrote", path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    proposal()
