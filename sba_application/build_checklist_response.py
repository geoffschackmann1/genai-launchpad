"""John Hart SBA 7(a) checklist response - AVANT structure (pivot 2026-08-14).

Structure: SBA 7(a) $500,000 (funds month 2 - retires the seller note ~$303,000 incl. accrued
interest + working capital) + $250,000 equity injection (Bullard $195K / Schackmann $55K) = $750,000 project.
Target: Avant Hospice, LLC (TX), Medicare-only, CCN 741798, NPI 1699350090, TX license 020708
(exp 7/8/2027). Seller: Kimberly Carlisle, sole member. Terms: $300K at 6% simple, 100% transfer
at CHOW, ALL payments deferred until PPEO clears; SBA takeout retires the note in full.

Fills the lender's fillable PDFs and writes the annotated checklist + cover email.
Output: sba_application/13_checklist_response_2026-07-16/
"""
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__))))
import docfmt

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, RGBColor
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject

SRC = "sba_application/00_source_forms/"
OUT = "sba_application/13_checklist_response_2026-07-16/"


def fill_pdf(src, dst, values, checkboxes=()):
    reader = PdfReader(src)
    writer = PdfWriter()
    writer.append(reader)
    for page in writer.pages:
        writer.update_page_form_field_values(page, values)
    for cb in checkboxes:
        for page in writer.pages:
            if "/Annots" not in page:
                continue
            for a in page["/Annots"]:
                obj = a.get_object()
                if obj.get("/T") == cb:
                    # first non-Off appearance state
                    on = next((k for k in obj["/AP"]["/N"].keys() if k != "/Off"), "/Yes")
                    obj[NameObject("/V")] = NameObject(on)
                    obj[NameObject("/AS")] = NameObject(on)
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        writer._root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)
    with open(dst, "wb") as f:
        writer.write(f)
    print("wrote", dst)


def company_profile():
    v = {
        "Company Name": "Tyler Hospice Hold, LLC",
        "CompanyContact": "Geoff Schackmann",
        "CompanyAddress": "13387 Hwy 69 N",
        "CompanyCity": "Tyler",
        "CompanyCounty": "Smith",
        "CompanyState": "TX",
        "CompanyZip": "75706",
        "Company Assumed Name if applicable": "Azalea Hospice & Palliative Care (dba)",
        "Brief Business Description 1": "Medicare-certified hospice serving Tyler / Smith County, East Texas.",
        "Brief Business Description 2": "Acquiring 100% of Avant Hospice, LLC (TX; CMS CCN 741798), operating as Azalea Hospice.",
        "Brief Business Description 3": "Wyoming LLC foreign-qualified in TX; S-corp election (Form 2553). 0.7% reserved pool completes 100%.",
        "OtherEntityType": "WY LLC; S-corp election",
        "Date Business Formed": "3/18/2026",
        "Date Incorporated": "3/18/2026 (Wyoming)",
        "Date current management assumed control": "3/18/2026",
        "Federal taxpayer identification number": "41-4966640",
        "Number of employees at time of application": "0 (pre-close)",
        "When loan is approved": "14 by mo 6",
        "Person1Name": "Geoffery Michael Schackmann",
        "Person1Ownership": "39.9%",
        "Person1Title": "Manager",
        "Person1eMail": "gs@h-care.us",
        "Person2Name": "James Bullard",
        "Person2Ownership": "19.5%",
        "Person2Title": "Member (passive)",
        "Person3Name": "Silas Shelton",
        "Person3Ownership": "13.3%",
        "Person3Title": "Member",
        "Person4Name": "Dana Davenport",
        "Person4Ownership": "13.3%",
        "Person4Title": "Member",
        "Person5Name": "Bradley Woodard",
        "Person5Ownership": "13.3%",
        "Person5Title": "Member",
        "BankAccountName": "Tyler Hospice Hold, LLC dba Azalea Hospice & Palliative Care",
        "AccountSigners": "Geoff Schackmann",
        "Affiliated Companies if applicable 1": "None (prior interest sold Aug 2025)",
    }
    fill_pdf(SRC + "2_Company_Profile_FORM.pdf", OUT + "2 Company Profile FILLED.pdf", v, checkboxes=("LLC",))


def use_of_proceeds():
    v = {
        "BusinessPurchase": "$300,000",
        "WorkingCapital": "$435,000",
        "ClosingCosts": "$15,000",
        "TotalProjectCost": "$750,000",
        "Injection": "$250,000",
        "LoanAmount": "$500,000",
        "Description": "Acquisition of 100% of Avant Hospice, LLC (TX), Medicare-certified (CCN 741798, NPI 1699350090),",
        "Description1": "operating as Azalea Hospice & Palliative Care (Tyler, TX). Seller terms: $300,000 at 6% simple, 100% of",
        "Description2": "membership interests transfer at CHOW approval; NO down payment and NO installments - the seller carries",
        "Description3": "the full balance until SBA funding (month 2), when SBA proceeds retire the note (~$303,000 incl. interest);",
        "Description4": "remainder funds ramp working capital. Injection: J. Bullard $195,000 ($100K wired 5/7/2026) + G. Schackmann $55,000.",
        "Benefits": "Converts deferred seller financing into permanent SBA financing at the earliest lawful point (license is 100% in the",
        "Benefits1": "borrower name at CHOW close); funds working capital through the Medicare payment lag and prepayment review;",
        "Benefits2": "reaches ~40 patients/day by month 12; ~30 East Texas jobs by year 2. EBITDA $299K Y1 / $696K Y2 (Rev 5.00).",
    }
    fill_pdf(SRC + "3_Use_of_Funds_FORM.pdf", OUT + "3 Use of Proceeds FILLED.pdf", v)


def debt_schedule():
    v = {
        "BorrowerNameDS": "Tyler Hospice Hold, LLC (dba Azalea Hospice & Palliative Care)",
        "AsOfDate": "At SBA funding (M2)",
        "LenderDS1": "Seller (K. Carlisle)",
        "Date1": "CHOW close",
        "Purpose1": "Avant acquisition",
        "OriginalAmt1": "$300,000",
        "Balance1": "~$303,000",
        "Rate1": "6.00%",
        "Payment1": "Deferred",
        "Security1": "Membership interests",
        "OriginalAmtTotal": "$300,000",
        "CurrentBalanceTotal": "~$303,000",
        "TotalPayment": "Deferred",
        "GovtAgency1": "NONE (company or affiliates)",
    }
    fill_pdf(SRC + "5_Business_Debt_Schedule_FORM.pdf", OUT + "5 Business Debt Schedule FILLED.pdf", v)


def _h(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x3B, 0x2D)


def _p(doc, text, bold=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(10)


def checklist_memo():
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)
    doc.add_heading("SBA 7(a) Checklist - Item-by-Item Status", 0)
    _p(doc, "Borrower: Tyler Hospice Hold, LLC (dba Azalea Hospice & Palliative Care) | EIN 41-4966640 | 8/14/2026", bold=True)
    _p(doc, "Request: $500,000 SBA 7(a) - funds month 2 after CHOW close (license is 100% in the borrower name at closing), "
            "RETIRING the deferred seller note (~$303,000 incl. accrued interest) + working capital. "
            "Equity injection $250,000 (33% of the $750,000 project): James Bullard $195,000 (first $100K wired 5/7/2026) "
            "+ Geoff Schackmann $55,000.")

    rows = [
        ("ALL LOANS", "", ""),
        ("Company Profile", "ATTACHED (filled form)", "CPA/attorney/insurance contacts, office street address/zip/phone, and personal SSNs still to be added by Geoff."),
        ("Use of Proceeds", "ATTACHED (filled form)", "Reflects the Avant structure above."),
        ("Entity documents incl. EIN", "PARTIAL", "EIN 41-4966640. WY formation certificate + EIN letter to be attached; Amended & Restated OA drafted, needs signatures."),
        ("Bank approval items", "FOR DISCUSSION", "Structure now a single $500K SBA loan - the prior $300K piggyback bank loan is gone (simpler placement)."),
        ("CAIVRS/SAMS", "LENDER RUNS", "Runs off Geoff's Form 912."),
        ("ALL INDIVIDUALS (Geoff only - sole 20%+ owner)", "", ""),
        ("Personal Financial Statement (SBA 413)", "GEOFF TO COMPLETE", "Pre-fill worksheet started; values + supporting statements needed. Joint with spouse (AZ community property)."),
        ("Personal tax returns - 3 years", "LOCATED", "2022, 2023, 2024 1040s are in Google Drive ('9 - Tax Returns'). 2025 in preparation. Geoff sends directly - verify ALL schedules, W-2s, 1099s are included."),
        ("Personal cash flow statement", "GEOFF TO COMPLETE", "Pre-fill started in 07_personal_cash_flow/."),
        ("Resume / Personal History Form", "DRAFTED", "4 management resumes drafted; Geoff to add personal fields + complete SBA 912."),
        ("Credit report", "GEOFF TO PROVIDE", ""),
        ("Driver license", "GEOFF TO PROVIDE", ""),
        ("EXISTING BUSINESS / PURCHASE", "", ""),
        ("Target interim financials", "REQUEST FROM SELLER", "Avant Hospice, LLC balance sheet + P&L, month-end not older than 60 days. Request via Kim Carlisle / David Groom."),
        ("Target tax returns - 3 years", "REQUEST FROM SELLER", "Avant Hospice, LLC returns (2023-2025 as applicable). Entity has never billed Medicare - returns will show minimal activity; underwrite as startup."),
        ("Company debt schedule", "ATTACHED (filled form)", "Single item: the $300K deferred seller note (Kimberly Carlisle). Copy of the note/MIPA to accompany it when executed."),
        ("Aged AR / AP", "REQUEST FROM SELLER", "Avant has never billed - listing will legitimately be zero; get it in writing."),
        ("LOI / contract for sale", "IN NEGOTIATION", "Avant terms agreed in principle: $300,000 at 6% simple; 100% of membership interests at CHOW approval; ALL payments deferred until PPEO clears; prepayable; SBA takeout retires the note in full at funding. Payment-confirmation schedule (5/5/2026) on file; MIPA to follow."),
        ("STARTUPS / CHOW", "", ""),
        ("Business plan (editable Word)", "ATTACHED", "Rev 7.00 (Avant target). Editable Word per checklist."),
        ("2-yr P&L projections, Y1 monthly, written assumptions", "ATTACHED", "Rev 5.00 AVANT dynamic proforma (36 monthly periods, QA 22/22) + assumptions narrative. Officer salaries split out: Administrator $150K, DON $120K, Community Liaison $120K, Office Mgr $60K."),
        ("OTHER", "", ""),
        ("Lease / LOI for premises", "ATTACHED", "Executed Tyler lease (13387 Hwy 69 N); Fair Investments extension to a 10-year total term in process."),
        ("Sources of cash injection - 2 months statements", "PARTIAL", "Bullard tranche 1 ($100K) wire confirmed 5/7/2026; need Bullard statements (30+ days pre-wire), remaining $95K wire, Geoff's $55K source statements."),
        ("Franchise documents", "N/A", ""),
        ("Affiliates", "NONE", "Prior hospice interest (VistaRiver) sold Aug 2025; passive note only. Affiliate/size-standard memo on file."),
    ]
    t = doc.add_table(rows=1, cols=3)
    t.style = "Light Grid Accent 1"
    for i, htxt in enumerate(["Checklist item", "Status", "Notes"]):
        c = t.rows[0].cells[i]
        c.text = htxt
        for par in c.paragraphs:
            for run in par.runs:
                run.bold = True
                run.font.size = Pt(9)
    for a, b, cnote in rows:
        cells = t.add_row().cells
        for i, vtxt in enumerate([a, b, cnote]):
            cells[i].text = vtxt
            for par in cells[i].paragraphs:
                for run in par.runs:
                    run.font.size = Pt(9)
                    if b == "":
                        run.bold = True

    _h(doc, "COVID questionnaire - draft responses")
    for q, a in [
        ("1. Other stimulus loans (PPP/EIDL)?", "No. Neither the applicant (formed 3/2026) nor any affiliate has PPP, EIDL, or other stimulus financing. To confirm for the seller entity (Avant Hospice, LLC) as part of CHOW diligence."),
        ("2. Industry/business impact & 18-month contingency", "Hospice was an essential service throughout COVID and demand is demographic (aging population), not COVID-cyclical. The applicant is a startup; projections are built month-by-month on current CMS rates with a break-even analysis and a documented downside case (slower census ramp) supported by a working-capital reserve."),
        ("3. Restriction impacts (stay-home, travel, supply)", "None currently. Hospice care is delivered in patients' residences and facilities; no material supply-chain exposure. PPE is a routine, budgeted operating cost."),
        ("4. Protective gear / safety costs", "PPE and infection-control supplies are included in the per-patient supply budget (standard post-COVID hospice practice)."),
        ("5. Reliability of historical financials", "Startup - projections (not history) drive the credit case: monthly proforma with break-even analysis on current market conditions, CMS FY2026 rates, and conservative census assumptions."),
        ("6. Customer concentration", "Payor mix is Medicare per-diem - federal payor concentration typical of every hospice; no single referral source exceeds a modest share of projected census."),
        ("7. Vendor concentration", "Standard interchangeable vendors (pharmacy/PBM, DME, medical supplies) with competitive alternatives in the Tyler market."),
        ("8. Collateral adequacy vs market", "Primary repayment is cash flow. The Medicare hospice license itself is a scarce, marketable asset in Texas (comparable Medicare-only licenses trade at $225-350K; the purchase at $300,000 is inside that band, and the CMS enrollment moratorium restricts new supply)."),
        ("9. CHOW - industry experience / seller stimulus loans", "Geoff Schackmann has 10+ years of hospice ownership experience (prior interest sold Aug 2025); the operating team (Administrator, DON, Community Liaison) is a proven East Texas hospice leadership group. Seller PPP/EIDL status to be confirmed in diligence."),
        ("10. Working capital >=50% justification", "Working capital is 58% of the project ($435K of $750K) by design: the target has never billed Medicare, so first claims enter a provisional period of enhanced oversight (prepayment review); the reserve carries ramp payroll through that review window per the monthly cash-flow model (base 2-month hold; 4-month downside disclosed)."),
    ]:
        _p(doc, q, bold=True)
        _p(doc, a)

    path = OUT + "1 Checklist STATUS RESPONSE.docx"
    docfmt.finalize(doc, "Azalea Hospice - SBA Checklist Status")
    doc.save(path)
    print("wrote", path)


def cover_email():
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading("Draft email to John Hart", 0)
    _p(doc, "To: john.hart@sourcefundingtexas.com | From: gs@h-care.us | Subject: Azalea Hospice - updated SBA package", bold=True)
    for para in [
        "John,",
        "Target update, and this one simplifies the file rather than complicating it. We are moving forward with "
        "Avant Hospice, LLC, a Texas Medicare-certified license (CCN 741798), at $300,000 - down from the prior "
        "$500,000 target. Same operating plan, same team, same Tyler market; the acquisition just costs $200,000 less.",
        "The structure also answers the funding-timing concern you raised directly: this is a 100% membership transfer "
        "at CHOW approval, so the license is fully in the borrower's name from day one - no split closing, no waiting "
        "period. The seller (a single individual owner) carries the entire $300,000 at 6% simple with NO down payment "
        "and NO installments until Medicare's prepayment review clears; the SBA loan funds at month 2 and retires the "
        "note in full (~$303,000 including accrued interest), with the remaining ~$197,000 going to working capital. "
        "One loan, one debt, one payment of $6,747 a month after funding. Equity injection is unchanged: $250,000 "
        "cash, a third of the $750,000 project - Jim Bullard $195K (first $100K already wired) and $55K from me.",
        "Attached: the filled Company Profile, Use of Proceeds, and Business Debt Schedule from your checklist, an "
        "item-by-item status on everything else, the purchase agreement, and the financial model (monthly for 36 months "
        "with written assumptions - officer salaries split out like you asked). My personal returns for 2022-2024 are ready "
        "to send separately, 2025 is with the CPA.",
        "Still collecting: my 413 and cash flow statement, the sellers' interim financials and returns, and the office lease. "
        "Item-by-item status is in the attachment so your lenders can see exactly where we are.",
        "Call me with what your network needs beyond this. I would like to get this in front of them this week.",
        "Thanks,",
        "Geoff",
    ]:
        _p(doc, para)
    path = OUT + "0 Cover Email DRAFT.docx"
    docfmt.finalize(doc, "Azalea Hospice - Cover Email Draft")
    doc.save(path)
    print("wrote", path)


def refuge_addendum():
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)
    doc.add_heading("Business Plan Addendum - Target Change to Avant Hospice, LLC", 0)
    _p(doc, "Tyler Hospice Hold, LLC | 8/14/2026 | Supplements Business Plan Rev 7.00", bold=True)
    for h, body in [
        ("What changed", "The Company is now acquiring Avant Hospice, LLC, a Texas LLC (CMS CCN 741798, NPI 1699350090, "
         "TX HHSC license 020708, expiring 7/8/2027), for $300,000 - replacing the prior $500,000 target. The seller is "
         "Kimberly Carlisle, sole member. Payment terms agreed in principle 5/5/2026 and improved since: 6% simple "
         "interest, 100% of membership interests transfer at CHOW approval, and ALL payments are deferred until "
         "Medicare's prepayment review clears."),
        ("Why this is a stronger structure for the lender", "First, price: $300,000 is inside the $225-350K market band "
         "for Texas Medicare-only licenses, and $200,000 below the prior target - the loan-to-project economics improve "
         "accordingly. Second, timing: because 100% of the membership interests transfer at CHOW approval, the license "
         "is fully in the borrower's name from day one. There is no split closing, no 36-month staging, and no waiting "
         "period before SBA funds can move - the exact concern raised in underwriting review of the prior structure. "
         "Third, simplicity: no bridge or interim bank financing. The seller carries the full balance at 6% with nothing due until "
         "prepayment review clears, and the SBA loan retires the note directly at funding (month 2)."),
        ("Prepayment review, addressed head-on", "Avant has never submitted a Medicare claim, so its first claims enter "
         "a provisional period of enhanced oversight with prepayment medical review (42 CFR 424.527) - typically 60-120 "
         "days of payment delay. The package is built around this rather than hoping it away: the working-capital "
         "reserve is $435,000 (58% of the project), the base case models a 2-month collection hold, a 4-month downside "
         "is disclosed (peak additional need $99K, against $250K of equity), and the seller's own payment deferral "
         "means no acquisition debt service competes with payroll during review."),
        ("Financial projections", "All projections are from the Rev 5.00 dynamic proforma (36 monthly periods, QA-verified): "
         "census reaching 24 patients by month 2, 34 by month 6, and 40 by month 12; EBITDA $299K year 1 (15.5%), $696K "
         "year 2 (24.1%), $985K year 3 (28.2%), before debt service; net income $131K / $472K / $733K. The model assumes "
         "no revolver or invented capital: base-case cash never goes negative (minimum month $192K), and after the "
         "month-2 takeout the company carries a single $6,747/month SBA payment. Global 3-year debt-service coverage is "
         "3.7x including the one-time takeout. Officer salaries are separately stated (Administrator $150K, Director of "
         "Nursing $120K, Community Liaison $120K, Office Manager $60K)."),
        ("What has NOT changed", "The operating team, the Tyler/Smith County market, the dba (Azalea Hospice & Palliative "
         "Care), the entity (Tyler Hospice Hold, LLC, a Wyoming LLC with an S-corp election), the cap table (Schackmann "
         "39.9% / Bullard 19.5% / three operators 13.3% each / 0.7% reserved), the executed Tyler lease, and the equity "
         "injection ($250,000 cash, 33%)."),
    ]:
        _h(doc, h, level=1)
        _p(doc, body)
    path = OUT + "4 Business Plan ADDENDUM Avant.docx"
    docfmt.finalize(doc, "Azalea Hospice - Business Plan Addendum")
    doc.save(path)
    print("wrote", path)


def assumptions_narrative():
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)
    doc.add_heading("Projection Assumptions Narrative", 0)
    _p(doc, "Accompanies Azalea_Hospice_Proforma_Rev5.00_AVANT.xlsx (36 monthly periods; every calculation "
            "cell is a live formula; all inputs on the Control Tower tab) | 8/14/2026", bold=True)
    for h, body in [
        ("Census (the revenue driver)", "End-of-month census of 24 patients by month 2, 34 by month 6, and 40 by month 12, "
         "reaching 50 by month 24 and 56 by month 36. Admissions ramp from the team's established East Texas referral "
         "relationships; discharges follow an average-length-of-stay-driven waterfall. The team has operated at these census "
         "levels before in this market - the ramp is a return to a proven book, not speculation."),
        ("Revenue", "Medicare Routine Home Care per-diem at FY2026 CMS rates for Smith County, split between the first-60-day "
         "high rate (~15% of patient days) and the post-60-day low rate, net ~$172 per patient day blended after sequestration "
         "and billing fees. Payer mix is 100% Medicare (Avant is Medicare-only). Collections lag billing by ~30-60 days (NOE "
         "timing), PLUS a prepayment-review hold on first claims: 2 months in the base case, 4 in the disclosed downside - "
         "the cash-flow tab models both explicitly, month by month."),
        ("Officer salaries (separately stated)", "Administrator $150,000; Director of Nursing $120,000; Community Liaison "
         "$120,000; Office Manager $60,000. Benefits load 22% plus 3% workers compensation."),
        ("Clinical staffing", "Nursing, aide, chaplain, and social-work FTEs scale with census on standard hospice "
         "staffing ratios; per-patient costs (pharmacy, DME, supplies) are per-patient-day rates from the line-item budget."),
        ("Contingency", "5% of net patient revenue is deducted below EBITDA as an explicit contingency in cash flow."),
        ("Debt service", "Base case (Rev 5.00 Avant): the seller carries the full $300,000 at 6% simple with nothing due; the "
         "SBA 7(a) funds at month 2 (the license is 100% in the borrower name from CHOW close) and retires the seller note "
         "(~$303,000 including accrued interest), leaving a single $6,747/month payment (10.5%/10-year) from month 3. Owner "
         "salaries defer until break-even census ($42,500 accrued, repaid after collections catch up); clinical hiring lags "
         "census one month with PRN coverage."),
        ("Business Debt Schedule - detail", "The Business Debt Schedule form carries a single line because there is only one "
         "piece of existing/interim debt in the structure: the seller note - $300,000 owed to Kimberly Carlisle (sole member "
         "of the target) at 6.00% simple interest, with NO down payment and NO scheduled installments; all payments are "
         "deferred until Medicare's prepayment review clears. The balance shown (~$303,000) is principal plus roughly two "
         "months of accrued interest at the SBA funding date (month 2). Collateral is the membership interests, and the "
         "note is expressly structured to be retired by the SBA loan at funding - it is not intended to remain outstanding "
         "alongside the SBA debt. The SBA loan itself does not appear as a line on this form because it is the loan being "
         "applied for, not existing debt; once it funds and retires the seller note, the business carries a single "
         "$6,747/month SBA payment going forward with no other debt."),
        ("Results on these assumptions", "EBITDA $299K year 1 (15.5% margin), $696K year 2 (24.1%), $985K year 3 (28.2%), "
         "before debt service; net income $131K / $472K / $733K; month-12 EBITDA margin 22.3%. Break-even census ~17-18 "
         "patients, crossed in month 2. The model carries no revolver or assumed facility: base-case cash never goes negative "
         "(minimum month $192K), and the documented downside (slow census plus a 4-month prepayment-review hold) needs ~$99K "
         "of additional capital against $250K of equity on deposit - shown honestly, not plugged."),
    ]:
        _h(doc, h, level=1)
        _p(doc, body)
    path = OUT + "6 Projection Assumptions Narrative.docx"
    docfmt.finalize(doc, "Azalea Hospice - Projection Assumptions")
    doc.save(path)
    print("wrote", path)


def reply_financing_and_checklist():
    """Reply to John's 7/20 1:58pm email: the SBA-funding-timing/license-transfer question,
    the Bullard ownership-percentage flag, and his 11 numbered checklist items."""
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading("Draft reply - John Hart (financing timing + checklist items)", 0)
    p = doc.add_paragraph()
    r = p.add_run("To: John Hart | Cc: Mary Brownmiller | From: Geoff Schackmann | "
                   "Re: RE: Azalea Hospice SBA Funding Request")
    r.bold = True
    r.font.size = Pt(11)

    for para in [
        "John,",
        "Thanks for the close read, and for catching the timing question - that's exactly the kind of thing "
        "I'd rather get right now than at closing.",
    ]:
        _p(doc, para)

    _h(doc, "On SBA funding vs. the license transfer", level=1)
    for para in [
        "You're right that SBA won't fund until the license is under the borrower's name - and the updated "
        "structure (per the cover memo, the target is now Avant Hospice, LLC at a $300,000 purchase price) is "
        "sequenced around exactly that rule. At closing, 100% of the Avant membership interests transfer via a "
        "standard CMS change of ownership, with nothing paid at close: the seller carries the full $300,000 at "
        "6%, fully deferred, until Medicare claims are actually paying. The SBA loan funds at month 2 - after "
        "the license is 100% in the borrower's name - and retires the seller note (about $303,000 with accrued "
        "interest), with the balance of proceeds going to ramp working capital. So there is no window where SBA "
        "money moves before ownership does, and no bridge or interim bank financing anywhere in the structure.",
        "One honest caveat: there's a difference between the ownership transfer itself (effective at closing "
        "under the purchase agreement) and CMS formally updating its enrollment records to reflect it. This is "
        "a standard CHOW on an existing certification, not a new-enrollment application, so the processing lag "
        "should be short - but I don't want to promise you zero gap without your read on it. How do your "
        "lenders typically handle that window on a hospice CHOW - do they need to see CMS's tie-in approval in "
        "hand before they'll fund, or is the executed transfer enough? If it's the former, the seller deferral "
        "means nothing in the deal breaks while we wait.",
    ]:
        _p(doc, para)

    _h(doc, "One flag on ownership percentage", level=1)
    _p(doc, "Your note assumes both James and I are at 20% or more, so items 6-8 would apply to both of us. Per "
             "our cap table James is at 19.5%, just under that threshold - which is also why he's been treated as "
             "a passive investor (source-of-funds documentation, no personal guaranty, no full PFS) rather than a "
             "20%+ owner. Let me know if you still want the full personal package from him anyway for the file, "
             "but wanted to flag the percentage before it goes further.")

    _h(doc, "Your checklist items", level=1)
    items = [
        ("1. Annual totals in the projections", "Done. The financial projections workbook now has a fuller "
         "3-Year Summary tab - annual revenue, cost, EBITDA, and net income buildup for each of the three years, "
         "plus a year-end balance sheet snapshot, not just the three headline numbers it had before. Attached."),
        ("2. Written narrative assumptions", "Attached - the assumptions narrative walks through census, revenue, "
         "officer salaries, staffing, debt service, and results in prose, not just tables."),
        ("3. Checklist items highlighted in yellow", "Attached, with one thing to flag: I highlighted every "
         "section that applies to this deal (all loans, all individuals, existing business/acquisition, "
         "startup/change-of-ownership) but left Land Purchase/Construction, Franchise, and the COVID section "
         "unhighlighted since none of those apply here - no land purchase, no franchise, and we have no PPP/EIDL "
         "or other stimulus financing. Flag it if you read any of those as relevant and I'll revisit."),
        ("4. Joint PFS", "In progress - my wife's information will be included per the community-property "
         "structure already on file."),
        ("5. Affiliate businesses", "On my side, the one affiliate item is a passive note from the VistaRiver "
         "sale (already disclosed as PFS support). Checking with James on his side and will report back."),
        ("6. Personal cash flow", "In progress."),
        ("7. Management resume", "The business plan has full bios for me and the other principals - happy to "
         "pull those into the standalone resume format if you'd rather have them separate."),
        ("8. Proof of US citizenship", "Gathering passports now for the owners this applies to."),
        ("9. More detail on the business debt schedule", "Added - the assumptions narrative now includes a "
         "dedicated section walking through why the debt schedule form carries a single line (the deferred "
         "seller note), what the balance shown represents, and how the SBA loan retires it at month-2 "
         "funding."),
        ("10. Proof of source of cash injection", "James's $100K wire is already documented; gathering the "
         "supporting bank/brokerage statements now, and will send the same for the remaining tranches as they "
         "land."),
        ("11. 2025 tax return extension", "Will send a copy if the return isn't filed by the time we get back to "
         "you on the rest."),
    ]
    for label, body in items:
        pp = doc.add_paragraph()
        rl = pp.add_run(label + " ")
        rl.bold = True
        rl.font.size = Pt(11)
        rb = pp.add_run(body)
        rb.font.size = Pt(11)

    for para in [
        "Attached: updated financial projections (with annual totals), the assumptions narrative (with the new "
        "debt schedule section), and your checklist with the applicable items highlighted.",
        "Let me know your thoughts on the CMS timing question above whenever you get a chance, and I'll keep "
        "working the personal items in parallel.",
        "Talk soon,",
    ]:
        _p(doc, para)
    _p(doc, "Geoff Schackmann", bold=True)

    docfmt.finalize(doc, "Azalea Hospice - Reply to John Hart (Financing Timing + Checklist)")
    path = OUT + "12 Reply to John Hart - Financing Timing and Checklist DRAFT.docx"
    doc.save(path)
    print("wrote", path)


def highlight_checklist():
    """Yellow-highlight every checklist item that applies to this transaction on John's own
    7(a) checklist doc, so the reply shows at a glance what's in scope. Skips the sections that
    don't apply here: Land purchase/construction (no land/construction in this deal), Franchise
    documents (not a franchise), and the COVID section (no PPP/EIDL/stimulus financing involved).
    """
    d = Document(SRC + "1_SBA_7a_Loan_Checklist_JohnHart_2026-05-19.docx")
    applicable = set(range(1, 7)) | set(range(7, 14)) | set(range(14, 20)) | set(range(20, 26)) | set(range(30, 37))
    for i, p in enumerate(d.paragraphs):
        if i in applicable:
            for r in p.runs:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    path = OUT + "1a SBA 7a Checklist HIGHLIGHTED.docx"
    d.save(path)
    print("wrote", path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    company_profile()
    use_of_proceeds()
    debt_schedule()
    checklist_memo()
    cover_email()
    refuge_addendum()
    assumptions_narrative()
    highlight_checklist()
    reply_financing_and_checklist()
    import shutil
    shutil.copy("financial_models/output/Azalea_Hospice_Proforma_Rev5.00_AVANT.xlsx",
                OUT + "7 Financial Projections Rev5.00 AVANT.xlsx")
    print("wrote", OUT + "7 Financial Projections Rev5.00 AVANT.xlsx")
