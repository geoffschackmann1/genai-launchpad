"""Generate the draftable SBA application documents for Azalea Hospice.

Fills everything derivable from the business plan + operating model. Personal
data (SSNs, addresses, personal financials) is left as clearly marked blanks.
Clean, professional business tone; no machine-generated tells.
"""
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

BASE = os.path.dirname(os.path.abspath(__file__))
NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x60, 0x60, 0x60)
TBD = "__________ (to provide)"

# ------------------------------------------------------------------ helpers
def new_doc(title):
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Calibri"; st.font.size = Pt(10.5)
    return d

def h1(d, text):
    p = d.add_paragraph(); r = p.add_run(text)
    r.bold = True; r.font.size = Pt(15); r.font.color.rgb = NAVY
    p.space_after = Pt(4)
    return p

def h2(d, text):
    p = d.add_paragraph(); r = p.add_run(text)
    r.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = NAVY
    p.space_before = Pt(8); p.space_after = Pt(2)
    return p

def para(d, text, italic=False, size=10.5, color=None, bold=False):
    p = d.add_paragraph(); r = p.add_run(text)
    r.italic = italic; r.bold = bold; r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return p

def field(d, label, value):
    p = d.add_paragraph()
    r = p.add_run(f"{label}:  "); r.bold = True
    p.add_run(value if value else TBD)
    return p

def table(d, headers, rows, widths=None):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for j, htext in enumerate(headers):
        c = t.rows[0].cells[j]; c.text = ""
        r = c.paragraphs[0].add_run(htext); r.bold = True; r.font.size = Pt(9.5)
    for row in rows:
        cells = t.add_row().cells
        for j, val in enumerate(row):
            cells[j].text = ""
            r = cells[j].paragraphs[0].add_run(str(val)); r.font.size = Pt(9.5)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Inches(w)
    return t

def footer_note(d):
    para(d, "")
    para(d, "Prepared for the SBA 7(a) loan application of Tyler Hospice Hold LLC "
            "(dba Azalea Hospice & Palliative Care). Confidential.", italic=True,
         size=8, color=GREY)

def save(d, subpath):
    out = os.path.join(BASE, subpath)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    d.save(out)
    print("  wrote", subpath)


# ================================================================ 1. COMPANY PROFILE
def company_profile():
    d = new_doc("Company Profile")
    h1(d, "Company Profile")
    para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC", color=GREY, size=9)
    para(d, "")
    field(d, "Company Name", "Tyler Hospice Hold LLC")
    field(d, "Company Assumed Name (dba)", "Azalea Hospice & Palliative Care")
    field(d, "Brief Business Description",
          "Medicare-certified hospice and palliative care agency serving the Tyler / "
          "East-Texas market, operated through wholly owned subsidiary Hickory Hospice LLC "
          "via a change-of-ownership acquisition of an existing provider number.")
    field(d, "Business Entity", "Corporation (X)  |  Limited Liability Company")
    field(d, "Date Entity Formed", "March 18, 2026 (Wyoming LLC)")
    field(d, "Date current management assumed control", "At CHOW close (see business plan, Section 10)")
    field(d, "Federal Taxpayer Identification Number (EIN)", "41-4966640")
    field(d, "Web Address", TBD)
    field(d, "Address", TBD + "  (principal office, Tyler, TX area)")
    field(d, "City / County / State / Zip", "Tyler / Smith / TX / " + TBD)
    field(d, "Contact / Phone / Fax", TBD)
    field(d, "Number of employees at application", "9 (Year-1 staggered roster)")
    field(d, "Number of employees when loan approved", "Scales to ~25+ FT by Year 3 on census triggers")
    para(d, "")
    h2(d, "Company Ownership (100% must be shown)")
    table(d,
          ["#", "Name and Address", "SSN / EIN", "Ownership %", "Company Title", "Email"],
          [["1", "Adeline & Lilah, LLC (AZ) - Geoff Schackmann, sole member", TBD, "39.9%", "Managing Member", TBD],
           ["2", "James Bullard", TBD, "19.9%*", "Minority investor member (passive; no management or control role)", TBD],
           ["3", "Silas R. Shelton", TBD, "13.3%", "Executive Director", TBD],
           ["4", "Dana L. Davenport", TBD, "13.3%", "Director of Nursing", TBD],
           ["5", "Bradley G. Woodard", TBD, "13.3%", "Director of Sales", TBD]],
          widths=[0.3, 2.4, 1.0, 0.8, 1.4, 1.2])
    para(d, "*James Bullard holds a direct, fully funded minority equity interest of 19.9% - below the 20% "
            "threshold - as a passive investor. He has no management authority, no voting control, and no side "
            "agreement granting him control of the business; the Managing Member (Adeline & Lilah, LLC) retains "
            "control. Consistent with SOP 50 10 8 and 13 CFR 120.160, an equity holder of less than 20% in a "
            "complete change of ownership is not required to provide a personal guaranty or Personal Financial "
            "Statement. Mr. Bullard's $250,000 cash capital contribution is the source of the equity injection "
            "for this transaction (see Use of Proceeds). Unissued/reserved pool: 0.3%.", italic=True, size=9)
    para(d, "")
    field(d, "Name to be used on business checking account", "Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)")
    field(d, "Designated signers on checking account", "Geoff Schackmann (Managing Member, sole signer)")
    para(d, "")
    h2(d, "Professional Advisors")
    field(d, "CPA Name / Firm / Phone", TBD)
    field(d, "Attorney Name / Firm / Phone", TBD)
    field(d, "Insurance Agent Name / Firm / Phone", TBD)
    para(d, "")
    h2(d, "Affiliated Companies")
    table(d, ["Affiliate", "Relationship", "# Employees"],
          [["Hickory Hospice LLC (TX)", "Wholly owned operating subsidiary (the acquired agency)", "Operating staff"],
           ["Adeline & Lilah, LLC (AZ)", "Holding entity for Managing Member's interest", "0"]],
          widths=[2.2, 3.6, 1.0])
    footer_note(d)
    save(d, "01_company_profile/Company_Profile_DRAFT.docx")


# ================================================================ 2. USE OF FUNDS
def use_of_funds():
    d = new_doc("Use of Funds")
    h1(d, "Use of Proceeds")
    para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC", color=GREY, size=9)
    para(d, "")
    h2(d, "Total Project Sources and Uses")
    table(d, ["Use of funds", "Amount", "Funding source"],
          [["Purchase of existing business - Hickory Medicare license (CHOW)", "$300,000", "Seller note (36 mo, 6%)"],
           ["Startup one-time costs (855A filing, TX licensure, EMR, supply stock, legal, contingency)", "$63,000", "SBA + equity"],
           ["Equipment purchase (computers, office furniture)", "$15,000", "SBA + equity"],
           ["Working-capital reserve (opening cash; funds ramp payroll and dual debt service)", "$672,000", "SBA + equity"],
           ["Total funds required", "$1,050,000", ""]],
          widths=[4.0, 1.1, 1.7])
    para(d, "")
    h2(d, "Loan Request Reconciliation")
    table(d, ["Line", "Amount"],
          [["Total funds required", "$1,050,000"],
           ["Less: borrower equity injection (cash capital contribution)", "($250,000)"],
           ["Less: seller financing (Hickory license note, non-SBA)", "($300,000)"],
           ["Total SBA 7(a) loan request", "$500,000"]],
          widths=[4.5, 1.5])
    para(d, "")
    h2(d, "Details on the transaction and use of working capital")
    para(d, "This is a change-of-ownership (CHOW) acquisition of Hickory Hospice LLC, an established, "
            "Medicare-certified hospice, structured as a purchase of 100% of its membership interests. The "
            "$300,000 purchase price for the existing Medicare provider number is fully seller-financed, so no "
            "cash leaves at close for the acquisition and the working-capital reserve is preserved. The acquired "
            "provider number, state license, and CHAP/ACHC accreditation transfer in the CHOW, so the agency is "
            "billing-ready from day one.")
    para(d, "The SBA loan proceeds and the equity injection fund the $672,000 working-capital reserve plus "
            "$78,000 of startup costs and equipment. The reserve carries ramp-period payroll and combined debt "
            "service (SBA loan plus seller note) while accounts receivable normalize to the ~45-day Medicare "
            "cycle, and holds a $25,000 minimum cash floor with the $100,000 working-capital line undrawn in the "
            "base case.")
    para(d, "Source of the borrower's equity injection: $250,000 in cash, provided as a capital contribution by "
            "minority member James Bullard (19.9% non-controlling interest) in exchange for his equity interest. "
            "This is a bona fide, non-borrowed equity injection equal to approximately 23.8% of total project "
            "cost - well above the 10% minimum required for a complete change of ownership under SOP 50 10 8. The "
            "injection is verified by Mr. Bullard's bank / brokerage statements and the capital-contribution wire "
            "into the company's account. Because Mr. Bullard is a passive investor holding less than 20% with no "
            "management or control rights and no side agreement, his contribution counts as qualifying equity "
            "without triggering a personal guaranty (consistent with SOP 50 10 8 and 13 CFR 120.160). The "
            "$300,000 Hickory seller note is not on standby and is therefore treated as acquisition debt within "
            "combined debt service, not as part of the equity injection.", italic=True, size=9.5)
    para(d, "")
    h2(d, "How will this loan benefit the company?")
    para(d, "The loan capitalizes the acquisition and relaunch of an established, billing-ready hospice on "
            "validated local economics, led by an experienced East-Texas clinical and business-development team "
            "with established referral relationships. It funds the working capital required to operate from day "
            "one through the receivables-normalization period and to grow census with the team's referral "
            "pipeline, producing combined debt-service coverage that clears the 1.25x floor in every year and "
            "strengthens as census builds.")
    footer_note(d)
    save(d, "02_use_of_funds/Use_of_Funds_DRAFT.docx")


# ================================================================ 4. BUSINESS DEBT SCHEDULE
def debt_schedule():
    d = new_doc("Business Debt Schedule")
    h1(d, "Schedule of Business Debts")
    para(d, "Borrower: Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)", color=GREY, size=9)
    field(d, "As of", TBD + "  (most recent month-end, not older than 60 days)")
    para(d, "")
    para(d, "Tyler Hospice Hold LLC is a newly formed holding entity (Wyoming, March 18, 2026) and carries "
            "no existing business debt prior to this transaction. The Hickory license seller note arises at "
            "close and is shown below as proposed closing debt.", italic=True)
    para(d, "")
    table(d, ["Lender", "Orig. date", "Purpose", "Orig. amount", "Current balance", "Rate", "Payment", "Collateral"],
          [["Seller (Hickory CHOW) - PROPOSED at close", "At close", "Acquire Medicare provider number (CHOW)",
            "$300,000", "$300,000", "6.00%", "$9,126.58/mo (36 mo)", "Membership interests / provider number"],
           ["(No other existing business debt)", "-", "-", "$0", "$0", "-", "$0", "-"]],
          widths=[1.5, 0.7, 1.6, 0.9, 0.9, 0.5, 1.0, 1.4])
    para(d, "")
    h2(d, "Government Financing")
    para(d, "Tyler Hospice Hold LLC, its owners, and affiliated companies have no prior or outstanding "
            "government financing (no PPP, EIDL, or other federal stimulus or government loans). "
            "Confirm for each 20%+ owner and affiliate.", italic=True)
    table(d, ["Government agency", "Orig. amount", "Date of request", "Current balance", "Status"],
          [["None", "-", "-", "-", "-"]], widths=[2.5, 1.2, 1.2, 1.2, 1.4])
    footer_note(d)
    save(d, "04_business_debt_schedule/Business_Debt_Schedule_DRAFT.docx")


# ================================================================ 6/8. MANAGEMENT RESUMES
RESUMES = {
 "Geoff_Schackmann": dict(
   name="Geoff Schackmann", title="Managing Member (39.9% via Adeline & Lilah, LLC)",
   summary="Multi-hospice operator and transaction-led growth leader with operational responsibility for "
           "Medicare-certified hospice and palliative-care agencies across multiple states. Sole member of "
           "Adeline & Lilah, LLC (AZ), the holding entity through which equity in Tyler Hospice Hold LLC is held. "
           "Leads transaction structuring, financing, capital allocation, and post-close integration.",
   experience=[
     ("Multi-state hospice operations", "Present", "Owner / Operator",
      "Operational responsibility for Medicare-certified hospice and palliative-care agencies across multiple "
      "states; change-of-ownership (CHOW) transactions and post-CHOW enrollment oversight; census growth from "
      "sub-30 to 100+ ADC under existing Medicare provider numbers; multi-site clinical operations under Texas "
      "HCSSA and Oregon hospice licensure; CHAP and Joint Commission accreditation; acquisition due diligence "
      "(clinical, financial, regulatory); lender and investor relationships across SBA 7(a), conventional, "
      "mezzanine, and convertible-note structures."),
   ],
   education=[("__________ (to provide)", "", "", "")]),
 "Bradley_Woodard": dict(
   name="Bradley G. Woodard", title="Director of Sales (13.3%)",
   summary="More than 25 years of East-Texas hospice business development and administration, with deep referral "
           "relationships across area hospitals, LTC/LTAC facilities, home-health agencies, and physician "
           "practices. Licensed Nursing Home Administrator (LNFA #7136).",
   experience=[
     ("East-Texas hospice (VP of Business Development)", "2021 - Present", "VP, Business Development",
      "Sustained a 40+ ADC referral book for 4.5 years across the East-Texas market."),
     ("Grace Hospice of East Texas", "(to provide)", "Census-building / BD leadership",
      "Grew census from 8 to 145+ patients."),
     ("Start-up hospices, Lufkin / Nacogdoches", "(to provide)", "Administrator / BD",
      "Built start-up census from 0 to 48 in six months; separate engagement 12 to 70+."),
     ("Tyler-market hospice (founding administrator)", "Licensed 2004", "Founding Administrator",
      "Founding administrator of one of the Tyler market's largest hospices; grew to 220 patients, 109 "
      "employees, and 62 volunteers, becoming the area's flagship operation."),
   ],
   education=[("Texas A&M University", "", "B.S.", "Bachelor of Science"),
              ("Licensed Nursing Home Administrator", "", "LNFA #7136", "License")]),
 "Silas_Shelton": dict(
   name="Silas R. Shelton", title="Executive Director (13.3%)",
   summary="Operational leadership of Azalea Hospice & Palliative Care, with day-to-day responsibility for "
           "regulatory compliance under Texas HCSSA and the Medicare Conditions of Participation, payer-mix "
           "management, and referral-source partnerships across Smith County and surrounding East-Texas counties.",
   experience=[
     ("Azalea Hospice & Palliative Care", "Present", "Executive Director",
      "Site leadership; direct reports include the Director of Nursing, the Director of Sales, and the full "
      "clinical and operational team. Prior experience to provide."),
   ],
   education=[("__________ (to provide)", "", "", "")]),
 "Dana_Davenport": dict(
   name="Dana L. Davenport", title="Director of Nursing (13.3%)",
   summary="Clinical leadership of the Azalea nursing team, responsible for Medicare Conditions of Participation "
           "compliance, interdisciplinary group (IDG) oversight, plan-of-care management, and clinical quality "
           "(QAPI). Leads RN case managers, hospice aides, social workers, and chaplains across the service area.",
   experience=[
     ("Azalea Hospice & Palliative Care", "Present", "Director of Nursing",
      "Clinical and IDG leadership; CoP compliance and QAPI. Nursing licensure and prior clinical roles to provide."),
   ],
   education=[("__________ (to provide)", "", "", "")]),
}

def management_resumes():
    for key, info in RESUMES.items():
        d = new_doc("Management Resume")
        h1(d, "Management Resume")
        para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)",
             color=GREY, size=9)
        para(d, "Complete one form per owner. Personal-identifier fields are left blank for the individual "
                "to complete.", italic=True, size=9)
        para(d, "")
        field(d, "Name (First, Middle, Last)", info["name"])
        field(d, "Company Title", info["title"])
        field(d, "Social Security Number", TBD)
        field(d, "Date of Birth / Place of Birth", TBD)
        field(d, "Residence Telephone / Business Phone", TBD)
        field(d, "Residence Address", TBD)
        field(d, "Previous Address (and dates)", TBD)
        field(d, "Spouse Name / SSN / DOB", TBD)
        field(d, "Military Service (branch / dates / rank)", TBD)
        para(d, "")
        h2(d, "Professional Summary")
        para(d, info["summary"])
        h2(d, "Work Experience (most recent first)")
        table(d, ["Company / Location", "From - To", "Title", "Duties / accomplishments"],
              [[c, dt, ti, du] for (c, dt, ti, du) in info["experience"]],
              widths=[1.6, 1.0, 1.3, 3.0])
        h2(d, "Education")
        table(d, ["Institution", "Dates", "Major / Credential", "Degree / Certificate"],
              [[a, b, c, e] for (a, b, c, e) in info["education"]],
              widths=[2.2, 1.0, 1.8, 1.6])
        footer_note(d)
        save(d, f"06_personal_history_resume/Management_Resume_{key}_DRAFT.docx")


# ================================================================ RECONCILIATION + BREAK-EVEN MEMO
def reconciliation_memo():
    d = new_doc("Reconciliation Memo")
    h1(d, "Memo: Business Plan vs. Operating Model - Financial Reconciliation")
    para(d, "Internal note - resolve before submitting to lender", color=GREY, size=9)
    para(d, "")
    para(d, "Issue: Business Plan Rev 4.00 (June 2, 2026) was generated from an earlier version of the "
            "operating model. Net revenue and census match the current model, but the cost base - and therefore "
            "EBITDA, net income, and DSCR - differ, because the current model carries a fuller, more defensible "
            "direct-care staffing build (PRN social work / LVN / chaplain converted to full-time, six RNs and six "
            "CNAs by Year 3, plus Quality/Compliance, Intake, and Volunteer Coordinator) and outsourced billing "
            "at 1.5% of revenue. The plan and the workbook must show the same figures before they go to the lender.")
    h2(d, "Side-by-side (annual)")
    table(d, ["Metric", "Plan Rev 4.00", "Current model", "Note"],
          [["Net revenue Y1 / Y2 / Y3", "$1.44M / $2.72M / $3.73M", "$1.44M / $2.73M / $3.73M", "Match"],
           ["EBITDA Y1", "$311,204 (21.6%)", "$302,522 (21.0%)", "Close"],
           ["EBITDA Y2", "$814,391 (29.9%)", "$679,524 (24.9%)", "Lower - richer staffing"],
           ["EBITDA Y3", "$1,185,619 (31.8%)", "$1,007,520 (27.0%)", "Lower - richer staffing"],
           ["Net income Y1 / Y2 / Y3", "$199K / $706K / $1,084K", "$190K / $571K / $905K", "Lower"],
           ["Combined DSCR Y1 / Y2 / Y3", "1.61x / 4.20x / 6.12x", "1.56x / 3.50x / 5.20x", "Lower, still strong"],
           ["Global 3-yr DSCR", "3.97x", "3.42x", "Both well above 1.25x"]],
          widths=[1.9, 1.7, 1.7, 1.5])
    h2(d, "Recommendation")
    para(d, "Update the Business Plan financials (Sections 11.4 and 11.6, plus the Executive Summary repayment "
            "tiles) to the current model. The current model is the version that is fully reconciled, error-free, "
            "and reflects realistic staffing for a 50+ ADC agency - the more defensible position with an SBA "
            "underwriter. The story holds: every year is profitable, EBITDA still grows from $303K to $1.0M, and "
            "combined DSCR clears 1.25x in every year (1.56x / 3.50x / 5.20x; global 3.42x), springing higher once "
            "the seller note retires after Year 3. The alternative (reverting the model to the leaner roster) would "
            "raise margins but invites the question of how a 54-ADC agency runs on so few clinical staff.")
    h2(d, "Corrected three-year P&L (drop-in for Section 11.4)")
    table(d, ["Line item", "Year 1", "Year 2", "Year 3"],
          [["Average daily census (ADC)", "21.8", "40.2", "53.8"],
           ["Net patient revenue", "$1,442,930", "$2,728,642", "$3,734,934"],
           ["Patient-related COGS", "$71,549", "$132,002", "$176,276"],
           ["Payroll & related", "$868,172", "$1,686,864", "$2,297,748"],
           ["Operating G&A", "$200,687", "$230,251", "$253,390"],
           ["EBITDA", "$302,522", "$679,524", "$1,007,520"],
           ["EBITDA margin", "21.0%", "24.9%", "27.0%"],
           ["Depreciation & amortization", "$35,600", "$35,600", "$35,600"],
           ["Interest (SBA + seller note)", "$71,479", "$62,558", "$52,623"],
           ["TX franchise tax", "$5,411", "$10,232", "$14,006"],
           ["Net income", "$190,032", "$571,134", "$905,290"]],
          widths=[2.6, 1.4, 1.4, 1.4])
    h2(d, "Corrected lender summary (drop-in for Section 11.6)")
    table(d, ["Metric", "Year 1", "Year 2", "Year 3"],
          [["EBITDA", "$302,522", "$679,524", "$1,007,520"],
           ["Combined debt service", "$193,876", "$193,876", "$193,876"],
           ["Combined DSCR", "1.56x", "3.50x", "5.20x"],
           ["Net income", "$190,032", "$571,134", "$905,290"]],
          widths=[2.6, 1.4, 1.4, 1.4])
    para(d, "Global 3-year combined DSCR: 3.42x (vs. 1.25x floor). SBA-only coverage springs far higher once "
            "the seller note retires at the end of Year 3.", bold=True)
    h2(d, "Break-even (answers Business Plan Guide VIII.3 - currently missing)")
    para(d, "Contribution margin is $160.52 per patient-day (net rate $181.30 less $20.78 of variable cost per "
            "patient-day: supplies/DME/pharmacy plus PRN visit labor plus the 1.5% billing and 0.75% QR fees). "
            "Against the Year-1 fixed-cost base (the lean opening roster plus fixed G&A), the operation breaks "
            "even at approximately 16 ADC, or about $86,000 per month of net revenue. The validated opening "
            "census of ~22 ADC is already above break-even, and Year-1 covers combined debt service at 1.56x.")
    footer_note(d)
    save(d, "03_business_plan/RECONCILIATION_and_Breakeven_MEMO.docx")


# ================================================================ INFORMATION NEEDED
def information_needed():
    d = new_doc("Information Needed")
    h1(d, "Information Still Needed to Complete the SBA Package")
    para(d, "Everything below is data only you or your advisors can supply. Drafts are pre-filled with "
            "everything derivable from the business plan and model; these are the remaining blanks.",
         color=GREY, size=9)
    h2(d, "A. Entity / company (for Company Profile)")
    for x in ["Principal office street address, city, zip, phone, fax, web address",
              "CPA name / firm / phone", "Attorney name / firm / phone",
              "Insurance agent name / firm / phone",
              "Confirm business checking account name and signers"]:
        para(d, "  - " + x)
    h2(d, "B. For each owner of 20% or more (Geoff Schackmann via Adeline & Lilah, LLC at 39.9%)")
    para(d, "Note: only owners of 20% or more need the personal items below. James Bullard holds 19.9% as a "
            "passive minority investor and is below the 20% threshold, so no personal guaranty, PFS, or Form 912 "
            "is required of him. The lender will still verify the source of his $250,000 injection via his bank / "
            "brokerage statements (item below), which is standard for any equity-injection source.", italic=True, size=9)
    for x in ["SSN, date of birth, place of birth, home address (and prior address with dates)",
              "Personal Financial Statement (SBA Form 413) - assets, liabilities, income",
              "Personal cash flow statement (template 7a)",
              "Personal History Form (SBA Form 912) - the legal/background questions",
              "Three years of personal tax returns (all schedules, W-2s, 1099s)",
              "Resume / Management Resume personal fields (SSN, DOB, addresses, dates, spouse, military, education)",
              "Credit report and driver license",
              "Two months of bank / brokerage statements showing the equity injection source"]:
        para(d, "  - " + x)
    h2(d, "C. Transaction documents")
    for x in ["Hickory CHOW purchase agreement (Asset/Membership Purchase Agreement)",
              "Executed Hickory seller note ($300,000, 36 months, 6%) and security documentation",
              "Letter of intent / contract for the Hickory purchase",
              "Office lease or LOI (term should match the loan term, with options)"]:
        para(d, "  - " + x)
    h2(d, "D. Decisions for you / John Hart")
    for x in ["Approve updating the Business Plan financials to the current model (see Reconciliation memo) - recommended",
              "The 'bank approval - a couple items to discuss' note on the checklist: what are John's items?",
              "Whether the lender wants the COVID questionnaire answered (we can draft brief responses)"]:
        para(d, "  - " + x)
    footer_note(d)
    save(d, "INFORMATION_NEEDED_FROM_YOU.docx")


# ================================================================ PERSONAL FORMS GUIDE
def personal_forms_guide():
    d = new_doc("Personal Forms Guide")
    h1(d, "How to Complete the Personal SBA Forms")
    para(d, "These three forms require personal financial and background information and cannot be drafted "
            "from the business records. Guidance below; the blank forms are in 00_source_forms.", color=GREY, size=9)
    h2(d, "SBA Form 413 - Personal Financial Statement (folder 05)")
    para(d, "Required from each owner of 20% or more (and spouse if jointly filing). Lists personal assets "
            "(cash, savings, retirement, real estate, autos, other) and liabilities (mortgages, loans, credit "
            "cards), plus a personal income/expense section. Attach supporting bank statements and a recent pay "
            "stub. The lender uses this to assess the personal guaranty and the equity-injection source.")
    h2(d, "Personal Cash Flow (template 7a, folder 07)")
    para(d, "Monthly personal income vs. personal expenses for each 20%+ owner. Demonstrates the owner can meet "
            "personal obligations independent of the business during ramp.")
    h2(d, "SBA Form 912 - Personal History Statement / Personal History Form (folder 06)")
    para(d, "Background questionnaire for each 20%+ owner: citizenship, residency, and the character questions "
            "(arrests, charges, convictions). Answer exactly and truthfully; any 'yes' requires a written "
            "explanation and may require additional clearance. This drives the CAIVRS/SAM screening on the checklist.")
    h2(d, "Tip")
    para(d, "Complete the Management Resume drafts first (folder 06) - they capture the work-history and "
            "education content; you only add personal identifiers. Then Form 912 and Form 413 reuse much of the "
            "same identity data.")
    footer_note(d)
    save(d, "05_personal_financial_statement/HOW_TO_COMPLETE_personal_forms.docx")


# ================================================================ EQUITY INJECTION / SBA STRUCTURE MEMO
def equity_injection_memo():
    d = new_doc("Equity Injection Memo")
    h1(d, "Equity Injection & Capital Structure - SOP 50 10 8 Compliance")
    para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)",
         color=GREY, size=9)
    para(d, "")
    para(d, "This memo documents how the transaction satisfies the equity-injection requirements of SOP 50 10 8 "
            "(effective June 1, 2025) for a complete change of ownership, and how the investor's contribution is "
            "structured so that it qualifies as equity without triggering a personal guaranty.")
    h2(d, "Capital structure")
    table(d, ["Source", "Amount", "% of project", "Treatment"],
          [["SBA 7(a) loan", "$500,000", "47.6%", "Senior debt"],
           ["Equity injection - cash (James Bullard, 19.9% member)", "$250,000", "23.8%", "Qualifying equity injection"],
           ["Seller note (Hickory CHOW, 36 mo, 6%)", "$300,000", "28.6%", "Acquisition debt (not on standby; not equity)"],
           ["Total project cost", "$1,050,000", "100%", ""]],
          widths=[3.2, 1.1, 1.1, 1.8])
    h2(d, "1. The 10% minimum injection is met more than twice over")
    para(d, "SOP 50 10 8 requires a minimum equity injection of at least 10% of total project cost for a complete "
            "change of ownership - here, $105,000. The transaction injects $250,000 of cash equity (23.8% of "
            "project cost), comfortably above the floor. The injection is non-borrowed cash, verified by the "
            "investor's bank / brokerage statements and the capital-contribution wire into the company's account, "
            "with funds seasoned in the contributor's account (lender to retain 30+ days of statements per the SOP "
            "verification standard).")
    h2(d, "2. The seller note does not need to - and does not - count as equity")
    para(d, "Under SOP 50 10 8, seller debt counts toward the injection only if it is on full standby (no "
            "principal or interest) for the life of the SBA loan, and even then for no more than 50% of the "
            "required injection. The $300,000 Hickory seller note is amortizing (36 months at 6%) and is therefore "
            "treated as acquisition debt serviced within combined debt service - not as part of the equity "
            "injection. Because the cash injection alone clears the 10% requirement by a wide margin, no standby "
            "of the seller note is required.")
    h2(d, "3. The investor's equity qualifies without a personal guaranty")
    para(d, "The $250,000 is contributed by James Bullard as a capital contribution in exchange for a direct, "
            "fully funded 19.9% membership interest. In a complete change of ownership, SOP 50 10 8 does not "
            "require a personal guaranty or Personal Financial Statement from an equity holder of less than 20%. "
            "Mr. Bullard's interest is held below that threshold, so his cash qualifies as equity injection while "
            "he remains a non-guarantor.")
    para(d, "To preserve eligibility, the structure deliberately avoids the practice SBA flags as disqualifying - "
            "an investor taking less than 20% to dodge a guaranty while using a side agreement to control the "
            "business. Mr. Bullard is a passive investor: he holds no management authority, no voting control, and "
            "no side agreement, option, or convertible instrument that would give him control or push his interest "
            "to 20% or more. Control of the business rests entirely with the Managing Member, Adeline & Lilah, LLC "
            "(Geoff Schackmann, 39.9%), who provides the unconditional personal guaranty as the sole 20%+ owner.")
    h2(d, "4. Guaranty and disclosure summary")
    table(d, ["Owner", "Interest", "20%+?", "Personal guaranty / PFS / Form 912"],
          [["Adeline & Lilah, LLC (Geoff Schackmann)", "39.9%", "Yes", "Required - provided"],
           ["James Bullard (passive investor)", "19.9%", "No", "Not required (source-of-funds verification only)"],
           ["Silas R. Shelton", "13.3%", "No", "Not required"],
           ["Dana L. Davenport", "13.3%", "No", "Not required"],
           ["Bradley G. Woodard", "13.3%", "No", "Not required"]],
          widths=[2.8, 1.0, 0.8, 2.6])
    para(d, "All equity owners are disclosed regardless of percentage, and all are U.S. citizens or lawful "
            "permanent residents (to be verified by the lender per SOP 50 10 8).", italic=True, size=9)
    footer_note(d)
    save(d, "02_use_of_funds/Equity_Injection_and_SBA_Structure_MEMO.docx")


if __name__ == "__main__":
    company_profile()
    use_of_funds()
    equity_injection_memo()
    debt_schedule()
    management_resumes()
    reconciliation_memo()
    information_needed()
    personal_forms_guide()
    print("Done.")
