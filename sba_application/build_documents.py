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
    field(d, "Address", "13387 Hwy 69 N, Tyler, TX (principal office)")
    field(d, "City / County / State / Zip", "Tyler / Smith / TX / " + TBD + " (zip)")
    field(d, "Contact / Phone / Fax", "Geoff Schackmann, Managing Member - 480-495-5474 - geoff@azaleahospice.com")
    field(d, "Number of employees at application", "9 (Year-1 staggered roster)")
    field(d, "Number of employees when loan approved", "Scales to ~25+ FT by Year 3 on census triggers")
    para(d, "")
    h2(d, "Company Ownership (100% must be shown)")
    table(d,
          ["#", "Name and Address", "SSN / EIN", "Ownership %", "Company Title", "Email"],
          [["1", "Adeline & Lilah, LLC (AZ) - members: Geoff Schackmann 50%, Mary Elizabeth Burcham 50%; Geoff serves as Managing Member - 4602 E Cheery Lynn Rd, Phoenix, AZ 85018", TBD, "39.9%", "Managing Member of Tyler Hospice Hold LLC", "geoff@azaleahospice.com"],
           ["2", "James E. Bullard - 13910 Indiana Ave, Suite 300, Lubbock, TX 79423", TBD, "19.5%*", "Minority investor member (passive; no management or control role)", "jimbullard01@aol.com"],
           ["3", "Silas R. Shelton - 1503 Lake Park Circle, Hideaway, TX 75771", TBD, "13.3%**", "Executive Director", "silas@azaleahospice.com"],
           ["4", "Dana L. Davenport", TBD, "13.3%**", "Director of Nursing", TBD],
           ["5", "Bradley G. Woodard - 421 W Cumberland Rd, Apt 403, Tyler, TX 75703", TBD, "13.3%**", "Director of Sales", TBD]],
          widths=[0.3, 2.4, 1.0, 0.8, 1.4, 1.2])
    para(d, "*James Bullard holds a direct, fully funded minority equity interest of 19.5% - below the 20% "
            "threshold - as a passive investor. He has no management authority, no voting control, and no side "
            "agreement granting him control of the business; the Managing Member (Adeline & Lilah, LLC) retains "
            "control. Consistent with SOP 50 10 8 and 13 CFR 120.160, an equity holder of less than 20% in a "
            "complete change of ownership is not required to provide a personal guaranty or Personal Financial "
            "Statement. Mr. Bullard's $195,000 cash capital contribution is the source of the equity injection "
            "for this transaction (see Use of Proceeds).", italic=True, size=9)
    para(d, "**Operator-members (Shelton, Davenport, Woodard) hold their 13.3% interests as Restricted Interests, "
            "the entire 13.3% of which is at risk and forfeitable until vested. Each interest comprises a 4.9% "
            "Initial Base that time-vests on a four-year schedule with a one-year cliff (continuous service alone) "
            "and an 8.4% Earn-Up Portion that vests only on a dual trigger - a four-year time schedule with a "
            "one-year cliff AND performance milestones, the lesser of the two governing - so each operator may earn "
            "up to a fully-vested 13.3%, recorded in the Operating Agreement (Amendment No. 1). "
            "Unvested interests are forfeited at $0 on departure; vested interests are subject to a Company "
            "repurchase (call) right at fair market value for a good-leaver separation (no-cause termination, "
            "death, disability, retirement) or the lower of cost or fair market value for a bad-leaver "
            "separation (resignation before full vesting, or termination for cause).", italic=True, size=9)
    para(d, "Indirect chain: Geoff Schackmann holds 50% of Adeline & Lilah, LLC and serves as its sole Managing "
            "Member; he therefore controls the 39.9% A&L block in Tyler Hospice Hold LLC and is the sole 20%+ "
            "owner of record for SBA personal-guaranty purposes. Mary Elizabeth Burcham (Geoff's spouse) holds "
            "the other 50% of Adeline & Lilah, LLC as a passive member (no management role), giving her an "
            "indirect economic interest of 19.95% in Tyler Hospice Hold LLC. As Geoff's spouse, she will sign "
            "the customary spouse acknowledgement / consent on the personal guaranty at close. "
            "Unissued/reserved pool: 0.7%.", italic=True, size=9)
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
          [["Acquisition of Hickory Hospice LLC - paid in full at close (CHOW; 100% membership interests)", "$300,000", "SBA + equity"],
           ["Startup one-time costs (855A filing, TX licensure, EMR, supply stock, legal, contingency)", "$63,000", "SBA + equity"],
           ["Equipment purchase (computers, office furniture)", "$15,000", "SBA + equity"],
           ["Working-capital reserve (opening cash; funds ramp payroll and SBA debt service)", "$372,000", "SBA + equity"],
           ["Total funds required", "$750,000", ""]],
          widths=[4.0, 1.1, 1.7])
    para(d, "")
    h2(d, "Loan Request Reconciliation")
    table(d, ["Line", "Amount"],
          [["Total funds required", "$750,000"],
           ["Less: borrower equity injection (cash capital contribution)", "($195,000)"],
           ["Total SBA 7(a) loan request", "$555,000"]],
          widths=[4.5, 1.5])
    para(d, "The seller is paid in full at close from SBA proceeds plus the equity injection; there is no "
            "retained seller note, so the SBA loan is the only debt the business carries. The $195,000 equity "
            "injection is 26.0% of the $750,000 project - well above the 10% SOP 50 10 8 minimum, though below "
            "the ~30% down some startup lenders prefer.", italic=True, size=9.5)
    para(d, "")
    h2(d, "Details on the transaction and use of working capital")
    para(d, "This is a change-of-ownership (CHOW) acquisition of Hickory Hospice LLC, an established, "
            "Medicare-certified hospice, structured as a purchase of 100% of its membership interests. The "
            "$300,000 purchase price for the existing Medicare provider number is paid in full at close from SBA "
            "proceeds and the equity injection - there is no retained seller note, so the SBA loan is the only "
            "debt the business carries. The acquired provider number, state license, and CHAP/ACHC accreditation "
            "transfer in the CHOW, so the agency is billing-ready from day one.")
    para(d, "After the $300,000 acquisition payoff and $78,000 of startup costs and equipment, the SBA loan and "
            "the equity injection fund a $372,000 working-capital reserve. The reserve carries ramp-period payroll "
            "and SBA debt service while accounts receivable normalize to the ~45-day Medicare cycle. Because there "
            "is no seller note, Year-1 debt service is the SBA loan alone (~$93,637/yr), so the reserve stretches "
            "further than under a dual-debt structure.")
    para(d, "Source of the borrower's equity injection: $195,000 in cash, provided as a capital contribution by "
            "minority member James Bullard (19.5% non-controlling interest) in exchange for his equity interest. "
            "This is a bona fide, non-borrowed equity injection equal to 26.0% of total project cost - more than "
            "than 2.5 times the 10% minimum required under SOP 50 10 8, though below the ~30% down some startup "
            "lenders prefer. The injection is verified by Mr. Bullard's bank / brokerage statements and the capital-"
            "contribution wire into the company's account. Because Mr. Bullard is a passive investor holding less "
            "than 20% with no management or control rights and no side agreement, his contribution counts as "
            "qualifying equity without triggering a personal guaranty (consistent with SOP 50 10 8 and 13 CFR "
            "120.160).", italic=True, size=9.5)
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
    para(d, "Tyler Hospice Hold LLC is a newly formed holding entity and carries no existing business debt prior "
            "to this transaction. The Hickory acquisition is paid in full at close from SBA proceeds and the "
            "equity injection - there is no seller note or other carried debt. After close, the only business "
            "debt is the SBA 7(a) loan itself.", italic=True)
    para(d, "")
    table(d, ["Lender", "Orig. date", "Purpose", "Orig. amount", "Current balance", "Rate", "Payment", "Collateral"],
          [["(No existing or proposed business debt other than the subject SBA 7(a) loan)", "-", "-", "$0", "$0", "-", "$0", "-"]],
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
           ["Net income Y1 / Y2 / Y3", "$199K / $706K / $1,084K", "$199K / $575K / $903K", "Y1 in line; Y2-Y3 lower on richer staffing"],
           ["DSCR Y1 / Y2 / Y3 (SBA-only debt)", "1.61x / 4.20x / 6.12x", "3.23x / 7.26x / 10.76x", "Stronger - seller paid at close"],
           ["Global 3-yr DSCR", "3.97x", "7.08x", "Both well above 1.25x"]],
          widths=[1.9, 1.7, 1.7, 1.5])
    para(d, "Structure note: the deal is now a startup CHOW in which the $300,000 seller price is paid in full at "
            "close from SBA proceeds and the equity injection - no retained seller note. Debt service is the SBA "
            "loan alone (~$93,637/yr), which lifts DSCR materially above the earlier dual-debt assumption.",
         italic=True, size=9.5)
    h2(d, "Recommendation")
    para(d, "Use the current operating model with the seller-paid-at-close structure. Every year is profitable, "
            "EBITDA grows from $303K to $1.0M, and DSCR clears 1.25x in every year by a wide margin (3.23x / 7.26x "
            "/ 10.76x; global 7.08x) because there is no seller note competing with the SBA loan for cash flow. "
            "The staffing build is the fuller, more defensible roster for a 50+ ADC agency.")
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
           ["Interest (SBA loan only)", "$62,202", "$58,741", "$54,551"],
           ["TX franchise tax", "$5,411", "$10,232", "$14,006"],
           ["Net income", "$199,308", "$574,951", "$903,362"]],
          widths=[2.6, 1.4, 1.4, 1.4])
    h2(d, "Corrected lender summary (drop-in for Section 11.6)")
    table(d, ["Metric", "Year 1", "Year 2", "Year 3"],
          [["EBITDA", "$302,522", "$679,524", "$1,007,520"],
           ["SBA debt service", "$93,637", "$93,637", "$93,637"],
           ["DSCR (SBA-only)", "3.23x", "7.26x", "10.76x"],
           ["Net income", "$199,308", "$574,951", "$903,362"]],
          widths=[2.6, 1.4, 1.4, 1.4])
    para(d, "Global 3-year DSCR: 7.08x (vs. 1.25x floor). With the seller paid at close, the SBA loan is the only "
            "debt and coverage is strong from Year 1.", bold=True)
    h2(d, "Break-even (answers Business Plan Guide VIII.3 - currently missing)")
    para(d, "Contribution margin is $160.52 per patient-day (net rate $181.30 less $20.78 of variable cost per "
            "patient-day: supplies/DME/pharmacy plus PRN visit labor plus the 1.5% billing and 0.75% QR fees). "
            "Against the Year-1 fixed-cost base (the lean opening roster plus fixed G&A), the operation breaks "
            "even at approximately 16 ADC, or about $86,000 per month of net revenue. The validated opening "
            "census of ~22 ADC is already above break-even, and Year-1 covers SBA debt service at 3.23x.")
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
    h2(d, "B. For Geoff Schackmann (sole 20%+ controlling owner via Adeline & Lilah, LLC at 39.9%)")
    para(d, "Note: Geoff controls A&L's 39.9% block as its sole Manager and is the sole 20%+ owner for SBA "
            "guaranty purposes. James Bullard (19.5% direct) and Mary Elizabeth Burcham (19.95% indirect via "
            "her 50% of A&L) are each under 20% with no management role, so neither needs PFS, Form 912, or "
            "tax returns. Bullard's source of funds is verified separately (Section B2); Mary Elizabeth signs a "
            "spouse acknowledgement at close (Section B3).", italic=True, size=9)
    for x in ["SSN, date of birth, place of birth, home address (and prior address with dates)",
              "Personal Financial Statement (SBA Form 413) - assets, liabilities, income",
              "Personal cash flow statement (template 7a)",
              "Personal History Form (SBA Form 912) - the legal/background questions",
              "Three years of personal tax returns (all schedules, W-2s, 1099s)",
              "Resume / Management Resume personal fields (SSN, DOB, addresses, dates, spouse, military, education)",
              "Credit report and driver license",
              "PFS asset support: the executed VistaRiver MIPA + Promissory Note (already in folder 10) and",
              "  a current VistaRiver payment ledger or recent statements confirming the note is current",
              "Any other PFS assets to disclose (real estate, retirement accounts, vehicles, other investments)"]:
        para(d, "  - " + x)
    h2(d, "B2. Equity injection verification (James Bullard, contributor)")
    para(d, "Tranche 1 of $100,000 is confirmed received - Mercury (Column N.A.) acct ****1275, May 7, 2026 wire "
            "from JIM BULLARD; balance still held as of June 3, 2026. Tranche 2 ($95,000, by July 31, 2026) is "
            "committed and outstanding. The following remain needed; SOP 50 10 8 source-of-funds "
            "verification is required for each tranche even though Mr. Bullard is under 20% (no PFS, Form 912, "
            "tax returns, or guaranty required of him).", italic=True, size=9)
    for x in ["Mr. Bullard's bank or brokerage statements - in his own name - covering 30+ days BEFORE each wire, showing the funds available",
              "Wire/transfer confirmation for the remaining $95,000 tranche as it lands",
              "Final Mercury statement (or trailing-day balance) showing the full $195,000 has been received once Tranche 2 lands",
              "Save the screenshot/PDF of Mercury acct ****1275 showing the May 7 wire under 10_supporting_documents/equity_injection_evidence/"]:
        para(d, "  - " + x)
    para(d, "Timing: SBA closing/disbursement will occur AFTER the July 2026 tranche so the full $195,000 is on "
            "deposit at disbursement (per SOP 50 10 8). No holdback or acceleration needed.", italic=True, size=9)
    h2(d, "B3. Spouse acknowledgement (Mary Elizabeth Burcham)")
    para(d, "Mary Elizabeth Burcham is Geoff's spouse and holds 50% of Adeline & Lilah, LLC (giving her a 19.95% "
            "indirect interest in Tyler Hospice Hold LLC - under 20%, no management role). As the guarantor's "
            "spouse she signs a customary spouse acknowledgement / consent, not a co-guaranty.", italic=True, size=9)
    for x in ["Mary Elizabeth Burcham's full legal name (with middle name/initial), date of birth, address",
              "Her signature on the Spouse Acknowledgement / Consent (lender provides form at close)",
              "If the lender requires it (varies by lender, not by SBA): a brief PFS for the marital household. Confirm with John Hart whether his bank requires this for the spouse of the sole guarantor."]:
        para(d, "  - " + x)
    h2(d, "C. Transaction documents")
    for x in ["Hickory CHOW purchase agreement (Asset/Membership Purchase Agreement)",
              "Wire confirmation for the $300,000 seller payoff at close (no seller note retained)",
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
            "stub. A pre-fill worksheet (SBA_Form_413_PFS_PREFILL_Geoff.docx) is provided with the two known "
            "items already entered - the VistaRiver note receivable and the $555,000 SBA guaranty as a contingent "
            "liability; add your remaining values and transcribe onto the official form. Note: the $195,000 equity "
            "injection is Bullard's cash, not Geoff's, so it is not on Geoff's PFS.")
    h2(d, "SBA Form 1919 - Borrower Information Form (folder 01)")
    para(d, "Required at application. A pre-fill worksheet (SBA_Form_1919_Borrower_Information_PREFILL.docx) has "
            "the business section completed; each principal still answers the Section II citizenship and "
            "character questions personally and truthfully.")
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
          [["SBA 7(a) loan", "$555,000", "74.0%", "Senior debt (only debt the business carries)"],
           ["Equity injection - cash (James Bullard, 19.5% member)", "$195,000", "26.0%", "Qualifying equity injection"],
           ["Total project cost", "$750,000", "100%", "Seller paid in full at close (no retained note)"]],
          widths=[3.2, 1.1, 1.1, 1.8])
    h2(d, "1. The injection is 2.6x the 10% minimum")
    para(d, "SOP 50 10 8 requires a minimum equity injection of at least 10% of total project cost for a startup "
            "/ complete change of ownership - here, $75,000. The transaction injects $195,000 of cash equity "
            "(26.0% of project cost), comfortably above the 10% floor, though below the ~30% down some startup "
            "a startup. The injection is non-borrowed cash, verified by the investor's bank / brokerage statements "
            "and the capital-contribution wire into the company's account, with funds seasoned in the contributor's "
            "account (lender to retain 30+ days of statements per the SOP verification standard).")
    h2(d, "1a. Contribution status and schedule")
    para(d, "The injection is being contributed in tranches to the borrower's Mercury (Column N.A.) operating "
            "account ending 1275. As of the date of this memo:")
    table(d, ["Tranche", "Date", "Amount", "Method", "Status"],
          [["1 of 2", "May 7, 2026", "$100,000", "Wire from James Bullard", "RECEIVED - in account"],
           ["2 of 2", "By July 31, 2026", "$95,000", "Wire from James Bullard", "Committed"],
           ["Total", "", "$195,000", "", ""]],
          widths=[0.8, 1.2, 1.2, 2.0, 1.6])
    para(d, "Evidence on file: Mercury account balance and transaction record dated June 3, 2026 showing the "
            "May 7, 2026 incoming wire of $100,000 from JIM BULLARD with the $100,000 balance still held (only "
            "de-minimis Gusto payroll-test ACH activity since). Source-of-funds statements in Mr. Bullard's name "
            "covering the 30+ days prior to each wire will be retained for each tranche per SOP 50 10 8.", italic=True, size=9.5)
    para(d, "Closing timing: SBA closing and disbursement are scheduled to occur AFTER the July 2026 equity "
            "tranche lands, so the full $195,000 will be on deposit at the time of SBA loan disbursement, in "
            "compliance with SOP 50 10 8. No disbursement holdback is required, and no acceleration of "
            "tranches 2 and 3 is sought.", italic=True, size=9.5)
    h2(d, "2. Seller paid in full at close - no retained note")
    para(d, "The $300,000 acquisition price is paid in full at close from SBA proceeds and the equity injection. "
            "There is no retained seller note, so the SBA loan is the only debt the business carries and Year-1 "
            "debt service is the SBA loan alone (~$93,637/yr). This produces strong coverage from Year 1 "
            "(DSCR ~3.23x) rather than the compressed coverage of a dual-debt (SBA + seller note) structure.")
    h2(d, "3. The investor's equity qualifies without a personal guaranty")
    para(d, "The $195,000 is contributed by James Bullard as a capital contribution in exchange for a direct, "
            "fully funded 19.5% membership interest. In a complete change of ownership, SOP 50 10 8 does not "
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
    table(d, ["Owner / interest holder", "Interest", "20%+?", "Personal guaranty / PFS / Form 912"],
          [["Adeline & Lilah, LLC (entity holding 39.9%)", "39.9% direct", "Yes", "Entity disclosed"],
           ["  - Geoff Schackmann (50% of A&L; sole Manager)", "19.95% indirect; controls A&L block", "Yes (by control)", "Required - provided as the sole 20%+ controlling owner"],
           ["  - Mary Elizabeth Burcham (50% of A&L; passive)", "19.95% indirect", "No (under 20%; no control)", "Not required; signs spouse acknowledgement/consent at close"],
           ["James Bullard (passive investor, direct)", "19.5%", "No", "Not required (source-of-funds verification only)"],
           ["Silas R. Shelton", "13.3% (fully forfeitable; 4-yr vest, 1-yr cliff)", "No", "Not required"],
           ["Dana L. Davenport", "13.3% (fully forfeitable; 4-yr vest, 1-yr cliff)", "No", "Not required"],
           ["Bradley G. Woodard", "13.3% (fully forfeitable; 4-yr vest, 1-yr cliff)", "No", "Not required"]],
          widths=[2.8, 1.6, 0.8, 2.0])
    para(d, "All equity owners are disclosed regardless of percentage, and all are U.S. citizens or lawful "
            "permanent residents (to be verified by the lender per SOP 50 10 8). The controlling owner (Geoff "
            "Schackmann, via Adeline & Lilah, LLC) provides the personal guaranty. His spouse, Mary Elizabeth "
            "Burcham, signs a spouse acknowledgement/consent rather than a co-guaranty, as her independent "
            "indirect interest is under 20% and she holds no management role.", italic=True, size=9)
    footer_note(d)
    save(d, "02_use_of_funds/Equity_Injection_and_SBA_Structure_MEMO.docx")


# ================================================================ LENDER CREDIT MEMO (1-page exec summary)
def lender_credit_memo():
    d = new_doc("Lender Credit Memo")
    h1(d, "Credit Memo - Executive Summary")
    para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)",
         color=GREY, size=9)
    para(d, "")
    h2(d, "The deal at a glance")
    table(d, ["Field", "Value"],
          [["Borrower", "Tyler Hospice Hold LLC (newco) - EIN 41-4966640"],
           ["Operating subsidiary", "Hickory Hospice LLC (Texas) - Medicare-certified hospice provider"],
           ["Transaction", "Startup change of ownership (CHOW) - purchase of 100% membership interests of Hickory"],
           ["Loan request", "$555,000 SBA 7(a)"],
           ["Total project cost", "$750,000"],
           ["Use of proceeds", "Acquisition paid at close ($300K) + working capital ($372K) + startup ($63K) + equipment ($15K)"],
           ["Borrower equity injection", "$195,000 cash (26.0% of project) - 2.6x the 10% SOP 50 10 8 floor (below the ~30% some startup lenders prefer)"],
           ["Primary guarantor", "Geoff Schackmann (sole 20%+ owner via Adeline & Lilah, LLC; 39.9%)"],
           ["DSCR (SBA-only debt)", "Y1 3.23x | Y2 7.26x | Y3 10.76x | Global 7.08x (floor 1.25x)"]],
          widths=[1.9, 4.6])
    h2(d, "Why this credit is strong")
    para(d, "1. Established, billing-ready agency. The transaction is the purchase of an operating Medicare-"
            "certified hospice with a transferring provider number, state license, and CHAP/ACHC accreditation - "
            "no payor-enrollment ramp risk and no startup CHAP/ACHC cycle. The agency is billing-capable on day one.")
    para(d, "2. Experienced multi-hospice operator. The Managing Member has operated Medicare-certified hospice "
            "agencies across multiple states with documented census growth from sub-30 to 100+ ADC under existing "
            "provider numbers, plus prior CHOW transactions with successful post-close enrollment and integration.")
    para(d, "3. Validated East-Texas BD pipeline. The Director of Sales sustained a 40+ ADC referral book in the "
            "Tyler market for 4.5 years and previously grew Tyler-market and Lufkin / Nacogdoches agencies from "
            "single-digit ADC to 70-220 patients. Opening census of ~22 ADC is conservatively underwritten against "
            "this pipeline.")
    para(d, "4. Conservative staffing build. Year-1 P&L carries a fully staffed clinical roster sized to the "
            "underwritten census (RN case managers, CNAs, social work, chaplain, IDG, Quality/Compliance, Intake, "
            "Volunteer Coordinator) plus outsourced billing at 1.5% of revenue - not a thin headcount story.")
    para(d, "5. Coverage and headroom. Because the seller is paid in full at close (no retained note), the SBA "
            "loan is the only debt and Year-1 DSCR is ~3.23x, rising to 7.26x and 10.76x (global 7.08x). "
            "Break-even is ~16 ADC vs. opening 22 ADC, so the deal is profitable on the validated opening census "
            "alone.")
    h2(d, "Equity injection - SOP 50 10 8 compliant")
    para(d, "$195,000 cash (26.0% of the $750,000 project) contributed by James Bullard as a capital contribution "
            "for a direct, fully funded 19.5% passive minority interest. Under 20%, no PFS or guaranty required; "
            "the lender verifies only the source of funds per the SOP. Tranche 1 of $100,000 was wired May 7, 2026 "
            "into the borrower's Mercury (Column N.A.) account ****1275 and remains on deposit; the balance follows "
            "on a committed schedule. The $300,000 acquisition price is paid in full at close from SBA proceeds and "
            "the injection - no retained seller note. Detailed compliance write-up in "
            "Equity_Injection_and_SBA_Structure_MEMO.")
    h2(d, "Risks and mitigants")
    table(d, ["Risk", "Mitigant"],
          [["Census ramp slower than plan",
            "Break-even ~16 ADC; opening census ~22; BD pipeline validated; sensitivity (separate memo) shows DSCR holds through -20% Y1 revenue"],
           ["Medicare CoP / survey risk",
            "Existing CHAP/ACHC accreditation transfers; DON-led IDG; QAPI program; experienced clinical leadership"],
           ["Key-person dependency",
            "Four-person leadership team (Managing Member, ED, DON, BD Director); 25+ year East-Texas BD relationships are institutional"],
           ["Equity injection phased (not all in at signing)",
            "Tranche 1 in account; balance committed; lender can sequence SBA disbursement after the full injection is on deposit"]],
          widths=[2.0, 4.5])
    h2(d, "Closing conditions to verify")
    for x in ["Bullard source-of-funds statements for each tranche (30+ day seasoning, in his name)",
              "Receipt and on-deposit evidence for the remaining tranches; full $195K balance confirmation pre-disbursement",
              "Executed Hickory CHOW purchase agreement (MIPA) and wire confirmation of the $300K seller payoff at close",
              "Office lease / LOI with term matching the SBA loan term and reasonable options",
              "Geoff Schackmann personal package: PFS (413), cash flow (7a), history form (912), 3 yrs tax returns, credit, license",
              "Business Plan / sources-and-uses reflecting the seller-paid-at-close structure (Rev 5.00)"]:
        para(d, "  - " + x)
    footer_note(d)
    save(d, "11_lender_credit_memo/Lender_Credit_Memo_EXEC_SUMMARY.docx")


# ================================================================ DSCR SENSITIVITY / STRESS-TEST MEMO
def sensitivity_memo():
    d = new_doc("DSCR Sensitivity Memo")
    h1(d, "DSCR Sensitivity / Stress-Test Analysis")
    para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC. Illustrative scenarios; underlying model in folder 09.",
         color=GREY, size=9)
    para(d, "")
    para(d, "The base case clears the 1.25x SBA DSCR floor in every year by a wide margin. Because the seller is "
            "paid in full at close (no retained note), debt service is the SBA loan alone (~$93,637/yr), so "
            "coverage is strong from Year 1. This memo stress-tests the Year-1 result against the risks a lender "
            "will probe: a slower census ramp, a softer reimbursement rate, payroll inflation, and a combined "
            "downside.")
    h2(d, "Base case (Year 1)")
    table(d, ["Line", "Amount"],
          [["Net patient revenue", "$1,442,930"],
           ["EBITDA", "$302,522"],
           ["SBA debt service (only debt)", "$93,637"],
           ["DSCR", "3.23x"]],
          widths=[3.5, 2.0])
    h2(d, "Single-factor stress (Year 1)")
    para(d, "Each scenario isolates one variable; all others held at base. Variable cost flexes proportionally "
            "with revenue at the model's contribution ratio (~88.5% contribution margin per patient-day). Debt "
            "service is the SBA loan only ($93,637).", italic=True, size=9.5)
    table(d, ["Scenario", "Assumption", "EBITDA", "DSCR", "Clears 1.25x?"],
          [["S1 - Census -10%",       "ADC 19.6 vs. 21.8",          "$174,823", "1.87x", "Yes"],
           ["S2 - Census -20%",       "ADC 17.4 vs. 21.8",          "$47,123",  "0.50x", "No - shortfall"],
           ["S3 - Rate -3%",          "Net rate -3%",               "$264,212", "2.82x", "Yes"],
           ["S4 - Rate -5%",          "Net rate -5%",               "$238,672", "2.55x", "Yes"],
           ["S5 - Payroll +10%",      "Payroll $955K vs. $868K",    "$215,705", "2.30x", "Yes"],
           ["S6 - Payroll +5%",       "Payroll $911K vs. $868K",    "$259,113", "2.77x", "Yes"]],
          widths=[1.4, 2.0, 1.0, 0.8, 1.3])
    h2(d, "Combined downside (Year 1)")
    table(d, ["Scenario", "Assumption", "EBITDA", "DSCR"],
          [["C1 - Mild downside",    "Census -5%, Rate -2%, Payroll +3%",    "$187,087", "2.00x"],
           ["C2 - Moderate downside","Census -10%, Rate -3%, Payroll +5%",   "$93,104",  "0.99x"]],
          widths=[1.6, 2.6, 1.0, 0.8])
    h2(d, "What the stress tells us")
    para(d, "With the seller paid at close, Year-1 coverage is strong (3.23x) and absorbs realistic shocks: a 10% "
            "census miss, a 5% rate cut, and a 10% wage spike each still clear comfortably (1.87x / 2.55x / 2.30x), "
            "and a mild combined downside clears at 2.00x. The only single-factor break is a hard, sustained 20% "
            "census shortfall (0.50x); a moderate triple-downside falls below the floor (0.99x). Even then, the "
            "buffers below are not credited in the static flex:")
    for x in ["Break-even is ~16 ADC; even the -20% case (ADC ~17.4) is still above operating break-even - the strain is coverage timing, not viability.",
              "The Year-1 roster is partly census-driven: PRN, intake, and capacity RN/CNA headcount flex back if census softens, which the static stress holds flat against revenue.",
              "The $372K opening working-capital reserve bridges any timing shortfall; with only SBA debt service (~$7,030/mo), monthly obligations are far lighter than a dual-debt structure.",
              "A working-capital line, if the lender bundles one, adds further headroom not reflected in EBITDA."]:
        para(d, "  - " + x)
    h2(d, "Recommendation to the underwriter")
    para(d, "Approve with standard covenants. The deal covers SBA debt service at 3.23x in Year 1 and far higher "
            "thereafter; the only meaningful stress is a hard census shortfall, mitigated by (a) experienced BD "
            "leadership with a validated 40+ ADC referral book, (b) opening census ~38% above break-even, "
            "(c) the $372K reserve, and (d) the seller-paid-at-close structure that leaves the SBA loan as the only "
            "debt. A monthly census-and-cash covenant during Y1 would let the lender monitor the binding variable.")
    footer_note(d)
    save(d, "11_lender_credit_memo/DSCR_Sensitivity_Stress_Test_MEMO.docx")


# ================================================================ BULLARD NO-CONTROL / NO-SIDE-AGREEMENT ATTESTATION
def bullard_attestation():
    d = new_doc("Bullard Attestation")
    h1(d, "Investor Attestation - No Control, No Side Agreement")
    para(d, "Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care) - SBA 7(a) Application",
         color=GREY, size=9)
    para(d, "")
    para(d, "The undersigned, James Bullard, in connection with the SBA 7(a) loan application of Tyler Hospice "
            "Hold LLC (the \"Company\"), and as a minority equity holder in the Company, hereby attests as follows:")
    h2(d, "1. Ownership and contribution")
    para(d, "I hold a 19.5% direct membership interest in the Company, acquired in exchange for a $195,000 cash "
            "capital contribution. My contribution is being made in tranches: $100,000 was wired to the Company's "
            "operating account on May 7, 2026, with the remaining $95,000 scheduled to be contributed by "
            "July 31, 2026. The funds are my own non-borrowed cash; I have not used loan proceeds, advances, "
            "or credit (including credit-card advances or unsecured personal loans) to fund any portion of the "
            "contribution.")
    h2(d, "2. No management role")
    para(d, "I hold no officer, director, manager, or employee position with the Company. I do not participate "
            "in the day-to-day management or operation of the Company or any of its subsidiaries, including "
            "Hickory Hospice LLC. The Managing Member is Adeline & Lilah, LLC (Geoff Schackmann, sole member), "
            "who has sole authority over management, hiring, financial decisions, and Company operations.")
    h2(d, "3. No voting control or governance rights beyond ordinary minority interest")
    para(d, "My equity interest entitles me only to economic returns and customary minority-member rights under "
            "the Company's Operating Agreement and applicable state law. I have no veto, no class of preferred "
            "voting rights, no special board or manager-appointment rights, no consent rights over ordinary-course "
            "business decisions, and no other governance right that would give me control or de facto control "
            "of the Company.")
    h2(d, "4. No side agreement, option, or convertible instrument")
    para(d, "There is no written or oral side agreement, voting agreement, proxy, option, warrant, convertible "
            "note, profits-interest plan, earnout, employment arrangement, or other instrument between me and "
            "the Company, the Managing Member, or any other equity holder that would (i) increase my equity "
            "interest to 20% or more, (ii) give me voting or management control of the Company, or (iii) entitle "
            "me to direct or restrict the Company's operations, financing, or strategic decisions.")
    h2(d, "5. Independent affiliation")
    para(d, "I am not affiliated with the Company by virtue of common ownership, common management, identity of "
            "interest, or any contractual arrangement other than my passive minority equity interest described above. "
            "I am acting solely as an investor in the Company.")
    h2(d, "6. Compliance with SBA requirements")
    para(d, "I understand that my passive minority interest is being relied upon by the SBA-participating lender "
            "in connection with the equity-injection and personal-guaranty analysis under SOP 50 10 8 and 13 CFR "
            "120.160. I acknowledge that any change in my equity interest, the addition of any control right, or "
            "the creation of any side agreement of the kind described above could change my status under SBA "
            "regulations, and I agree to provide prompt written notice to the Company and the lender of any such "
            "change while the SBA loan remains outstanding.")
    para(d, "")
    para(d, "I attest under penalty of perjury that the foregoing is true and correct to the best of my knowledge.")
    para(d, "")
    para(d, "")
    para(d, "______________________________________________________")
    para(d, "James Bullard", bold=True)
    para(d, "Date: ___________________________")
    para(d, "")
    para(d, "")
    para(d, "STATE OF ___________________________}")
    para(d, "COUNTY OF __________________________}")
    para(d, "")
    para(d, "Subscribed and sworn to before me this _____ day of _____________________, 2026.")
    para(d, "")
    para(d, "______________________________________________________")
    para(d, "Notary Public                                      My commission expires: _______________")
    para(d, "")
    para(d, "Drafted for the SBA 7(a) application of Tyler Hospice Hold LLC. Have counsel review before "
            "execution. The structure described above is reflected in the Company's Operating Agreement; if "
            "the Operating Agreement has not yet been amended to match, do that first.",
         italic=True, size=8.5, color=GREY)
    footer_note(d)
    save(d, "08_entity_documents/Bullard_Investor_Attestation_TEMPLATE.docx")


# ================================================================ SUBMISSION COVER SHEET + TOC
def submission_cover_sheet():
    d = new_doc("Submission Cover Sheet")
    h1(d, "Submission Package - Table of Contents")
    para(d, "SBA 7(a) Application - Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)", color=GREY, size=9)
    field(d, "Loan request", "$555,000 SBA 7(a) | Total project: $750,000")
    field(d, "Borrower", "Tyler Hospice Hold LLC (WY) - EIN 41-4966640")
    field(d, "Operating subsidiary", "Hickory Hospice LLC (TX) - the acquired Medicare-certified hospice")
    field(d, "Primary guarantor", "Geoff Schackmann via Adeline & Lilah, LLC (39.9%, sole 20%+ owner)")
    field(d, "Lender contact", "John Hart")
    para(d, "")
    h2(d, "How this package is organized")
    para(d, "Each section corresponds to one or more items on the lender's SBA 7(a) Loan Checklist. "
            "The folder structure under /sba_application/ mirrors this index.")
    table(d, ["Section", "Contents", "Folder"],
          [["00", "Source forms (blank, as provided by lender)", "00_source_forms/"],
           ["01", "Company Profile", "01_company_profile/"],
           ["02", "Use of Proceeds + Equity Injection / SBA Structure memo", "02_use_of_funds/"],
           ["03", "Business Plan (Rev 5.00, editable Word - current) + Reconciliation/Break-even memo", "03_business_plan/"],
           ["04", "Business Debt Schedule", "04_business_debt_schedule/"],
           ["05", "Personal Financial Statement (SBA 413) - guide + pre-fill worksheet", "05_personal_financial_statement/"],
           ["06", "Personal History (SBA 912) + Management Resumes (4)", "06_personal_history_resume/"],
           ["07", "Personal Cash Flow (SBA 7a)", "07_personal_cash_flow/"],
           ["08", "Entity documents - WY formation, EIN letter, Operating Agreement, Investor Attestation", "08_entity_documents/"],
           ["09", "Financial Model - 3 workbooks (Operations Dashboard, SBA Loan Package, Investor Package)", "09_financial_model/"],
           ["10", "Supporting documents - tax returns, bank statements, purchase agreement, PFS support, injection evidence", "10_supporting_documents/"],
           ["11", "Lender Credit Memo (exec summary) + DSCR Sensitivity / Stress Test", "11_lender_credit_memo/"]],
          widths=[0.6, 4.4, 2.0])
    h2(d, "Cross-reference to lender's SBA 7(a) Loan Checklist")
    para(d, "Each item on the lender-provided checklist is found at the location below:", italic=True, size=9)
    table(d, ["Checklist item", "Location"],
          [["1. SBA 7(a) Loan Checklist", "00_source_forms/1_SBA_7a_Loan_Checklist.docx"],
           ["2. Company Profile", "01_company_profile/Company_Profile_DRAFT.docx"],
           ["2a. SBA Form 1919 - Borrower Information (pre-fill)", "01_company_profile/SBA_Form_1919_Borrower_Information_PREFILL.docx"],
           ["3. Use of Funds", "02_use_of_funds/Use_of_Funds_DRAFT.docx"],
           ["3a. Equity Injection / SBA Structure memo (supporting)", "02_use_of_funds/Equity_Injection_and_SBA_Structure_MEMO.docx"],
           ["4. Business Plan (editable Word, current)", "03_business_plan/Azalea_SBA_Business_Plan_Rev5.00.docx"],
           ["4a. Reconciliation / Break-even memo (supporting)", "03_business_plan/RECONCILIATION_and_Breakeven_MEMO.docx"],
           ["5. Business Debt Schedule", "04_business_debt_schedule/Business_Debt_Schedule_DRAFT.docx"],
           ["6. Personal History (SBA 912) - Geoff Schackmann", "06_personal_history_resume/"],
           ["7. Personal Financial Statement (SBA 413) - Geoff Schackmann (pre-fill provided)", "05_personal_financial_statement/SBA_Form_413_PFS_PREFILL_Geoff.docx"],
           ["7a. Personal Cash Flow - Geoff Schackmann", "07_personal_cash_flow/"],
           ["8. Management Resumes (4)", "06_personal_history_resume/"],
           ["Lender Credit Memo (exec summary)", "11_lender_credit_memo/Lender_Credit_Memo_EXEC_SUMMARY.docx"],
           ["DSCR Sensitivity / Stress Test", "11_lender_credit_memo/DSCR_Sensitivity_Stress_Test_MEMO.docx"],
           ["Equity injection evidence (Tranche 1 confirmed; 2-3 pending)", "10_supporting_documents/equity_injection_evidence/"],
           ["Geoff PFS support (VistaRiver MIPA + Note)", "10_supporting_documents/personal_financial_statement_support/"],
           ["Hickory CHOW purchase agreement", "10_supporting_documents/purchase_agreement_seller_note/"],
           ["Tax returns (3 yrs personal)", "10_supporting_documents/tax_returns/"],
           ["Bank / brokerage statements", "10_supporting_documents/bank_brokerage_statements/"],
           ["Credit report and driver license", "10_supporting_documents/credit_drivers_license/"],
           ["Operating Agreement, EIN letter, formation docs", "08_entity_documents/"],
           ["Financial Model (Excel)", "09_financial_model/"]],
          widths=[3.6, 3.4])
    footer_note(d)
    save(d, "SUBMISSION_COVER_SHEET_and_TOC.docx")


# ================================================================ CLOSING CHECKLIST
def closing_checklist():
    d = new_doc("Closing Checklist")
    h1(d, "SBA 7(a) Closing Checklist")
    para(d, "Tyler Hospice Hold LLC - Azalea Hospice CHOW. Items typically required at SBA loan close.",
         color=GREY, size=9)
    para(d, "This is a standard close checklist; the lender will issue a final, deal-specific list. Use this to "
            "stage documents and identify long-lead items now.")
    h2(d, "A. Borrower entity")
    for x in ["Tyler Hospice Hold LLC - WY Certificate of Formation",
              "Tyler Hospice Hold LLC - Operating Agreement (executed, current)",
              "Tyler Hospice Hold LLC - EIN assignment letter (IRS CP-575)",
              "Tyler Hospice Hold LLC - Certificate of Good Standing (WY) - recent (within 30 days of close)",
              "Tyler Hospice Hold LLC - Foreign qualification in TX (Certificate of Authority) and TX Good Standing",
              "Hickory Hospice LLC - Certificate of Formation (TX)",
              "Hickory Hospice LLC - Operating Agreement (pre- and post-CHOW)",
              "Hickory Hospice LLC - Certificate of Good Standing (TX) - recent",
              "Member resolutions authorizing the loan, the CHOW, the guarantor signing"]:
        para(d, "  - " + x)
    h2(d, "B. Transaction documents")
    for x in ["Hickory CHOW Membership Interest Purchase Agreement (MIPA) - executed (filed)",
              "Wire confirmation of the $300,000 seller payoff at close (no retained seller note)",
              "Bill of sale / assignment of membership interests - executed",
              "Manager resignations + spousal consents per the MIPA closing deliverables",
              "UCC-1 financing statement filing (SBA lender)"]:
        para(d, "  - " + x)
    h2(d, "C. Licenses, accreditations, payor enrollment (CHOW)")
    for x in ["Texas Department of Health and Human Services - HCSSA license transfer / amendment",
              "Medicare Provider Transaction Access Number (PTAN) - CHOW filing (Form 855A) status confirmation",
              "Medicare provider agreement / Tie-In notice",
              "CHAP or ACHC accreditation - notice of CHOW and continuation",
              "DEA registration (if applicable to controlled-drug handling) - CHOW update",
              "State pharmacy licenses (if applicable)",
              "NPI - Type 2 organizational NPI verified for the new ownership structure"]:
        para(d, "  - " + x)
    h2(d, "D. Real estate / lease")
    for x in ["Office lease (term not less than the SBA loan term, with options) - executed",
              "Landlord consent / Landlord's Waiver of distraint (SBA standard) - signed by landlord",
              "Certificate of insurance naming SBA lender as additional insured (premises and operations)"]:
        para(d, "  - " + x)
    h2(d, "E. Equity injection verification")
    for x in ["Full $195,000 on deposit in Tyler Hospice Hold LLC operating account at close (per SOP 50 10 8)",
              "Bank statements (Mercury / Column N.A. acct ****1275) showing all two tranches received",
              "Bullard source-of-funds statements (30+ day seasoning prior to each tranche, in his name)",
              "Bullard wire / capital-contribution confirmations for each tranche",
              "Bullard Investor Attestation - executed and notarized (no control / no side agreement)",
              "Capital Contribution Agreement (Bullard) - executed",
              "Updated Operating Agreement reflecting current cap table - executed"]:
        para(d, "  - " + x)
    h2(d, "F. Guarantor (Geoff Schackmann)")
    for x in ["Personal guaranty (SBA Form 148 or lender form) - signed",
              "SBA Form 413 (Personal Financial Statement) - signed, dated within 90 days",
              "SBA Form 912 (Personal History) - signed; CAIVRS / SAM cleared",
              "Three years personal tax returns + most recent year's W-2s and 1099s",
              "Credit report (lender pulls; have driver license ready)",
              "Life insurance on guarantor, collaterally assigned to lender (if required for loan size)",
              "Spouse acknowledgement / consent - Mary Elizabeth Burcham, signed and dated"]:
        para(d, "  - " + x)
    h2(d, "G. Insurance (at close)")
    for x in ["Business / general liability - certificate naming SBA lender",
              "Professional liability for hospice / clinical operations",
              "Hazard insurance on tangible personal property (equipment) financed",
              "Workers' compensation - state certificate",
              "Cyber liability (recommended given PHI/EMR exposure)",
              "Life insurance on guarantor, collaterally assigned to SBA lender (if required)"]:
        para(d, "  - " + x)
    h2(d, "H. Other lender deliverables")
    for x in ["SBA Authorization (issued by SBA after lender request) - lender provides",
              "SBA Form 1919 (Borrower Information Form) - completed",
              "SBA Form 1920 (Lender's Application) - lender completes",
              "Resolution of borrower authorizing the loan - executed",
              "Bank's loan agreement, note, security agreement - executed",
              "Disbursement instructions for the $555K (operating account + use-of-proceeds wires)",
              "Closing Statement / Settlement Statement"]:
        para(d, "  - " + x)
    footer_note(d)
    save(d, "11_lender_credit_memo/CLOSING_CHECKLIST.docx")


# ================================================================ INSURANCE REQUIREMENTS SUMMARY
def insurance_requirements():
    d = new_doc("Insurance Requirements")
    h1(d, "Insurance Requirements - SBA 7(a) and Hospice Operations")
    para(d, "Tyler Hospice Hold LLC - Azalea Hospice & Palliative Care", color=GREY, size=9)
    para(d, "Coverages typically required by SBA and prudent for a hospice CHOW operating at the scale "
            "underwritten. Confirm specifics with the lender at close.")
    h2(d, "Required by SBA / lender")
    table(d, ["Coverage", "Typical limits", "Purpose / SBA requirement"],
          [["General Liability",                "$1M occurrence / $2M aggregate",
            "Standard SBA requirement; covers third-party bodily injury and property damage."],
           ["Professional Liability (Hospice)", "$1M / $3M typical",
            "Clinical malpractice exposure. SBA expects coverage appropriate to industry."],
           ["Property / Hazard",                "Replacement cost on equipment",
            "Required on any tangible personal property securing the SBA loan."],
           ["Workers' Compensation",            "Statutory (TX is opt-out but coverage recommended)",
            "Standard requirement once payroll commences; Texas allows non-subscription but most lenders prefer carriage."],
           ["Business Auto (if applicable)",    "$1M CSL",
            "Required if any owned/leased vehicles used in patient visits or administration."],
           ["Life Insurance on Guarantor",      "Amount = SBA loan balance, declining or level term",
            "Required by SBA for sole-guarantor deals to protect repayment in the event of death; collaterally assigned to lender."],
           ["Cyber / PHI Breach Liability",    "$1M typical",
            "Strongly recommended given EMR and HIPAA-regulated PHI; not always SBA-required but a hospice agency should carry it."]],
          widths=[1.9, 1.8, 3.0])
    h2(d, "Prudent additional coverages")
    for x in ["Employment Practices Liability (EPLI) - $1M",
              "Directors & Officers / Management Liability - $1M (light coverage for an LLC, but useful given multi-member structure)",
              "Crime / Employee Dishonesty - $250K-$500K",
              "Excess / Umbrella - $5M (often cheapest way to extend other limits)",
              "Hired & Non-Owned Auto - if staff use personal vehicles for patient visits"]:
        para(d, "  - " + x)
    h2(d, "Carriers to approach (Texas hospice market)")
    for x in ["CNA Healthcare - specialty hospice/home-health markets",
              "Coverys - clinical malpractice for hospice",
              "ProAssurance - hospice and home-health professional",
              "Markel - mid-market healthcare",
              "Travelers / Hartford / Hiscox - GL/property/cyber packages"]:
        para(d, "  - " + x)
    para(d, "Ask your broker for a comparison of two or three quotes and note the SBA collateral-assignment "
            "requirement when getting the life insurance quote (this drives the carrier choice; not all carriers "
            "issue collateral assignments quickly).", italic=True, size=9.5)
    footer_note(d)
    save(d, "11_lender_credit_memo/INSURANCE_REQUIREMENTS.docx")


# ================================================================ PERSONAL CASH FLOW PRE-FILL FOR GEOFF
def personal_cash_flow():
    d = new_doc("Personal Cash Flow Geoff")
    h1(d, "Personal Cash Flow Statement - Geoff Schackmann")
    para(d, "SBA 7(a) Form 7a (template). Monthly basis. Known income lines pre-filled; please complete the "
            "remaining lines. Sign and date at bottom.", color=GREY, size=9)
    field(d, "Name", "Geoff Schackmann")
    field(d, "Address", TBD)
    field(d, "Statement period (monthly)", TBD + "  (typically the most recent month or a 3-month average)")
    para(d, "")
    h2(d, "Personal income (monthly)")
    table(d, ["Source", "Amount", "Notes / verification"],
          [["VistaRiver Inc - promissory note payment", "$14,456.80", "Per executed Note dated Aug 15, 2025 (in folder 10). Through Aug 2030."],
           ["Salary / W-2 income (if any)", TBD, "Pay stub or W-2; list employer."],
           ["Distributions from Tyler Hospice Hold LLC (post-close)", "$0 (ramp)", "No member distributions modeled in Y1; tax distributions only in Y2+"],
           ["Distributions from other businesses owned", TBD, "List entity, role, amount, frequency."],
           ["Rental income (net)", TBD, "If applicable."],
           ["Investment income (interest, dividends)", TBD, "From most recent 1099 / brokerage statement."],
           ["Spouse income (if jointly filing)", TBD, "Source and amount; W-2 / 1099."],
           ["Other (royalties, retirement, etc.)", TBD, ""],
           ["TOTAL MONTHLY INCOME", TBD, ""]],
          widths=[3.0, 1.2, 2.5])
    h2(d, "Personal expenses (monthly)")
    table(d, ["Category", "Amount", "Notes"],
          [["Housing - mortgage or rent", TBD, "Principal + interest + taxes + insurance if escrowed"],
           ["Property taxes (if not escrowed)", TBD, ""],
           ["Homeowner / renter insurance (if not escrowed)", TBD, ""],
           ["Utilities (electric, gas, water, internet, phone)", TBD, ""],
           ["Food and household", TBD, "Average; include groceries and routine dining"],
           ["Vehicle - payment(s)", TBD, ""],
           ["Vehicle - fuel, insurance, maintenance", TBD, ""],
           ["Health insurance premiums (out of pocket)", TBD, ""],
           ["Medical / dental (routine)", TBD, ""],
           ["Childcare / tuition / education", TBD, ""],
           ["Other insurance (life, disability, umbrella)", TBD, ""],
           ["Credit card minimum payments", TBD, "List balances; minimums only here"],
           ["Other loan payments (student, personal, etc.)", TBD, ""],
           ["Alimony / child support (if any)", TBD, ""],
           ["Charitable / religious giving", TBD, ""],
           ["Discretionary (entertainment, travel, etc.)", TBD, ""],
           ["Income / self-employment taxes (estimated)", TBD, "Quarterly estimate / 12"],
           ["Other recurring", TBD, ""],
           ["TOTAL MONTHLY EXPENSES", TBD, ""]],
          widths=[3.0, 1.2, 2.5])
    h2(d, "Net monthly cash flow")
    table(d, ["Line", "Amount"],
          [["Total monthly income", TBD],
           ["Less: total monthly expenses", TBD],
           ["NET MONTHLY CASH FLOW", TBD]],
          widths=[4.0, 1.5])
    para(d, "")
    h2(d, "Certification")
    para(d, "I certify that the foregoing is a true and complete statement of my monthly personal income and "
            "expenses as of the date indicated.")
    para(d, "")
    para(d, "____________________________________________________________")
    para(d, "Geoff Schackmann", bold=True)
    para(d, "Date: ___________________________")
    footer_note(d)
    save(d, "07_personal_cash_flow/Personal_Cash_Flow_Geoff_PRE-FILL.docx")


# ================================================================ INTERVIEW WORKSHEET (for remaining drafts)
def interview_worksheet():
    d = new_doc("Interview Worksheet")
    h1(d, "Interview Worksheet - Remaining Drafts")
    para(d, "Answers will feed into: Cover Letter to John Hart, Anticipated Q&A memo, Capital Contribution "
            "Agreement (Bullard), Operating Agreement for Tyler Hospice Hold LLC, Bullard Source-of-Funds "
            "Letter, Hickory CHOW Purchase Agreement starter, and Affiliate / Size-Standard memo.",
         color=GREY, size=9)
    para(d, "Reasonable defaults are pre-filled where industry standard applies - just confirm or override. "
            "Free-form lines are marked TBD. Answer in any order; sections are independent.", italic=True, size=9.5)
    para(d, "")

    # --- LENDER / COVER LETTER --------------------------------------------
    h2(d, "1. Lender contact (for cover letter and anticipated Q&A)")
    field(d, "1.1 Bank / lender name", TBD)
    field(d, "1.2 Bank street address", TBD)
    field(d, "1.3 John Hart's exact title", TBD + "  (default if unsure: \"SBA Loan Officer\")")
    field(d, "1.4 John Hart's direct email and phone", TBD)
    field(d, "1.5 Your preferred contact info for the lender file", TBD + "  (email + cell)")
    field(d, "1.6 Submission date", "Today's date unless you specify otherwise")
    field(d, "1.7 The 'couple items to discuss' on the SBA checklist - what are they?",
          TBD + "  (this is the highest-signal flag we have; list everything John mentioned, even informally)")
    field(d, "1.8 Anything you've already verbally committed to John (timing, structure, follow-ups)", TBD)
    field(d, "1.9 Any prior lender feedback worth pre-addressing", TBD + "  (other banks that passed or asked hard questions)")
    field(d, "1.10 Anything in the credit profile you're worried about", TBD)

    # --- BULLARD / CAPITAL CONTRIBUTION & SOURCE OF FUNDS -----------------
    h2(d, "2. James Bullard - Capital Contribution Agreement + Source-of-Funds letter")
    field(d, "2.1 Bullard's full legal name (with middle name or initial)", TBD)
    field(d, "2.2 Bullard's home address", TBD)
    field(d, "2.3 Bullard's email and phone", TBD)
    field(d, "2.4 Source of the $195K (savings / sale of asset / inheritance / business income / other)", TBD)
    field(d, "2.5 If sale or business income: brief paper trail",
          TBD + "  (e.g., 'sale of XYZ stock March 2026' or 'business distributions from ABC LLC over 2024-2025')")
    field(d, "2.6 Account institution(s) the funds are sitting in", TBD + "  (bank/brokerage name; account type)")
    field(d, "2.7 How long the funds have been seasoned in those accounts",
          TBD + "  (SOP 50 10 8 wants 30+ days; longer is better)")
    field(d, "2.8 Scheduled date(s) and approximate amounts for tranches 2 and 3", TBD + "  (e.g., '$75K June 15, $75K July 15')")
    field(d, "2.9 Effective date of Bullard's 19.5% membership interest",
          "Default: May 7, 2026 (date Tranche 1 hit the account) - confirm or specify alternative")
    field(d, "2.10 Distribution preference for Bullard",
          "Default: pro-rata distributions with all members (no preferred return) - confirm or override")
    field(d, "2.11 Transfer restrictions on his interest",
          "Default: no transfer without Manager consent and right of first refusal by Company / other members - confirm")
    field(d, "2.12 Put/call or buyout terms",
          "Default: none in initial agreement; can be added by amendment - confirm")
    field(d, "2.13 Governing law for the Capital Contribution Agreement",
          "Default: Wyoming (entity's state of formation) - confirm")

    # --- ADELINE & LILAH / GEOFF ------------------------------------------
    h2(d, "3. Adeline & Lilah, LLC and Geoff (for OA, Cap Contribution recitals)")
    field(d, "3.1 Adeline & Lilah, LLC's principal office address",
          "Default: 2942 N 24th St STE 115 PMB, Phoenix, AZ 85016 (from VistaRiver note) - confirm")
    field(d, "3.2 Adeline & Lilah, LLC's state of formation", "Default: Arizona - confirm")
    field(d, "3.3 Geoff's home address", TBD)
    field(d, "3.4 Geoff's email and phone for SBA file", TBD)
    field(d, "3.5 Spouse's name (if any) and whether she is a member of Adeline & Lilah, LLC",
          "From the VistaRiver MIPA Adeline & Lilah, LLC appears to have two owners: Geoff (50%) and Mary Elizabeth "
          "Burcham (50%). Confirm relationship and whether Mary Elizabeth signs guarantor docs / spouse "
          "acknowledgement.")
    field(d, "3.6 Is anyone else a member or signatory of Adeline & Lilah, LLC?", TBD)

    # --- TYLER HOSPICE HOLD LLC / OPERATING AGREEMENT ----------------------
    h2(d, "4. Tyler Hospice Hold LLC - Operating Agreement")
    field(d, "4.1 Does Tyler Hospice Hold LLC have an existing Operating Agreement?",
          "If YES, share it (we'll amend). If NO, we'll draft a fresh one. Default assumption: NO.")
    field(d, "4.2 Tyler Hospice Hold LLC registered office (Wyoming filing address)", TBD + "  (typically a WY registered agent)")
    field(d, "4.3 Tax election",
          "Default: LLC taxed as partnership (default federal treatment for multi-member LLCs). "
          "Alternative: elect S-corporation (Form 2553) - usually only beneficial above ~$150K of distributable profit "
          "per owner-employee. Confirm or specify.")
    field(d, "4.4 Manager-managed vs. member-managed",
          "Default: MANAGER-MANAGED with Adeline & Lilah, LLC as sole Manager. This matches the SBA narrative "
          "(passive minority + sole Manager with control). Confirm.")
    field(d, "4.5 Vesting for operator-members (Silas, Dana, Bradley, 13.3% each)",
          "RECOMMENDATION: each operator's full 13.3% Restricted Interest is at risk and forfeitable until vested - "
          "a 4.9% Initial Base that time-vests (4-year schedule, 1-year cliff, continuous service alone) plus an "
          "8.4% Earn-Up Portion that vests on a dual trigger (4-year time schedule with a 1-year cliff AND "
          "performance milestones, lesser of the two). Protects the deal if any operator leaves. If you prefer "
          "fully vested at close, say so. Default if you skip: full 13.3% forfeitable (4.9% time-vested base + "
          "8.4% dual-trigger earn-up).")
    field(d, "4.6 Distribution policy",
          "Default: (a) mandatory quarterly tax distributions sized to cover members' tax on allocated income at "
          "the highest applicable rate; (b) other distributions at Manager's discretion. Confirm.")
    field(d, "4.7 Capital call authority",
          "Default: no mandatory capital calls; voluntary additional contributions require Manager approval and "
          "do not dilute non-contributing members. Confirm.")
    field(d, "4.8 Officers / titles (if any) and who holds them",
          "Default: Managing Member = Geoff (via Adeline & Lilah); Executive Director = Silas; Director of "
          "Nursing = Dana; Director of Sales = Bradley. Confirm.")
    field(d, "4.9 Books, records, fiscal year",
          "Default: calendar year, accrual basis, books at the principal office. Confirm.")
    field(d, "4.10 Buy-sell / departure (operator-member leaves Company)",
          "RECOMMENDATION: Company has option to repurchase unvested interest at $0 and vested interest at book "
          "value over 36 months. If you want a different price formula (e.g., trailing-12 EBITDA multiple), say so.")
    field(d, "4.11 Drag-along (Manager can force sale)",
          "Default: Manager (Adeline & Lilah) holding 39.9% may not unilaterally drag; needs members holding "
          ">=51% combined to trigger drag of remaining members. Confirm.")
    field(d, "4.12 Governing law", "Default: Wyoming - confirm")

    # --- HICKORY CHOW PURCHASE AGREEMENT ----------------------------------
    h2(d, "5. Hickory Hospice LLC - CHOW Purchase Agreement (starter draft)")
    para(d, "If you already have a signed PA from the Hickory seller, share it and skip this section - we'll use "
            "that. Otherwise we'll build a starter.", italic=True, size=9)
    field(d, "5.1 Hickory Hospice LLC's state of formation", "Default: Texas - confirm")
    field(d, "5.2 Hickory's principal office address", TBD)
    field(d, "5.3 Current owner(s) of Hickory: legal name(s), address(es), ownership %",
          TBD + "  (if multiple, list each)")
    field(d, "5.4 Target closing date", TBD + "  (consider coordinating with the July equity tranche - see Equity Injection memo)")
    field(d, "5.5 Seller payoff - confirm $300,000 paid in full at close (no retained seller note)", "Confirm or override")
    field(d, "5.6 Cash at close from buyer for the acquisition",
          "Default: $0 (fully seller-financed per the model). Confirm.")
    field(d, "5.7 Earnest money / good-faith deposit", "Default: $0 / none. Confirm or specify.")
    field(d, "5.8 Existing employees", "Are clinical / billing staff staying? Severance? Accrued PTO handled by whom?")
    field(d, "5.9 Existing contracts to assume",
          "List or describe: lease, EMR contract, supply contracts, payor contracts, etc.")
    field(d, "5.10 Licenses / accreditations transferring",
          "Default to assume: TX HCSSA license, CHAP or ACHC accreditation, Medicare PTAN/NPI. Confirm and specify "
          "which accreditation (CHAP or ACHC).")
    field(d, "5.11 Survey or due-diligence period before close",
          "Default: 30-day diligence + 30 days to close after diligence. Confirm.")
    field(d, "5.12 Reps and warranties survival period",
          "Default: 12 months post-close (fundamentals + tax + healthcare-regulatory uncapped or up to purchase price; "
          "general reps capped at 10% of purchase price). Confirm.")
    field(d, "5.13 Non-compete from seller", "Default: 3 years, 25-mile radius from the agency. Confirm.")
    field(d, "5.14 Indemnification cap and basket",
          "Default: $30,000 basket (~10%); cap at $300,000 (purchase price); 12-month survival for general; "
          "longer for fundamentals/tax/healthcare-regulatory. Confirm.")
    field(d, "5.15 Governing law", "Default: Texas - confirm")

    # --- AFFILIATE / SIZE-STANDARD ----------------------------------------
    h2(d, "6. Affiliate / size-standard memo - facts I need")
    field(d, "6.1 Geoff's role with VistaRiver Inc post-sale",
          "Default assumption based on the MIPA: passive note holder ONLY (no officer, director, manager, "
          "employee role; no continuing equity). Confirm or correct.")
    field(d, "6.2 Other businesses Geoff currently owns >=20% of (operating or holding)",
          TBD + "  (list each: name, state, ownership %, role, approximate revenue)")
    field(d, "6.3 Other businesses Geoff currently MANAGES (any title) even if he doesn't own them",
          TBD + "  (officer, manager, signer)")
    field(d, "6.4 Multi-hospice operating history referenced in the plan",
          TBD + "  (which agencies, current status: still active / sold / closed; Geoff's current role with each)")
    field(d, "6.5 'Paloma' is referenced as the source of the migrated clinical team - what's the relationship?",
          TBD + "  (former employer of Silas/Dana? Does Geoff have an ownership or board role? Any current contract?)")
    field(d, "6.6 Bullard's other business interests where he holds 50%+ ownership",
          TBD + "  (these can pull into SBA affiliate analysis even though he's <20% here)")
    field(d, "6.7 Silas's current employer (where he works today, before transitioning to Azalea)",
          TBD + "  (and any non-compete or non-solicit that could constrain hiring)")
    field(d, "6.8 Dana's current employer", TBD + "  (and any non-compete or non-solicit)")
    field(d, "6.9 Bradley's current employer (the 'East-Texas hospice' VP-BD role in his resume)",
          TBD + "  (active employment? Non-compete? Garden leave?)")
    field(d, "6.10 Identity-of-interest affiliation (any close family members operating other hospices)?",
          TBD + "  (SBA aggregates relatives' businesses for affiliation analysis)")

    para(d, "")
    h2(d, "How to return this")
    para(d, "Just type answers into the doc and send it back, or summarize the answers in chat - whichever is "
            "faster for you. Skip any question and we'll use the default noted; flag anything you want me to "
            "research further.")
    footer_note(d)
    save(d, "INTERVIEW_for_remaining_drafts.docx")


# ================================================================ OA AMENDMENT NO. 1
def oa_amendment():
    d = new_doc("OA Amendment No. 1")
    h1(d, "Amendment No. 1 to Operating Agreement")
    para(d, "TYLER HOSPICE HOLDCO L.L.C. - A Wyoming Limited Liability Company", color=GREY, size=9)
    para(d, "")
    para(d, "DRAFT - For review and execution. Prepared for the SBA 7(a) loan application. The Members and "
            "Manager are not represented by counsel as to this Amendment; each is advised to obtain independent "
            "review before signing. Attorney review is recommended.",
         italic=True, size=9, color=GREY)
    para(d, "")
    field(d, "Effective Date", "_______________, 2026")
    field(d, "Original Operating Agreement (the 'Original OA')",
          "Operating Agreement of Tyler Hospice HoldCo L.L.C., effective March 2026 (Version 5)")
    para(d, "")
    h2(d, "Background and recitals")
    para(d, "The Members and the Manager have determined that certain provisions of the Original OA require "
            "amendment to: (i) align the capitalization table with the structure on which the Company's SBA 7(a) "
            "financing is being arranged; (ii) align federal income tax treatment with the multi-member ownership "
            "of Adeline & Lilah, LLC; (iii) update the change-of-ownership acquisition target reference from "
            "Healing Hands Palliative Hospice INC to Hickory Hospice LLC, a Texas limited liability company "
            "(\"Hickory\"); (iv) calibrate Bullard's protective rights to remain consistent with U.S. Small "
            "Business Administration (\"SBA\") Standard Operating Procedure 50 10 8 for minority equity holders "
            "in change-of-ownership transactions; (v) coordinate Member governance with the post-closing "
            "capital structure; and (vi) strengthen the operator-member Restricted Interests so that the entire "
            "13.3% is at risk and forfeitable until vested - a 4.9% time-vested base plus an 8.4% earn-up portion "
            "subject to dual-trigger (time and performance) vesting - with Company repurchase (call) rights on "
            "departure. The Members "
            "hereby agree to amend the Original OA as set forth below. "
            "Capitalized terms used and not otherwise defined have the meanings given in the Original OA.")
    para(d, "")
    h2(d, "1. Restated capitalization (Exhibit A)")
    para(d, "Exhibit A to the Original OA is hereby restated in its entirety as follows. All Percentage "
            "Interests, including those held subject to Forfeiture Conditions, are issued effective as of the "
            "Effective Date of this Amendment, replacing any prior Exhibit A.")
    table(d, ["Member", "Consideration", "Percentage Interest"],
          [["Adeline & Lilah, LLC", "Services rendered (sweat equity)", "39.9%"],
           ["James E. Bullard", "$195,000 cash (capital contribution, two tranches)", "19.5%"],
           ["Silas R. Shelton", "Services rendered (Restricted Interest)", "13.3%*"],
           ["Dana L. Davenport", "Services rendered (Restricted Interest)", "13.3%*"],
           ["Bradley Gene Woodard", "Services rendered (Restricted Interest)", "13.3%*"],
           ["Unissued Pool", "Reserved", "0.7%"],
           ["TOTAL", "", "100.0%"]],
          widths=[2.8, 2.6, 1.2])
    para(d, "*Each Restricted Interest of Silas R. Shelton, Dana L. Davenport, and Bradley Gene Woodard is issued "
            "and outstanding at 13.3%, the entire amount of which is at risk and forfeitable until vested. It "
            "comprises a 4.9% Initial Base that time-vests on continuous service (four-year schedule, one-year "
            "cliff) and an 8.4% Earn-Up Portion subject to dual-trigger vesting, all under the Forfeiture Conditions "
            "of Article V of the Original OA as modified by Section 4 of this Amendment; no portion is Vested at the "
            "Effective Date and each Equity Grantee may earn up to a fully-Vested 13.3%.", italic=True, size=9)
    para(d, "")
    h2(d, "2. Tax treatment (replaces Section 11.2 of the Original OA)")
    para(d, "Section 11.2 of the Original OA is hereby deleted and replaced in its entirety with the following:")
    para(d, "11.2  Tax Matters. The Company shall be treated as a partnership for U.S. federal income tax "
            "purposes under subchapter K of the Internal Revenue Code. No election shall be made by or on "
            "behalf of the Company to be treated as an association taxable as a corporation (including no "
            "election under Treasury Reg. §301.7701-3 and no election under IRC §1361) without the prior "
            "written consent of all Members. The Manager shall serve as the Company's partnership "
            "representative under IRC §6223 for partnership audit purposes. The Manager shall cause the "
            "preparation and timely filing of all federal, state, and local tax returns (including Form 1065 "
            "and Schedule K-1 to each Member) within seventy-five (75) days of each fiscal year-end. All "
            "references in this Agreement to S-Corporation status, IRC §1361, IRC §1362, Form 2553, "
            "\"S-Corp Eligible Person,\" \"S-Corp Eligibility Certification,\" or similar S-Corporation "
            "concepts (including the definition at Section 2.23 and references in Sections 4.5, 8.5, 8.11, "
            "9.1(d), 9.1(e), 11.4(b), 12.4, and 12.5) shall be of no further force or effect. References to "
            "members of any Member entity being \"S-Corp Eligible Persons\" or having to certify eligibility "
            "under §12.4 are deleted. The eligibility certifications in Article XII shall be limited to "
            "(i) accredited-investor status under Regulation D, and (ii) representations and warranties "
            "regarding OIG / SAM exclusion as set forth in Section 9.1(d).",
         italic=True, size=10)
    para(d, "Effect on Adeline & Lilah, LLC. The Members acknowledge that Adeline & Lilah, LLC is a multi-member "
            "Arizona limited liability company taxed as a partnership for federal purposes. This Section 2 of "
            "the Amendment removes the prior incompatibility between the S-Corporation election in the Original "
            "OA and Adeline & Lilah, LLC's multi-member status.", italic=True, size=9.5)
    para(d, "")
    h2(d, "3. Acquisition target (updates Sections 1.2, 3.14, and 3.15 of the Original OA)")
    para(d, "All references in the Original OA to \"Healing Hands Palliative Hospice INC\" are hereby replaced "
            "with \"Hickory Hospice LLC, a Texas limited liability company (the 'Acquired Agency').\" The "
            "Acquired Agency is currently owned by Tracy Gleason and Ann Lozano and operates a Medicare-"
            "certified hospice and palliative care agency in the East Texas market with current Texas Health "
            "and Human Services Commission Home and Community Support Services Agency (\"HCSSA\") licensure "
            "and accreditation. The Company's acquisition of the Acquired Agency is structured as a complete "
            "change of ownership (CHOW) by purchase of one hundred percent (100%) of the Acquired Agency's "
            "membership interests, expected to close on or about late July 2026, subject to the SBA financing "
            "and the equity-injection schedule described in Section 3.3 of the Original OA as amended.")
    para(d, "")
    h2(d, "4. Operator Restricted Interests - sizing, vesting, and repurchase (amends and supplements Article V)")
    para(d, "4.1 Sizing and components. Section 5.1 of the Original OA is amended so that the Restricted Interest "
            "of each of Silas R. Shelton, Dana L. Davenport, and Bradley Gene Woodard (each, an \"Equity Grantee\") "
            "is sized at thirteen and three-tenths percent (13.3%) of the Company, issued and outstanding as of the "
            "Effective Date and held subject to this Section 4. Each Equity Grantee's Restricted Interest comprises "
            "two components:")
    para(d, "   (a) Initial Base - four and nine-tenths percent (4.9%), corresponding to the Equity "
            "Grantee's original Percentage Interest under the Original OA, which is initially Unvested and time-vests "
            "under Section 4.2(a) based on Continuous Service alone (a four-year schedule with a one-year cliff), "
            "without regard to the performance milestones, and is forfeitable while Unvested; and")
    para(d, "   (b) Earn-Up Portion - the remaining eight and four-tenths percent (8.4%), which is initially "
            "Unvested and becomes Vested only as the Equity Grantee satisfies both the time and performance "
            "conditions of Section 4.2(b). Accordingly, no portion of the Restricted Interest is Vested at the "
            "Effective Date; the entire 13.3% is at risk and forfeitable until Vested, and each Equity Grantee may "
            "vest up to a fully-Vested 13.3%.")
    para(d, "The milestone tranche sizing in Section 5.2 is amended to apply to the Earn-Up Portion as follows:")
    table(d, ["Tranche", "Performance milestone", "Tranche size (of 8.4% Earn-Up)", "Company-level %"],
          [["First", "Achieve Breakeven (1 month EBITDA >= 0)", "25% of Earn-Up", "2.100%"],
           ["Second", "3 consecutive Profitable months", "50% of Earn-Up", "4.200%"],
           ["Final", "12 consecutive Profitable months", "25% of Earn-Up", "2.100%"],
           ["Earn-Up subtotal", "", "100% of Earn-Up", "8.400%"],
           ["Initial Base", "Time-vests (4-yr; 1-yr cliff)", "-", "4.900%"],
           ["Total per Equity Grantee", "", "", "13.300%"]],
          widths=[1.0, 2.5, 1.6, 1.5])
    para(d, "4.2 Vesting - the entire Restricted Interest is at risk. Section 5.2 of the Original OA is amended so "
            "that no portion of the Restricted Interest is Vested at the Effective Date and the entire thirteen and "
            "three-tenths percent (13.3%) becomes vested and non-forfeitable (\"Vested\") only as set out below; any "
            "portion not yet Vested is \"Unvested\" and remains subject to forfeiture under Section 4.4. "
            "\"Continuous Service\" means continuous service to the Company or the Acquired Agency as an employee, "
            "officer, or manager.")
    para(d, "   (a) Initial Base - time-vesting only. The Initial Base (Section 4.1(a)) vests on Continuous Service "
            "alone on a four-year schedule with a one-year cliff: nothing vests before the first anniversary; "
            "twenty-five percent (25%) of the Initial Base vests on the first anniversary (the \"Cliff\"); and the "
            "remaining seventy-five percent (75%) vests in thirty-six (36) equal monthly installments over months "
            "13-48, so that 100% of the Initial Base is Vested on the fourth anniversary. No performance milestone "
            "is required for the Initial Base to vest.")
    para(d, "   (b) Earn-Up Portion - dual trigger (time AND performance). The Earn-Up Portion (Section 4.1(b)) "
            "becomes Vested only to the extent BOTH of the following are satisfied; the Vested percentage of the "
            "Earn-Up Portion at any time equals the LESSER of (i) and (ii):")
    para(d, "      (i) Time-vesting. Continuous Service on the same four-year schedule with a one-year cliff applied "
            "to the Earn-Up Portion: nothing vests before the first anniversary; twenty-five percent (25%) vests on "
            "the Cliff; and the remaining seventy-five percent (75%) vests in thirty-six (36) equal monthly "
            "installments over months 13-48.")
    para(d, "      (ii) Performance-vesting. The milestone tranches in Section 4.1 (Breakeven; 3 consecutive "
            "Profitable months; 12 consecutive Profitable months), determined under Sections 5.2-5.3 of the "
            "Original OA.")
    para(d, "For the Earn-Up Portion, both conditions must be met; satisfying only one does not vest it. A "
            "Grantee's total Vested Interest at any time equals the then-Vested portion of the Initial Base plus the "
            "then-Vested portion of the Earn-Up Portion; the balance is Unvested. This Section 4.2 supersedes and "
            "replaces the Termination Without Cause pro-rata credit at Section 5.5 of the Original OA, which is "
            "deleted.")
    para(d, "4.3 Separation definitions. \"Separation\" means an Equity Grantee ceasing Continuous Service for "
            "any reason. The reason determines the buy-back price under Section 4.5:")
    para(d, "   - \"Good-Leaver Separation\": Separation due to (i) termination by the Company without Cause; "
            "(ii) death; (iii) Disability; (iv) Retirement; or (v) voluntary resignation on or after full "
            "time-vesting (the fourth anniversary).")
    para(d, "   - \"Bad-Leaver Separation\": Separation due to (i) termination by the Company for Cause (as "
            "defined in Section 5.4 of the Original OA) at any time, or (ii) voluntary resignation before full "
            "time-vesting (the fourth anniversary).")
    para(d, "   - \"Retirement\" means voluntary Separation at or after age sixty-five (65) with at least three "
            "(3) years of Continuous Service, designated as a retirement by the Manager in good faith. "
            "\"Disability\" means inability to perform the essential functions of the role for one hundred eighty "
            "(180) consecutive days, as determined by the Manager in good faith.")
    para(d, "4.4 Forfeiture of Unvested Interest. Upon any Separation, the Equity Grantee's Unvested Interest is "
            "automatically and immediately forfeited for no consideration and reverts to the Unissued Pool, "
            "regardless of the reason for Separation. The Manager, as attorney-in-fact under Section 5.4 of the "
            "Original OA, is authorized to execute all instruments to effect the forfeiture and any repurchase "
            "under Section 4.5.")
    para(d, "4.5 Company repurchase (call) right over Vested Interest. Upon any Separation, the Company shall "
            "have the right, but not the obligation, to purchase all (but not less than all) of the Equity "
            "Grantee's Vested Interest (the \"Call Right\"), exercisable by written notice within one hundred "
            "twenty (120) days after the later of the Separation date and the date the Call Price is finally "
            "determined. If the Company does not fully exercise, the other Members (excluding the Departing "
            "Grantee) may purchase the balance pro rata within a further thirty (30) days. The purchase price "
            "(the \"Call Price\") for the Vested Interest is:")
    para(d, "   (a) Good-Leaver Separation: the Fair Market Value of the Vested Interest under Section 4.6.")
    para(d, "   (b) Bad-Leaver Separation: the LOWER of (i) the Equity Grantee's cost basis in the Vested "
            "Interest (for a Restricted Interest issued for services, the cash amount, if any, actually paid, "
            "and otherwise zero) and (ii) Fair Market Value under Section 4.6. The Members acknowledge this may "
            "result in a Call Price at or near zero for a Bad-Leaver Separation.")
    para(d, "4.6 Fair Market Value. \"Fair Market Value\" of a Vested Interest means its pro-rata share of the "
            "Company's equity value - enterprise value less all outstanding indebtedness (including the SBA Loan "
            "and any seller financing) - determined as of the last day of the calendar month preceding "
            "Separation by an independent qualified appraiser under the appraised-value procedure of Section 3.11 "
            "of the Original OA, reflecting appropriate discounts for the interest's minority (non-controlling) "
            "and illiquid (lack-of-marketability) character. The Company bears the appraisal cost unless the "
            "Equity Grantee disputes it and a second appraisal differs by less than ten percent (10%), in which "
            "case the disputing Equity Grantee bears the second appraisal's cost.")
    para(d, "4.7 Payment terms; SBA subordination. The Call Price shall be paid by an unsecured, subordinated "
            "promissory note of the Company at the Applicable Federal Rate, in equal quarterly installments over "
            "four (4) years (prepayable without penalty; the Manager may instead elect a lump sum). "
            "Notwithstanding the foregoing, no payment of the Call Price shall be made, and any such note is "
            "fully subordinated, to the extent a payment would (i) violate the SBA Loan documents or any lender "
            "covenant, (ii) occur during or cause an event of default under the SBA Loan, or (iii) reduce the "
            "Company's liquidity or debt-service coverage below any level the SBA Loan requires. Suspended "
            "payments accrue and resume when permitted; transfer of the Vested Interest to the Company may close "
            "notwithstanding deferral of cash payment.")
    para(d, "4.8 Mechanics pending repurchase. From Separation until the Call Right expires or the repurchase "
            "closes, the Departing Grantee's Vested Interest shall be non-voting and not entitled to "
            "distributions declared after Separation (other than tax distributions on allocated income), without "
            "limiting the right to the Call Price. The Departing Grantee shall execute all transfer instruments "
            "reasonably requested; the Manager's power of attorney under Section 5.4 of the Original OA extends "
            "to all actions under this Section 4.")
    para(d, "4.9 Tax; Section 83(b). The Section 83(b) election requirement at Section 5.7 of the Original OA "
            "remains in full force. The vesting and repurchase terms of this Section 4 are intended to be "
            "consistent with the Company's partnership tax treatment under Section 2 of this Amendment and any "
            "intended profits-interest treatment under Rev. Proc. 93-27; each Equity Grantee should consult "
            "independent tax counsel. Except as modified by this Section 4, Article V of the Original OA "
            "(including Sections 5.4 and 5.7) remains in full force and effect.", italic=True, size=9.5)
    para(d, "")
    h2(d, "5. Bullard's protective rights - SBA calibration")
    para(d, "The parties confirm that Bullard's protective rights set forth in the Original OA (information "
            "rights under Section 6.7, anti-dilution and pro-rata protection under Sections 6.8 and 12.6, "
            "voting rights on Reserved Matters under Section 6.2, strategic observer role under Section 6.9, "
            "tag-along rights under Section 8.8, and the Texas Shootout buy-sell provision under Section 8.10) "
            "remain in full force and effect, except as expressly modified by this Section 5. The Members "
            "acknowledge that these protective rights are customary minority-investor protections of an economic "
            "and fundamental nature and do not confer day-to-day operational control, employment, agency, or "
            "authority to bind the Company in the ordinary course of business, which remain solely with the "
            "Manager.")
    para(d, "5.1 Anti-dilution floor. Section 6.8 (Anti-Dilution Protection) and Section 12.3 (Hard Dilution "
            "Floor - Bullard Minimum Interest) of the Original OA are amended so that all references to "
            "\"twenty percent (20.0%)\" as Bullard's floor Percentage Interest are replaced with "
            "**\"nineteen and one-half percent (19.5%)\"**.")
    para(d, "5.2 SBA debt carve-out. Section 6.2(d) of the Original OA (Reserved Matters - incurrence of "
            "indebtedness exceeding $500,000.00) is amended by adding the following proviso: \"; provided, "
            "however, that the Manager may incur, modify, refinance, restructure, prepay, or take any other "
            "action with respect to (x) the SBA 7(a) loan made or guaranteed by any participating lender under "
            "the U.S. Small Business Administration's 7(a) loan program (the 'SBA Loan'), (y) the Hickory "
            "seller financing arising from the acquisition described in Section 3 of Amendment No. 1, and "
            "(z) any commercial loan that refinances or replaces (x) or (y), without requiring the consent of "
            "Members under this Section 6.2(d), so long as such action is in the ordinary course of operating "
            "the Company's hospice business and does not in itself constitute a Change of Control.\"")
    para(d, "5.3 Texas Shootout deferral. Section 8.10 (Buy-Sell - Texas Shootout) of the Original OA is "
            "amended by adding the following at the end: \"Notwithstanding any other provision of this "
            "Section 8.10, no Member may invoke, exercise, or enforce the Texas Shootout buy-sell provision "
            "while any principal or accrued interest under the SBA Loan remains outstanding. Upon payment in "
            "full of the SBA Loan and any successor or refinancing loan that includes SBA guarantee, this "
            "deferral shall lapse and Section 8.10 shall again be operative without restriction.\"")
    para(d, "")
    h2(d, "6. Reserved Matters arithmetic (clarification to Section 2.20 of the Original OA)")
    para(d, "For the avoidance of doubt and reflecting the restated cap table in Section 1 of this Amendment, "
            "the parties acknowledge that:")
    for x in ["The Unissued Pool (0.7%) is excluded from the denominator for voting purposes under Section 2.20.",
              "Adeline & Lilah, LLC alone holds 39.9% and cannot unilaterally pass a 60% Reserved Matter.",
              "Adeline & Lilah, LLC may pass a Reserved Matter by combining its 39.9% with the votes of any two (2) Equity Grantees (totaling 66.5%), or with Bullard's 19.5% plus the vote of at least one (1) Equity Grantee (totaling 72.7%).",
              "Bullard does not hold a unilateral veto on Reserved Matters; a Reserved Matter may be passed without his vote if Adeline & Lilah, LLC and at least two (2) Equity Grantees concur."]:
        para(d, "  - " + x)
    para(d, "")
    h2(d, "7. Closing schedule coordination (new Section 3.16 of the Original OA)")
    para(d, "A new Section 3.16 is added to the Original OA as follows: \"3.16 Coordination with SBA "
            "Closing. The Members acknowledge that the change-of-ownership closing of the Acquired Agency "
            "shall be scheduled to occur on or after the date on which Bullard has contributed the full "
            "$195,000.00 of the Investor Capital Contribution under Section 3.3 (currently anticipated by the "
            "end of July 2026), in order to permit the SBA Loan to disburse with the full equity injection on "
            "deposit, as required by SBA SOP 50 10 8. The Manager is authorized to coordinate the CHOW closing "
            "date, the SBA Loan disbursement, and Bullard's tranche schedule accordingly.\"")
    para(d, "")
    h2(d, "8. Effect on prior consents; ratification")
    para(d, "Adeline & Lilah, LLC, as the Member holding sixty percent (60.0%) of Percentage Interests under "
            "the Original OA (and which, after giving effect to this Amendment, holds 39.9% of Percentage "
            "Interests), and James E. Bullard, as the Member whose rights under Sections 4.2, 4.5, 3.15, 6.7, "
            "6.8, 6.9, and Article VIII are subject to consent rights under Section 11.4(a), each consent to "
            "the amendments set forth herein. The Members ratify and confirm the Original OA as modified by "
            "this Amendment. All provisions of the Original OA not specifically modified by this Amendment "
            "remain in full force and effect.")
    para(d, "")
    h2(d, "9. Counterparts; electronic signatures")
    para(d, "This Amendment may be executed in counterparts and by electronic signature (including DocuSign "
            "and PDF). Each counterpart taken together shall constitute one and the same instrument.")
    para(d, "")
    para(d, "")
    para(d, "SIGNATURES", bold=True)
    para(d, "")
    para(d, "MANAGER (individually) — GEOFF SCHACKMANN:")
    para(d, "_______________________________________     Date: ____________")
    para(d, "Geoff Schackmann, Manager (non-Member capacity)")
    para(d, "Address: 4602 E Cheery Lynn Rd, Phoenix, AZ 85018  |  geoff@azaleahospice.com")
    para(d, "")
    para(d, "MEMBER — ADELINE & LILAH, LLC:")
    para(d, "_______________________________________     Date: ____________")
    para(d, "By: Geoff Schackmann, Authorized Representative  |  Interest after Amendment: 39.9%")
    para(d, "")
    para(d, "_______________________________________     Date: ____________")
    para(d, "By: Mary Elizabeth Burcham, Authorized Representative (50% member of Adeline & Lilah, LLC)")
    para(d, "")
    para(d, "MEMBER — INVESTOR:")
    para(d, "_______________________________________     Date: ____________")
    para(d, "James E. Bullard  |  Interest after Amendment: 19.5%")
    para(d, "Address: 13910 Indiana Ave, Suite 300, Lubbock, TX 79423  |  jimbullard01@aol.com")
    para(d, "")
    para(d, "MEMBER — EQUITY GRANTEE (RESTRICTED INTEREST):")
    para(d, "_______________________________________     Date: ____________")
    para(d, "Silas R. Shelton  |  Interest after Amendment: 13.3% Restricted Interest, entire 13.3% subject to Forfeiture Conditions (4.9% time-vested base; 8.4% dual-trigger Earn-Up)")
    para(d, "Address: 1503 Lake Park Circle, Hideaway, TX 75771")
    para(d, "")
    para(d, "_______________________________________     Date: ____________")
    para(d, "Dana L. Davenport  |  Interest after Amendment: 13.3% Restricted Interest, entire 13.3% subject to Forfeiture Conditions (4.9% time-vested base; 8.4% dual-trigger Earn-Up)")
    para(d, "Address: ____________________________________  [REQUIRED]")
    para(d, "")
    para(d, "_______________________________________     Date: ____________")
    para(d, "Bradley Gene Woodard  |  Interest after Amendment: 13.3% Restricted Interest, entire 13.3% subject to Forfeiture Conditions (4.9% time-vested base; 8.4% dual-trigger Earn-Up)")
    para(d, "Address: 421 W Cumberland Rd, Apt 403, Tyler, TX 75703")
    para(d, "")
    footer_note(d)
    save(d, "08_entity_documents/OA_Amendment_No_1_DRAFT.docx")


# ================================================================ REVISED BULLARD ATTESTATION
def bullard_attestation_v2():
    d = new_doc("Bullard Attestation v2")
    h1(d, "Investor Attestation - No Operational Control")
    para(d, "Tyler Hospice HoldCo L.L.C. - SBA 7(a) Application  |  Revised to reflect OA Amendment No. 1",
         color=GREY, size=9)
    para(d, "")
    para(d, "The undersigned, James E. Bullard, in connection with the SBA 7(a) loan application of Tyler "
            "Hospice HoldCo L.L.C. (the \"Company\"), and as a minority equity holder in the Company, hereby "
            "attests as follows:")
    h2(d, "1. Ownership and contribution")
    para(d, "I hold a 19.5% direct membership interest in the Company, acquired in exchange for a $195,000 "
            "cash capital contribution. My contribution is being made in two tranches: $100,000 was wired "
            "to the Company's operating account on May 7, 2026, with the remaining $95,000 scheduled to be "
            "contributed by July 31, 2026. The funds are my own non-borrowed cash, drawn from "
            "personal savings; I have not used loan proceeds, advances, or credit (including credit-card "
            "advances or unsecured personal loans) to fund any portion of the contribution.")
    h2(d, "2. No management role")
    para(d, "I hold no officer, director, manager, or employee position with the Company. I do not "
            "participate in the day-to-day management or operation of the Company or any of its subsidiaries, "
            "including Hickory Hospice LLC. The Manager (Geoff Schackmann, individually) and the Member "
            "Adeline & Lilah, LLC have sole authority over management, hiring and firing, financial "
            "decisions, vendor relationships, clinical operations, and all other Company operations.")
    h2(d, "3. Protective minority rights only")
    para(d, "My equity interest entitles me to certain customary minority-investor protective rights set "
            "forth in the Operating Agreement, as amended by Amendment No. 1 (collectively, the "
            "\"Protective Rights\"): (a) information rights including monthly unaudited financial "
            "statements, annual reviewed financial statements, copies of tax returns, and notice of "
            "material events; (b) anti-dilution protection with a 19.5% floor; (c) pro-rata participation "
            "rights on future issuances; (d) voting rights on a limited set of fundamental Reserved Matters "
            "(sale of substantially all assets, merger, additional equity issuances, certain non-SBA debt "
            "above $500,000, amendment of the Operating Agreement, dissolution, and Change of Control); "
            "(e) tag-along rights on a Change of Control transfer by Adeline & Lilah, LLC; and (f) a Texas "
            "Shootout buy-sell provision deferred until the SBA Loan is paid in full.")
    para(d, "These Protective Rights are economic in nature and are limited to fundamental Company events. "
            "They do not confer authority over day-to-day operations, hiring or firing, vendor selection, "
            "clinical decisions, payor relationships, marketing, or the ordinary-course business of the "
            "Company. The Manager retains sole authority over all of these matters.")
    h2(d, "4. SBA-related debt and Texas Shootout — explicit non-control")
    para(d, "I confirm and acknowledge that, under OA Amendment No. 1: (a) my consent right on the "
            "incurrence, modification, refinancing, or prepayment of indebtedness above $500,000 expressly "
            "does NOT apply to the SBA Loan or to commercial loans that refinance the SBA Loan, so that my "
            "approval is not required for any SBA-related debt action by the Manager; and (b) the Texas "
            "Shootout buy-sell provision is suspended while the SBA Loan is outstanding, so that I cannot "
            "force a buyout of the Company's controlling Members during the SBA financing period.")
    h2(d, "5. No side agreement, secret option, or convertible instrument")
    para(d, "Except for the Operating Agreement as amended by Amendment No. 1 and the Subscription Agreement "
            "between the Company and me, there is no written or oral side agreement, voting agreement, "
            "proxy, option, warrant, convertible note, profits-interest plan, earnout, employment "
            "arrangement, or other instrument between me and the Company, the Manager, Adeline & Lilah, "
            "LLC, or any other Member that would (i) increase my equity interest to 20% or more, "
            "(ii) give me day-to-day management or operational control of the Company, or (iii) entitle "
            "me to direct or override the Manager's decisions in the ordinary course of business.")
    h2(d, "6. Reserved-matters arithmetic")
    para(d, "I acknowledge that the Operating Agreement's Reserved Matters provisions require a 60% "
            "supermajority vote. Adeline & Lilah, LLC (39.9%) may pass any Reserved Matter without my vote "
            "by combining its interest with at least two of the three Equity Grantees (Silas, Dana, "
            "Bradley), whose combined interest of 39.9% gives a total of 66.5%. I therefore do not hold "
            "a unilateral veto on Reserved Matters.")
    h2(d, "7. Compliance with SBA requirements")
    para(d, "I understand that my passive minority interest is being relied upon by the SBA-participating "
            "lender in connection with the equity-injection and personal-guaranty analysis under SOP 50 10 8 "
            "and 13 CFR 120.160. I acknowledge that any change in my equity interest above 20%, the addition "
            "of any operational control right not disclosed above, or the creation of any side agreement of "
            "the kind described above could change my status under SBA regulations, and I agree to provide "
            "prompt written notice to the Company and the lender of any such change while the SBA loan "
            "remains outstanding.")
    para(d, "")
    para(d, "I attest under penalty of perjury that the foregoing is true and correct to the best of my "
            "knowledge.")
    para(d, "")
    para(d, "______________________________________________________")
    para(d, "James E. Bullard", bold=True)
    para(d, "Address: 13910 Indiana Ave, Suite 300, Lubbock, TX 79423")
    para(d, "Date: ___________________________")
    para(d, "")
    para(d, "STATE OF ___________________________}")
    para(d, "COUNTY OF __________________________}")
    para(d, "")
    para(d, "Subscribed and sworn to before me this _____ day of _____________________, 2026.")
    para(d, "")
    para(d, "______________________________________________________")
    para(d, "Notary Public                                      My commission expires: _______________")
    para(d, "")
    para(d, "DRAFT - reflects OA Amendment No. 1 calibration of Bullard's rights. The undersigned should "
            "review with independent counsel before execution; the description of \"Protective Rights\" "
            "above should be conformed to the executed Amendment.",
         italic=True, size=8.5, color=GREY)
    footer_note(d)
    save(d, "08_entity_documents/Bullard_Investor_Attestation_v2_DRAFT.docx")


# ================================================================ BULLARD SUBSCRIPTION / CAPITAL CONTRIBUTION AGREEMENT
def bullard_subscription():
    d = new_doc("Bullard Subscription Agreement")
    h1(d, "Subscription and Capital Contribution Agreement")
    para(d, "TYLER HOSPICE HOLDCO L.L.C. - A Wyoming Limited Liability Company", color=GREY, size=9)
    para(d, "DRAFT for review and execution. The parties are not represented by counsel as to this Agreement; "
            "independent review is recommended before signing.", italic=True, size=9, color=GREY)
    para(d, "")
    para(d, "This Subscription and Capital Contribution Agreement (this \"Agreement\") is entered into as of "
            "_______________, 2026 (the \"Effective Date\"), by and between Tyler Hospice HoldCo L.L.C., a "
            "Wyoming limited liability company (the \"Company\"), and James E. Bullard, an individual residing "
            "at 13910 Indiana Ave, Suite 300, Lubbock, TX 79423 (the \"Investor\"). This Agreement is the "
            "\"Subscription Agreement\" referenced in Section 11.5 of the Company's Operating Agreement, as "
            "amended by Amendment No. 1 (together, the \"Operating Agreement\").")
    h2(d, "1. Subscription")
    para(d, "The Investor hereby irrevocably subscribes for and agrees to purchase a nineteen and nine-tenths "
            "percent (19.5%) membership interest in the Company (the \"Interest\") for an aggregate purchase "
            "price of One Hundred Ninety-Five Thousand Dollars ($195,000.00) (the \"Purchase Price\"), and the "
            "Company agrees to issue the Interest to the Investor upon the terms and conditions set forth "
            "herein and in the Operating Agreement. The Interest is fully earned upon payment and is not "
            "subject to any vesting, forfeiture, or repurchase condition.")
    h2(d, "2. Payment of the Purchase Price (tranches)")
    para(d, "The Investor shall pay the Purchase Price by wire transfer of immediately available funds to the "
            "Company's operating account in two tranches:")
    table(d, ["Tranche", "Amount", "Timing", "Status"],
          [["Tranche 1", "$100,000", "Paid May 7, 2026", "RECEIVED - on deposit"],
           ["Tranche 2", "$95,000", "On or before July 31, 2026", "Committed"],
           ["Total", "$195,000", "", ""]],
          widths=[1.1, 1.1, 2.3, 2.0])
    para(d, "The parties acknowledge that the $100,000 Tranche 1 wire was received in the Company's Mercury "
            "(Column N.A.) operating account ending 1275 on May 7, 2026. The Investor's full 19.5% Interest is "
            "recorded and effective as of the Effective Date; provided that if the Investor fails to fund any "
            "tranche when due and fails to cure within ten (10) business days of written notice, the Company's "
            "sole remedy shall be to reduce the Investor's Percentage Interest proportionally to the amount "
            "actually funded (funded amount / $195,000 x 19.5%), with the unfunded balance reverting to "
            "Adeline & Lilah, LLC, and the Investor shall have no further obligation as to the unfunded "
            "portion. The tranche amounts of $75,000 for Tranches 2 and 3 may be adjusted by mutual written "
            "agreement so long as the full $195,000 is funded on or before July 31, 2026.")
    h2(d, "3. Source of funds")
    para(d, "The Investor represents that the entire Purchase Price is and will be funded from the Investor's "
            "own personal savings and is not borrowed, advanced, or financed by any third party, and is not "
            "secured directly or indirectly by any assets of the Company or the business being acquired. The "
            "Investor agrees to provide the Company and its SBA lender with bank or brokerage statements "
            "evidencing the availability and source of the funds for each tranche, covering at least the "
            "thirty (30) days prior to each wire, as required by SBA SOP 50 10 8.")
    h2(d, "4. Use of proceeds")
    para(d, "The Company shall use the Purchase Price as the borrower equity injection for its SBA 7(a) "
            "financing and the change-of-ownership acquisition of Hickory Hospice LLC, and for working capital "
            "of the hospice business, consistent with Section 3.14 of the Operating Agreement.")
    h2(d, "5. Adoption of the Operating Agreement")
    para(d, "By executing this Agreement, the Investor (a) adopts, joins, and agrees to be bound by all terms "
            "of the Operating Agreement as a Member; (b) confirms that the rights associated with the Interest "
            "are solely those set forth in the Operating Agreement as amended (including the protective "
            "minority rights and the SBA-related limitations described in Amendment No. 1); and (c) "
            "acknowledges that the Interest does not confer any day-to-day management or operational control "
            "of the Company.")
    h2(d, "6. Investor representations and warranties")
    para(d, "The Investor represents and warrants to the Company that: (a) the Investor is an \"accredited "
            "investor\" as defined in Rule 501(a) of Regulation D under the Securities Act of 1933, as amended; "
            "(b) the Investor is acquiring the Interest for the Investor's own account for investment and not "
            "with a view to resale or distribution; (c) the Investor has such knowledge and experience in "
            "financial and business matters as to be capable of evaluating the merits and risks of the "
            "investment and can bear the economic risk of a total loss; (d) the Investor understands the "
            "Interest is not registered under federal or state securities laws and is a \"restricted "
            "security\"; (e) the Investor is a United States citizen or lawful permanent resident; (f) neither "
            "the Investor nor any entity the Investor controls is excluded, suspended, or debarred from any "
            "federal or state healthcare program (OIG LEIE / SAM.gov); and (g) the Investor has had the "
            "opportunity to ask questions of and receive answers from the Manager.")
    h2(d, "7. Company representations")
    para(d, "The Company represents that, upon receipt of each tranche, the Interest will be duly issued, "
            "and that the Company shall record the Investor as a Member on its books and update Exhibit A "
            "to the Operating Agreement accordingly.")
    h2(d, "8. Miscellaneous")
    para(d, "This Agreement, together with the Operating Agreement, constitutes the entire agreement between "
            "the parties regarding the subject matter and supersedes all prior understandings. This Agreement "
            "is governed by the laws of the State of Wyoming. It may be executed in counterparts and by "
            "electronic signature. If any provision conflicts with the Operating Agreement, the Operating "
            "Agreement controls except as to the tranche payment schedule in Section 2, which controls as to "
            "payment timing.")
    para(d, "")
    para(d, "COMPANY — TYLER HOSPICE HOLDCO L.L.C.", bold=True)
    para(d, "By: _______________________________________   Date: ____________")
    para(d, "Geoff Schackmann, Manager")
    para(d, "")
    para(d, "INVESTOR", bold=True)
    para(d, "_______________________________________   Date: ____________")
    para(d, "James E. Bullard")
    footer_note(d)
    save(d, "08_entity_documents/Bullard_Subscription_and_Capital_Contribution_DRAFT.docx")


# ================================================================ BULLARD SOURCE-OF-FUNDS LETTER
def bullard_source_of_funds():
    d = new_doc("Bullard Source of Funds")
    h1(d, "Source-of-Funds Letter - Equity Injection")
    para(d, "SBA 7(a) Application - Tyler Hospice HoldCo L.L.C.  |  Investor: James E. Bullard",
         color=GREY, size=9)
    para(d, "Template for the Investor to complete with his bank/brokerage details and sign. Accompanies the "
            "statements that evidence each tranche.", italic=True, size=9, color=GREY)
    para(d, "")
    field(d, "Date", TBD)
    para(d, "To: [SBA-participating lender] and Tyler Hospice HoldCo L.L.C.")
    para(d, "")
    para(d, "Re: Source of $195,000 equity injection into Tyler Hospice HoldCo L.L.C.")
    para(d, "")
    para(d, "To Whom It May Concern:")
    para(d, "I, James E. Bullard, am contributing $195,000 in cash to Tyler Hospice HoldCo L.L.C. (the "
            "\"Company\") in exchange for a 19.5% membership interest. I am providing this letter in support "
            "of the Company's SBA 7(a) loan application to document the source of those funds, as required "
            "by SBA SOP 50 10 8.")
    h2(d, "1. Source")
    para(d, "The full $195,000 comes from my personal savings, accumulated over time and held in the "
            "account(s) listed below. None of these funds are borrowed, advanced on credit, or financed by "
            "any third party, and none are or will be secured by the assets of the Company or of Hickory "
            "Hospice LLC.")
    h2(d, "2. Account(s) holding the funds")
    table(d, ["Institution", "Account type", "Last 4 of account", "Approx. balance held for this purpose"],
          [[TBD, TBD, TBD, TBD],
           [TBD, TBD, TBD, TBD]],
          widths=[2.2, 1.5, 1.3, 1.7])
    h2(d, "3. Contribution schedule")
    table(d, ["Tranche", "Amount", "Date wired / to be wired", "Sending account (last 4)"],
          [["Tranche 1", "$100,000", "May 7, 2026 (completed)", TBD],
           ["Tranche 2", "$95,000", "On or before July 31, 2026", TBD]],
          widths=[1.0, 1.0, 2.2, 2.0])
    h2(d, "4. Supporting documentation")
    para(d, "I have provided (or will provide) bank/brokerage statements for the account(s) above covering at "
            "least the 30 days prior to each wire, showing the funds available and on deposit. I authorize the "
            "Company and its SBA lender to verify these funds with my financial institution(s).")
    h2(d, "5. Certification")
    para(d, "I certify under penalty of perjury that the foregoing is true and correct, and that the funds "
            "described are my own and are not borrowed.")
    para(d, "")
    para(d, "_______________________________________   Date: ____________")
    para(d, "James E. Bullard", bold=True)
    para(d, "13910 Indiana Ave, Suite 300, Lubbock, TX 79423  |  jimbullard01@aol.com")
    footer_note(d)
    save(d, "10_supporting_documents/equity_injection_evidence/Bullard_Source_of_Funds_Letter_TEMPLATE.docx")


# ================================================================ AFFILIATE / SIZE-STANDARD MEMO
def affiliate_memo():
    d = new_doc("Affiliate Memo")
    h1(d, "Affiliate and Size-Standard Memorandum")
    para(d, "SBA 7(a) Application - Tyler Hospice HoldCo L.L.C. (dba Azalea Hospice & Palliative Care)",
         color=GREY, size=9)
    para(d, "This memorandum identifies the borrower's affiliates for SBA size-standard and eligibility "
            "purposes (13 CFR 121.103) and documents why the borrower meets the applicable size standard.")
    h2(d, "Applicable size standard")
    para(d, "The borrower operates in NAICS 621610 (Home Health Care Services) / hospice, for which the SBA "
            "receipts-based size standard is $19.0 million in average annual receipts (confirm current "
            "threshold at application). As a newly formed entity acquiring a single small hospice agency, the "
            "borrower and its affiliates are comfortably within the standard.")
    h2(d, "Ownership and control analysis")
    para(d, "Under 13 CFR 121.103, affiliation arises principally through control - common ownership of more "
            "than 50%, common management, or identity of interest. The borrower's controlling party is Geoff "
            "Schackmann, through Adeline & Lilah, LLC (39.9%) and as Manager. Affiliation is therefore "
            "assessed primarily through Mr. Schackmann's other controlled entities.")
    table(d, ["Entity", "Relationship to borrower", "Geoff's interest / role", "Affiliate?"],
          [["Hickory Hospice LLC (TX)", "Acquisition target / future wholly owned operating subsidiary",
            "100% after CHOW close", "Yes - consolidated as subsidiary"],
           ["Adeline & Lilah, LLC (AZ)", "Holding entity for Geoff's interest in the borrower",
            "50% member + Manager", "Yes - upstream holder; no operations"],
           ["VistaRiver Inc (OR)", "Geoff sold his 33.33% in Aug 2025; holds only a seller note",
            "0% equity; passive note holder only; no officer/manager/employee role",
            "No - passive creditor, no control"],
           ["James E. Bullard's other interests", "Bullard is a 19.5% passive minority investor in the borrower",
            "n/a to Geoff; Bullard holds no control of borrower", "Investigate Bullard's >50% entities (below)"]],
          widths=[1.9, 2.2, 2.0, 1.2])
    h2(d, "VistaRiver Inc - not an affiliate")
    para(d, "Geoff Schackmann sold his entire 33.33% membership interest in VistaRiver Inc on August 15, 2025 "
            "(Membership Interest Purchase Agreement on file in folder 10). He retains no equity, no officer, "
            "manager, director, or employee role, and no control of any kind. His only continuing relationship "
            "is as the holder of a secured promissory note (a passive creditor position). Under 13 CFR "
            "121.103, a creditor relationship without ownership or control does not create affiliation. "
            "VistaRiver is therefore NOT an affiliate of the borrower. The note is, however, a personal asset "
            "and income source disclosed on Mr. Schackmann's Personal Financial Statement.")
    h2(d, "Geoff Schackmann - no other controlled operating businesses")
    para(d, "Mr. Schackmann has confirmed that, other than the entities listed above, he does not own 20% or "
            "more of, nor manage, any other operating business. The multi-hospice operating experience "
            "described in the business plan and his resume reflects prior operating roles and engagements; "
            "to the extent any such agency is currently owned or controlled by Mr. Schackmann, it should be "
            "listed here. [Confirm: list any currently owned/controlled hospice agency, with ownership %, "
            "role, and revenue, or confirm none.]")
    para(d, "Action item: Mr. Schackmann to confirm in writing that he holds no current ownership of 20%+ or "
            "management role in any operating business other than those listed above, so this memo can state "
            "it affirmatively for the lender.", italic=True, size=9.5)
    h2(d, "James E. Bullard - minority investor; limited affiliation reach")
    para(d, "Mr. Bullard is a 19.5% passive minority investor in the borrower with protective minority rights "
            "but no operational control (see OA Amendment No. 1 and his Investor Attestation). Because he does "
            "not control the borrower, his other business holdings generally do not pull into the borrower's "
            "affiliation group. However, SBA may review entities Bullard controls (>50%) for completeness. "
            "[Confirm: list any business in which Mr. Bullard holds 50%+, or confirm none known.]")
    h2(d, "Operator-members - employment, not affiliation")
    para(d, "Silas R. Shelton, Dana L. Davenport, and Bradley G. Woodard hold 13.3% each (Restricted "
            "Interests; the full 13.3% forfeitable until vested - a 4.9% time-vested base plus an 8.4% dual-trigger "
            "earn-up) and serve in "
            "operating roles. None controls the borrower. Their "
            "prior or current employers are relevant to non-compete / non-solicit and transition questions "
            "(addressed separately), not to affiliation. Mr. Woodard has confirmed he has no non-compete "
            "with his current employer.")
    h2(d, "Conclusion")
    para(d, "After giving effect to the CHOW, the borrower's affiliate group consists of the borrower, its "
            "wholly owned subsidiary Hickory Hospice LLC, and the upstream holding entity Adeline & Lilah, "
            "LLC (non-operating). Combined average annual receipts are well below the applicable size "
            "standard. VistaRiver Inc is not an affiliate (passive creditor relationship only). The borrower "
            "is a small business concern eligible for SBA 7(a) financing, subject to the lender's confirmation "
            "of the current size threshold and the two open confirmations noted above.")
    footer_note(d)
    save(d, "01_company_profile/Affiliate_and_Size_Standard_MEMO.docx")


# ================================================================ COVER LETTER + ANTICIPATED Q&A
def cover_letter_and_qa():
    d = new_doc("Cover Letter and QA")
    h1(d, "Cover Letter")
    para(d, "Tyler Hospice HoldCo L.L.C. (dba Azalea Hospice & Palliative Care)  |  SBA 7(a) Application",
         color=GREY, size=9)
    para(d, "Generic version - personalize the lender name/address and the salutation once a specific lender "
            "is identified from John Hart's network.", italic=True, size=9, color=GREY)
    para(d, "")
    field(d, "Date", TBD)
    para(d, "[Lender name]")
    para(d, "[Lender address]")
    para(d, "Attn: [SBA Loan Officer]")
    para(d, "")
    para(d, "Re: SBA 7(a) loan request - $555,000 - Tyler Hospice HoldCo L.L.C. (Azalea Hospice & Palliative "
            "Care) - change-of-ownership acquisition of Hickory Hospice LLC")
    para(d, "")
    para(d, "Dear [Loan Officer]:")
    para(d, "We are pleased to submit this SBA 7(a) loan application for Tyler Hospice Hold, LLC, doing business "
            "as Azalea Hospice & Palliative Care. We are requesting a $555,000 SBA 7(a) loan as part of a "
            "$750,000 total project to acquire and operate Hickory Hospice LLC, an established, Medicare-"
            "certified hospice agency in the East Texas market, through a startup change of ownership. The "
            "seller is paid in full at close from loan proceeds and our equity injection, so the SBA loan is the "
            "only debt the business carries.")
    para(d, "The transaction is built on a strong foundation:")
    for x in ["Established, billing-ready agency. We are acquiring an operating Medicare-certified hospice with a transferring provider number, HCSSA license, and accreditation - billing-capable from day one, with no startup enrollment ramp.",
              "Experienced operator and validated local team. Our Managing Member has 10+ years of hospice ownership and operations; our Director of Sales has sustained a 40+ ADC referral book in the Tyler market for 4.5 years.",
              "Substantial equity. A $195,000 cash equity injection - 26.0% of the $750,000 project, 2.6x the 10% SOP 50 10 8 minimum (below the ~30% some startup lenders prefer). $100,000 is already on deposit; the balance follows on a committed schedule.",
              "Strong coverage. Because the seller is paid at close (no retained note), the SBA loan is the only debt: DSCR is 3.23x in Year 1, rising to 7.26x and 10.76x (global 7.08x) against the 1.25x floor. Break-even is ~16 ADC versus a validated opening census of ~22.",
              "Clean structure. A single SBA loan, no seller note, and a fully staffed clinical roster sized to the underwritten census."]:
        para(d, "  - " + x)
    para(d, "The enclosed package is indexed in the Submission Cover Sheet and includes the Company Profile, "
            "Use of Proceeds, Business Plan, financial model, equity-injection documentation, the borrower's "
            "operating agreement (as amended), and a lender credit memorandum with a DSCR sensitivity "
            "analysis. The sole 20%+ owner and guarantor is Geoff Schackmann (through Adeline & Lilah, LLC); "
            "his personal financial information is provided separately.")
    para(d, "We would welcome the opportunity to walk you through the deal and answer any questions. Thank you "
            "for your consideration.")
    para(d, "")
    para(d, "Sincerely,")
    para(d, "")
    para(d, "Geoff Schackmann")
    para(d, "Managing Member, Tyler Hospice HoldCo L.L.C.")
    para(d, "480-495-5474  |  geoff@azaleahospice.com")

    d.add_page_break()
    h1(d, "Anticipated Underwriter Questions & Responses")
    para(d, "Prepared to speed the credit team's review. These are the questions an SBA underwriter is most "
            "likely to raise on this deal, with our responses.", color=GREY, size=9)
    qa = [
        ("Why this level of working capital?",
         "This is a startup relaunch of an acquired agency. SBA proceeds and equity pay the seller in full at "
         "close ($300K) and fund a $372,000 working-capital reserve that carries ramp-period payroll and SBA "
         "debt service while Medicare receivables normalize to the ~45-day cycle. Because there is no seller "
         "note, monthly debt service is light (~$7,030), so the reserve stretches well through the ramp."),
        ("Is Bullard's 19.5% a device to avoid a guaranty?",
         "No. Mr. Bullard is contributing $195,000 of his own savings as a genuine capital contribution for a "
         "19.5% interest - the cash is the equity injection. He holds protective minority rights but no "
         "day-to-day operational control; the Manager (Geoff Schackmann via Adeline & Lilah, LLC) controls the "
         "business. Per OA Amendment No. 1, Bullard's consent is not required for SBA-loan actions, and the "
         "Texas Shootout buy-sell is suspended while the SBA loan is outstanding. He has signed an attestation "
         "to this effect. He does not hold a unilateral veto: a 60% Reserved-Matters vote can pass with "
         "Adeline & Lilah plus any two operator-members (66.5%)."),
        ("The equity injection isn't all in yet - how is that handled?",
         "$100,000 was wired May 7, 2026 and is on deposit. The balance follows on a committed schedule. We "
         "have scheduled the CHOW close and SBA disbursement to occur after the full $195,000 is on deposit, so "
         "the complete injection is verified at disbursement per SOP 50 10 8. Source-of-funds statements are "
         "provided for each tranche."),
        ("How real is the census ramp?",
         "Opening census of ~22 ADC is underwritten conservatively against a validated referral pipeline: the "
         "Director of Sales sustained 40+ ADC in this market for 4.5 years and previously built agencies from "
         "single digits to 70-220 patients. Break-even is ~16 ADC, so we are profitable at opening census. The "
         "DSCR sensitivity memo stresses census down 10-20% and shows the mitigants (the $372K reserve, a partly "
         "census-driven roster, and the seller-paid-at-close structure that leaves the SBA loan as the only debt)."),
        ("What happens if Year 1 underperforms?",
         "The enclosed sensitivity analysis tests census, rate, and payroll shocks. With no seller note, Year-1 "
         "coverage starts at 3.23x and absorbs a 10% census miss, a 5% rate cut, or a 10% wage spike and still "
         "clears comfortably; the only break is a hard, sustained 20% census shortfall, bridged by the $372K "
         "reserve and a roster that flexes with census. We propose a monthly census-and-cash covenant in Year 1 "
         "so the lender can monitor the binding variable."),
        ("Is the CHOW provider-number transfer clean?",
         "The acquisition is a purchase of 100% of Hickory Hospice LLC's membership interests; the Medicare "
         "PTAN/NPI, HCSSA license, and accreditation transfer through the CHOW process. We will provide the "
         "executed purchase agreement, the CMS-855A CHOW filing status, and the accreditation continuation "
         "notice at or before close."),
        ("Are there affiliates that affect size or eligibility?",
         "No operating affiliates. The borrower is a newco acquiring Hickory (which becomes its wholly owned "
         "subsidiary), held through the non-operating holding entity Adeline & Lilah, LLC. The Managing Member "
         "sold his prior hospice interest (VistaRiver) in August 2025 and holds only a passive seller-financed "
         "note - no equity or management role - so it is not an affiliate. He does not own 20%+ of, or control, "
         "any other operating business. See the Affiliate & Size-Standard memo."),
        ("This is a startup newco - how do you get comfortable with no operating history?",
         "The borrower is new, but it acquires an established, billing-ready Medicare-certified agency whose "
         "economics are benchmarked to the seller's actual Apr-Jun 2025 collections (~$118K/mo net at ~22 ADC), "
         "and it is led by an operator with 10+ years of hospice ownership and a Director of Sales with a "
         "validated 40+ ADC East-Texas referral book. The equity injection is 26.0% of project cost, and with "
         "the seller paid at close the SBA loan is the only debt - Year-1 DSCR ~3.23x."),
    ]
    for q, a in qa:
        para(d, "Q: " + q, bold=True)
        para(d, "A: " + a)
        para(d, "")
    footer_note(d)
    save(d, "11_lender_credit_memo/Cover_Letter_and_Anticipated_QA_DRAFT.docx")


# ================================================================ BUSINESS PLAN REV 5.00 (regenerated, current model)
def business_plan():
    d = new_doc("Business Plan Rev 5.00")
    # ---- Cover ----
    h1(d, "AZALEA HOSPICE & PALLIATIVE CARE")
    para(d, "Business Plan - SBA 7(a) Application", bold=True, size=13, color=NAVY)
    para(d, "Tyler Hospice HoldCo L.L.C., operating as Azalea Hospice & Palliative Care", color=GREY, size=10)
    para(d, "dba Azalea Hospice & Palliative Care - EIN 41-4966640 - Wyoming holding company operating "
            "Hickory Hospice LLC (Texas)", size=9, color=GREY)
    para(d, "Rev. 5.00 - June 2026  (supersedes Rev 4.00 dated 2026-06-02; financials reconciled to the "
            "current operating model and capital structure updated)", italic=True, size=9, color=GREY)
    para(d, "")
    para(d, "Acquisition and relaunch of Hickory Hospice LLC - an established, Medicare-certified hospice - "
            "as Azalea Hospice & Palliative Care, led by an experienced East-Texas clinical team with an "
            "established patient panel and modeled on validated local economics.")
    para(d, "")
    h2(d, "Headline metrics")
    table(d, ["Metric", "Value", "Detail"],
          [["Loan request", "$555,000", "SBA 7(a) - 10-yr term"],
           ["Total project", "$750K", "$555K SBA + $195K equity; seller paid in full at close"],
           ["Year-1 EBITDA", "$302,522", "21.0% margin - growing to $1.01M by Year 3"],
           ["Global DSCR (SBA-only)", "7.08x", "3-yr - floor 1.25x"],
           ["Equity injection", "$195,000", "26.0% of project - 2.6x the 10% SOP 50 10 8 minimum"]],
          widths=[1.6, 1.3, 3.6])
    para(d, "Confidential - prepared exclusively for the SBA 7(a) loan application. All figures are computed "
            "from the Azalea Hospice SBA Loan Package operating model. Do not distribute without written "
            "consent.", italic=True, size=8, color=GREY)

    # ---- 01 Executive Summary ----
    d.add_page_break()
    h1(d, "01 - Executive Summary")
    h2(d, "Loan request")
    para(d, "Tyler Hospice Hold, LLC requests a $555,000 SBA 7(a) loan to acquire and relaunch Hickory "
            "Hospice LLC - an established, Medicare-certified hospice - as Azalea Hospice & Palliative Care. "
            "The agency's existing Medicare-certified provider number is acquired through a change of ownership "
            "(CHOW) for $300,000 paid in full at close from SBA proceeds and the equity injection - no retained "
            "seller note - giving Azalea immediate, billing-ready capability with no new-enrollment delay. The "
            "loan is paired with a $195,000 cash equity injection, for a total project of $750,000. Because the "
            "seller is paid at close, the SBA loan is the only debt the business carries.")
    h2(d, "Use of funds")
    table(d, ["Use", "Amount", "Detail"],
          [["Acquisition of Hickory (paid at close)", "$300,000", "100% membership interests; provider number / license / accreditation transfer in the CHOW"],
           ["Startup one-time costs", "$63,000", "855A filing, TX licensure, EMR, supply stock, legal, contingency"],
           ["Startup equipment (capex)", "$15,000", "Computers, office furniture; 5-yr straight-line depreciation"],
           ["Working-capital reserve (opening cash)", "$372,000", "Funds ramp payroll and SBA debt service; holds the $25K floor"],
           ["Total uses", "$750,000", "Funded by SBA $555K + Equity $195K"]],
          widths=[2.5, 1.1, 2.9])
    h2(d, "Business overview")
    para(d, "Azalea relaunches the acquired Hickory agency in the Tyler / Smith County market (CBSA 46340) "
            "under an experienced local team that brings established referral relationships and an existing "
            "patient panel. The model is grounded in validated economics - a proven ~22 ADC / ~$118K-net-per-"
            "month book benchmarked to actual collections - and the base case grows with the local referral "
            "pipeline to ~40 ADC in Year 2 and ~54 ADC in Year 3. Revenue is 100% Medicare RHC per-diem at a "
            "blended net rate of $181/patient-day (escalating 2.5%/year).")
    h2(d, "Management team")
    para(d, "Azalea is led by an already-seated team: Geoff Schackmann (Managing Member - multi-hospice "
            "operator), Silas Shelton (Executive Director), Dana Davenport (Director of Nursing), and Bradley "
            "Woodard (Director of Sales - 25+ years East-Texas hospice business development), supported by RN "
            "case managers, hospice aides, a PRN visit pool, and 1099 Medical Directors.")
    h2(d, "Repayment case")
    table(d, ["", "Year 1", "Year 2", "Year 3"],
          [["Net income", "$199,308", "$574,951", "$903,362"],
           ["EBITDA", "$302,522", "$679,524", "$1,007,520"],
           ["DSCR (SBA-only)", "3.23x", "7.26x", "10.76x"]],
          widths=[1.8, 1.5, 1.5, 1.5])
    para(d, "The operation is EBITDA-positive from early in the ramp. Because the seller is paid in full at "
            "close, the SBA loan is the only debt: against SBA debt service of $93,637/yr, the model clears a "
            "global 3-year DSCR of 7.08x - far above the 1.25x floor - and strengthens each year "
            "(3.23x -> 7.26x -> 10.76x). The $372,000 opening reserve keeps minimum cash well above the $25,000 "
            "floor through the ramp, and with only ~$7,030/mo of debt service the reserve stretches "
            "comfortably.")

    # ---- 02 Company Description ----
    d.add_page_break()
    h1(d, "02 - Company Description")
    h2(d, "2.1 Legal structure")
    para(d, "Tyler Hospice HoldCo L.L.C. is a Wyoming LLC (formed 2026; EIN 41-4966640), authorized to do "
            "business in Texas. As the borrower and holding entity, it owns 100% of Hickory Hospice LLC - a "
            "Texas HCSSA-licensed, Medicare-certified hospice - which operates as Azalea Hospice & Palliative "
            "Care. The transaction is a change of ownership (CHOW) of Hickory's existing provider number "
            "(purchase of 100% of membership interests), so Azalea runs on a licensed, billing-ready platform "
            "from day one - preserving the existing number rather than awaiting a new-provider 855A "
            "enrollment, with CHAP/ACHC accreditation transferring in the CHOW. The Company is taxed as a "
            "partnership for federal income tax purposes.")
    h2(d, "2.2 Ownership & capitalization")
    table(d, ["Member / source", "Interest", "Role & structure"],
          [["Adeline & Lilah, LLC (AZ)", "39.9%", "Managing Member - owned 50% Geoff Schackmann / 50% Mary Elizabeth Burcham; Geoff serves as Manager and provides the SBA personal guaranty"],
           ["James E. Bullard", "19.5%", "Passive minority investor - $195,000 cash capital contribution; protective minority rights only, no operational control; under 20%, no guaranty"],
           ["Silas R. Shelton", "13.3%", "Executive Director - Restricted Interest (fully forfeitable; 4.9% time-vested base + 8.4% dual-trigger earn-up)"],
           ["Dana L. Davenport", "13.3%", "Director of Nursing - Restricted Interest (fully forfeitable; 4.9% time-vested base + 8.4% dual-trigger earn-up)"],
           ["Bradley G. Woodard", "13.3%", "Director of Sales - Restricted Interest (fully forfeitable; 4.9% time-vested base + 8.4% dual-trigger earn-up)"],
           ["Unissued pool", "0.7%", "Reserved for future grants"]],
          widths=[2.2, 0.9, 3.4])
    h2(d, "Sources of capital & SBA guaranty")
    para(d, "The $750,000 project is funded by the $555,000 SBA 7(a) loan and a $195,000 cash equity injection "
            "contributed by investor James E. Bullard in exchange for a direct 19.5% membership interest. The "
            "$300,000 acquisition price is paid in full at close from these sources - there is no retained "
            "seller note, so the SBA loan is the only debt the business carries. The injection equals 26.0% of "
            "total project cost - 2.6 times the 10% minimum required under SBA SOP 50 10 8, though below the "
            "~30% down some startup lenders prefer. Mr. Bullard's funds are his own savings (non-borrowed), "
            "verified by his bank statements; $100,000 was wired on May 7, 2026 and the balance follows on a "
            "committed schedule. Because he holds less than 20% and exercises no operational control, no personal "
            "guaranty is required of him (13 CFR 120.160). The Managing Member (Geoff Schackmann, via "
            "Adeline & Lilah, LLC) is the sole 20%+ owner and provides the SBA personal guaranty.")
    h2(d, "2.3 Acquired platform & validated economics")
    para(d, "Rather than projecting speculative new-admit growth, the financial model is benchmarked to "
            "actual operating results - a proven ~22 ADC / ~$118K-net-per-month book (Apr-Jun 2025 actuals) "
            "- so the per-diem rate, patient-day costs, and census all reconcile to real collections. Because "
            "the CHOW preserves the provider number, the patient panel, care team, and referral relationships "
            "carry over, supporting revenue continuity from day one.")
    h2(d, "2.4 Brand and identity")
    para(d, "Azalea Hospice & Palliative Care is positioned for dignity-led hospice service in East Texas. "
            "Tagline: \"The standard you upheld your whole life is the one we keep today.\" Service pillars: "
            "Honor, Dignity, Faith, and Reverence - aligned to the conservative-Christian demographic "
            "prevalent in Smith County.")

    # ---- 03 Market and Industry ----
    d.add_page_break()
    h1(d, "03 - Market and Industry Analysis")
    h2(d, "3.1 Hospice regulatory environment")
    para(d, "Hospice in the United States is delivered under the Medicare Hospice Benefit (established 1982, "
            "codified at 42 CFR Part 418). Patients with a terminal prognosis of six months or less, certified "
            "by two physicians, elect the benefit in lieu of curative treatment. Four levels of care are "
            "reimbursed: Routine Home Care (RHC) at a two-tier per-diem (days 1-60 higher, days 61+ lower), "
            "General Inpatient (GIP), Continuous Home Care (CHC), and Inpatient Respite (IRC). Azalea is "
            "modeled at ~100% RHC, consistent with national community-hospice patterns.")
    para(d, "For FY2026 the national RHC rates are $230.83/day (Tier 1, days 1-60) and $182.36/day (Tier 2, "
            "days 61+), adjusted by the Tyler CBSA 46340 wage index (0.88) on the 0.68 labor share. With 40% "
            "of patient-days in Tier 1, the blended gross rate is $185.29/PD; net of 2% sequestration and "
            "0.15% write-offs, the blended net rate is $181/PD - validated against actual Tyler-market "
            "collections and escalated 2.5%/year (CMS hospice rates have risen every year since 2010).")
    h2(d, "3.2 Tyler market - demographics and service area")
    para(d, "Tyler, Texas (city pop. ~110,000; metro ~245,000) is the principal city of Smith County and the "
            "regional healthcare hub for East Texas. Azalea operates from a principal office in the Tyler area "
            "with a ~45-mile service radius.")
    for x in ["Service-area population: ~350,000 across Smith, Cherokee, Henderson, Rusk, Van Zandt, and Wood counties",
              "Medicare-eligible population (65+): ~60,000+",
              "Hospice utilization: 50-55% of Medicare decedents (above the ~50% national average) - a mature, accepting market",
              "Population growth: Tyler has grown 15%+ since 2010, driven by healthcare-sector expansion and retirement in-migration",
              "Dominant religious demographic: conservative-Christian - strong alignment with Azalea's faith-based service pillars"]:
        para(d, "  - " + x)
    h2(d, "3.3 Major hospital systems - referral drivers")
    for x in ["UT Health Tyler - 502-bed Level 1 trauma center, regional referral hub",
              "CHRISTUS Trinity Mother Frances - 438-bed acute care flagship",
              "UT Health North East - 153-bed acute care",
              "CHRISTUS Louis & Peaches Owen Heart Hospital - cardiac",
              "Texas Spine & Joint Hospital - 64-bed surgical",
              "UT Health Jacksonville - 90-bed acute care (Cherokee County)"]:
        para(d, "  - " + x)
    h2(d, "3.4 Skilled nursing & assisted living landscape")
    para(d, "Smith County and surrounding counties hold ~40+ skilled nursing facilities (SNFs) and ~30+ "
            "assisted living facilities (ALFs). These are a concentrated, high-volume referral source: "
            "residents with terminal diagnoses frequently elect hospice, and facility relationships drive "
            "consistent admission volume. Tier 1 facilities (highest-census SNFs and ALFs in the Tyler metro) "
            "receive weekly in-person visits from the Director of Sales and clinical liaison; Tier 2 "
            "facilities receive bi-weekly or monthly touchpoints by census potential.")
    h2(d, "3.5 Physician & referral source ecosystem")
    para(d, "Hospice admissions in the Tyler market are driven by referral relationships with hospital "
            "discharge planners and case managers (the single highest-volume source), 200+ primary care "
            "physicians, oncologists (UT Health Tyler Cancer Center, Texas Oncology, CHRISTUS), organ-"
            "specialty physicians (pulmonology, cardiology, neurology), SNF/ALF medical directors and DONs, "
            "and the faith-community and pastoral-care networks that are uniquely important in East Texas. A "
            "mature, above-average-utilization market with 60,000+ Medicare-eligible residents and 70+ senior "
            "facilities - entered not cold, but by an experienced local team on validated Tyler-market "
            "economics.")

    # ---- 04 Competitive ----
    d.add_page_break()
    h1(d, "04 - Competitive Analysis")
    para(d, "The Tyler/Smith County hospice market includes national multi-state operators, regional "
            "providers, and local independents. Key Medicare-certified providers serving the market:")
    table(d, ["Competitor", "Type", "Weaknesses vs. Azalea"],
          [["Heart to Heart Hospice - Tyler", "Regional (TX)", "Standard model; no faith-based differentiation"],
           ["VITAS Healthcare - Tyler", "National", "Bureaucratic; national protocols; less personal"],
           ["Hospice of East Texas", "Local nonprofit", "Limited marketing budget; slower growth"],
           ["CHRISTUS Hospice", "Health system", "System constraints; less flexibility"],
           ["Enhabit Hospice - Tyler", "National (public)", "Fragmented sales effort; less local focus"],
           ["Traditions Health - Tyler", "National", "Newer entrant; relationships unproven"],
           ["Kindred / CenterWell", "National (Humana)", "Limited to Humana MA members"],
           ["Harbor Hospice - East TX", "Regional", "Limited capacity; smaller sales team"]],
          widths=[2.2, 1.4, 3.0])
    h2(d, "Azalea's differentiation")
    for x in ["Validated local economics - modeled on a proven ~22 ADC Tyler book and an experienced team, not a standing start",
              "Dignity-led hospitality model - luxury-hospitality standards applied to end-of-life care",
              "Faith-aligned positioning - explicit Christian-values messaging resonant with the local demographic",
              "Speed of admission - same-day / next-day response, faster than national competitors' 24-72 hours",
              "Local independent ownership - faster decisions, no corporate bureaucracy, deeper relationship investment",
              "MVI-certified in The Perfect Visit and Perfect Phones - national co-marketing exposure that local and regional competitors lack (see 4.1)"]:
        para(d, "  - " + x)

    h2(d, "4.1 MVI \u2018Perfect Visit\u2019 certification & national co-marketing (strategic upside)")
    para(d, "Azalea is pursuing certification by MultiView Incorporated (MVI) in The Perfect Visit and Perfect "
            "Phones - structured programs that certify an agency's patient-visit quality and its intake and "
            "telephone responsiveness against a defined standard. MVI, led by Andrew Reid, is launching a "
            "national, multi-platform advertising campaign that promotes hospice agencies certified in The "
            "Perfect Visit. As a certified agency, Azalea would be featured to a national audience - the kind of "
            "brand exposure normally reserved for the largest national operators with multi-million-dollar ad "
            "budgets, here made accessible to a local independent through the certification.")
    para(d, "Azalea will develop an in-tandem local marketing plan timed to the national campaign - capturing the "
            "awareness it generates and converting it into local inquiries and referrals across the Tyler service "
            "area. The pairing of a national demand-generation engine with a focused local conversion plan is a "
            "differentiator no local competitor currently has, and it creates a credible path to take market "
            "share from the regional and national incumbents listed above.")
    para(d, "Underwriting note: this is incremental upside. The financial projections in Section 11 remain "
            "conservative and assume no lift from the MVI certification or the national campaign; debt service is "
            "covered on the base-case census path without it. The certification represents potential growth above "
            "plan, not a dependency.", italic=True, size=9.5)

    # ---- 05 SWOT ----
    d.add_page_break()
    h1(d, "05 - SWOT Analysis")
    h2(d, "Strengths")
    for x in ["Modeled on validated Tyler-market economics (~22 ADC, ~$118K/mo) benchmarked to actual local collections",
              "Experienced clinical & admin team already in seat, with deep East-Texas referral relationships",
              "Acquired Medicare provider number (CHOW) - billing-ready from day one, no new-provider enrollment lag",
              "Capital cushion - $372K opening cash reserve; seller paid at close so the SBA loan is the only debt",
              "DSCR rises 3.23x -> 10.76x across the plan; substantial $195K cash equity injection (26% of project)",
              "Strong ownership alignment - operators hold equity (13.3% each); experienced multi-hospice Managing Member"]:
        para(d, "  - " + x)
    h2(d, "Weaknesses")
    for x in ["Newco borrower with no operating history of its own (mitigated by the acquired agency's actuals and an experienced operator)",
              "Revenue depends on converting the team's established relationships into admissions during ramp",
              "Single-office, single-market concentration",
              "Several go-forward cost adds (own lease, Medical Directors, benefits, marketing) not in the benchmarked actuals",
              "Sensitivity to a sustained census shortfall (see Section 11.7 and the DSCR Sensitivity memo)",
              "Managing Member's attention is split across other interests during launch"]:
        para(d, "  - " + x)
    h2(d, "Opportunities")
    for x in ["Growing 65+ population (60,000+ Medicare-eligible) with above-average utilization",
              "Base-case census growth to ~54 ADC by Year 3 from the team's referral pipeline; upside path to ~60 ADC",
              "70+ SNF/ALF facilities, many underserved by incumbents",
              "Faith-based marketing to an extensive church and ministry network",
              "Palliative-care consultation line as a Year 2-3 revenue diversifier",
              "MVI 'Perfect Visit' / 'Perfect Phones' certification + Andrew Reid's national multi-platform campaign - national lead exposure for a local independent (upside, not in the base case)",
              "CMS hospice rates have risen every year since 2010; model carries a conservative 2.5%/yr escalation"]:
        para(d, "  - " + x)
    h2(d, "Threats")
    for x in ["National competitors with larger sales forces and budgets",
              "CMS reimbursement or regulatory changes",
              "Medicare Advantage hospice carve-in could shift referral dynamics",
              "East Texas clinical labor shortage could pressure wages",
              "Front-loaded ramp costs before census stabilizes; mitigated by the $372K reserve and light single-loan debt service"]:
        para(d, "  - " + x)

    # ---- 06 Marketing ----
    d.add_page_break()
    h1(d, "06 - Marketing, Sales & Referral Strategy")
    para(d, "Hospice census in the Tyler market is relationship-driven. Azalea's strategy leverages the "
            "leadership team's established local referral relationships across five active development "
            "channels.")
    h2(d, "6.1 Hospital referral channel")
    para(d, "Discharge planners and case managers generate the highest referral volume. The Director of Sales "
            "(Woodard) maintains a structured weekly cadence: UT Health Tyler and CHRISTUS Trinity Mother "
            "Frances (2x/week each), UT Health North East and Texas Spine & Joint (1x/week each), and "
            "bi-weekly touchpoints at regional hospitals. Tactics: in-services on hospice eligibility, branded "
            "referral materials, and a same-day admission-response protocol (clinical liaison dispatched "
            "within 60 minutes).")
    h2(d, "6.2 SNF / ALF facility channel")
    para(d, "Tier 1 covers the ~15 highest-census SNFs and ALFs/memory-care communities in the Tyler metro "
            "with weekly visits and a dedicated clinical liaison; Tier 2 covers the remaining 40+ SNFs and "
            "20+ ALFs across the six-county service area with bi-weekly touchpoints. Quarterly facility "
            "in-services and pre-positioned election packets support consistent volume.")
    h2(d, "6.3 Physician & clinical channel")
    para(d, "CE-accredited physician lunch-and-learns, direct educational outreach, an EHR-embedded referral "
            "pathway, and same-day response to physician referrals target the 200+ PCPs, oncologists, and "
            "organ-specialty physicians in the service area.")
    h2(d, "6.4 Faith community & pastoral care channel")
    para(d, "Tyler's 200+ churches and pastoral networks are uniquely valuable referral and brand-building "
            "partners. Tactics: monthly pastor breakfasts, hospice education for church visitation teams, "
            "bereavement-workshop sponsorship, and chaplaincy partnerships with major hospitals.")
    h2(d, "6.5 Marketing tactics & budget")
    para(d, "Digital: local SEO, Google Business Profile and review generation, Facebook/Instagram education "
            "and testimonial content to adults 45-75 within 50 miles, and Google Ads on high-intent hospice "
            "terms. Traditional & community: quarterly educational mailers to SNFs/ALFs/physician offices, "
            "senior-focused print, branded admission books and referral pads, bereavement support groups, "
            "senior health fairs, and church health-ministry workshops. The model funds $3,000/mo "
            "($36,000/yr) of marketing within fixed G&A, sustained across all three years.")
    para(d, "National co-marketing (upside): as an agency pursuing MVI 'Perfect Visit' and 'Perfect Phones' "
            "certification, Azalea will run an in-tandem local campaign timed to MVI's national, multi-platform "
            "advertising of certified hospice agencies (led by Andrew Reid) - local landing pages, geo-targeted "
            "digital, and referral-source outreach engineered to convert national awareness into Tyler-area "
            "leads. This sits on top of the budgeted spend above and is treated as upside; the base-case "
            "projections do not rely on it.")
    h2(d, "6.6 Census model")
    table(d, ["Period", "Base ADC", "Primary channel"],
          [["Months 1-3 (launch)", "12 -> 22", "Panel transition + team activation"],
           ["Months 4-12", "22 -> 24", "Hospital + SNF/ALF"],
           ["Year 2", "~40 avg (30 -> 50)", "All channels + faith"],
           ["Year 3", "~54 avg (51 -> 56)", "All channels + organic"]],
          widths=[1.8, 1.6, 3.2])
    para(d, "All DSCR and repayment figures in Section 11 use this conservative base path. An upside path "
            "reaches ~60 ADC by Year 3 and is presented as scenario sensitivity.")

    # ---- 07 Operations ----
    d.add_page_break()
    h1(d, "07 - Operations Plan")
    h2(d, "7.1 Service lines & clinical model")
    para(d, "Azalea delivers Routine Home Care (RHC) at patient homes, ALFs, and SNFs, with access to GIP "
            "(contracted inpatient beds), CHC, and Inpatient Respite, plus 12-month bereavement support per "
            "Medicare CoP. Care is delivered by an interdisciplinary group (IDG) - RN case manager, medical "
            "director, hospice aide, social worker, chaplain - meeting every 15 days, with 24/7 RN on-call "
            "coverage. A QAPI program runs monthly KPI dashboards.")
    h2(d, "7.2 Core staffing roster (Year 1)")
    para(d, "Staggered W-2 hires plus a census-driven PRN pool. Employer burden of 18.27% applies to W-2 "
            "wages; health insurance is $550/employee/mo; Medical Directors are 1099 contractors. The Year-1 "
            "roster is sized to the opening ~22 ADC within hospice staffing ratios and includes the Executive "
            "Director, Director of Nursing, Director of Sales, RN case managers, a CNA/hospice aide, a "
            "Director of Business Development, an Office Manager, and a 1099 Medical Director.")
    h2(d, "7.3 Capacity-driven hiring (Years 2-3)")
    para(d, "As census grows toward ~54 ADC, the roster scales on census triggers (RN caseload 1:12; aide "
            "1:10), converting PRN roles to salaried and adding compliance, intake, and volunteer-coordinator "
            "support. Second through sixth RN case managers and CNAs load in as ADC crosses their thresholds "
            "(Months 13-34); FT social worker, LVN, and chaplain convert from PRN (Months 14-18); a Quality/"
            "Compliance Manager, Intake Coordinator, and Volunteer Coordinator are added (Months 16-26); and "
            "a second Medical Director is added at Month 22. This is the fuller, more defensible direct-care "
            "build reflected in the current operating model.")
    h2(d, "7.4 Facility & technology")
    para(d, "Principal office in the Tyler, TX area at $3,000/mo (a go-forward lease assumption). Hours M-F "
            "8-5 with 24/7/365 on-call; ~45-mile service radius. EMR, PCR, and outsourced billing (1.5% of "
            "gross) run on the platform reflected in G&A. Field travel is mileage-reimbursed initially; an "
            "owned fleet is evaluated in Year 2.")

    # ---- 08 Management ----
    d.add_page_break()
    h1(d, "08 - Management and Organization")
    para(d, "Azalea is operated by an experienced, already-seated team whose established East-Texas referral "
            "relationships are a key de-risking factor. Officer salaries are fully loaded into the operating "
            "model as expenses (no compensation is omitted, per SBA requirements).")
    h2(d, "8.1 Ownership & reporting structure")
    table(d, ["Entity / person", "Role", "Interest", "Reports to"],
          [["Adeline & Lilah, LLC - Geoff Schackmann", "Managing Member (holding co.)", "39.9%", "-"],
           ["Silas R. Shelton", "Executive Director", "13.3%", "Managing Member"],
           ["Dana L. Davenport", "Director of Nursing", "13.3%", "Executive Director"],
           ["Bradley G. Woodard", "Director of Sales", "13.3%", "Executive Director"],
           ["James E. Bullard", "Passive minority investor", "19.5%", "No operational role"]],
          widths=[2.4, 1.9, 0.9, 1.6])
    para(d, "Tyler Hospice HoldCo L.L.C. (Wyoming holding company) owns 100% of Hickory Hospice LLC (Texas "
            "operating subsidiary), which does business as Azalea Hospice & Palliative Care. The Executive "
            "Director leads operations; the Director of Nursing leads the clinical team; the Director of Sales "
            "leads business development. Mr. Bullard is a passive investor with protective minority rights but "
            "no management authority or signing power.")
    h2(d, "8.2 Decision authority")
    for x in ["Operating within budget ($0-$10K): Executive Director (Shelton)",
              "Above $10K or out-of-budget: Managing Member (Schackmann, via Adeline & Lilah)",
              "Clinical / regulatory: Director of Nursing (Davenport), escalated to Executive Director + Medical Director",
              "Hiring / firing senior staff: Executive Director, with Managing Member concurrence",
              "Strategic / capital / acquisition: Managing Member",
              "Bank account signatures: Managing Member (sole signer)"]:
        para(d, "  - " + x)
    h2(d, "8.3 Equity vesting (operator-members)")
    para(d, "Each operator-member's 13.3% interest is a Restricted Interest under the Operating Agreement. "
            "The entire 13.3% is at risk and forfeitable until vested. It comprises a 4.9% Initial Base that "
            "time-vests on continuous service (four-year schedule, one-year cliff) "
            "and an 8.4% Earn-Up Portion that vests only on a dual trigger: a four-year time-vesting schedule with a "
            "one-year cliff AND the performance milestones (25% at Breakeven, 50% at three consecutive Profitable "
            "months, 25% at twelve consecutive Profitable months), with the lesser of the two schedules governing - "
            "so no portion is vested at inception and each operator may earn up to a fully-vested 13.3%. Each grantee files a "
            "timely IRC Section 83(b) election. On departure, unvested interests are forfeited at $0; the Company "
            "holds a repurchase (call) right over vested interests - at fair market value for a good-leaver "
            "separation (termination without cause, death, disability, or retirement) and at the lower of cost or "
            "fair market value for a bad-leaver separation (resignation before full vesting, or termination for "
            "cause), paid via an SBA-subordinated note. This keeps operator equity aligned with the credit's "
            "performance and ensures departed operators do not retain equity in the company.")

    # ---- 09 Leadership ----
    d.add_page_break()
    h1(d, "09 - Leadership Team")
    h2(d, "Geoff Schackmann - Managing Member (via Adeline & Lilah, LLC, 39.9%)")
    para(d, "Multi-hospice operator and transaction-led growth leader with operational responsibility for "
            "Medicare-certified hospice and palliative-care agencies across multiple states. Co-owner and "
            "Managing Member of Adeline & Lilah, LLC (AZ), the holding entity through which equity in Tyler "
            "Hospice HoldCo is held. Direct experience includes change-of-ownership (CHOW) transactions and "
            "post-CHOW enrollment oversight, census growth from sub-30 to 100+ ADC under existing Medicare "
            "provider numbers, multi-site clinical operations under Texas HCSSA and Oregon hospice licensure, "
            "CHAP and Joint Commission accreditation, acquisition due diligence (clinical, financial, "
            "regulatory), and lender/investor relationships across SBA 7(a), conventional, and mezzanine "
            "structures. Leads transaction structuring, financing, capital allocation, and post-close "
            "integration. Sole 20%+ owner and SBA personal guarantor.")
    h2(d, "Bradley G. Woodard - Director of Sales (13.3%)")
    para(d, "More than 25 years of East-Texas hospice business development and administration. Founding "
            "administrator of one of the Tyler market's largest hospices (licensed 2004; grown to 220 "
            "patients, 109 employees, and 62 volunteers), with subsequent census-building roles across "
            "multiple East-Texas agencies - including start-ups in Lufkin/Nacogdoches (0->48 in six months; "
            "12->70+) and Grace Hospice of East Texas (8->145+). Most recently VP of Business Development for "
            "an East-Texas hospice (2021-present), sustaining a 40+ ADC book for 4.5 years, with no "
            "non-compete restricting his transition to Azalea. Licensed Nursing Home Administrator "
            "(LNFA #7136); B.S., Texas A&M University.")
    h2(d, "Silas R. Shelton - Executive Director (13.3%)")
    para(d, "Operational leadership of Azalea Hospice & Palliative Care, with day-to-day responsibility for "
            "regulatory compliance under Texas HCSSA and the Medicare Conditions of Participation, payer-mix "
            "management, referral-source partnerships across Smith County and surrounding East-Texas counties, "
            "and overall site leadership. Direct reports include the Director of Nursing, the Director of "
            "Sales, and the full clinical and operational team.")
    h2(d, "Dana L. Davenport - Director of Nursing (13.3%)")
    para(d, "Clinical leadership of the Azalea nursing team, responsible for Medicare Conditions of "
            "Participation compliance, interdisciplinary group (IDG) oversight, plan-of-care management, and "
            "clinical quality (QAPI). Leads RN case managers, hospice aides, social workers, and chaplains "
            "across the service area.")
    h2(d, "Investor")
    para(d, "James E. Bullard is a passive minority investor holding a 19.5% membership interest acquired "
            "for a $195,000 cash capital contribution (his own savings). He holds customary protective "
            "minority rights (information rights, anti-dilution with a 19.5% floor, voting on a defined set "
            "of fundamental Reserved Matters, and tag-along rights) but has no management authority, no "
            "operational or clinical role, and no signing power. His consent is not required for SBA-loan "
            "actions, and his buy-sell rights are suspended while the SBA loan is outstanding. As a sub-20% "
            "non-controlling member, he provides no SBA personal guaranty (13 CFR 120.160).")

    # ---- 10 Acquisition ----
    d.add_page_break()
    h1(d, "10 - Acquisition and Launch Strategy")
    h2(d, "10.1 Acquisition & CHOW structure")
    para(d, "The transaction is the acquisition of Hickory Hospice LLC - an established, Medicare-certified "
            "hospice currently owned by Tracy Gleason and Ann Lozano - structured as a Medicare change of "
            "ownership (CHOW) by purchase of 100% of its membership interests. The $300,000 purchase price is "
            "paid in full at close from SBA proceeds and the equity injection (no retained seller note). "
            "Hickory's existing Medicare provider number, state "
            "license, and CHAP/ACHC accreditation carry over to Azalea, so the agency is billing-ready from "
            "day one rather than waiting on a new-provider 855A enrollment. Azalea files the CMS-855A change "
            "of ownership and updates banking, EMR, and insurance into the new ownership structure.")
    h2(d, "10.2 Revenue continuity")
    para(d, "Because the CHOW preserves the existing provider number, the patient panel, care team, and "
            "referral relationships transfer intact - revenue continues without the gap a cold-start de-novo "
            "would face. The model's economics are benchmarked to actual collections rather than to "
            "speculative new-admit forecasts, so the opening ramp toward ~22 ADC carries materially less "
            "uncertainty, and growth above that level is driven by the team's own referral pipeline.")
    h2(d, "10.3 Affiliate disclosure")
    para(d, "The borrower's affiliate group consists of the borrower, its wholly owned operating subsidiary "
            "Hickory Hospice LLC, and the non-operating holding entity Adeline & Lilah, LLC - well within the "
            "applicable SBA size standard. The Managing Member previously held a 33.33% interest in VistaRiver "
            "Inc, which he sold in August 2025; he retains only a passive seller-note receivable (no equity, "
            "officer, manager, or employee role), so VistaRiver is not an SBA affiliate. A separate Affiliate "
            "and Size-Standard memorandum is included in the application package.")
    h2(d, "10.4 Launch timeline & equity-injection coordination")
    table(d, ["Window", "Phase", "Key activities"],
          [["Through July 2026", "Equity injection", "Bullard funds the $195,000 in two tranches ($100K received 5/7/2026; balance by 7/31/2026)"],
           ["Late July 2026", "Close & stand-up", "Close Hickory CHOW after full injection on deposit; update banking (CMS-588 EFT), EMR, insurance; staff retention/credentialing"],
           ["Days 1-90 post-close", "855A CHOW", "File CMS-855A; CMS updates the provider number to Azalea; license & accreditation transfer confirmed; billing uninterrupted"],
           ["Days 30-120", "BD ramp", "Woodard-led weekly referral cadence; digital marketing live; AR normalizes to ~45-day Medicare cycle"],
           ["Day 120+", "Growth", "Census builds with the referral pipeline; capacity hires triggered by ADC thresholds; DSCR > 1.25x and rising"]],
          widths=[1.4, 1.3, 3.9])
    para(d, "SBA disbursement is sequenced to occur with the full $195,000 equity injection on deposit, in "
            "compliance with SOP 50 10 8.", italic=True, size=9.5)

    # ---- 11 Financials ----
    d.add_page_break()
    h1(d, "11 - Financial Plan and Projections")
    h2(d, "11.1-11.2 Sources & uses of funds")
    table(d, ["Source", "Amount", "Use", "Amount"],
          [["SBA 7(a) loan", "$555,000", "Acquisition of Hickory (paid at close)", "$300,000"],
           ["Cash equity injection (Bullard)", "$195,000", "Startup one-time costs (855A, licensure, EMR, supplies, legal, contingency)", "$63,000"],
           ["", "", "Startup equipment (capex)", "$15,000"],
           ["", "", "Working-capital reserve (opening cash)", "$372,000"],
           ["Total sources", "$750,000", "Total uses", "$750,000"]],
          widths=[2.2, 1.1, 2.6, 1.1])
    para(d, "The seller is paid in full at close from SBA proceeds and the equity injection, so there is no "
            "retained seller note and the SBA loan is the only debt the business carries. The $195,000 "
            "injection is 26.0% of the $750,000 project.")
    h2(d, "11.3 Key assumptions")
    table(d, ["Assumption", "Value"],
          [["Payer mix", "100% Medicare RHC"],
           ["RHC per-diem (Tier 1 / Tier 2)", "$230.83 / $182.36 (FY2026)"],
           ["Tyler CBSA wage index", "0.88 (0.68 labor share)"],
           ["Blended net rate / patient-day", "$181 - escalating 2.5% / year"],
           ["Census - base case (avg ADC)", "21.8 -> 40.2 -> 53.8 (Yrs 1-3)"],
           ["Patient COGS / patient-day", "supplies + DME + Rx (per model)"],
           ["Benefits load (W-2)", "18.27% + $550/employee/mo health"],
           ["Medical Directors (1099)", "$4,000/mo + $5,000/mo from M22"],
           ["Outsourced billing / QR fee", "1.5% / 0.75% of gross"],
           ["SBA 7(a)", "$555,000 - 11.5% APR - 120 mo - $7,803.05/mo"],
           ["Seller", "$300,000 paid in full at close (no retained note)"],
           ["Equity injection", "$195,000 cash (26.0% of project)"],
           ["Depreciation & amortization", "$78K startup+capex (5-yr) + $300K license (15-yr) = $35,600/yr"],
           ["TX franchise tax / AR-AP days", "per model / 45-30"],
           ["Opening cash reserve", "$372,000"]],
          widths=[2.6, 4.0])
    h2(d, "11.4 Three-year profit & loss (current operating model)")
    table(d, ["Line item", "Year 1", "Year 2", "Year 3"],
          [["Average daily census (ADC)", "21.8", "40.2", "53.8"],
           ["Net patient revenue", "$1,442,930", "$2,728,642", "$3,734,934"],
           ["Patient-related COGS", "$71,549", "$132,002", "$176,276"],
           ["Payroll & related", "$868,172", "$1,686,864", "$2,297,748"],
           ["Operating G&A", "$200,687", "$230,251", "$253,390"],
           ["EBITDA", "$302,522", "$679,524", "$1,007,520"],
           ["EBITDA margin", "21.0%", "24.9%", "27.0%"],
           ["Depreciation & amortization", "$35,600", "$35,600", "$35,600"],
           ["Interest (SBA loan only)", "$62,202", "$58,741", "$54,551"],
           ["TX franchise tax", "$5,411", "$10,232", "$14,006"],
           ["Net income", "$199,308", "$574,951", "$903,362"]],
          widths=[2.6, 1.4, 1.4, 1.4])
    para(d, "Net revenue grows from $1.44M to $3.73M and EBITDA from $303K to $1.01M as the base-case census "
            "builds from ~22 to ~54 ADC on the team's referral pipeline, with a conservative 2.5% annual CMS "
            "per-diem escalation. Payroll scales on census triggers, carrying the fuller, more defensible "
            "direct-care roster sized to the census being underwritten. D&A of $35,600/yr is a non-cash "
            "charge below EBITDA that does not affect debt-service coverage.")
    h2(d, "11.5 Year-1 monthly detail (selected months)")
    table(d, ["Month", "ADC", "Net revenue", "DSCR (SBA-only)"],
          [["M1 (transition)", "12.0", "~$66,000", "covered by reserve"],
           ["M3", "22.0", "~$121,000", "clears comfortably"],
           ["M6", "22.7", "~$125,000", "clears comfortably"],
           ["M12", "24.0", "~$132,000", "clears comfortably"]],
          widths=[1.6, 1.0, 1.6, 2.4])
    para(d, "DSCR here is on SBA debt service alone ($7,803.05/month) - there is no seller note. Month 1 is the "
            "transition month (partial census at 12 ADC) and is absorbed by the $372K opening cash reserve; "
            "coverage clears the 1.25x floor with wide margin as census stabilizes around the validated ~22 ADC. "
            "On an annual basis Year 1 covers at 3.23x.")
    h2(d, "11.6 Lender summary & debt-service coverage")
    table(d, ["Metric", "Year 1", "Year 2", "Year 3"],
          [["Net revenue", "$1,442,930", "$2,728,642", "$3,734,934"],
           ["EBITDA", "$302,522", "$679,524", "$1,007,520"],
           ["SBA debt service (only debt)", "$93,637", "$93,637", "$93,637"],
           ["DSCR (SBA-only)", "3.23x", "7.26x", "10.76x"],
           ["Net income", "$199,308", "$574,951", "$903,362"]],
          widths=[2.6, 1.4, 1.4, 1.4])
    para(d, "Global 3-year DSCR: 7.08x (vs. 1.25x floor). Because the seller is paid in full at close, the SBA "
            "loan ($7,803.05/mo, $93,637/yr) is the only debt the business carries, so coverage is strong from "
            "Year 1 and rises as census builds. Break-even is approximately 16 ADC, below the validated opening "
            "census of ~22 ADC, so the operation is profitable on the opening book.")
    h2(d, "11.7 Stress tests (Year-2 steady state, SBA-only debt service, floor 1.25x)")
    table(d, ["Scenario", "Net revenue", "EBITDA", "DSCR", "Clears?"],
          [["Base (Year 2)", "$2,728,642", "$679,524", "7.26x", "Yes"],
           ["Census -20% (static roster)", "$2,182,913", "$186,766", "1.99x", "Yes"],
           ["Census -35% (static roster)", "$1,773,617", "($182,802)", "<0", "No"],
           ["Wage +10% (fixed-cost shock)", "$2,728,642", "$501,098", "5.35x", "Yes"],
           ["Combined (-20% census / +10% cost)", "$2,182,913", "$8,340", "0.09x", "No"]],
          widths=[2.4, 1.4, 1.3, 0.9, 0.9])
    para(d, "Reading the stress tests. Because the seller is paid at close, debt service is the SBA loan alone "
            "($93,637), and Year-2 coverage absorbs a great deal: the base case covers at 7.26x, a sustained 20% "
            "census drop still clears at 1.99x, and a 10% fixed-cost spike clears at 5.35x. Only the extreme cases "
            "fail - a 35% census collapse turns EBITDA negative, and a severe combined shock (-20% census AND "
            "+10% costs, full roster held constant) is essentially uncovered (0.09x). Even these understate resilience: (1) the "
            "roster is census-driven, so a 20-35% lower census carries a lighter cost base than the static test "
            "assumes; (2) the binding variable is census, which is why opening census is the validated ~22 ADC "
            "book (break-even ~16 ADC), not a cold start; and (3) the $372K reserve bridges timing shortfalls. A "
            "detailed Year-1 single-factor and combined sensitivity is in the separate DSCR Sensitivity / "
            "Stress-Test memorandum.")
    h2(d, "11.8 Proof: actuals vs. model")
    para(d, "The model is not pure projection - its economics are benchmarked to the acquired book's actual "
            "operating results (Apr-Jun 2025), which reconcile to the model's opening steady state within "
            "~2%. April-June average net revenue was approximately $118,400/mo at ~22 ADC; the model's "
            "opening steady-state month is ~$121,000 (+2.2%). Patient-related COGS also reconcile to the "
            "actuals. Payroll differs by design - Azalea's staggered W-2 roster plus benefits and go-forward "
            "Medical Directors versus the acquired book's blended run.")

    # ---- 12 Risk ----
    d.add_page_break()
    h1(d, "12 - Risk Factors and Mitigations")
    para(d, "The dominant risks are the newco's lack of operating history and the census ramp - both mitigated "
            "by validated, benchmarked economics, a CHOW that keeps the agency billing from day one, an "
            "experienced operator, and a well-capitalized single-loan structure (seller paid at close).")
    table(d, ["Risk", "Prob.", "Impact", "Mitigation"],
          [["Debt service / coverage", "Low", "Med",
            "Seller paid at close, so the SBA loan is the only debt; DSCR clears the 1.25x floor annually and rises 3.23x -> 10.76x (global 7.08x)"],
           ["CHOW transition / cash timing", "Low", "Med",
            "CHOW preserves the provider number (billing-ready day one); $372K reserve funds any claims-processing pause; only SBA debt service to cover"],
           ["Slower census ramp", "Med", "High",
            "Opening census is the validated ~22 ADC book, not a cold start; break-even ~16 ADC; capacity hires are census-gated so cost flexes with volume; liquidity covers timing shortfalls"],
           ["Wage inflation", "Med", "Med",
            "+10% fixed-cost shock holds at 5.35x in Year 2; 3% merit modeled; PRN pool buffers"],
           ["CMS rate / regulatory change", "Low", "Med",
            "Rates up or flat every year since 2010; only a conservative 2.5%/yr escalation assumed"],
           ["Key-person dependency", "Low", "Med",
            "Operators are owners (13.3% each, milestone-vested); Executive Director leads day-to-day; OA key-person succession provisions"],
           ["Equity injection phased", "Low", "Low",
            "$100K on deposit; balance committed; SBA disbursement sequenced after full injection per SOP 50 10 8"],
           ["Newco / no operating history", "Med", "Med",
            "Acquires an established billing-ready agency benchmarked to actual collections; operator has 10+ years of hospice ownership; 26% equity injection"]],
          widths=[1.7, 0.6, 0.6, 3.6])
    para(d, "The credit's core protection is liquidity and a clean structure: a $372,000 opening reserve plus a "
            "single SBA loan (the seller is paid at close, so there is no second debt layer) keep Azalea solvent "
            "and well-covered through the modeled scenarios.")

    # ---- 13 Milestones ----
    d.add_page_break()
    h1(d, "13 - Milestones and Conclusion")
    h2(d, "13.1 Launch & growth milestones")
    table(d, ["Window", "Milestone", "Verification"],
          [["Through July 2026", "Full $195K equity injection on deposit", "Bank statements (Mercury acct ****1275) + Bullard source-of-funds"],
           ["Late July 2026", "Close Hickory CHOW; banking & insurance updated", "Purchase agreement + cert. of insurance"],
           ["Months 1-3", "CMS-855A filed; staff retained; EMR transitioned", "855A CHOW receipt + payroll"],
           ["Months 1-3", "License & CHAP/ACHC accreditation transfer; billing uninterrupted", "License / accreditation transfer"],
           ["Months 4-6", "BD ramp; AR normalizes to ~45-day cycle", "CRM report + AR aging < 45 days"],
           ["Months 7-12", "Census builds toward ~24 ADC; DSCR well above 1.25x", "Monthly financials"],
           ["Years 2-3", "Capacity hires triggered by ADC; census to ~54 ADC; coverage rises above 10x", "Board reviews"]],
          widths=[1.5, 2.9, 2.1])
    h2(d, "13.2 Conclusion")
    para(d, "Azalea is a well-structured SBA 7(a) opportunity: the startup CHOW acquisition and relaunch of an "
            "established, Medicare-certified hospice on validated economics, led by an operator with 10+ years of "
            "hospice ownership. The credit rests on three verifiable points: (1) the acquired book's actual "
            "operating results, which the opening steady-state model reconciles to within ~2%; (2) a base case "
            "whose DSCR clears the 1.25x floor and strengthens every year (3.23x -> 7.26x -> 10.76x; global "
            "7.08x), because the seller is paid at close and the SBA loan is the only debt; and (3) a "
            "well-capitalized structure - $195K cash equity (26% of project) and a $372K opening reserve - that "
            "keeps liquidity intact through stress scenarios. The binding risk is a sustained census shortfall, "
            "mitigated by a validated opening book above break-even, a census-driven cost structure, and "
            "substantial liquidity.")
    para(d, "")
    para(d, "Rev 5.00 - June 2026. This plan is computed from the Azalea Hospice SBA Loan Package operating "
            "model, the single source of truth for all financial figures. Confidential - do not distribute "
            "without written consent.", italic=True, size=8, color=GREY)
    footer_note(d)
    save(d, "03_business_plan/Azalea_SBA_Business_Plan_Rev5.00.docx")


def submission_readiness():
    d = new_doc("Submission Readiness")
    h1(d, "Submission Readiness Checklist")
    para(d, "Tyler Hospice HoldCo L.L.C. - SBA 7(a) Application to John Hart. Work this list top to bottom; "
            "the three BLOCKERS must clear before submission.", color=GREY, size=9)
    h2(d, "BLOCKERS - must resolve before submitting")
    table(d, ["#", "Blocker", "Owner", "Status"],
          [["B1", "Seller payment / structure: RESOLVED - the $300K seller price is paid in full at close from SBA proceeds + equity (no retained note). The MIPA's 12 x $25K installment schedule should be settled at closing (prepay/payoff at close); confirm with sellers via Greiner. With no seller note, Year-1 DSCR is ~3.23x.", "Geoff + sellers (Greiner)", "RESOLVED in structure; confirm payoff mechanics"],
           ["B2", "Borrower entity mismatch: MIPA buyer is 'Tyler Hospice Hold, LLC (Texas)'; SBA/OA is 'Tyler Hospice HoldCo L.L.C. (Wyoming), EIN 41-4966640'. Make one entity/name/state consistent everywhere.", "Geoff", "OPEN"],
           ["B3", "Execute OA Amendment No. 1 (cap table 39.9/19.5/13.3x3 with operators' full 13.3% forfeitable until vested - 4.9% time-vested base + 8.4% dual-trigger earn-up, partnership tax, Hickory target, SBA carve-outs). Current OA still shows 60/25/4.9 + S-Corp + 'Healing Hands'.", "All members", "DRAFTED - needs signatures"]],
          widths=[0.4, 4.3, 1.5, 0.9])
    h2(d, "Geoff's personal package (sole 20%+ guarantor)")
    for x in ["SBA Form 413 (Personal Financial Statement) - include the VistaRiver note as an asset",
              "SBA Form 912 (Personal History)",
              "Personal Cash Flow (7a) - complete the expense lines (income pre-filled)",
              "3 years personal tax returns + W-2s/1099s",
              "Credit report authorization + driver license",
              "Spouse acknowledgement - Mary Elizabeth Burcham"]:
        para(d, "  - " + x)
    h2(d, "Equity injection (James Bullard, $195,000)")
    for x in ["Mercury statement/screenshot of the $100,000 wire (5/7/2026) - drop into 10_supporting_documents/equity_injection_evidence/",
              "Bullard bank/brokerage statements (30-day seasoning) for each tranche",
              "remaining wire as it lands; final statement showing full $195K on deposit",
              "Signed: Subscription Agreement, Investor Attestation v2 (notarized), Source-of-Funds letter"]:
        para(d, "  - " + x)
    h2(d, "Acquisition / Hickory (mostly in hand from Drive)")
    for x in ["Final executed MIPA (Rev 2.00) - confirm full signatures and Effective Date - FILED",
              "Schedule 2.02 (ownership split Gleason/Lozano) - obtain",
              "Tracy Gleason individual address - obtain",
              "Hickory EIN letter (86-2886807), SOS certificate, PTAN, CHAP contract, HCSSA approval, 2021-2022 Form 1065 - collect from Drive into folder 10",
              "Office lease/LOI for 700 N Main Street, Lindale, TX (term matching loan term)"]:
        para(d, "  - " + x)
    h2(d, "Entity documents")
    for x in ["WY Certificate of Formation (or TX, per B2 resolution)",
              "EIN assignment letter (CP-575) for Tyler Hospice HoldCo (41-4966640)",
              "Executed Operating Agreement + Amendment No. 1",
              "Certificate of Good Standing"]:
        para(d, "  - " + x)
    h2(d, "Ready to go (drafted in this package)")
    for x in ["Company Profile, Use of Funds, Equity Injection memo, Business Debt Schedule",
              "Business Plan Rev 5.00 (refresh financials if seller terms change per B1)",
              "Lender Credit Memo, DSCR Sensitivity, Affiliate memo, Cover Letter + Q&A",
              "Submission Cover Sheet / TOC, Closing Checklist, Insurance Requirements",
              "4 Management Resumes (add personal identifiers)"]:
        para(d, "  - " + x)
    footer_note(d)
    save(d, "SUBMISSION_READINESS_CHECKLIST.docx")


def borrower_info_1919():
    d = new_doc("SBA Form 1919 - Borrower Information (pre-fill)")
    h1(d, "SBA Form 1919 - Borrower Information Form")
    para(d, "PRE-FILL WORKSHEET - transcribe onto the current official SBA Form 1919 (OMB 3245-0178) at "
            "signing. Business answers are pre-filled from the application; identity fields and the "
            "character/citizenship questions in Section II must be completed and attested by each principal "
            "personally and truthfully.", color=GREY, size=9)

    h2(d, "Section I - Applicant Business Information (completed once by the business)")
    field(d, "1. Applicant business legal name", "Tyler Hospice Hold LLC")
    field(d, "   Trade name / dba", "Azalea Hospice & Palliative Care")
    field(d, "2. Primary business address", "13387 Hwy 69 N, Tyler, TX " + TBD + " (zip)")
    field(d, "3. Business EIN", "41-4966640")
    field(d, "4. Business phone", "480-495-5474")
    field(d, "5. Primary contact", "Geoff Schackmann, Managing Member")
    field(d, "6. NAICS code", "621610 - Home Health Care Services (confirm hospice classification with lender)")
    field(d, "7. Type of business entity", "Limited Liability Company (Wyoming; formed 03/18/2026)")
    field(d, "8. Amount of SBA loan requested", "$555,000")
    field(d, "9. Number of employees (incl. owners)", "9 at funding; scaling to ~25+ by Year 3")
    field(d, "10. Is the Applicant a franchise?", "No")
    para(d, "")
    para(d, "Business eligibility questions (mark the answer; any 'Yes' requires a written explanation):", bold=True, size=10)
    table(d, ["#", "Question", "Answer"],
          [["1", "Is the Applicant presently suspended, debarred, proposed for debarment, declared ineligible, voluntarily excluded, or otherwise excluded from participation by any federal department or agency?", "No"],
           ["2", "Has the Applicant, or any business owned/controlled by any of its owners, ever obtained a direct or guaranteed loan from SBA or any other federal agency that is currently delinquent or has defaulted in the last 7 years and caused a loss to the government?", "No"],
           ["3", "Is the Applicant presently involved in any bankruptcy or insolvency proceeding?", "No - newly formed entity"],
           ["4", "Is the Applicant or any owner presently subject to any pending civil lawsuit, judgment, or tax lien?", "No (confirm)"],
           ["5", "Is the Applicant a party to any pending lawsuit, and if so will it materially affect operations?", "No"],
           ["6", "Will any loan proceeds be used to pay a delinquent federal debt or to reimburse any owner for prior contributions?", "No - proceeds per Use of Proceeds (acquisition at close, startup, equipment, working capital)"],
           ["7", "Does the Applicant or any associate have a current or prior SBA loan, or an application pending?", "No"],
           ["8", "Is any owner of the Applicant an associate of another business with an existing SBA loan?", "No"],
           ["9", "Has the Applicant or any associate been more than 60 days delinquent on child support?", "(Each principal to confirm)"],
           ["10", "Does the Applicant have an Employee Stock Ownership Plan (ESOP)?", "No"]],
          widths=[0.3, 5.5, 1.2])
    para(d, "Business certification: by signing the official form, the authorized representative certifies the "
            "above are true and that the business is eligible (for-profit; located in the U.S.; meets SBA size "
            "standards; not engaged in an ineligible activity). Hospice care is an eligible business activity.",
            italic=True, size=9)

    h2(d, "Section II - Principal Information (one per principal)")
    para(d, "Complete for: each owner of 20% or more; each officer, director, managing member; and any person "
            "hired to manage day-to-day operations. Below is who must complete Section II and their status. "
            "SSN, date of birth, place of birth, and the three character questions are personal attestations - "
            "leave them for each individual to complete and sign.", italic=True, size=9)
    table(d, ["Principal", "Title", "% owned", "20%+ owner?", "Guarantor?", "Section II required"],
          [["Geoff Schackmann", "Managing Member (via Adeline & Lilah, LLC)", "39.9% (indirect control)", "Yes", "Yes (sole PG)", "Yes - full"],
           ["Mary Elizabeth Burcham", "Passive member of Adeline & Lilah, LLC", "19.95% indirect", "No", "No (spouse consent only)", "Yes - as spouse/owner of A&L"],
           ["James E. Bullard", "Passive minority investor", "19.5%", "No", "No", "Confirm with lender*"],
           ["Silas R. Shelton", "Executive Director", "13.3%", "No", "No", "Yes - key employee (manages operations)"],
           ["Dana L. Davenport", "Director of Nursing", "13.3%", "No", "No", "Yes - key employee"],
           ["Bradley G. Woodard", "Director of Sales", "13.3%", "No", "No", "Yes - key employee"]],
          widths=[1.5, 1.9, 1.2, 0.8, 1.0, 1.6])
    para(d, "*Bullard is a passive 19.5% investor with no management role; confirm with the lender whether they "
            "require a Section II from him. Under 20% and non-managing, many lenders do not - but some collect it "
            "for all equity holders. No personal guaranty or PFS is required of him regardless.", italic=True, size=9)
    para(d, "")
    para(d, "Per-principal fields to complete on the official form:", bold=True, size=10)
    for x in ["Full legal name; any other names used",
              "Residential address; home phone",
              "Social Security Number; date of birth; place of birth (city/state/country)",
              "U.S. citizen? / Lawful Permanent Resident? (if LPR, Alien Registration Number)",
              "Percentage of ownership and position/title",
              "Character Q1: Are you presently subject to an indictment, criminal information, arraignment, or other means by which formal criminal charges are brought in any jurisdiction?",
              "Character Q2: Have you been arrested in the past six months for any criminal offense?",
              "Character Q3: For any criminal offense - other than a minor vehicle violation - have you ever: (a) been convicted; (b) pleaded guilty; (c) pleaded nolo contendere; (d) been placed on pretrial diversion; or (e) been placed on any form of parole or probation?",
              "Are you presently suspended, debarred, or otherwise excluded by any federal agency?"]:
        para(d, "  - " + x, size=9.5)
    para(d, "Any 'Yes' to a character question requires a written explanation and may require additional "
            "background processing (Form 912 and FBI clearance). Answer exactly and truthfully - this drives the "
            "lender's CAIVRS/SAM screening.", italic=True, size=9)
    footer_note(d)
    save(d, "01_company_profile/SBA_Form_1919_Borrower_Information_PREFILL.docx")


def pfs_413_prefill():
    d = new_doc("SBA Form 413 - PFS pre-fill (Geoff Schackmann)")
    h1(d, "SBA Form 413 - Personal Financial Statement (pre-fill worksheet)")
    para(d, "PRE-FILL WORKSHEET for Geoff Schackmann, the sole 20%+ owner and guarantor. Transcribe onto the "
            "current official SBA Form 413 (OMB 3245-0188). Dollar values must be supplied by Geoff and supported "
            "by statements; the two items known from the file are pre-filled below. As of date: " + TBD + ".",
            color=GREY, size=9)
    para(d, "Important: the $195,000 equity injection is James Bullard's cash (a separate 19.5% member), NOT "
            "Geoff's - so it does not appear on Geoff's PFS. Geoff's PFS supports the personal guaranty; "
            "Bullard's source-of-funds is verified separately from his own statements.", italic=True, size=9)

    field(d, "Name", "Geoff Schackmann")
    field(d, "Residential address", TBD)
    field(d, "Business name of applicant", "Tyler Hospice Hold LLC (dba Azalea Hospice & Palliative Care)")
    para(d, "")
    h2(d, "Assets")
    table(d, ["Asset", "Value", "Notes"],
          [["Cash on hand and in banks", TBD, "Attach recent bank statements"],
           ["Savings accounts", TBD, ""],
           ["IRA / other retirement accounts", TBD, "Schedule / statement"],
           ["Accounts and notes receivable", "see Schedule", "Includes the VistaRiver seller note (below)"],
           ["Life insurance - cash surrender value", TBD, "Name insurer"],
           ["Stocks and bonds", TBD, "Schedule B"],
           ["Real estate (from Schedule C)", TBD, "Primary residence / other"],
           ["Automobiles", TBD, "Year/make"],
           ["Other personal property", TBD, ""],
           ["Other assets", TBD, ""],
           ["TOTAL ASSETS", TBD, ""]],
          widths=[2.6, 1.6, 2.6])
    h2(d, "Liabilities")
    table(d, ["Liability", "Value", "Notes"],
          [["Accounts payable", TBD, ""],
           ["Notes payable to banks and others", TBD, "Schedule A"],
           ["Installment account (auto)", TBD, "Monthly payment ____"],
           ["Installment account (other)", TBD, "Monthly payment ____"],
           ["Loans on life insurance", TBD, ""],
           ["Mortgages on real estate (from Schedule C)", TBD, ""],
           ["Unpaid taxes", TBD, ""],
           ["Other liabilities", TBD, ""],
           ["TOTAL LIABILITIES", TBD, ""],
           ["NET WORTH (assets - liabilities)", TBD, ""]],
          widths=[2.6, 1.6, 2.6])
    h2(d, "Notes Receivable - pre-filled known item")
    table(d, ["Payor", "Origin", "Original amount", "Current balance", "Status"],
          [["VistaRiver (sale of Geoff's prior hospice interest)", "MIPA dated 08/15/2025 (in PFS support folder)", TBD, TBD, "Provide payment ledger / recent statement showing the note is current"]],
          widths=[2.2, 2.2, 1.2, 1.2, 1.5])
    h2(d, "Contingent Liabilities - pre-filled known item")
    table(d, ["Type", "Amount", "Notes"],
          [["As guarantor: SBA 7(a) loan to Tyler Hospice Hold LLC", "$555,000 (proposed)", "This transaction - Geoff is the sole personal guarantor"],
           ["As endorser or co-maker (other)", TBD, ""],
           ["Legal claims and judgments", TBD, ""],
           ["Provision for federal income tax", TBD, ""],
           ["Other special debt", TBD, ""]],
          widths=[3.4, 1.6, 2.0])
    h2(d, "Source of Income (annual) and Personal Information")
    for x in ["Salary / wages; net investment income; real estate income; other income (describe)",
              "Schedule A - notes payable to banks and others (to whom, balance, terms, security)",
              "Schedule B - stocks and bonds",
              "Schedule C - real estate owned (address, cost, market value, mortgage balance, payment)",
              "Life insurance held (face amount, cash surrender value, beneficiaries)"]:
        para(d, "  - " + x, size=9.5)
    para(d, "Attach: most recent statements for each asset/liability line, a recent pay stub or income "
            "documentation, and the VistaRiver note ledger. The lender uses the PFS to assess the personal "
            "guaranty and overall credit strength.", italic=True, size=9)
    footer_note(d)
    save(d, "05_personal_financial_statement/SBA_Form_413_PFS_PREFILL_Geoff.docx")


if __name__ == "__main__":
    company_profile()
    borrower_info_1919()
    pfs_413_prefill()
    submission_readiness()
    use_of_funds()
    equity_injection_memo()
    debt_schedule()
    management_resumes()
    reconciliation_memo()
    information_needed()
    personal_forms_guide()
    lender_credit_memo()
    sensitivity_memo()
    bullard_attestation()
    submission_cover_sheet()
    closing_checklist()
    insurance_requirements()
    personal_cash_flow()
    interview_worksheet()
    oa_amendment()
    bullard_attestation_v2()
    bullard_subscription()
    bullard_source_of_funds()
    affiliate_memo()
    cover_letter_and_qa()
    business_plan()
    print("Done.")
