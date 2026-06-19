"""
Amended & Restated Operating Agreement of TYLER HOSPICE HOLD, LLC (Wyoming).

Self-contained builder (does not import build_documents.py, which is being edited
concurrently). Integrates the provided OA v7 with every amendment/decision locked
on 2026-06-16:

  - Entity: Tyler Hospice Hold, LLC, a Wyoming LLC (EIN 41-4966640), foreign-qualified
    in Texas; owns 100% of Hickory Hospice LLC (dba Azalea Hospice & Palliative Care).
  - Tax: S corporation under IRC §1361 (single class of stock; strictly pro-rata distributions).
  - Cap table: Geoff Schackmann 39.9 / Bullard 19.5 / Shelton-Davenport-Woodard 13.3 each / pool 0.7.
    Geoff holds the 39.9% directly as an individual (an eligible S-corp shareholder).
  - Operator Restricted Interests: full 13.3% at risk; 4.9% time-vested Initial Base +
    8.4% dual-trigger Earn-Up; vesting from the Effective Date; good/bad-leaver call.
  - Bullard floor 19.5%; SBA debt carve-out; Texas Shootout deferred while SBA debt out.
  - Unvested Restricted Interests vote.
  - Acquisition: MIPA purchase of 100% of Hickory; SBA $450K + bank $300K + equity $195K;
    Hickory corporate guaranty of the SBA loan; Bullard guarantees the bank loan only.
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor as RGB, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

NAVY = RGB(0x1F, 0x3A, 0x5F)
GREY = RGB(0x6B, 0x6B, 0x6B)
BLACK = RGB(0x00, 0x00, 0x00)
OUT = os.path.join(os.path.dirname(__file__), "08_entity_documents",
                   "Tyler_Hospice_Hold_LLC_AMENDED_and_RESTATED_OA.docx")


def doc_new():
    d = Document()
    st = d.styles["Normal"].font
    st.name = "Calibri"; st.size = Pt(10.5)
    d.styles["Normal"].paragraph_format.space_after = Pt(6)
    for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(d.sections[0], m, Inches(1.0))
    return d


def title(d, text, size=15):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.bold = True; r.font.size = Pt(size); r.font.color.rgb = NAVY
    return p


def sub(d, text, size=10, color=GREY, italic=True, center=True):
    p = d.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.italic = italic; r.font.size = Pt(size); r.font.color.rgb = color
    return p


def art(d, text):
    d.add_paragraph()
    p = d.add_paragraph()
    r = p.add_run(text); r.bold = True; r.font.size = Pt(12); r.font.color.rgb = NAVY
    return p


def sec(d, num, heading, body=None):
    p = d.add_paragraph()
    r = p.add_run("%s  %s  " % (num, heading)); r.bold = True; r.font.size = Pt(10.5)
    if body:
        r2 = p.add_run(body); r2.font.size = Pt(10.5)
    return p


def para(d, text, italic=False, size=10.5, bold=False, indent=False):
    p = d.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Inches(0.3)
    r = p.add_run(text); r.italic = italic; r.bold = bold; r.font.size = Pt(size)
    return p


def rule(d):
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("─" * 60); r.font.color.rgb = GREY
    return p


def captable(d):
    rows = [
        ("Member", "Consideration", "Percentage Interest"),
        ("Geoff Schackmann (individual)", "Services rendered (sweat equity)", "39.9%"),
        ("James E. Bullard", "$195,000.00 cash (capital contribution)", "19.5%"),
        ("Silas R. Shelton", "Services (Restricted Interest)", "13.3%*"),
        ("Dana L. Davenport", "Services (Restricted Interest)", "13.3%*"),
        ("Bradley Gene Woodard", "Services (Restricted Interest)", "13.3%*"),
        ("Unissued Pool", "Reserved for future grants", "0.7%"),
        ("TOTAL", "", "100.0%"),
    ]
    t = d.add_table(rows=len(rows), cols=3); t.style = "Light Grid Accent 1"
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.rows[i].cells[j]; c.text = val
            for pr in c.paragraphs:
                for rr in pr.runs:
                    rr.font.size = Pt(9.5)
                    if i == 0 or row[0] == "TOTAL":
                        rr.font.bold = True
    return t


def build():
    d = doc_new()

    # ---------- Title ----------
    title(d, "TYLER HOSPICE HOLD, LLC")
    sub(d, "AMENDED AND RESTATED OPERATING AGREEMENT", size=12, color=NAVY, italic=False)
    sub(d, "A Wyoming Limited Liability Company  (EIN 41-4966640)")
    sub(d, "Electing S-Corporation Tax Treatment Under IRC §1361")
    sub(d, "Amendment & Restatement No. 1  |  Effective as of ________________, 2026")
    sub(d, "DRAFT FOR COUNSEL REVIEW AND EXECUTION - the parties are advised to obtain "
           "independent legal and tax advice before signing.", size=9)
    sub(d, "Geoff Schackmann holds the 39.9% Membership Interest directly, as an individual (an "
           "eligible S-corporation shareholder), replacing the former Adeline & Lilah, LLC holder.", size=9)
    rule(d)

    # ---------- Preamble & recitals ----------
    para(d,
         "This Amended and Restated Operating Agreement (this “Agreement”) of Tyler "
         "Hospice Hold, LLC, a Wyoming limited liability company (the “Company”), is "
         "effective as of the date last signed below (the “Effective Date”), by and among "
         "Geoff Schackmann, an individual (“Geoff Schackmann”); "
         "James E. Bullard (“Bullard” or the “Investor”); Silas R. Shelton "
         "(“Silas”); Dana L. Davenport (“Dana”); and Bradley Gene Woodard "
         "(“Brad”) (collectively, the “Members” and each a “Member”). Geoff "
         "Schackmann, who is a Member, also serves as the Company's manager (in that capacity, the "
         "“Manager”).")
    para(d, "RECITALS", bold=True)
    para(d, "A.  The Company and the Members entered into (or agreed to the form of) that certain "
            "Operating Agreement of the Company (styled “Tyler Hospice HoldCo L.L.C.,” Version "
            "5, March 2026) (the “Original Agreement”). The Members now desire to amend and "
            "restate the Original Agreement in its entirety, and this Agreement supersedes and "
            "replaces the Original Agreement in full as of the Effective Date. To the extent the "
            "Original Agreement was not executed or did not become effective, this Agreement "
            "constitutes the initial Operating Agreement of the Company.")
    para(d, "B.  This Agreement (i) conforms the Company's exact legal name to “Tyler Hospice "
            "Hold, LLC”; (ii) confirms the Company's election to be taxed as an S corporation "
            "under IRC §1361 and conforms the Company's membership to the S-corporation eligibility "
            "rules (the 39.9% interest is held directly by Geoff Schackmann, an eligible individual, and "
            "the Company maintains a single class of stock with strictly pro-rata distributions); (iii) restates "
            "the capitalization table; (iv) updates "
            "the acquisition target to Hickory Hospice LLC; (v) revises the operator-member "
            "Restricted Interests so the entire interest is earned through time- and "
            "performance-based vesting; (vi) calibrates the Investor's minority-protection rights to "
            "remain consistent with U.S. Small Business Administration (“SBA”) Standard "
            "Operating Procedure 50 10; and (vii) coordinates the Company's governance with its SBA "
            "7(a) financing and the related bank acquisition loan.")
    para(d, "C.  The Company is, or upon the closing described in Section 1.7 will be, the sole "
            "member and 100% owner of Hickory Hospice LLC, a Texas limited liability company that "
            "holds a Medicare-certified hospice provider number and operates (and after the closing "
            "will continue to operate, doing business as “Azalea Hospice & Palliative Care”) "
            "as the Company's wholly owned operating subsidiary.")
    para(d, "NOW, THEREFORE, in consideration of the mutual covenants herein and other good and "
            "valuable consideration, the receipt and sufficiency of which are acknowledged, the "
            "parties agree as follows:")

    # ================= ARTICLE I =================
    art(d, "ARTICLE I - THE COMPANY")
    sec(d, "1.1", "Company Name and Formation.",
        "The Company is organized as Tyler Hospice Hold, LLC pursuant to the Wyoming Limited "
        "Liability Company Act (W.S. §17-29-101 et seq.) (the “Act”). The Company's "
        "registered agent in Wyoming is Northwest Registered Agent, or such other agent as the "
        "Manager designates. The Company is, or shall promptly be, foreign-qualified to transact "
        "business in the State of Texas, where it operates through its subsidiary.")
    sec(d, "1.2", "Sole Permitted Business.",
        "The Company shall carry on the sole and exclusive business of acquiring, owning, and "
        "operating hospice and related healthcare businesses (the “Business”), including "
        "owning the membership interests of one or more operating subsidiaries licensed under the "
        "Texas Health and Human Services Commission Home and Community Support Services Agency "
        "(“HCSSA”) rules and enrolled in the Medicare and Medicaid programs. The Company "
        "may engage in all activities incidental to the Business.")
    sec(d, "1.3", "Powers and Privileges.",
        "The Company may exercise all powers granted by the Act and applicable law, including the "
        "power to contract, to incur and guarantee indebtedness and grant security interests in its "
        "property and in the equity and assets of its subsidiaries, and to take all action "
        "incidental to the Business under the laws of Wyoming, Texas, or any other applicable "
        "jurisdiction.")
    sec(d, "1.4", "Principal Place of Business.",
        "The principal place of business shall be in Tyler, Texas, or such other location as the "
        "Manager designates in writing.")
    sec(d, "1.5", "Term.", "The Company's term commenced upon filing of its Certificate of "
        "Organization and continues in perpetuity until dissolved under Article VII.")
    sec(d, "1.6", "Fiscal Year.", "The fiscal year is the calendar year. All milestone "
        "calculations are based on calendar months within the applicable fiscal year.")
    sec(d, "1.7", "Acquisition of the Operating Subsidiary; SBA Financing.",
        "The Company has entered into a Membership Interest Purchase Agreement (as amended, the "
        "“MIPA”) to purchase one hundred percent (100%) of the membership interests of "
        "Hickory Hospice LLC (the “Acquired Agency”), a Texas limited liability company, in "
        "a change-of-ownership transaction under which the Acquired Agency retains its Medicare "
        "provider number and operates as “Azalea Hospice & Palliative Care.” The parties "
        "acknowledge that the MIPA currently names the Buyer as a Texas limited liability company "
        "and shall be amended so that the Buyer is “Tyler Hospice Hold, LLC, a Wyoming limited "
        "liability company.” The acquisition and the Company's working capital are financed by "
        "(a) an SBA 7(a) loan to the Company (the “SBA Loan”); (b) a bank term loan used to "
        "fund the acquisition (the “Bank Acquisition Loan”), which is subordinate to the SBA "
        "Loan and which the Investor personally guarantees; and (c) the Investor's equity injection "
        "under Section 3.3. The Acquired Agency shall provide a corporate guaranty of, and a lien on "
        "its assets securing, the SBA Loan. The closing shall be coordinated under Section 3.16.")

    # ================= ARTICLE II =================
    art(d, "ARTICLE II - DEFINITIONS")
    defs = [
        ("2.1", "“Act”", "the Wyoming Limited Liability Company Act (W.S. §17-29-101 et seq.), as amended."),
        ("2.2", "“Acquired Agency”", "Hickory Hospice LLC, a Texas limited liability company, the Company's wholly owned operating subsidiary (doing business as Azalea Hospice & Palliative Care)."),
        ("2.3", "“Bank Acquisition Loan”", "the bank term loan (approximately $300,000, 6% per annum, three-year amortization) obtained to fund the purchase of the Acquired Agency, guaranteed by the Investor and subordinate to the SBA Loan."),
        ("2.4", "“Capital Account” and “Capital Contributions”", "have the meanings in Article III."),
        ("2.5", "“Change of Control”", "any transaction or series of related transactions in which (a) any person or group acquires more than fifty percent (50%) of outstanding Percentage Interests; (b) the Company sells all or substantially all of its assets; or (c) the Company merges or consolidates such that existing Members hold less than fifty percent (50%) of the surviving entity's voting interests."),
        ("2.6", "“Code”", "the Internal Revenue Code of 1986, as amended."),
        ("2.7", "“Continuous Service”", "continuous service to the Company or the Acquired Agency as an employee, officer, or manager."),
        ("2.8", "“EBITDA”", "earnings before interest, taxes, depreciation, and amortization, determined under GAAP applied consistently."),
        ("2.9", "“Breakeven”", "the first calendar month in which the Company's EBITDA is greater than or equal to zero."),
        ("2.10", "“Profitable”", "any calendar month in which the Company's EBITDA is greater than zero."),
        ("2.11", "“Equity Grantee”", "each of Silas, Dana, and Brad, as a holder of a Restricted Interest under Article V."),
        ("2.12", "“Event of Insolvency”", "has the meaning customary for an order for relief, assignment for the benefit of creditors, or unsatisfied receivership, attachment, or charging order continuing for thirty (30) days, with respect to the Company or a Member."),
        ("2.13", "“Forfeiture Conditions”", "the time- and performance-based vesting conditions applicable to Restricted Interests under Article V; an interest is “Unvested” until, and “Vested” to the extent, such conditions are satisfied."),
        ("2.14", "“Percentage Interest”", "each Member's ownership percentage as set forth in Exhibit A, including all rights to capital, voting, profits, and distributions, subject to this Agreement."),
        ("2.15", "“Person”", "any natural person or legal entity."),
        ("2.16", "“Prime Rate”", "the prime rate published from time to time by The Wall Street Journal."),
        ("2.17", "“Profits or Losses”", "all items of income, gain, loss, expense, and deduction determined under the Company's method of accounting in accordance with the Code."),
        ("2.18", "“Regulations”", "the Treasury Regulations under the Code, as amended."),
        ("2.19", "“Restricted Interest”", "a membership interest granted to an Equity Grantee under Article V that is subject to the Forfeiture Conditions."),
        ("2.20", "“SBA Loan”", "the loan made to the Company under the SBA 7(a) loan program by a participating lender, together with any successor, replacement, or refinancing loan that carries an SBA guarantee."),
        ("2.21", "“Transfer”", "any sale, assignment, pledge, hypothecation, gift, encumbrance, or other disposition of a membership interest, voluntary or involuntary."),
        ("2.22", "“Unissued Pool”", "the seven-tenths of one percent (0.7%) of membership interest reserved for future grants, plus any interests later forfeited and reverted under Article V, available for future grants at the Manager's discretion, subject to Section 6.2(c)."),
        ("2.23", "“Vesting Commencement Date”", "the Effective Date of this Agreement; all vesting under Article V is measured from this date and no service before this date is credited."),
        ("2.24", "“Vote” or “Prevailing Vote”", "a vote by Members holding a simple majority (greater than fifty percent (50%)) of the total Percentage Interests then outstanding. Unissued Pool interests are excluded from the denominator. Unvested Restricted Interests carry full voting rights and are included in the denominator (subject only to the post-Separation rule in Section 5.9). Unless otherwise stated, a Prevailing Vote approves any action submitted to the Members."),
        ("2.25", "“Withdrawing Member”", "has the meaning in Section 8.6."),
        ("2.26", "“Undeployed Acquisition Funds”", "Investor capital contribution funds designated by the Manager for, but not yet expended toward, the acquisition described in Section 3.14 at the time a refund is requested under Section 3.15."),
        ("2.27", "“S-Corp Eligible Person”", "a person eligible to be a shareholder of an S corporation under IRC §1361(b)(1) - namely (a) a U.S. citizen or resident-alien individual, (b) an estate, (c) a trust described in IRC §1361(c)(2), or (d) a single-member limited liability company that is disregarded for federal tax purposes and is wholly owned by a person described in (a)-(c). No nonresident alien, partnership, multi-member LLC, or C corporation may hold a Percentage Interest."),
    ]
    for n, term, body in defs:
        sec(d, n, term + " means", body)

    # ================= ARTICLE III =================
    art(d, "ARTICLE III - CAPITAL")
    sec(d, "3.1", "Capital Accounts.",
        "The Company shall maintain a Capital Account for each Member in accordance with Treasury "
        "Reg. §1.704-1(b)(2)(iv), increased by contributions and allocations of income/gain and "
        "decreased by distributions and allocations of loss/expense, as determined by the Manager "
        "with the Company's accountants.")
    sec(d, "3.2", "Capital Contributions.",
        "The Members' capital contributions and Percentage Interests are set forth in Exhibit A.")
    sec(d, "3.3", "Investor Capital Contribution.",
        "Bullard purchases, and the Company issues to Bullard, a nineteen and one-half percent "
        "(19.5%) Percentage Interest for a purchase price of One Hundred Ninety-Five Thousand "
        "Dollars ($195,000.00), payable in cash by wire transfer (and which may be funded in "
        "tranches on a schedule approved by the Manager). Of the purchase price, $180,000 is "
        "designated as the Company's equity injection for the SBA Loan. Bullard's Percentage "
        "Interest is fully earned upon payment and is not subject to any Forfeiture Condition. "
        "Bullard's obligation to fund shall not arise until the Manager delivers (i) a certified "
        "copy of the Company's Wyoming Certificate of Organization confirming good standing, and "
        "(ii) the Company's federal Employer Identification Number.")
    sec(d, "3.4", "Manager (Geoff Schackmann) Capital Contribution.",
        "Geoff Schackmann's thirty-nine and nine-tenths percent (39.9%) Percentage Interest is "
        "issued in consideration of services rendered and to be rendered by Geoff Schackmann as "
        "Manager, including organizing and operating the Company, managing the acquisition, and "
        "building clinical and operational infrastructure.")
    sec(d, "3.5", "Equity Grantee Capital Contributions.",
        "The Restricted Interests granted to Silas, Dana, and Brad under Article V are issued as "
        "compensation for services. The parties acknowledge that the fair market value of each "
        "Restricted Interest as of the Effective Date is nominal, reflecting the startup nature of "
        "the Company.")
    sec(d, "3.6", "Additional Capital Contributions.",
        "The Manager may call additional capital contributions (each, an “Additional Capital "
        "Contribution”) at such times and amounts as the Manager determines in the Company's "
        "best interests, on written notice stating the due date (not more than sixty (60) days "
        "out), each Member's proportional amount, and the business purpose. Members contribute pro "
        "rata to their Percentage Interests. No Member is required to contribute without the consent "
        "of Members holding at least sixty percent (60%) of outstanding Percentage Interests; any "
        "Member may request the Manager to call an Additional Capital Contribution, and if the "
        "Manager fails to act within twenty (20) days, the requesting Member may call a meeting for "
        "a Prevailing Vote.")
    sec(d, "3.7", "Delinquent Members.",
        "A Member who timely contributes all of an Additional Capital Contribution is a "
        "“Contributing Member”; one who fails to contribute timely all or any portion is a "
        "“Delinquent Member.” Upon such failure, either (a) Contributing Members elect to "
        "make a Contribution Loan covering the unpaid amount; (b) no Contribution Loan is made, in "
        "which case Section 3.10 governs; or (c) a Contribution Loan covers part, with the balance "
        "governed by Section 3.10.")
    sec(d, "3.8", "Contribution Loan.",
        "A Contributing Member may, at the Manager's election, advance in cash the amount owed by a "
        "Delinquent Member as a nonrecourse loan (a “Contribution Loan”); multiple "
        "Contributing Members participate pro rata. A Contribution Loan bears interest at the lesser "
        "of the Prime Rate plus three percent (3%) or the maximum lawful rate, is due the earlier of "
        "six (6) months after advance or sale/dissolution, and is secured by the Delinquent "
        "Member's Percentage Interest, which the Delinquent Member pledges and as to which the "
        "Delinquent Member appoints the Contributing Member(s) attorney-in-fact to perfect.")
    sec(d, "3.9", "Failure to Repay Contribution Loan.",
        "If a Delinquent Member fails to timely repay, the Contributing Members may (a) extend the "
        "term six (6) months or (b) contribute the outstanding principal and interest to Company "
        "capital and dilute the Delinquent Member under Section 3.10. Failure to elect within thirty "
        "(30) days of the due date is deemed an election to extend six (6) months.")
    sec(d, "3.10", "Dilution of Delinquent Member.",
        "If a Contributing Member contributes the Contribution Loan amount to capital, the Delinquent "
        "Member's Percentage Interest is diluted by (amount contributed) ÷ (Delinquent Member's "
        "Percentage Interest × appraised value of the Company) × (Delinquent Member's Percentage "
        "Interest), but not below 0.01%, and the Delinquent Member is then released from that "
        "Contribution Loan obligation.")
    sec(d, "3.11", "Determination of Appraised Value.",
        "For Section 3.10, appraised value is determined by an independent qualified appraiser "
        "appointed by the Manager (or, failing agreement, by a court in Smith County, Texas on any "
        "Member's petition), as the greater of the Company's going-concern value or the fair market "
        "value of its assets in an open-market sale, in each case net of liabilities and reasonable "
        "reserves.")
    sec(d, "3.12", "No Interest on Capital.",
        "No Member accrues interest on capital contributions.")
    sec(d, "3.13", "Withdrawal of Capital; Limitation on Distributions.",
        "No Member may withdraw or reduce any Capital Account except on dissolution or as allowed by "
        "law, demand distributions except as provided herein, bring a partition action, cause "
        "dissolution except as set forth herein, demand property other than cash, or claim priority "
        "over any other Member as to capital or distributions, except as expressly provided herein.")
    sec(d, "3.14", "Use of Investor Proceeds.",
        "The Company shall use Bullard's $195,000 contribution, together with the SBA Loan and Bank "
        "Acquisition Loan proceeds, for: (a) the purchase of 100% of the membership interests of the "
        "Acquired Agency (Hickory Hospice LLC), including transfer, legal, and regulatory costs; "
        "(b) working capital for hospice operations; (c) facility and infrastructure costs; and "
        "(d) other operational expenses as the Manager deems necessary. The Manager shall maintain "
        "use-of-proceeds records available to Bullard on request.")
    sec(d, "3.15", "Acquisition Failure.",
        "If the acquisition of the Acquired Agency is not completed within twelve (12) months of the "
        "Effective Date, Bullard may request return of any Undeployed Acquisition Funds. Within "
        "thirty (30) days, the Manager shall return such funds, which shall reduce Bullard's "
        "Percentage Interest by (Returned Amount ÷ $195,000) × 19.5%, with the corresponding "
        "percentage restored to Geoff Schackmann.")
    sec(d, "3.16", "Coordination with SBA Closing.",
        "The change-of-ownership closing shall be scheduled on or after the date the Investor has "
        "contributed the full $195,000 under Section 3.3 (anticipated by the end of July 2026), so "
        "the SBA Loan may disburse with the full equity injection on deposit as required by SBA SOP "
        "50 10. The Manager is authorized to coordinate the closing date, the SBA Loan and Bank "
        "Acquisition Loan disbursements, the Acquired Agency's corporate guaranty and collateral, "
        "any intercreditor or subordination agreement, and the Investor's tranche schedule.")
    sec(d, "3.17", "Enforceability.",
        "The Members acknowledge that the capital remedies in this Article III are fair and "
        "reasonable, do not constitute a penalty or forfeiture, and that each Member has had the "
        "opportunity to consult independent counsel.")

    # ================= ARTICLE IV =================
    art(d, "ARTICLE IV - DISTRIBUTIONS AND ALLOCATIONS")
    sec(d, "4.1", "Discretionary Distributions.",
        "The Manager may distribute funds in excess of reasonable working-capital needs and reserves "
        "pro rata to Percentage Interests, subject to Section 4.3 and to any restriction in the SBA "
        "Loan documents. Restricted Interests are counted for distribution purposes once a timely "
        "Section 83(b) election is filed; any distribution on an Unvested interest bears the same "
        "forfeiture risk as the underlying interest, and the Company has no obligation to recover "
        "prior distributions if the interest is later forfeited.")
    sec(d, "4.2", "Mandatory Tax Distributions.",
        "No later than seventy-five (75) days after each fiscal year in which the Company has taxable "
        "income, the Company shall distribute to each Member, pro rata, an amount sufficient to pay "
        "the estimated federal and state income tax on the Member's allocable share of taxable "
        "income, computed at the highest combined marginal rate applicable to any Member. Tax "
        "Distributions are advances against future distributions and are subject to Section 4.3 and "
        "to any SBA Loan restriction; the Manager shall make the maximum partial distribution "
        "consistent with solvency and the SBA Loan and notify Members of any shortfall.")
    sec(d, "4.3", "Limitations on Distributions.",
        "No distribution shall be made if, after giving effect to it, the Company could not pay its "
        "debts as they come due (W.S. §17-29-405), or if it would violate the SBA Loan documents "
        "or any lender covenant.")
    sec(d, "4.4", "Allocations.",
        "As required for an S corporation with a single class of stock under IRC §1361(b)(1)(D), all "
        "items of income, gain, loss, deduction, and credit are allocated among the Members strictly "
        "in proportion to their Percentage Interests, and all distributions shall be made strictly pro "
        "rata to Percentage Interests; no Member shall receive any allocation or distribution that "
        "differs in timing or amount per Percentage Interest from any other Member.")
    sec(d, "4.5", "Liquidating Distributions.",
        "Upon dissolution, after payment of all liabilities (including the SBA Loan and Bank "
        "Acquisition Loan), remaining assets are distributed to the Members pro rata to their "
        "Percentage Interests. All Percentage Interests constitute a single class of stock and are "
        "treated identically for liquidation purposes in accordance with IRC §1361(b)(1)(D) and "
        "Treasury Reg. §1.1361-1(l), so that no Member has any priority or preference over any other "
        "as to the return of capital or liquidation proceeds, except as expressly provided herein.")

    # ================= ARTICLE V =================
    art(d, "ARTICLE V - RESTRICTED INTERESTS, VESTING, AND REPURCHASE")
    sec(d, "5.1", "Grant and Components.",
        "The Company grants to each Equity Grantee (Silas, Dana, and Brad) a Restricted Interest of "
        "thirteen and three-tenths percent (13.3%) of the Company, issued and outstanding as of the "
        "Effective Date, identical in economic and voting rights to all other membership interests "
        "except for the Forfeiture Conditions, and held subject to this Article V. Each Restricted "
        "Interest "
        "comprises: (a) an “Initial Base” of four and nine-tenths percent (4.9%), "
        "corresponding to the Equity Grantee's original interest, and (b) an “Earn-Up "
        "Portion” of eight and four-tenths percent (8.4%). No portion is Vested at the Effective "
        "Date; the entire 13.3% is at risk and forfeitable until Vested under Section 5.2, and each "
        "Equity Grantee may vest up to a fully-Vested 13.3%.")
    sec(d, "5.2", "Vesting - the Entire Restricted Interest Is at Risk.",
        "No portion of a Restricted Interest is Vested at the Effective Date; an interest becomes "
        "vested and non-forfeitable (“Vested”) only as set out below, and any portion not "
        "yet Vested is “Unvested” and subject to forfeiture under Section 5.4. All vesting "
        "is measured from the Vesting Commencement Date.")
    para(d, "(a) Initial Base - time-vesting only. The Initial Base (4.9%) vests on Continuous "
            "Service alone on a four-year schedule with a one-year cliff: nothing vests before the "
            "first anniversary; 25% of the Initial Base vests on the first anniversary (the "
            "“Cliff”); and the remaining 75% vests in thirty-six (36) equal monthly "
            "installments over months 13-48. No performance milestone is required for the Initial "
            "Base to vest.", indent=True)
    para(d, "(b) Earn-Up Portion - dual trigger (time AND performance). The Earn-Up Portion (8.4%) "
            "becomes Vested only to the extent BOTH (i) and (ii) are satisfied; the Vested "
            "percentage equals the LESSER of: (i) the same four-year/one-year-cliff time schedule "
            "applied to the Earn-Up Portion; and (ii) the performance milestones - 25% at Breakeven, "
            "50% upon three (3) consecutive Profitable months, and 25% upon twelve (12) consecutive "
            "Profitable months. Satisfying only one of the two conditions does not vest any portion "
            "of the Earn-Up Portion.", indent=True)
    para(d, "A Grantee's total Vested Interest at any time equals the then-Vested portion of the "
            "Initial Base plus the then-Vested portion of the Earn-Up Portion; the balance is "
            "Unvested.")
    sec(d, "5.3", "Determination of Vesting.",
        "The Manager shall determine in good faith whether and to what extent vesting has occurred "
        "and shall notify each Equity Grantee within thirty (30) days after the end of any month in "
        "which vesting occurs. An Equity Grantee may request an independent accounting review at the "
        "grantee's expense within thirty (30) days.")
    sec(d, "5.4", "Forfeiture of Unvested Interest.",
        "Upon an Equity Grantee ceasing Continuous Service for any reason (a “Separation”), "
        "the Equity Grantee's Unvested Interest is automatically and immediately forfeited for no "
        "consideration and reverts to the Unissued Pool, regardless of the reason for Separation. "
        "Each Equity Grantee irrevocably appoints the Manager as attorney-in-fact, coupled with an "
        "interest, to execute instruments effecting the forfeiture and any repurchase under Section "
        "5.6.")
    sec(d, "5.5", "Separation Categories.",
        "A “Good-Leaver Separation” means Separation due to (i) termination by the Company "
        "without Cause; (ii) death; (iii) Disability; (iv) Retirement; or (v) voluntary resignation "
        "on or after full time-vesting (the fourth anniversary). A “Bad-Leaver Separation” "
        "means (i) termination for Cause at any time, or (ii) voluntary resignation before full "
        "time-vesting. “Cause,” “Disability,” and “Retirement” have the "
        "meanings customary for such terms and as reasonably determined by the Manager in good "
        "faith (Cause including material breach of duties, conviction of a felony, fraud or "
        "dishonesty, or willful misconduct).")
    sec(d, "5.6", "Company Repurchase (Call) Right Over Vested Interest.",
        "Upon any Separation, the Company has the right (not the obligation) to purchase all (but not "
        "less than all) of the Equity Grantee's Vested Interest (the “Call Right”) by "
        "written notice within one hundred twenty (120) days after the later of the Separation date "
        "and the date the Call Price is finally determined; if the Company does not fully exercise, "
        "the other Members may purchase the balance pro rata within a further thirty (30) days. The "
        "purchase price (the “Call Price”) is: (a) for a Good-Leaver Separation, the Fair "
        "Market Value of the Vested Interest; and (b) for a Bad-Leaver Separation, the LOWER of "
        "(i) the Equity Grantee's cost basis (which, for an interest issued for services, is the cash "
        "amount actually paid, if any, and otherwise zero) and (ii) Fair Market Value - which the "
        "Members acknowledge may be at or near zero.")
    sec(d, "5.7", "Fair Market Value.",
        "“Fair Market Value” of a Vested Interest means its pro-rata share of the Company's "
        "equity value (enterprise value less all outstanding indebtedness, including the SBA Loan "
        "and Bank Acquisition Loan), determined as of the last day of the calendar month preceding "
        "Separation by an independent qualified appraiser, reflecting appropriate discounts for the "
        "interest's minority and illiquid character. The Company bears the appraisal cost unless the "
        "Equity Grantee disputes it and a second appraisal differs by less than ten percent (10%).")
    sec(d, "5.8", "Payment Terms; SBA Subordination.",
        "The Call Price is payable by an unsecured, subordinated promissory note of the Company at "
        "the Applicable Federal Rate in equal quarterly installments over four (4) years (prepayable "
        "without penalty; the Manager may elect a lump sum). No payment shall be made, and any such "
        "note is fully subordinated, to the extent it would (i) violate the SBA Loan documents or any "
        "lender covenant, (ii) cause or occur during an event of default under the SBA Loan, or "
        "(iii) reduce liquidity or debt-service coverage below any level the SBA Loan requires. "
        "Suspended payments accrue and resume when permitted.")
    sec(d, "5.9", "Mechanics Pending Repurchase.",
        "From Separation until the Call Right expires or the repurchase closes, the departing "
        "Grantee's Vested Interest is non-voting and not entitled to distributions declared after "
        "Separation (other than tax distributions on allocated income). The Manager's power of "
        "attorney extends to all actions under this Article V.")
    sec(d, "5.10", "Section 83(b) Election; Tax.",
        "As a mandatory condition of receiving a Restricted Interest, each Equity Grantee shall file "
        "a timely election under IRC §83(b) within thirty (30) days of the Effective Date by "
        "certified mail and deliver a copy to the Company within five (5) business days; failure to "
        "file by the deadline results in automatic forfeiture of the entire Restricted Interest to "
        "the Unissued Pool. The Restricted Interests are restricted shares of the Company's single "
        "class of stock issued for services; the Section 83(b) election fixes their nominal fair "
        "market value at the Effective Date and starts the holding period, and is also intended to "
        "ensure each Restricted Interest is treated as outstanding stock so as to preserve the "
        "S-corporation election under IRC §1361. The differential good-leaver / bad-leaver repurchase "
        "terms in this Article V are a bona fide buy-sell arrangement and are not intended to create a "
        "second class of stock under Treasury Reg. §1.1361-1(l)(2)(iii). Each Equity Grantee should "
        "consult independent tax counsel; nothing herein is tax advice from the "
        "Company or the Manager.")
    sec(d, "5.11", "Investor Not Subject to Forfeiture.",
        "Bullard's Percentage Interest is fully earned upon payment and is not subject to any "
        "Forfeiture Condition. Bullard's engagement is limited to the strategic-observer role in "
        "Section 6.9; Bullard has no management authority and serves in no operational, clinical, "
        "administrative, or employment capacity.")

    # ================= ARTICLE VI =================
    art(d, "ARTICLE VI - MANAGEMENT")
    sec(d, "6.1", "Manager-Managed.",
        "The Company is manager-managed. Geoff Schackmann, individually, is the Manager and has sole "
        "authority over day-to-day operations, including hiring/termination, clinical operations, "
        "use of proceeds, budgets, regulatory compliance, vendor contracts, and all matters "
        "necessary to operate the Business, and over the governance of the Acquired Agency as its "
        "sole member's representative. The Manager is entitled to reasonable, market-rate "
        "compensation disclosed to Members annually.")
    sec(d, "6.2", "Reserved Matters Requiring Member Approval.",
        "The Manager shall not, without the prior written consent of Members holding at least sixty "
        "percent (60%) of outstanding Percentage Interests, take any of the following actions; the "
        "parties acknowledge that, after the restatement, no single Member holds 60% and Geoff "
        "Schackmann (39.9%) cannot unilaterally pass a Reserved Matter:")
    for x in [
        "(a) Sell all or substantially all of the Company's assets;",
        "(b) Merge or consolidate the Company with any other entity;",
        "(c) Issue additional membership interests beyond those authorized in Article XII;",
        "(d) Incur indebtedness in excess of $500,000.00 in the aggregate; provided, however, that "
        "the Manager may incur, modify, refinance, restructure, or prepay (x) the SBA Loan, (y) the "
        "Bank Acquisition Loan and the Acquired Agency's related guaranty/collateral, and (z) any "
        "loan that refinances or replaces (x) or (y), without Member consent under this Section "
        "6.2(d), so long as the action is in the ordinary course of operating the Business and is "
        "not itself a Change of Control;",
        "(e) Amend this Agreement, except as provided in Section 11.4;",
        "(f) Dissolve or wind up the Company, except as provided in Article VII; or",
        "(g) Make any Change of Control transaction.",
    ]:
        para(d, x, indent=True)
    para(d, "For the avoidance of doubt, Geoff Schackmann may pass a Reserved Matter only by "
            "combining his 39.9% with the votes of at least two (2) Equity Grantees (66.5% total) or "
            "with Bullard's 19.5% plus at least one (1) Equity Grantee (72.7% total); Bullard holds "
            "no unilateral veto on Reserved Matters.")
    sec(d, "6.3", "Meetings and Votes.",
        "The Manager may call meetings on twenty (20) days' notice; any Member holding at least 25% "
        "may request a meeting. The Investor receives notice of all meetings and has voting rights "
        "on Reserved Matters. Minutes and adopted resolutions are provided to all Members within "
        "five (5) business days.")
    sec(d, "6.4", "No Competing Activity; Business Opportunities.",
        "Each of the Manager, Silas, Dana, and Brad (each an “Operational Member”; this "
        "Section does not apply to Bullard) shall not, within a seventy-five (75) mile radius of "
        "Tyler, Texas, compete with the Business, solicit its patients or referral sources, or "
        "solicit its employees, during membership and for twenty-four (24) months thereafter. "
        "Violation, fraud, willful misconduct, gross negligence, misappropriation, or conviction of "
        "a crime of moral turpitude permits the Manager to expel the Violating Member, whose "
        "Percentage Interest is redistributed pro rata among the remaining non-expelled Members. "
        "Operational Members may engage in hospice activities outside the 75-mile radius.")
    sec(d, "6.5", "Confidential Information.",
        "Each Member shall protect the Company's Confidential Information and not use or disclose it "
        "except as authorized, as public through no fault of the Member, or as legally compelled "
        "with notice. This obligation survives departure for five (5) years.")
    sec(d, "6.6", "Key Person and Succession.",
        "On a Key Person Event affecting Geoff Schackmann (death, permanent disability, or "
        "incapacity), the Company continues under interim management designated by Geoff Schackmann, "
        "LLC's authorized successor for up to ninety (90) days, after which the Members elect a "
        "successor Manager by majority vote. Bullard is notified within five (5) business days and "
        "votes in the successor election; during the interim period, material financial decisions "
        "exceeding $25,000 require Bullard's written approval.")
    sec(d, "6.7", "Investor Information Rights.",
        "Bullard shall receive monthly unaudited financials within thirty (30) days of month-end, "
        "annual reviewed financials within ninety (90) days of year-end, tax returns within fifteen "
        "(15) days of filing, and prompt notice of any material event (regulatory actions, loss of "
        "licensure, or claims over $10,000). Bullard may attend all Member meetings as a voting "
        "Member on Reserved Matters and a non-voting observer otherwise.")
    sec(d, "6.8", "Anti-Dilution Protection.",
        "Bullard's 19.5% interest is calculated on a fully-diluted basis inclusive of the operator "
        "Restricted Interests (13.3% each to Silas, Dana, and Brad) and the 0.7% Unissued Pool. No "
        "additional membership interests shall be issued to new investors that would reduce "
        "Bullard's Percentage Interest below nineteen and one-half percent (19.5%) without Bullard's "
        "prior written consent. Any new investor admitted under Article XII is subject to Bullard's "
        "pro-rata participation right, allowing Bullard to participate at the same price and terms to "
        "maintain his Percentage Interest.")
    sec(d, "6.9", "Investor Strategic Observer Role.",
        "Bullard's role is advisory and commercial only: he may attend meetings (Section 6.7), make "
        "Manager-approved referral-source introductions in compliance with federal and Texas "
        "anti-kickback, Stark, and marketing laws, and receive quarterly strategic briefings. "
        "Nothing in this Section confers management authority, employee status, agency, or authority "
        "to bind the Company. The Manager may suspend or modify Bullard's introduction activities at "
        "any time without affecting his Percentage Interest.")

    # ================= ARTICLE VII =================
    art(d, "ARTICLE VII - DISSOLUTION AND LIQUIDATION")
    sec(d, "7.1", "Dissolution.",
        "The Company dissolves upon (a) an Event of Insolvency of the Company or withdrawal of the "
        "Manager without replacement within sixty (60) days; (b) sale of all or substantially all "
        "assets; (c) a judicial decree of dissolution under Wyoming law; or (d) a Prevailing Vote to "
        "dissolve. The Company continues until its affairs are wound up.")
    sec(d, "7.2", "Winding Up.",
        "Upon dissolution the Manager (or a liquidating trustee) shall complete unfinished business, "
        "collect amounts owed, pay or provide for all debts and obligations (including the SBA Loan, "
        "the Bank Acquisition Loan, and regulatory wind-down obligations under HCSSA, Medicare, and "
        "Medicaid), and distribute remaining assets under Section 4.5.")
    sec(d, "7.3", "No Withdrawal.",
        "No Member may withdraw without a Prevailing Vote; wrongful withdrawal entitles the Company "
        "to all legal remedies, including damages.")
    sec(d, "7.4", "Articles of Dissolution.",
        "Upon completion of winding up, the Manager or liquidating trustee shall file Articles of "
        "Dissolution with the Wyoming Secretary of State under W.S. §17-29-707.")

    # ================= ARTICLE VIII =================
    art(d, "ARTICLE VIII - TRANSFERS OF PERCENTAGE INTEREST")
    sec(d, "8.1", "Non-Complying Transfers Void.",
        "Any Transfer not in complete compliance with this Article VIII is void and not recognized "
        "by the Company.")
    sec(d, "8.2", "Compliance with Law.",
        "No Transfer may violate the Securities Act of 1933 or any securities law; the Manager may "
        "require legal opinions and other documents.")
    sec(d, "8.3", "Lock-Up Period.",
        "For twenty-four (24) months from the Effective Date, no Member may Transfer any interest "
        "without the Manager's consent and a Prevailing Vote.")
    sec(d, "8.4", "Right of First Refusal on Sale.",
        "After the Lock-Up Period, a Member proposing to Transfer to a third party must first offer "
        "the interest to the Company (30 days), then to the other Members pro rata (30 days), on the "
        "same bona fide arm's-length terms, before completing the Transfer within ninety (90) days "
        "on terms no more favorable.")
    sec(d, "8.5", "Permitted Family/Trust Transfers.",
        "A Member may assign, by gift or upon death, to a spouse, child, or revocable living trust, "
        "provided (a) for a trust, the Member retains sole voting control; (b) Geoff Schackmann "
        "remains Manager; and (c) Bullard receives written notice within five (5) business days. No "
        "other exception applies without Bullard's written consent.")
    sec(d, "8.6", "Mandatory Buyout Upon Death, Divorce, Bankruptcy, or Incapacity.",
        "Upon a Triggering Event (death, Event of Insolvency, permanent incapacity, dissolution of a "
        "Member entity, or marital dissolution awarding an interest to a non-Member), the Company "
        "(30 days), then the remaining Members pro rata (30 days), may elect to purchase the "
        "Withdrawing Member's interest; if neither elects, the Company is obligated to purchase. The "
        "price is (FMV of Company assets less liabilities) × Percentage Interest, less amounts "
        "owed to the Company, paid 10% down (20% for death/incapacity) with the balance over four "
        "(4) annual installments at the Prime Rate, closing within one hundred twenty (120) days. "
        "Any such payment is subordinated to the SBA Loan to the extent required by its documents.")
    sec(d, "8.7", "ROFR on Loan Default.",
        "A security interest a Member grants in its interest must give the Company and other Members "
        "the option, on default, to purchase on the terms of Section 8.4.")
    sec(d, "8.8", "Tag-Along Rights.",
        "If Geoff Schackmann proposes a Transfer constituting a Change of Control, Bullard may "
        "participate pro rata on the same terms, with at least twenty (20) days' notice.")
    sec(d, "8.9", "Drag-Along Rights.",
        "If Members holding more than 60% propose a bona fide arm's-length Change of Control, they "
        "may require the remaining Members to Transfer on the same per-unit terms (with "
        "representations limited to title and authority) on at least thirty (30) days' notice.")
    sec(d, "8.10", "Buy-Sell - Texas Shootout.",
        "After the Lock-Up Period, if Schackmann (for himself and Geoff Schackmann) and Bullard "
        "cannot resolve a material dispute after thirty (30) days of good-faith negotiation, either "
        "may invoke the buy-sell by naming an enterprise valuation, from which buyout amounts are "
        "computed pro rata; Unvested Restricted Interests are forfeited at closing without "
        "compensation and Vested interests are treated as fully vested subject to the ROFR. "
        "Notwithstanding the foregoing, NO Member may invoke or enforce this Section 8.10 while any "
        "principal or interest under the SBA Loan remains outstanding; upon payment in full of the "
        "SBA Loan (and any SBA-guaranteed successor/refinancing), this deferral lapses.")
    sec(d, "8.11", "Conditions to Assignment.",
        "A Member may Transfer only if (a) a written assignment satisfactory to the Manager is "
        "delivered; (b) the transferee accepts this Agreement in writing; (c) the Manager obtains a "
        "legal opinion that the Transfer will not violate securities laws and complies with this "
        "Agreement; and (d) a transfer fee covering reasonable expenses is paid.")

    # ================= ARTICLE IX =================
    art(d, "ARTICLE IX - REPRESENTATIONS AND WARRANTIES")
    sec(d, "9.1", "Representations of All Members.",
        "Each Member represents and warrants, as of the Effective Date and continuously: (a) "
        "Authority - full capacity and authority to execute and perform, and that this Agreement is "
        "binding; (b) No Conflicts - execution does not violate any agreement, law, or obligation; "
        "(c) No Litigation - no pending or threatened proceeding that would materially affect the "
        "Company or the Member's performance; and (d) OIG and SAM Exclusion - no Member, and no "
        "individual with a beneficial ownership interest in any Member entity, is excluded, "
        "suspended, debarred, or otherwise ineligible to participate in any federal or state "
        "healthcare program (per the OIG LEIE or SAM.gov databases), with a duty to notify the "
        "Manager within five (5) business days of any threatened or actual exclusion; and (e) S-Corp "
        "Eligibility - each Member is, and shall remain, an S-Corp Eligible Person as defined in "
        "Section 2.27, and shall notify the Manager within five (5) business days of any change that "
        "would cause it to cease to be one.")
    sec(d, "9.2", "Investor Representation - Accredited Investor.",
        "Bullard represents that he is an “accredited investor” under Rule 501(a) of "
        "Regulation D; is acquiring his interest for investment and not with a view to distribution; "
        "has the knowledge and experience to evaluate the investment; understands the interest is a "
        "restricted, unregistered security; can bear the economic risk including total loss; and has "
        "had the opportunity to ask questions of the Manager.")

    # ================= ARTICLE X =================
    art(d, "ARTICLE X - INDEMNIFICATION")
    sec(d, "10.1", "Indemnification of Manager.",
        "The Company shall indemnify and hold harmless the Manager and Geoff Schackmann and "
        "their respective members, managers, officers, employees, and agents from claims arising "
        "from the management or operation of the Company or good-faith actions under this Agreement, "
        "except for fraud, willful misconduct, gross negligence, improper personal benefit, or "
        "breach of this Agreement.")
    sec(d, "10.2", "Procedure.",
        "An Indemnified Party shall give prompt notice; the Company may assume the defense with "
        "reasonably acceptable counsel; no settlement without the Company's consent (not "
        "unreasonably withheld).")
    sec(d, "10.3", "Limitation of Liability.",
        "No Manager is personally liable to the Company or any Member for monetary damages for acts "
        "or omissions as Manager, absent fraud, willful misconduct, gross negligence, or a knowing "
        "violation of law.")

    # ================= ARTICLE XI =================
    art(d, "ARTICLE XI - MISCELLANEOUS")
    sec(d, "11.1", "Books and Records.",
        "The Manager shall maintain complete books and records at the principal place of business, "
        "including a current Member/Percentage-Interest list, the Certificate of Organization and "
        "this Agreement and amendments, minutes and consents, and financial statements and tax "
        "returns for the three most recent fiscal years. Each Member may inspect and copy on five "
        "(5) business days' notice.")
    sec(d, "11.2", "Tax Matters.",
        "The Company has elected, or shall elect, to be taxed as an S corporation for U.S. federal "
        "income tax purposes under IRC §1361 by timely filing (or obtaining late-election relief "
        "under Rev. Proc. 2013-30 for) IRS Form 2553, and the Members shall cooperate in executing "
        "all documents required to make and maintain the election. The Company shall at all times "
        "satisfy the S-corporation eligibility requirements, including: (a) every shareholder is an "
        "eligible shareholder (a U.S. individual, an estate, an eligible trust, or a single-member "
        "disregarded LLC owned by an eligible individual) - accordingly the 39.9% interest is held "
        "directly by Geoff Schackmann, an eligible individual; (b) the Company maintains a single "
        "class of stock with strictly pro-rata allocations and distributions (Sections 4.4-4.5); and "
        "(c) the number of shareholders does not exceed the statutory limit. No Transfer or admission "
        "may be made that would terminate the S-election (Sections 8.5, 8.11, and Article XII). The "
        "Manager shall serve as the Company's tax representative, shall cause the timely preparation "
        "and filing of all returns (including Form 1120-S and Schedule K-1 to each Member) within "
        "seventy-five (75) days of each fiscal year-end. The Acquired Agency, as a wholly owned "
        "single-member subsidiary, shall be a disregarded entity (or a qualified subchapter S "
        "subsidiary, if the Manager so elects) whose items are reported by the Company.")
    sec(d, "11.3", "Securities Compliance.",
        "The membership interests are unregistered, issued in reliance on Section 4(a)(2) and/or Rule "
        "506(b) of Regulation D. The Company shall make required Form D and state notice filings. No "
        "Member may Transfer interests in violation of securities laws.")
    sec(d, "11.4", "Amendment.",
        "This Agreement may be amended by a Prevailing Vote; provided that (a) any amendment "
        "adversely affecting Bullard's rights under Sections 4.2, 4.5, 3.15, 6.7, 6.8, or 6.9, or "
        "the Article VIII transfer restrictions, requires Bullard's prior written consent; (b) any "
        "amendment causing Members to lose limited liability, or changing the Company's tax "
        "classification, requires the written consent of all Members; and (c) no amendment may "
        "reduce any Member's Percentage Interest without that Member's express written consent.")
    sec(d, "11.5", "Entire Agreement.",
        "This Agreement, together with the Subscription Agreement between the Company and Bullard "
        "(incorporated by reference), constitutes the entire agreement of the parties as to the "
        "Company's internal governance and supersedes all prior agreements, including the Original "
        "Agreement. The MIPA governs the purchase of the Acquired Agency and is a separate "
        "transaction document; in the event of a conflict between this Agreement and the MIPA as to "
        "the Company's internal affairs, this Agreement controls.")
    sec(d, "11.6", "Further Assurances; Certificates and Filings.",
        "The Members shall execute all documents required to comply with the Act and applicable law, "
        "including any SBA, lender, intercreditor, subordination, and guaranty documents reasonably "
        "necessary to effect the financing described in Sections 1.7 and 3.16.")
    sec(d, "11.7", "Partial Invalidity.", "Invalidity of any provision does not affect the "
        "remainder, which remains in full force and effect.")
    sec(d, "11.8", "Governing Law; Dispute Resolution.",
        "This Agreement is governed by Wyoming law, without regard to conflicts principles. Disputes "
        "not resolved within thirty (30) days of written notice shall be submitted to binding "
        "arbitration in Smith County, Texas (Tyler) under the AAA Commercial Arbitration Rules.")
    sec(d, "11.9", "Attorney's Fees.", "The prevailing party in any proceeding to enforce this "
        "Agreement is entitled to reasonable attorney's fees and costs at trial and on appeal.")
    sec(d, "11.10", "Injunctive Relief.", "Breach of Sections 6.4 or 6.5 entitles the Company to "
        "seek injunctive relief without bond.")
    sec(d, "11.11", "Notices.", "Notices shall be in writing to the physical street addresses (not "
        "P.O. boxes) or emails in the signature block; hand-delivered and email notices are "
        "effective on delivery, mailed notices three (3) days after deposit.")
    sec(d, "11.12", "Counterparts; Electronic Signatures.", "This Agreement may be executed in "
        "counterparts and by electronic signature (including DocuSign and PDF).")
    sec(d, "11.13", "Titles; Pronouns.", "Captions are for convenience only; pronouns include all "
        "genders and the singular and plural as context requires.")

    # ================= ARTICLE XII =================
    art(d, "ARTICLE XII - FUTURE INVESTOR ADMISSION")
    sec(d, "12.1", "Manager Authority to Admit New Investors.",
        "The Manager may, in the Manager's sole discretion and without a Member vote, admit one or "
        "more new investors (each a “New Investor”), subject only to (a) the dilution floor "
        "in Section 12.3 and (b) the anti-dilution and pro-rata rights of Bullard under Section 6.8.")
    sec(d, "12.2", "Admission Mechanics.",
        "A New Investor must (a) execute a counterpart signature page; (b) execute a subscription "
        "agreement in a form approved by the Manager, including accredited-investor and OIG/SAM "
        "representations consistent with Section 9.1(d) and 9.2; (c) pay the agreed purchase price; "
        "and (d) receive a written admission notice. The Manager shall then update Exhibit A.")
    sec(d, "12.3", "Hard Dilution Floor - Bullard Minimum Interest.",
        "Notwithstanding Section 12.1, the Manager may not admit any New Investor or issue any "
        "additional Percentage Interests if doing so would reduce Bullard's Percentage Interest "
        "below nineteen and one-half percent (19.5%) without Bullard's prior written consent. This "
        "floor applies cumulatively to all admissions.")
    sec(d, "12.4", "Healthcare-Eligibility Certification.",
        "As a condition of admission, each New Investor shall certify, and maintain on a continuing "
        "basis, that it (and its beneficial owners) is not excluded, suspended, debarred, or "
        "otherwise ineligible to participate in any federal or state healthcare program, and shall "
        "notify the Manager within five (5) business days of any change. A breach permits the "
        "Manager to require a mandatory transfer or buyout under Section 8.6.")
    sec(d, "12.5", "Pro-Rata Participation Right.",
        "Before admitting any New Investor, the Manager shall give Bullard at least fifteen (15) "
        "days' notice specifying the proposed interest, price, and terms; Bullard may elect within "
        "fifteen (15) days to purchase up to his pro-rata share on the same terms to maintain his "
        "Percentage Interest. Non-response waives the right for that issuance only.")
    sec(d, "12.6", "New Investor Rights.",
        "Unless otherwise agreed in writing at admission, New Investors receive the reporting rights "
        "of Section 6.7 and voting rights on Reserved Matters consistent with their Percentage "
        "Interest, but not the strategic-observer rights of Section 6.9, the anti-dilution "
        "protection of Section 6.8, or the dilution floor of Section 12.3.")

    # ================= SIGNATURES =================
    art(d, "SIGNATURES")
    para(d, "IN WITNESS WHEREOF, the Members have executed this Amended and Restated Operating "
            "Agreement effective as of the date set forth above.")
    para(d, "")
    blocks = [
        ("MEMBER AND MANAGER - GEOFF SCHACKMANN (Interest: 39.9%)",
         "Geoff Schackmann, individually, as a Member and as the Company's Manager",
         "4602 E Cheery Lynn Rd, Phoenix, Arizona 85018  |  geoff@azaleahospice.com  |  480-495-5474"),
        ("SPOUSAL / COMMUNITY-PROPERTY CONSENT",
         "Mary Elizabeth Burcham, spouse of Geoff Schackmann - consents to this Agreement and to Geoff Schackmann's direct ownership of the 39.9% Membership Interest, and waives any community-property claim inconsistent with this Agreement and the Company's S-corporation eligibility (single eligible shareholder).",
         "Address: ______________________________________________"),
        ("MEMBER / INVESTOR - JAMES E. BULLARD (Interest: 19.5%)",
         "James E. Bullard",
         "13910 Indiana Ave, Suite 300, Lubbock, Texas 79423  |  jimbullard01@aol.com"),
        ("EQUITY GRANTEE - SILAS R. SHELTON (Interest: 13.3% Restricted; entire interest subject to Forfeiture Conditions)",
         "Silas R. Shelton, Executive Director - I acknowledge my obligation to file a timely §83(b) election within 30 days of the Effective Date.",
         "1503 Lake Park Circle, Hideaway, Texas 75771"),
        ("EQUITY GRANTEE - DANA L. DAVENPORT (Interest: 13.3% Restricted; entire interest subject to Forfeiture Conditions)",
         "Dana L. Davenport, Director of Nursing - I acknowledge my obligation to file a timely §83(b) election within 30 days of the Effective Date.",
         "Address: ______________________________________  [REQUIRED: physical street address - a P.O. box is not valid notice under §11.11]"),
        ("EQUITY GRANTEE - BRADLEY GENE WOODARD (Interest: 13.3% Restricted; entire interest subject to Forfeiture Conditions)",
         "Bradley Gene Woodard, Director of Sales - I acknowledge my obligation to file a timely §83(b) election within 30 days of the Effective Date.",
         "421 W Cumberland Rd, Apt 403, Tyler, Texas 75703"),
    ]
    for head, name, addr in blocks:
        para(d, head, bold=True, size=10)
        para(d, "By: _________________________________________   Date: ______________", size=10)
        para(d, name, size=9.5)
        para(d, addr, size=9.5)
        para(d, "")

    # ================= EXHIBIT A =================
    art(d, "EXHIBIT A - CAPITALIZATION TABLE")
    captable(d)
    para(d, "*Each Restricted Interest of Silas R. Shelton, Dana L. Davenport, and Bradley Gene "
            "Woodard is issued and outstanding at 13.3%, the entire amount of which is at risk and "
            "forfeitable until Vested under Article V: a 4.9% Initial Base that time-vests on "
            "continuous service (four-year schedule, one-year cliff) and an 8.4% Earn-Up Portion "
            "subject to dual-trigger (time and performance) vesting. No portion is Vested at the "
            "Effective Date. The 0.7% Unissued Pool is reserved for future grants and admissions "
            "under Article XII. Bullard's 19.5% is calculated on a fully-diluted basis inclusive of "
            "the operator Restricted Interests and the Unissued Pool.", italic=True, size=9)

    d.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    build()
