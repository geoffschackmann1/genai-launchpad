"""
Single-Member Operating Agreement of HICKORY HOSPICE LLC (Texas), the wholly owned
operating subsidiary of Tyler Hospice Hold, LLC, adopted at the CHOW closing.

Reuses the docx helpers from build_oa_restatement.py (same directory, self-contained).
"""
import os
from build_oa_restatement import (doc_new, title, sub, art, sec, para, rule, NAVY, GREY)

OUT = os.path.join(os.path.dirname(__file__), "08_entity_documents",
                   "Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx")


def build():
    d = doc_new()
    title(d, "HICKORY HOSPICE LLC")
    sub(d, "OPERATING AGREEMENT (SINGLE MEMBER)", size=12, color=NAVY, italic=False)
    sub(d, "A Texas Limited Liability Company  -  d/b/a Azalea Hospice & Palliative Care")
    sub(d, "Sole Member: Tyler Hospice Hold, LLC (a Wyoming limited liability company)")
    sub(d, "Adopted effective as of the Closing under the MIPA  |  ________________, 2026")
    sub(d, "DRAFT FOR COUNSEL REVIEW AND EXECUTION.", size=9)
    rule(d)

    para(d, "This Operating Agreement (this “Agreement”) of Hickory Hospice LLC, a Texas "
            "limited liability company (the “Company”), is adopted effective as of the "
            "closing of the change-of-ownership transaction under the Membership Interest Purchase "
            "Agreement (the “MIPA”), by Tyler Hospice Hold, LLC, a Wyoming limited liability "
            "company, as the sole member of the Company (the “Sole Member”).")
    para(d, "RECITALS", bold=True)
    para(d, "A.  Pursuant to the MIPA, the Sole Member purchased one hundred percent (100%) of the "
            "membership interests of the Company in a change-of-ownership (CHOW) transaction, "
            "following which the Company retains its Medicare-certified hospice provider number and "
            "its HCSSA licensure and operates, doing business as “Azalea Hospice & Palliative "
            "Care.”")
    para(d, "B.  The Sole Member desires to set forth the governance of the Company as a "
            "single-member, manager-managed Texas limited liability company, and to authorize the "
            "Company's guaranty of, and grant of collateral for, the Sole Member's SBA 7(a) loan.")

    art(d, "ARTICLE I - FORMATION")
    sec(d, "1.1", "Formation and Name.",
        "The Company is organized under the Texas Business Organizations Code (the “TBOC”) "
        "as “Hickory Hospice LLC.” The Company shall maintain its assumed name (d/b/a) "
        "“Azalea Hospice & Palliative Care” by filing the required assumed-name "
        "certificates with the Texas Secretary of State and the appropriate county.")
    sec(d, "1.2", "Registered Office and Agent.",
        "The Company shall continuously maintain a registered office and registered agent in Texas "
        "as required by the TBOC, as designated by the Sole Member.")
    sec(d, "1.3", "Purpose.",
        "The Company's purpose is to own and operate a Medicare- and Medicaid-certified hospice and "
        "palliative care agency licensed as an HCSSA under Texas law, and to engage in any lawful "
        "activity incidental thereto.")
    sec(d, "1.4", "Principal Office; Term.",
        "The Company's principal office is in Tyler, Texas, or as the Manager designates. The "
        "Company has perpetual existence until dissolved under Article VII.")
    sec(d, "1.5", "Fiscal Year.", "The fiscal year is the calendar year.")

    art(d, "ARTICLE II - SOLE MEMBER; CAPITAL")
    sec(d, "2.1", "Sole Member.",
        "Tyler Hospice Hold, LLC is the sole member of the Company and owns 100% of the membership "
        "interests. No other person has any membership interest, economic interest, or right to "
        "become a member except as the Sole Member admits in writing.")
    sec(d, "2.2", "Capital Contribution.",
        "The Sole Member's capital contribution consists of the consideration paid for the "
        "membership interests under the MIPA and any additional contributions the Sole Member elects "
        "to make. The Sole Member is not obligated to make additional contributions.")
    sec(d, "2.3", "Limited Liability.",
        "The Sole Member is not personally liable for the Company's debts, obligations, or "
        "liabilities solely by reason of being a member, except as expressly assumed in writing or "
        "as required by law.")
    sec(d, "2.4", "Distributions.",
        "The Company may distribute cash and property to the Sole Member at such times and in such "
        "amounts as the Sole Member determines, subject to the TBOC solvency limitation, any "
        "restriction in the SBA Loan documents and lender covenants, and the maintenance of adequate "
        "reserves for clinical, regulatory, and wind-down obligations under HCSSA, Medicare, and "
        "Medicaid.")

    art(d, "ARTICLE III - MANAGEMENT")
    sec(d, "3.1", "Manager-Managed.",
        "The Company is manager-managed. The initial Manager is Geoff Schackmann, who also serves as "
        "Manager of the Sole Member, or such other person as the Sole Member designates. The Manager "
        "has full authority to conduct the Company's business, including clinical and administrative "
        "operations, maintenance of licensure and Medicare/Medicaid enrollment, hiring and "
        "termination, contracts, and regulatory compliance.")
    sec(d, "3.2", "Reserved to the Sole Member.",
        "Notwithstanding Section 3.1, the following require the Sole Member's approval: (a) sale of "
        "all or substantially all of the Company's assets; (b) merger, conversion, or dissolution; "
        "(c) incurrence of indebtedness outside the ordinary course other than the SBA Loan, the "
        "Bank Acquisition Loan, and refinancings thereof; (d) admission of any new member; and "
        "(e) any amendment of this Agreement.")
    sec(d, "3.3", "Officers.",
        "The Manager may appoint officers (including an Administrator and a Director of Nursing as "
        "required by HCSSA rules) and delegate authority to them, subject to the Manager's "
        "supervision.")
    sec(d, "3.4", "Indemnification; Exculpation.",
        "The Company shall indemnify the Manager, the Sole Member, and their respective officers, "
        "managers, employees, and agents to the fullest extent permitted by the TBOC, except for "
        "fraud, willful misconduct, gross negligence, or a knowing violation of law. No such person "
        "is liable to the Company for monetary damages except for the foregoing excluded conduct.")

    art(d, "ARTICLE IV - TAX")
    sec(d, "4.1", "Disregarded Entity.",
        "The Company is a single-member limited liability company that is disregarded as an entity "
        "separate from its owner for U.S. federal income tax purposes under Treasury Reg. "
        "§301.7701-3, and its items of income, gain, loss, deduction, and credit are reported by "
        "the Sole Member (which is taxed as a partnership). No election to be treated as an "
        "association taxable as a corporation shall be made without the Sole Member's written "
        "direction and consultation with the Company's tax advisors.")

    art(d, "ARTICLE V - SBA LOAN GUARANTY AND COLLATERAL AUTHORIZATION")
    sec(d, "5.1", "Authorization of Guaranty and Lien.",
        "The Company is hereby authorized and directed to (a) execute and deliver an unconditional "
        "guaranty of the Sole Member's SBA 7(a) loan (the “SBA Loan”); (b) grant to the SBA "
        "lender a first-priority security interest in and lien on substantially all of the Company's "
        "assets, including accounts receivable and Medicare/Medicaid payment rights to the extent "
        "permitted by applicable law and program rules; and (c) execute all related loan, security, "
        "control, and SBA authorization documents. The Manager is authorized to execute and deliver "
        "all such documents on the Company's behalf.")
    sec(d, "5.2", "Corporate Benefit.",
        "The Sole Member and the Manager have determined that the SBA Loan and the related Bank "
        "Acquisition Loan provide direct and substantial benefit to the Company by funding the "
        "Company's acquisition, working capital, and operations, and that the guaranty and collateral "
        "in Section 5.1 are in the Company's best interest and supported by adequate consideration.")
    sec(d, "5.3", "Subordination of the Bank Acquisition Loan.",
        "Any guaranty, lien, or obligation of the Company in respect of the Bank Acquisition Loan "
        "shall be subordinate to the SBA Loan to the extent required by the SBA lender, and the "
        "Manager is authorized to execute any intercreditor or subordination agreement to effect "
        "that priority.")

    art(d, "ARTICLE VI - BOOKS, RECORDS, AND COMPLIANCE")
    sec(d, "6.1", "Books and Records.",
        "The Manager shall maintain complete books and records and all records required by the TBOC, "
        "HCSSA, Medicare, and Medicaid, available to the Sole Member on request.")
    sec(d, "6.2", "Healthcare Compliance.",
        "The Company shall maintain its licensure, certification, and enrollment, and shall comply "
        "with all applicable federal and Texas healthcare laws, including the Anti-Kickback Statute, "
        "the Stark Law, and HIPAA, and shall ensure no owner, officer, or manager is excluded under "
        "the OIG LEIE or SAM.gov databases.")

    art(d, "ARTICLE VII - DISSOLUTION")
    sec(d, "7.1", "Dissolution.",
        "The Company dissolves upon the written election of the Sole Member or entry of a judicial "
        "decree of dissolution. Upon dissolution, the Manager shall wind up the Company's affairs, "
        "pay or provide for all debts and obligations (including the SBA Loan, the Bank Acquisition "
        "Loan, and regulatory wind-down obligations), and distribute remaining assets to the Sole "
        "Member, and shall file a certificate of termination with the Texas Secretary of State.")

    art(d, "ARTICLE VIII - MISCELLANEOUS")
    sec(d, "8.1", "Governing Law.",
        "This Agreement is governed by the laws of the State of Texas, without regard to conflicts "
        "principles.")
    sec(d, "8.2", "Amendment.",
        "This Agreement may be amended only by a written instrument executed by the Sole Member.")
    sec(d, "8.3", "Entire Agreement; Severability; Counterparts.",
        "This Agreement is the entire agreement of the Sole Member as to the Company's governance; "
        "if any provision is invalid, the remainder continues in effect; and this Agreement may be "
        "executed in counterparts and by electronic signature.")

    art(d, "SIGNATURES")
    para(d, "IN WITNESS WHEREOF, the Sole Member has adopted this Agreement effective as of the date "
            "set forth above.")
    para(d, "")
    para(d, "SOLE MEMBER - TYLER HOSPICE HOLD, LLC", bold=True, size=10)
    para(d, "By: _________________________________________   Date: ______________", size=10)
    para(d, "Geoff Schackmann, Manager of Tyler Hospice Hold, LLC", size=9.5)
    para(d, "")
    para(d, "MANAGER - HICKORY HOSPICE LLC", bold=True, size=10)
    para(d, "By: _________________________________________   Date: ______________", size=10)
    para(d, "Geoff Schackmann, Manager", size=9.5)

    d.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    build()
