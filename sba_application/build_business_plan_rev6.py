"""Generate Azalea SBA Business Plan Rev 6.00.

Rev 6.00 updates the plan from the terminated Hickory Hospice acquisition to the
Refuge Hospice, LLC acquisition (MIPA dated 7/14/2026) and restructures the
financing to a single $500,000 SBA 7(a) loan plus a $250,000 equity injection
($750,000 total project). All financial figures are taken from the Rev 3.00
dynamic proforma (Azalea_Hospice_Proforma_Rev3.00_DYNAMIC.xlsx, QA 20/20):

  Census (EOM):        24 @ M2 / 34 @ M6 / 40 @ M12 / 50 @ M24 / 56 @ M36
  Net patient revenue: $1,932,877 / $2,891,641 / $3,490,623   (Y1/Y2/Y3)
  EBITDA:              $262,951 (13.6%) / $680,224 (23.5%) / $975,925 (28.0%)
  Month-12 EBITDA margin: 22.3%
  Net income:          $100,941 / $472,928 / $746,549
  Modeled debt:        $500K at 6% over 36 months, $15,211/mo from Oct 2026
                       (conservative; an actual SBA 7(a) at ~10.5% / 10 yr is
                       ~$6,745/mo)

Section structure, tone, and still-accurate content (team bios, market
demographics, service descriptions, operations, compliance) are ported from
Rev 5.00 (build_documents.py::business_plan()).

Run:  python3 sba_application/build_business_plan_rev6.py
"""
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))
NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x60, 0x60, 0x60)

OUT = "03_business_plan/Azalea_SBA_Business_Plan_Rev6.00.docx"


# ------------------------------------------------------------------ helpers
def new_doc():
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(10.5)
    return d


def h1(d, text):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = NAVY
    p.space_after = Pt(4)
    return p


def h2(d, text):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(11.5)
    r.font.color.rgb = NAVY
    p.space_before = Pt(8)
    p.space_after = Pt(2)
    return p


def para(d, text, italic=False, size=10.5, color=None, bold=False):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.italic = italic
    r.bold = bold
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def bullet(d, text):
    return para(d, "  - " + text)


def table(d, headers, rows, widths=None):
    t = d.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for j, htext in enumerate(headers):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(htext)
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


# ================================================================ the plan
def business_plan_rev6():
    d = new_doc()

    # ---- Cover ----
    h1(d, "AZALEA HOSPICE & PALLIATIVE CARE")
    para(d, "Business Plan - SBA 7(a) Application", bold=True, size=13, color=NAVY)
    para(d, "Tyler Hospice Hold, LLC, operating as Azalea Hospice & Palliative Care", color=GREY, size=10)
    para(d, "dba Azalea Hospice & Palliative Care - EIN 41-4966640 - Wyoming holding company acquiring "
            "Refuge Hospice, LLC (Texas)", size=9, color=GREY)
    para(d, "Rev. 6.00 - July 2026  (supersedes Rev 5.00 dated June 2026; acquisition target changed to "
            "Refuge Hospice, LLC, financing restructured to a single SBA 7(a) loan, and all financials "
            "reconciled to the Rev 3.00 dynamic proforma)", italic=True, size=9, color=GREY)
    para(d, "")
    para(d, "Acquisition of Refuge Hospice, LLC - a Texas hospice certified for both Medicare and Medicaid "
            "and ready to bill from day one - operating as Azalea Hospice & Palliative Care, led by an "
            "experienced East-Texas clinical team returning to a census level it has operated at before in "
            "this market.")
    para(d, "")
    h2(d, "Headline metrics")
    table(d, ["Metric", "Value", "Detail"],
          [["SBA loan request", "$500,000", "Single SBA 7(a) loan - retires the seller balance in September 2026 and funds ramp working capital"],
           ["Purchase price", "$500,000", "Refuge Hospice, LLC - dual Medicare + Medicaid certification (CMS CCN effective 1/8/2024); MIPA dated 7/14/2026"],
           ["Total project", "$750,000", "$500K SBA loan + $250K cash equity injection"],
           ["Year-1 EBITDA", "$262,951", "13.6% margin, before debt service - growing to $975,925 (28.0%) by Year 3; month-12 margin 22.3%"],
           ["Break-even census", "~17-18 patients", "Crossed in month 2-3 of the ramp"],
           ["Equity injection", "$250,000", "33% of project (Bullard $195K + Schackmann $55K) - well above the 10% SOP 50 10 8 minimum"]],
          widths=[1.6, 1.3, 3.6])
    para(d, "Confidential - prepared exclusively for the SBA 7(a) loan application. All figures are computed "
            "from the Rev 3.00 dynamic proforma (Azalea_Hospice_Proforma_Rev3.00_DYNAMIC.xlsx). Do not "
            "distribute without written consent.", italic=True, size=8, color=GREY)

    # ---- 01 Executive Summary ----
    d.add_page_break()
    h1(d, "01 - Executive Summary")
    h2(d, "Loan request")
    para(d, "Tyler Hospice Hold, LLC requests a $500,000 SBA 7(a) loan to complete the acquisition of Refuge "
            "Hospice, LLC - a Texas hospice certified for both Medicare and Medicaid (CMS CCN effective "
            "January 8, 2024) and ready to bill from day one - and to fund working capital through the census "
            "ramp, operating as Azalea Hospice & Palliative Care. The $500,000 purchase, under a Membership "
            "Interest Purchase Agreement dated July 14, 2026 (in final negotiation), is structured in two "
            "steps to comply with the Medicare 36-month rule at 42 CFR 424.550(b): 49% of the membership "
            "interests transfer at closing (approximately August 1, 2026) against a $125,000 down payment, "
            "and the remaining 51% transfer on January 15, 2027, the first date past the 36-month mark from "
            "initial certification. The $375,000 balance is seller-financed at 6% interest ($31,250 per month "
            "beginning September 1, 2026, with a final payment of approximately $258,489 on January 15, 2027) "
            "and is prepayable without penalty. SBA loan proceeds retire the seller balance in September "
            "2026, converting short-term seller financing into permanent SBA financing before the ramp "
            "deepens. The loan is paired with a $250,000 cash equity injection for a total project of "
            "$750,000 (a 33% injection). Once the seller balance is retired, the business carries a single "
            "debt: the SBA loan.")
    h2(d, "Use of funds")
    table(d, ["Use", "Amount", "Detail"],
          [["Acquisition of Refuge Hospice, LLC", "$500,000", "100% of membership interests in two steps per 42 CFR 424.550(b); $125,000 down at closing, $375,000 seller balance retired by SBA proceeds in September 2026"],
           ["Closing and startup costs", "$15,000", "Legal, licensure filings, EMR setup, initial supplies"],
           ["Working-capital reserve", "$235,000", "Funds ramp payroll ahead of the Medicare payment lag (NOE-to-cash ~30-60 days)"],
           ["Total uses", "$750,000", "Funded by SBA $500K + equity $250K"]],
          widths=[2.5, 1.1, 2.9])
    h2(d, "Business overview")
    para(d, "Azalea launches the acquired Refuge agency in the Tyler / Smith County market (CBSA 46340) "
            "under an experienced local team with established referral relationships. Refuge is certified "
            "for both Medicare and Medicaid - a scarce license position under the current Texas Medicaid "
            "enrollment freeze and the CMS nationwide hospice enrollment moratorium - and its Medicaid "
            "certification adds nursing-facility room and board billing, widening the referral base to "
            "nursing homes. Census builds from the team's referral pipeline to 24 patients by month 2, 34 by "
            "month 6, and 40 by month 12, reaching 50 by month 24 and 56 by month 36 - a return to census "
            "levels the team has operated at before in this market. Revenue is Medicare Routine Home Care "
            "per-diem at a blended net rate of approximately $172 per patient-day, with Medicaid "
            "nursing-facility room and board passed through at the standard rate.")
    h2(d, "Management team")
    para(d, "Azalea is led by an already-seated team: Geoff Schackmann (Manager - multi-hospice "
            "operator), Silas Shelton (Administrator / Executive Director), Dana Davenport (Director of "
            "Nursing), and Bradley Woodard (Community Liaison / Director of Sales - 25+ years East-Texas "
            "hospice business development), supported by RN case managers, hospice aides, a PRN visit pool, "
            "and 1099 Medical Directors. Officer salaries are separately stated in the projections: "
            "Administrator $150,000, Director of Nursing $120,000, Community Liaison $120,000, Office "
            "Manager $60,000.")
    h2(d, "Repayment case")
    table(d, ["", "Year 1", "Year 2", "Year 3"],
          [["Net patient revenue", "$1,932,877", "$2,891,641", "$3,490,623"],
           ["EBITDA (before debt service)", "$262,951", "$680,224", "$975,925"],
           ["EBITDA margin", "13.6%", "23.5%", "28.0%"],
           ["Net income", "$100,941", "$472,928", "$746,549"]],
          widths=[2.2, 1.4, 1.4, 1.4])
    para(d, "The operation crosses break-even census (~17-18 patients) in month 2-3 and reaches a 22.3% "
            "EBITDA margin by month 12. The projections deliberately model debt service far heavier than the "
            "actual SBA terms: $500,000 at 6% amortized over just 36 months ($15,211 per month from October "
            "2026, $182,532 per year). Even against that compressed schedule, EBITDA covers debt service "
            "1.8x in Year 1, 3.7x in Year 2, and 5.3x in Year 3 - all above the 1.25x floor. An actual SBA "
            "7(a) at ~10.5% over 10 years is approximately $6,745 per month ($80,940 per year), so real "
            "coverage will be materially better than modeled (Year 2 EBITDA covers actual-terms service "
            "roughly 8x). A separate 5% of net patient revenue contingency is deducted below EBITDA in the "
            "cash flow, and the $235,000 working-capital reserve plus a $250,000 revolver keep cash positive "
            "in the base case and in the documented downside case.")

    # ---- 02 Company Description ----
    d.add_page_break()
    h1(d, "02 - Company Description")
    h2(d, "2.1 Legal structure")
    para(d, "Tyler Hospice Hold, LLC is a Wyoming LLC (formed 2026; EIN 41-4966640), foreign-qualified to do "
            "business in Texas. As the borrower and holding entity, it is acquiring 100% of Refuge Hospice, "
            "LLC - a Texas LLC, HCSSA-licensed and certified for both Medicare and Medicaid (CMS CCN "
            "effective January 8, 2024) - which will operate as Azalea Hospice & Palliative Care. Because "
            "Medicare bars a change of majority ownership within 36 months of initial certification (42 CFR "
            "424.550(b)), the purchase transfers 49% of the membership interests at closing (approximately "
            "August 1, 2026) and the remaining 51% on January 15, 2027, the first date past the 36-month "
            "mark. Refuge's existing Medicare and Medicaid certifications and Texas license make the agency "
            "billing-ready from day one, with no new-provider 855A enrollment wait. The Company elects "
            "S-corporation tax treatment under IRC Section 1361 (IRS Form 2553) for federal income tax "
            "purposes (filing Form 1120-S, single class of stock, pro-rata distributions); Refuge Hospice, "
            "LLC is a disregarded entity (or QSub) of the S-corp.")
    h2(d, "2.2 Ownership & capitalization")
    table(d, ["Member / source", "Interest", "Role & structure"],
          [["Geoff Schackmann (individual)", "39.9%", "Manager - direct individual owner; 39.9% interest is Arizona community property (spouse Mary Elizabeth Burcham ~19.95% community-property interest, passive, under 20%); Geoff provides the SBA personal guaranty and contributes $55,000 of the equity injection"],
           ["James E. Bullard", "19.5%", "Passive minority investor - $195,000 cash capital contribution (first $100,000 wired 5/7/2026); protective minority rights only, no operational control; under 20%, no guaranty"],
           ["Silas R. Shelton", "13.3%", "Administrator / Executive Director - Restricted Interest (fully forfeitable; 4.9% time-vested base + 8.4% dual-trigger earn-up)"],
           ["Dana L. Davenport", "13.3%", "Director of Nursing - Restricted Interest (fully forfeitable; 4.9% time-vested base + 8.4% dual-trigger earn-up)"],
           ["Bradley G. Woodard", "13.3%", "Community Liaison / Director of Sales - Restricted Interest (fully forfeitable; 4.9% time-vested base + 8.4% dual-trigger earn-up)"],
           ["Unissued pool", "0.7%", "Reserved for future grants"]],
          widths=[2.2, 0.9, 3.4])
    h2(d, "Sources of capital & SBA guaranty")
    para(d, "The $750,000 project is funded by the $500,000 SBA 7(a) loan and a $250,000 cash equity "
            "injection: $195,000 from investor James E. Bullard in exchange for a direct 19.5% membership "
            "interest, and $55,000 from Geoff Schackmann. The injection equals 33% of total project cost - "
            "well above the 10% minimum required under SBA SOP 50 10 8. Mr. Bullard's funds are his own "
            "savings (non-borrowed), verified by his bank statements; $100,000 was wired on May 7, 2026 and "
            "the balance follows on a committed schedule. Because he holds less than 20% and exercises no "
            "operational control, no SBA personal guaranty is required of him (13 CFR 120.160). There is no "
            "companion bank loan: the prior dual-debt structure has been superseded, and after the SBA "
            "proceeds retire the seller balance in September 2026 the business carries a single debt. The "
            "SBA 7(a) loan is guaranteed by Geoff Schackmann (personal guaranty, controlling Manager and "
            "sole 20%+ owner - direct individual 39.9% owner) and by Refuge Hospice, LLC (corporate guaranty "
            "plus a lien on its assets and receivables, as the operating subsidiary).")
    h2(d, "2.3 Acquired platform & validated economics")
    para(d, "Refuge Hospice, LLC holds what the current regulatory environment has made scarce: a Texas "
            "hospice license certified for both Medicare and Medicaid, ready to bill immediately. Texas has "
            "an effective freeze on new Medicaid hospice enrollment, and CMS imposed a nationwide hospice "
            "enrollment moratorium (effective May 13, 2026) that names Texas - so dual-certified licenses "
            "cannot be newly created and comparable licenses trade at $400-500K and above. The Medicaid "
            "certification also adds nursing-facility room and board billing capability, widening the "
            "referral base to nursing homes. The operating economics are not speculative: the model's "
            "per-diem rates, patient-day costs, and census ramp are benchmarked to the leadership team's own "
            "prior Tyler-market book (~22 ADC and ~$118K net collections per month, April-June 2025 "
            "actuals), so the ramp is a return to a proven census level under the same team, in the same "
            "market.")
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
            "days 61+), wage-index adjusted for Smith County (Tyler CBSA 46340). With approximately 15% of "
            "patient-days at the higher first-60-day rate, the blended net rate is approximately $172 per "
            "patient-day after sequestration and billing fees - escalated 2.5% per year (CMS hospice rates "
            "have risen every year since 2010). Medicaid nursing-facility room and board is billed and "
            "passed through at the standard rate.")
    h2(d, "3.2 License scarcity - the enrollment moratorium")
    para(d, "Two regulatory actions have effectively closed the door on new entrants and made Refuge's "
            "dual certification a strategic asset. First, Texas maintains an effective freeze on new "
            "Medicaid hospice enrollment, so a new provider cannot obtain the Medicaid certification that "
            "unlocks nursing-facility room and board billing. Second, CMS imposed a nationwide hospice "
            "enrollment moratorium effective May 13, 2026, naming Texas among the states of concern - "
            "halting new Medicare hospice enrollments altogether. The only way into the market is to "
            "acquire an existing certified license, and dual Medicare + Medicaid licenses trade at "
            "$400-500K and above. Azalea's $500,000 purchase is therefore at market for the license alone, "
            "and the moratorium that constrains supply also protects Azalea from new competition during "
            "the ramp.")
    h2(d, "3.3 Tyler market - demographics and service area")
    para(d, "Tyler, Texas (city pop. ~110,000; metro ~245,000) is the principal city of Smith County and "
            "the regional healthcare hub for East Texas. Azalea operates from a principal office in the "
            "Tyler area with a ~45-mile service radius.")
    bullet(d, "Service-area population: ~350,000 across Smith, Cherokee, Henderson, Rusk, Van Zandt, and Wood counties")
    bullet(d, "Medicare-eligible population (65+): ~60,000+")
    bullet(d, "Hospice utilization: 50-55% of Medicare decedents (above the ~50% national average) - a mature, accepting market")
    bullet(d, "Population growth: Tyler has grown 15%+ since 2010, driven by healthcare-sector expansion and retirement in-migration")
    bullet(d, "Dominant religious demographic: conservative-Christian - strong alignment with Azalea's faith-based service pillars")
    h2(d, "3.4 Major hospital systems - referral drivers")
    bullet(d, "UT Health Tyler - 502-bed Level 1 trauma center, regional referral hub")
    bullet(d, "CHRISTUS Trinity Mother Frances - 438-bed acute care flagship")
    bullet(d, "UT Health North East - 153-bed acute care")
    bullet(d, "CHRISTUS Louis & Peaches Owen Heart Hospital - cardiac")
    bullet(d, "Texas Spine & Joint Hospital - 64-bed surgical")
    bullet(d, "UT Health Jacksonville - 90-bed acute care (Cherokee County)")
    h2(d, "3.5 Skilled nursing & assisted living landscape")
    para(d, "Smith County and surrounding counties hold ~40+ skilled nursing facilities (SNFs) and ~30+ "
            "assisted living facilities (ALFs). These are a concentrated, high-volume referral source: "
            "residents with terminal diagnoses frequently elect hospice, and facility relationships drive "
            "consistent admission volume. Azalea's Medicaid certification is a direct competitive lever in "
            "this channel - it allows Azalea to bill nursing-facility room and board for dual-eligible "
            "residents, which many facilities require of a hospice partner. Tier 1 facilities (highest-census "
            "SNFs and ALFs in the Tyler metro) receive weekly in-person visits from the Community Liaison "
            "and clinical liaison; Tier 2 facilities receive bi-weekly or monthly touchpoints by census "
            "potential.")
    h2(d, "3.6 Physician & referral source ecosystem")
    para(d, "Hospice admissions in the Tyler market are driven by referral relationships with hospital "
            "discharge planners and case managers (the single highest-volume source), 200+ primary care "
            "physicians, oncologists (UT Health Tyler Cancer Center, Texas Oncology, CHRISTUS), "
            "organ-specialty physicians (pulmonology, cardiology, neurology), SNF/ALF medical directors and "
            "DONs, and the faith-community and pastoral-care networks that are uniquely important in East "
            "Texas. A mature, above-average-utilization market with 60,000+ Medicare-eligible residents and "
            "70+ senior facilities - entered not cold, but by an experienced local team on validated "
            "Tyler-market economics, holding a dual-certified license new entrants cannot obtain.")

    # ---- 04 Competitive Analysis ----
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
          widths=[1.9, 1.3, 3.3])
    h2(d, "Azalea's differentiation")
    bullet(d, "Dual Medicare + Medicaid certification under an enrollment moratorium - a scarce license position that new entrants cannot replicate, with nursing-facility room and board billing that widens the SNF referral base")
    bullet(d, "Validated local economics - modeled on the leadership team's own proven ~22 ADC Tyler book, not a standing start")
    bullet(d, "Dignity-led hospitality model - luxury-hospitality standards applied to end-of-life care")
    bullet(d, "Faith-aligned positioning - explicit Christian-values messaging resonant with the local demographic")
    bullet(d, "Speed of admission - same-day / next-day response, faster than national competitors' 24-72 hours")
    bullet(d, "Local independent ownership - faster decisions, no corporate bureaucracy, deeper relationship investment")
    bullet(d, "MVI-certified in The Perfect Visit and Perfect Phones - national co-marketing exposure that local and regional competitors lack (see 4.1)")
    h2(d, "4.1 MVI 'Perfect Visit' certification & national co-marketing (strategic upside)")
    para(d, "Azalea is pursuing certification by MultiView Incorporated (MVI) in The Perfect Visit and "
            "Perfect Phones - structured programs that certify an agency's patient-visit quality and its "
            "intake and telephone responsiveness against a defined standard. MVI, led by Andrew Reid, is "
            "launching a national, multi-platform advertising campaign that promotes hospice agencies "
            "certified in The Perfect Visit. As a certified agency, Azalea would be featured to a national "
            "audience - the kind of brand exposure normally reserved for the largest national operators with "
            "multi-million-dollar ad budgets, here made accessible to a local independent through the "
            "certification.")
    para(d, "Azalea will develop an in-tandem local marketing plan timed to the national campaign - "
            "capturing the awareness it generates and converting it into local inquiries and referrals "
            "across the Tyler service area. The pairing of a national demand-generation engine with a "
            "focused local conversion plan is a differentiator no local competitor currently has, and it "
            "creates a credible path to take market share from the regional and national incumbents listed "
            "above.")
    para(d, "Underwriting note: this is incremental upside. The financial projections in Section 11 remain "
            "conservative and assume no lift from the MVI certification or the national campaign; debt "
            "service is covered on the base-case census path without it. The certification represents "
            "potential growth above plan, not a dependency.")

    # ---- 05 SWOT ----
    d.add_page_break()
    h1(d, "05 - SWOT Analysis")
    h2(d, "Strengths")
    bullet(d, "Dual Medicare + Medicaid certification (CCN effective 1/8/2024) - billing-ready day one, with nursing-facility room and board capability")
    bullet(d, "Scarce license: the Texas Medicaid enrollment freeze and the CMS nationwide enrollment moratorium (effective 5/13/2026, naming Texas) block new entrants; comparable licenses trade at $400-500K+")
    bullet(d, "Modeled on validated Tyler-market economics (~22 ADC, ~$118K/mo) benchmarked to the team's own actual collections")
    bullet(d, "Experienced clinical & admin team already in seat, with deep East-Texas referral relationships")
    bullet(d, "Simple, conservative capital structure - a single $500K SBA loan after the seller balance retires; projections model a compressed 36-month amortization and still cover")
    bullet(d, "$250K cash equity injection (33% of project) plus a $235K working-capital reserve and a $250K revolver facility")
    bullet(d, "Strong ownership alignment - operators hold equity (13.3% each); experienced multi-hospice Manager")
    h2(d, "Weaknesses")
    bullet(d, "Newco borrower with no operating history of its own (mitigated by the team's benchmarked prior book and an experienced operator)")
    bullet(d, "Refuge has minimal current operations - census must be built from the team's referral relationships during ramp")
    bullet(d, "Single-office, single-market concentration")
    bullet(d, "Seller-note window (August 2026 to SBA funding) concentrates payments early; mitigated by the $31,250/mo schedule, prepayment without penalty, and SBA takeout in September 2026")
    bullet(d, "Sensitivity to a sustained census shortfall (see Section 11.7)")
    bullet(d, "Manager's attention is split across other interests during launch")
    h2(d, "Opportunities")
    bullet(d, "Growing 65+ population (60,000+ Medicare-eligible) with above-average utilization")
    bullet(d, "Base-case census growth to 56 patients by month 36 from the team's referral pipeline")
    bullet(d, "70+ SNF/ALF facilities, many underserved by incumbents - and newly addressable through Medicaid nursing-facility room and board billing")
    bullet(d, "Faith-based marketing to an extensive church and ministry network")
    bullet(d, "Palliative-care consultation line as a Year 2-3 revenue diversifier")
    bullet(d, "MVI 'Perfect Visit' / 'Perfect Phones' certification + Andrew Reid's national multi-platform campaign - national lead exposure for a local independent (upside, not in the base case)")
    bullet(d, "CMS hospice rates have risen every year since 2010; model carries a conservative 2.5%/yr escalation")
    h2(d, "Threats")
    bullet(d, "National competitors with larger sales forces and budgets")
    bullet(d, "CMS reimbursement or regulatory changes")
    bullet(d, "Medicare Advantage hospice carve-in could shift referral dynamics")
    bullet(d, "East Texas clinical labor shortage could pressure wages")
    bullet(d, "Ramp-period cash timing ahead of Medicare collections (NOE-to-cash ~30-60 days); mitigated by the $235K reserve, the $250K revolver, and month-by-month cash-flow modeling")

    # ---- 06 Marketing ----
    d.add_page_break()
    h1(d, "06 - Marketing, Sales & Referral Strategy")
    para(d, "Hospice census in the Tyler market is relationship-driven. Azalea's strategy leverages the "
            "leadership team's established local referral relationships across five active development "
            "channels.")
    h2(d, "6.1 Hospital referral channel")
    para(d, "Discharge planners and case managers generate the highest referral volume. The Community "
            "Liaison (Woodard) maintains a structured weekly cadence: UT Health Tyler and CHRISTUS Trinity "
            "Mother Frances (2x/week each), UT Health North East and Texas Spine & Joint (1x/week each), and "
            "bi-weekly touchpoints at regional hospitals. Tactics: in-services on hospice eligibility, "
            "branded referral materials, and a same-day admission-response protocol (clinical liaison "
            "dispatched within 60 minutes).")
    h2(d, "6.2 SNF / ALF facility channel")
    para(d, "Tier 1 covers the ~15 highest-census SNFs and ALFs/memory-care communities in the Tyler metro "
            "with weekly visits and a dedicated clinical liaison; Tier 2 covers the remaining 40+ SNFs and "
            "20+ ALFs across the six-county service area with bi-weekly touchpoints. Quarterly facility "
            "in-services and pre-positioned election packets support consistent volume. Azalea's Medicaid "
            "certification lets it bill nursing-facility room and board for dual-eligible residents - a "
            "capability many facilities require of a hospice partner and a concrete door-opener in this "
            "channel.")
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
            "senior health fairs, and church health-ministry workshops. Marketing spend is funded within "
            "fixed G&A and sustained across all three years of the model.")
    para(d, "National co-marketing (upside): as an agency pursuing MVI 'Perfect Visit' and 'Perfect Phones' "
            "certification, Azalea will run an in-tandem local campaign timed to MVI's national, "
            "multi-platform advertising of certified hospice agencies (led by Andrew Reid) - local landing "
            "pages, geo-targeted digital, and referral-source outreach engineered to convert national "
            "awareness into Tyler-area leads. This sits on top of the budgeted spend above and is treated as "
            "upside; the base-case projections do not rely on it.")
    h2(d, "6.6 Census model")
    table(d, ["Period", "Census path (EOM)", "Primary channel"],
          [["Months 1-3 (launch)", "20 -> 24 -> 26", "Team activation + established referral relationships"],
           ["Months 4-12", "34 by month 6 -> 40 by month 12", "Hospital + SNF/ALF (incl. Medicaid room & board accounts)"],
           ["Year 2", "50 by month 24 (~45 avg ADC)", "All channels + faith"],
           ["Year 3", "56 by month 36 (~53 avg ADC)", "All channels + organic"]],
          widths=[1.6, 2.2, 2.7])
    para(d, "All coverage figures in Section 11 use this conservative base path. Admissions ramp from the "
            "team's established East Texas referral relationships; discharges follow an "
            "average-length-of-stay-driven waterfall. A documented downside case (slower ramp) is presented "
            "in Section 11.7.")

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
    para(d, "Staggered W-2 hires plus a census-driven PRN pool. A benefits load of 22% plus 3% workers "
            "compensation applies to W-2 wages; Medical Directors are 1099 contractors. Officer salaries are "
            "separately stated and fully loaded into the model: Administrator $150,000, Director of Nursing "
            "$120,000, Community Liaison $120,000, Office Manager $60,000. The Year-1 roster is sized to the "
            "ramp (24 patients by month 2, 40 by month 12) within standard hospice staffing ratios and "
            "includes the Administrator, Director of Nursing, Community Liaison, Office Manager, RN case "
            "managers, hospice aides, and a 1099 Medical Director, with chaplain and social-work coverage "
            "from the PRN pool.")
    h2(d, "7.3 Capacity-driven hiring (Years 2-3)")
    para(d, "As census grows toward 56 patients, the roster scales on census triggers (RN caseload 1:12; "
            "aide 1:10), converting PRN roles to salaried and adding compliance, intake, and "
            "volunteer-coordinator support. Additional RN case managers and CNAs load in as census crosses "
            "their thresholds; FT social worker and chaplain convert from PRN; and quality/compliance and "
            "intake support are added as volume warrants. Nursing, aide, chaplain, and social-work FTEs "
            "scale with census on standard hospice staffing ratios, and per-patient costs (pharmacy, DME, "
            "supplies) are per-patient-day rates from the line-item budget.")
    h2(d, "7.4 Facility & technology")
    para(d, "Principal office in the Tyler, TX area (go-forward lease assumption in the model). Hours M-F "
            "8-5 with 24/7/365 on-call; ~45-mile service radius. EMR, PCR, and outsourced billing run on the "
            "platform reflected in G&A. Field travel is mileage-reimbursed initially; an owned fleet is "
            "evaluated in Year 2.")

    # ---- 08 Management and Organization ----
    d.add_page_break()
    h1(d, "08 - Management and Organization")
    para(d, "Azalea is operated by an experienced, already-seated team whose established East-Texas referral "
            "relationships are a key de-risking factor. Officer salaries are fully loaded into the "
            "projections as expenses (no compensation is omitted, per SBA requirements).")
    h2(d, "8.1 Ownership & reporting structure")
    para(d, "Tyler Hospice Hold, LLC (Wyoming holding company) is acquiring 100% of Refuge Hospice, LLC "
            "(Texas operating subsidiary) in two steps per 42 CFR 424.550(b); the subsidiary does business "
            "as Azalea Hospice & Palliative Care. The Administrator leads operations; the Director of "
            "Nursing leads the clinical team; the Community Liaison leads business development. Mr. Bullard "
            "is a passive investor with protective minority rights but no management authority or signing "
            "power.")
    table(d, ["Entity / person", "Role", "Interest", "Reports to"],
          [["Geoff Schackmann", "Manager (direct individual owner)", "39.9%", "-"],
           ["Silas R. Shelton", "Administrator / Executive Director", "13.3%", "Manager"],
           ["Dana L. Davenport", "Director of Nursing", "13.3%", "Administrator"],
           ["Bradley G. Woodard", "Community Liaison / Director of Sales", "13.3%", "Administrator"],
           ["James E. Bullard", "Passive minority investor", "19.5%", "No operational role"]],
          widths=[1.7, 2.2, 0.8, 1.8])
    h2(d, "8.2 Decision authority")
    bullet(d, "Operating within budget ($0-$10K): Administrator (Shelton)")
    bullet(d, "Above $10K or out-of-budget: Manager (Schackmann, individually)")
    bullet(d, "Clinical / regulatory: Director of Nursing (Davenport), escalated to Administrator + Medical Director")
    bullet(d, "Hiring / firing senior staff: Administrator, with Manager concurrence")
    bullet(d, "Strategic / capital / acquisition: Manager")
    bullet(d, "Bank account signatures: Manager (sole signer)")
    h2(d, "8.3 Equity vesting (operator-members)")
    para(d, "Each operator-member's 13.3% interest is a Restricted Interest under the Operating Agreement. "
            "The entire 13.3% is at risk and forfeitable until vested. It comprises a 4.9% Initial Base that "
            "time-vests on continuous service (four-year schedule, one-year cliff) and an 8.4% Earn-Up "
            "Portion that vests only on a dual trigger: a four-year time-vesting schedule with a one-year "
            "cliff AND the performance milestones (25% at Breakeven, 50% at three consecutive Profitable "
            "months, 25% at twelve consecutive Profitable months), with the lesser of the two schedules "
            "governing - so no portion is vested at inception and each operator may earn up to a "
            "fully-vested 13.3%. Each grantee files a timely IRC Section 83(b) election. On departure, "
            "unvested interests are forfeited at $0; the Company holds a repurchase (call) right over vested "
            "interests - at fair market value for a good-leaver separation (termination without cause, "
            "death, disability, or retirement) and at the lower of cost or fair market value for a "
            "bad-leaver separation (resignation before full vesting, or termination for cause), paid via an "
            "SBA-subordinated note. This keeps operator equity aligned with the credit's performance and "
            "ensures departed operators do not retain equity in the company.")

    # ---- 09 Leadership Team ----
    d.add_page_break()
    h1(d, "09 - Leadership Team")
    h2(d, "Geoff Schackmann - Manager (direct individual owner, 39.9%)")
    para(d, "Multi-hospice operator and transaction-led growth leader with operational responsibility for "
            "Medicare-certified hospice and palliative-care agencies across multiple states. Direct "
            "individual owner of a 39.9% interest in Tyler Hospice Hold, LLC and its Manager. Direct "
            "experience includes change-of-ownership (CHOW) transactions and post-CHOW enrollment oversight, "
            "census growth from sub-30 to 100+ ADC under existing Medicare provider numbers, multi-site "
            "clinical operations under Texas HCSSA and Oregon hospice licensure, CHAP and Joint Commission "
            "accreditation, acquisition due diligence (clinical, financial, regulatory), and lender/investor "
            "relationships across SBA 7(a), conventional, and mezzanine structures. Leads transaction "
            "structuring, financing, capital allocation, and post-close integration. Sole 20%+ owner and SBA "
            "personal guarantor.")
    h2(d, "Bradley G. Woodard - Community Liaison / Director of Sales (13.3%)")
    para(d, "More than 25 years of East-Texas hospice business development and administration. Founding "
            "administrator of one of the Tyler market's largest hospices (licensed 2004; grown to 220 "
            "patients, 109 employees, and 62 volunteers), with subsequent census-building roles across "
            "multiple East-Texas agencies - including start-ups in Lufkin/Nacogdoches (0->48 in six months; "
            "12->70+) and Grace Hospice of East Texas (8->145+). Most recently VP of Business Development "
            "for an East-Texas hospice (2021-present), sustaining a 40+ ADC book for 4.5 years, with no "
            "non-compete restricting his transition to Azalea. Licensed Nursing Home Administrator (LNFA "
            "#7136); B.S., Texas A&M University.")
    h2(d, "Silas R. Shelton - Administrator / Executive Director (13.3%)")
    para(d, "Operational leadership of Azalea Hospice & Palliative Care, with day-to-day responsibility for "
            "regulatory compliance under Texas HCSSA and the Medicare Conditions of Participation, payer-mix "
            "management, referral-source partnerships across Smith County and surrounding East-Texas "
            "counties, and overall site leadership. Direct reports include the Director of Nursing, the "
            "Community Liaison, and the full clinical and operational team.")
    h2(d, "Dana L. Davenport - Director of Nursing (13.3%)")
    para(d, "Clinical leadership of the Azalea nursing team, responsible for Medicare Conditions of "
            "Participation compliance, interdisciplinary group (IDG) oversight, plan-of-care management, and "
            "clinical quality (QAPI). Leads RN case managers, hospice aides, social workers, and chaplains "
            "across the service area.")
    h2(d, "Investor")
    para(d, "James E. Bullard is a passive minority investor holding a 19.5% membership interest acquired "
            "for a $195,000 cash capital contribution (his own savings), part of the $250,000 total equity "
            "injection. He holds customary protective minority rights (information rights, anti-dilution "
            "with a 19.5% floor, voting on a defined set of fundamental Reserved Matters, and tag-along "
            "rights) but has no management authority, no operational or clinical role, and no signing power. "
            "His consent is not required for SBA-loan actions, and his buy-sell rights are suspended while "
            "the SBA loan is outstanding. As a sub-20% non-controlling member, he provides no SBA personal "
            "guaranty (13 CFR 120.160).")

    # ---- 10 Acquisition and Launch Strategy ----
    d.add_page_break()
    h1(d, "10 - Acquisition and Launch Strategy")
    h2(d, "10.1 The Refuge acquisition & 36-month-rule structure")
    para(d, "The transaction is the acquisition of 100% of the membership interests of Refuge Hospice, LLC, "
            "a Texas LLC certified for both Medicare and Medicaid (CMS CCN effective January 8, 2024), for "
            "$500,000 under a Membership Interest Purchase Agreement dated July 14, 2026 (in final "
            "negotiation). The sellers are Angel J. Resendiz and Dennis M. Hendrix, with payments made to "
            "DHJR, LP as the sellers' agent. Because 42 CFR 424.550(b) bars a change of majority ownership "
            "within 36 months of a hospice's initial Medicare certification, the purchase is structured in "
            "two steps: 49% of the membership interests transfer at closing (approximately August 1, 2026) "
            "against a $125,000 down payment, and the remaining 51% transfer on January 15, 2027 - the first "
            "date past the 36-month mark. The $375,000 balance carries 6% interest, payable $31,250 per "
            "month beginning September 1, 2026 with a final payment of approximately $258,489.46 on January "
            "15, 2027, and is prepayable without penalty. The requested SBA loan retires the seller balance "
            "in September 2026. Azalea files the applicable CMS-855A ownership-change updates and moves "
            "banking, EMR, and insurance into the new ownership structure at each step.")
    h2(d, "10.2 Prior target (history)")
    para(d, "The Company's previously contracted target, Hickory Hospice, LLC ($300,000 purchase price, "
            "Medicare-only certification), failed its state site survey and that transaction was terminated. "
            "The Refuge acquisition replaces it with a stronger asset: dual Medicare and Medicaid "
            "certification, a clean certification effective January 8, 2024, and immediate billing "
            "readiness. The prior dual-loan financing structure was terminated with that transaction and is "
            "superseded by the single SBA 7(a) request in this plan.")
    h2(d, "10.3 Billing readiness & revenue ramp")
    para(d, "Refuge is certified and ready to bill from day one - there is no new-provider 855A enrollment "
            "wait and no gap in billing capability. Refuge carries minimal current operations, so the "
            "revenue plan is a census build, not a panel transfer: admissions ramp from the leadership "
            "team's established East-Texas referral relationships to 24 patients by month 2, 34 by month 6, "
            "and 40 by month 12. The model's per-diem economics and cost structure are benchmarked to the "
            "team's own prior Tyler-market book (April-June 2025 actuals), so the ramp is a return to census "
            "levels this team has already operated at in this market.")
    h2(d, "10.4 Affiliate disclosure")
    para(d, "The borrower's affiliate group consists of the borrower and its operating subsidiary Refuge "
            "Hospice, LLC - well within the applicable SBA size standard. The controlling owner (Geoff "
            "Schackmann, direct individual 39.9% owner) previously held a 33.33% interest in VistaRiver Inc, "
            "which he sold in August 2025; he retains only a passive seller-note receivable (no equity, "
            "officer, manager, or employee role), so VistaRiver is not an SBA affiliate. A separate "
            "Affiliate and Size-Standard memorandum is included in the application package.")
    h2(d, "10.5 Launch timeline & equity-injection coordination")
    table(d, ["Window", "Phase", "Key activities"],
          [["Through July 2026", "Equity injection", "Full $250,000 on deposit: Bullard $195,000 ($100K wired 5/7/2026) + Schackmann $55,000"],
           ["~August 1, 2026", "Closing - step 1 (49%)", "49% of Refuge membership interests transfer; $125,000 down payment; banking, EMR, and insurance stand-up; staff credentialing"],
           ["September 2026", "SBA funding & seller payoff", "SBA loan funds; $375,000 seller balance retired (note prepayable without penalty); working-capital reserve in place"],
           ["Months 1-6", "BD ramp", "Woodard-led weekly referral cadence; census 24 by month 2, 34 by month 6; AR normalizes to ~45-day Medicare cycle"],
           ["January 15, 2027", "Closing - step 2 (51%)", "Remaining 51% transfers on the first date past the 36-month mark (42 CFR 424.550(b)); CMS-855A ownership updates"],
           ["Months 7-12+", "Growth", "Census to 40 by month 12; capacity hires triggered by census thresholds; coverage rising"]],
          widths=[1.3, 1.6, 3.6])
    para(d, "SBA disbursement is sequenced to occur with the full $250,000 equity injection on deposit, in "
            "compliance with SOP 50 10 8.")

    # ---- 11 Financial Plan ----
    d.add_page_break()
    h1(d, "11 - Financial Plan and Projections")
    para(d, "All projections are from the Rev 3.00 dynamic proforma (36 monthly periods; every calculation "
            "cell is a live formula; all inputs on a single Control Tower tab; QA-verified). The proforma is "
            "deliberately the conservative, lender-defensible basis: officer salaries fully loaded, a 22% "
            "benefits load plus 3% workers compensation, a 5% of net patient revenue contingency deducted "
            "below EBITDA, and debt service modeled on a compressed 36-month amortization far heavier than "
            "actual SBA terms.")
    h2(d, "11.1-11.2 Sources & uses of funds")
    table(d, ["Source", "Amount", "Use", "Amount"],
          [["SBA 7(a) loan", "$500,000", "Acquisition of Refuge Hospice, LLC ($125K down + $375K seller balance retired Sept 2026)", "$500,000"],
           ["Cash equity injection (Bullard $195K + Schackmann $55K)", "$250,000", "Closing and startup costs", "$15,000"],
           ["", "", "Working-capital reserve", "$235,000"],
           ["Total sources", "$750,000", "Total uses", "$750,000"]],
          widths=[2.2, 0.9, 2.5, 0.9])
    para(d, "The equity injection is 33% of the $750,000 project. The seller balance is retired by SBA "
            "proceeds in September 2026, so the going-forward balance sheet carries a single debt.")
    h2(d, "11.3 Key assumptions")
    table(d, ["Assumption", "Value"],
          [["Payer mix", "Medicare RHC per-diem; Medicaid nursing-facility room & board passed through at the standard rate"],
           ["RHC per-diem (Tier 1 / Tier 2)", "$230.83 / $182.36 (FY2026 national), wage-index adjusted for Smith County (CBSA 46340)"],
           ["High-rate (first 60 days) share", "~15% of patient-days"],
           ["Blended net rate / patient-day", "~$172 after sequestration and billing fees - escalating 2.5% / year"],
           ["Census (end of month)", "24 by month 2 / 34 by month 6 / 40 by month 12 / 50 by month 24 / 56 by month 36"],
           ["Officer salaries (separately stated)", "Administrator $150,000; Director of Nursing $120,000; Community Liaison $120,000; Office Manager $60,000"],
           ["Benefits load (W-2)", "22% + 3% workers compensation"],
           ["Clinical staffing", "RN, aide, chaplain, and social-work FTEs scale with census on standard hospice ratios; per-patient costs are per-patient-day rates"],
           ["Contingency", "5% of net patient revenue, deducted below EBITDA in cash flow"],
           ["Modeled debt service", "$500,000 at 6% amortized over 36 months - $15,211/mo from October 2026 ($182,532/yr)"],
           ["Actual SBA terms (for reference)", "~10.5% over 10 years is ~$6,745/mo ($80,940/yr) - real coverage materially better than modeled"],
           ["Seller note (pre-payoff)", "$375,000 at 6%; $31,250/mo from 9/1/2026; retired by SBA proceeds September 2026; prepayable without penalty"],
           ["Equity injection", "$250,000 cash (33% of the $750,000 project)"],
           ["Amortization", "Acquired license / intangibles amortized straight-line over 15 years (~$33,333/yr)"],
           ["Collections timing", "Collections lag billing by ~30-60 days (NOE timing), modeled month by month"],
           ["Liquidity backstop", "$235,000 working-capital reserve + $250,000 revolver facility"]],
          widths=[2.4, 4.1])
    h2(d, "11.4 Three-year profit & loss (Rev 3.00 proforma)")
    table(d, ["Line item", "Year 1", "Year 2", "Year 3"],
          [["Average daily census (ADC)", "31.0", "45.0", "53.0"],
           ["Net patient revenue", "$1,932,877", "$2,891,641", "$3,490,623"],
           ["Direct patient costs", "$180,568", "$263,288", "$309,983"],
           ["Clinical payroll (loaded)", "$739,682", "$1,036,237", "$1,218,374"],
           ["Indirect payroll (loaded)", "$596,875", "$688,812", "$709,477"],
           ["Facility", "$41,400", "$65,562", "$89,729"],
           ["Operating G&A", "$111,401", "$157,517", "$187,135"],
           ["EBITDA (before debt service)", "$262,951", "$680,224", "$975,925"],
           ["EBITDA margin", "13.6%", "23.5%", "28.0%"],
           ["Contingency (5% of NPR, below EBITDA)", "$96,644", "$144,582", "$174,531"],
           ["Amortization", "$33,333", "$33,333", "$33,333"],
           ["Interest (seller note + acquisition note + revolver)", "$26,499", "$18,537", "$8,422"],
           ["TX franchise tax", "$5,534", "$10,844", "$13,090"],
           ["Net income", "$100,941", "$472,928", "$746,549"]],
          widths=[2.6, 1.3, 1.3, 1.3])
    para(d, "Net patient revenue grows from $1.93M to $3.49M and EBITDA from $263K to $976K (before debt "
            "service) as the base-case census builds from 24 patients at month 2 to 56 at month 36 on the "
            "team's referral pipeline, with a conservative 2.5% annual CMS per-diem escalation. Payroll "
            "scales on census triggers, carrying the full, separately stated officer roster and a 22% + 3% "
            "benefits and workers compensation load. The explicit 5% contingency is carried below EBITDA and "
            "deducted in cash.")
    h2(d, "11.5 Year-1 monthly detail (selected months)")
    table(d, ["Month", "ADC", "Net revenue", "EBITDA (margin)"],
          [["M1 (launch)", "10.0", "$52,240", "($10,661) - absorbed by the working-capital reserve"],
           ["M2", "22.0", "$114,916", "($3,286) - census 24 EOM, at break-even"],
           ["M3", "25.2", "$131,893", "$4,291 - break-even census (~17-18) crossed"],
           ["M6", "32.7", "$171,073", "$23,313 (13.6%)"],
           ["M12", "39.5", "$206,376", "$46,099 (22.3%)"]],
          widths=[1.2, 0.9, 1.4, 3.0])
    para(d, "Month 1 is the launch month (10 ADC average) and is absorbed by the reserve; the agency "
            "crosses break-even census (~17-18 patients) in month 2-3 and exits Year 1 at a 22.3% EBITDA "
            "margin. Collections lag billing by ~30-60 days (NOE timing); the cash-flow tab models the lag "
            "explicitly, month by month, and base-case cash never breaches the $25,000 floor (peak revolver "
            "draw $78,710; month-12 cash $80,450; month-36 cash $936,793).")
    h2(d, "11.6 Lender summary & debt-service coverage")
    table(d, ["Metric", "Year 1", "Year 2", "Year 3"],
          [["Net patient revenue", "$1,932,877", "$2,891,641", "$3,490,623"],
           ["EBITDA (before debt service)", "$262,951", "$680,224", "$975,925"],
           ["Modeled debt service ($500K, 6%, 36-mo)", "$142,524", "$182,532", "$182,532"],
           ["EBITDA / modeled debt service", "1.8x", "3.7x", "5.3x"],
           ["EBITDA / actual-terms SBA service (~$80,940/yr)", "n/m (partial year)", "8.4x", "12.1x"],
           ["Net income", "$100,941", "$472,928", "$746,549"]],
          widths=[2.8, 1.3, 1.3, 1.3])
    para(d, "The model deliberately amortizes the full $500,000 over 36 months at $15,211 per month from "
            "October 2026 - a schedule roughly 2.3x heavier than an actual SBA 7(a) at ~10.5% over 10 years "
            "(~$6,745/mo). Even so, EBITDA covers the modeled service 1.8x in Year 1 (which also carries the "
            "seller-note interest), 3.7x in Year 2, and 5.3x in Year 3 - a 3-year aggregate of 3.8x against "
            "the 1.25x floor - and these coverages are computed before adding back the below-EBITDA "
            "contingency. On actual SBA terms, real coverage will be materially better than modeled. "
            "Break-even is approximately 17-18 patients, crossed in month 2-3 of the ramp.")
    h2(d, "11.7 Downside case")
    para(d, "The proforma carries a documented downside case (slower census ramp) as a scenario switch. In "
            "the downside, the $250,000 revolver is fully drawn and remains drawn, but cash stays positive "
            "throughout the 36 months - the working-capital reserve and revolver absorb the shortfall. A "
            "sustained downside beyond that envelope would require census recovery, additional equity, or a "
            "larger facility, and this is disclosed rather than modeled away. Three structural mitigants "
            "bound the risk: (1) the roster is census-driven, so a lower census carries a lighter cost base; "
            "(2) break-even is ~17-18 patients against a month-2 census of 24, so the plan operates above "
            "break-even from the second month; and (3) the modeled 36-month amortization overstates the real "
            "debt burden - on actual SBA terms, annual debt service falls by roughly $100,000, widening every "
            "downside margin.")
    h2(d, "11.8 Proof: benchmarked economics")
    para(d, "The model is not pure projection - its per-diem economics and cost structure are benchmarked "
            "to the leadership team's own prior Tyler-market book (April-June 2025 actuals: approximately "
            "$118,400/month net collections at ~22 ADC). The census path is a return to levels this team "
            "has already operated at in this market, and patient-related costs reconcile to those actuals. "
            "Payroll differs by design - Azalea's staggered W-2 roster with a 22% + 3% load, separately "
            "stated officer salaries, and go-forward Medical Directors versus the benchmarked book's "
            "blended run.")

    # ---- 12 Risk Factors ----
    d.add_page_break()
    h1(d, "12 - Risk Factors and Mitigations")
    para(d, "The dominant risks are the newco's lack of operating history and the census ramp - both "
            "mitigated by benchmarked economics, a dual-certified license that is billing-ready from day "
            "one, an experienced operator, and a capital structure with a 33% equity injection, a $235,000 "
            "working-capital reserve, and a $250,000 revolver. The seller-note window (closing to SBA "
            "funding) is short and is extinguished when SBA proceeds retire the balance in September 2026.")
    table(d, ["Risk", "Prob.", "Impact", "Mitigation"],
          [["Debt service / coverage", "Low", "Med", "Projections model a compressed 36-month amortization ($15,211/mo) and still cover 1.8x -> 3.7x -> 5.3x; actual SBA terms (~$6,745/mo) roughly halve the modeled burden"],
           ["Seller-note window / 36-month rule", "Low", "Med", "Two-step structure complies with 42 CFR 424.550(b); note is prepayable without penalty; SBA proceeds retire the $375K balance in September 2026; final 51% transfer is a fixed contractual date (1/15/2027)"],
           ["Slower census ramp", "Med", "High", "Break-even ~17-18 patients vs 24 by month 2; capacity hires are census-gated so cost flexes with volume; $235K reserve + $250K revolver keep cash positive in the documented downside case"],
           ["Ramp cash timing (Medicare NOE lag)", "Med", "Med", "Collections lag of ~30-60 days is modeled month by month; reserve sized to carry payroll ahead of collections"],
           ["Wage inflation", "Med", "Med", "22% + 3% benefits and workers compensation load fully modeled; census-gated hiring; PRN pool buffers"],
           ["CMS rate / regulatory change", "Low", "Med", "Rates up or flat every year since 2010; only a conservative 2.5%/yr escalation assumed; the enrollment moratorium that constrains supply also protects the license's value"],
           ["Key-person dependency", "Low", "Med", "Operators are owners (13.3% each, milestone-vested); Administrator leads day-to-day; OA key-person succession provisions"],
           ["Equity injection phased", "Low", "Low", "$100K wired 5/7/2026; balance committed (Bullard $95K + Schackmann $55K); SBA disbursement sequenced after full injection per SOP 50 10 8"],
           ["Newco / no operating history", "Med", "Med", "Dual-certified, billing-ready license; economics benchmarked to the team's own actual collections; operator has 10+ years of hospice ownership; 33% equity injection"]],
          widths=[1.7, 0.6, 0.7, 3.5])

    # ---- 13 Milestones and Conclusion ----
    d.add_page_break()
    h1(d, "13 - Milestones and Conclusion")
    h2(d, "13.1 Launch & growth milestones")
    table(d, ["Window", "Milestone", "Verification"],
          [["Through July 2026", "Full $250K equity injection on deposit (Bullard $195K + Schackmann $55K)", "Bank statements + source-of-funds documentation"],
           ["~August 1, 2026", "Close step 1: 49% of Refuge interests; $125K down", "Executed MIPA + closing statement"],
           ["September 2026", "SBA loan funds; $375K seller balance retired", "Payoff letter + loan statements"],
           ["Months 1-3", "CMS-855A ownership updates filed; team seated; EMR live; census 24 by month 2", "855A receipts + payroll + census reports"],
           ["Months 4-6", "BD ramp; census 34 by month 6; AR normalizes to ~45-day cycle", "CRM report + AR aging < 45 days"],
           ["January 15, 2027", "Close step 2: remaining 51% transfers (past the 36-month mark)", "Assignment of interests + CMS filings"],
           ["Months 7-12", "Census 40 by month 12; month-12 EBITDA margin 22.3%", "Monthly financials"],
           ["Years 2-3", "Census 50 by month 24 and 56 by month 36; coverage rising on census-gated hiring", "Board reviews of monthly financials"]],
          widths=[1.4, 3.0, 2.1])
    h2(d, "13.2 Conclusion")
    para(d, "Azalea is a well-structured SBA 7(a) opportunity: the acquisition of a scarce, dual-certified "
            "(Medicare and Medicaid) Texas hospice license that is ready to bill from day one, led by an "
            "operator with 10+ years of hospice ownership and a seated East-Texas team returning to census "
            "levels it has operated at before. The credit rests on four verifiable points: (1) the license "
            "itself - under the Texas Medicaid enrollment freeze and the CMS nationwide enrollment "
            "moratorium, comparable dual-certified licenses trade at $400-500K and above, supporting the "
            "$500,000 purchase price and constraining new competition; (2) benchmarked economics - the "
            "model's rates and costs reconcile to the team's own prior Tyler-market actuals, with break-even "
            "(~17-18 patients) crossed in month 2-3; (3) conservative coverage - EBITDA of $263K / $680K / "
            "$976K covers even a deliberately compressed 36-month amortization of the full $500,000 at 1.8x "
            "-> 3.7x -> 5.3x, and actual SBA terms roughly halve the modeled debt burden; and (4) a "
            "well-capitalized structure - a $250,000 cash injection (33% of the $750,000 project), a "
            "$235,000 working-capital reserve, and a $250,000 revolver that keep cash positive in the base "
            "case and in the documented downside case. The binding risk is a sustained census shortfall "
            "during the ramp; it is mitigated by an above-break-even census from month 2, a census-driven "
            "cost structure, substantial liquidity, and the retirement of the seller note in September 2026 "
            "that leaves a single, coverable debt.")

    para(d, "")
    para(d, "Rev 6.00 - July 2026. This plan is computed from the Rev 3.00 dynamic proforma "
            "(Azalea_Hospice_Proforma_Rev3.00_DYNAMIC.xlsx), the single source of truth for all financial "
            "figures. Confidential - do not distribute without written consent.", italic=True, size=8,
         color=GREY)
    para(d, "")
    para(d, "Prepared for the SBA 7(a) loan application of Tyler Hospice Hold, LLC (dba Azalea Hospice & "
            "Palliative Care). Confidential.", italic=True, size=8, color=GREY)

    out = os.path.join(BASE, OUT)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    d.save(out)
    print("  wrote", OUT)


if __name__ == "__main__":
    business_plan_rev6()
