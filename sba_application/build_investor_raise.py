"""Investor raise package: pitch deck (pptx) + teaser + FAQ + data room + scripts.

$500K promissory-note raise (10% / 3yr / interest-only quarterly / balloon /
prepayable). Warm-path 506(b) posture. All figures from the Rev 3.00 proforma
(QA 20/20): FY EBITDA $263K / $680K / $976K; M12 margin 22.3%; census 24/34/40.
Output: sba_application/12_investor_raise/
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as PRGB
from docx import Document
from docx.shared import Pt as DPt, RGBColor as DRGB, Inches as DIn

OUT = os.path.join(os.path.dirname(__file__), "12_investor_raise")
NAVY = PRGB(0x1F, 0x3A, 0x5F); DNAVY = DRGB(0x1F, 0x3A, 0x5F); DGREY = DRGB(0x6B, 0x6B, 0x6B)


# ============================ PITCH DECK ============================
def deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def slide(title, bullets=None, kicker=None, table=None, small=None):
        s = prs.slides.add_slide(blank)
        tb = s.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.1), Inches(1.0))
        p = tb.text_frame.paragraphs[0]; r = p.add_run(); r.text = title
        r.font.size = Pt(30); r.font.bold = True; r.font.color.rgb = NAVY
        y = 1.35
        if kicker:
            kb = s.shapes.add_textbox(Inches(0.6), Inches(y), Inches(12.1), Inches(0.6))
            kp = kb.text_frame.paragraphs[0]; kr = kp.add_run(); kr.text = kicker
            kr.font.size = Pt(16); kr.font.italic = True
            y += 0.7
        if bullets:
            bb = s.shapes.add_textbox(Inches(0.8), Inches(y), Inches(11.8), Inches(7.0 - y))
            tf = bb.text_frame; tf.word_wrap = True
            for i, b in enumerate(bullets):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                run = para.add_run(); run.text = b
                run.font.size = Pt(17)
                para.space_after = Pt(10)
        if table:
            heads, rows, width = table
            tw = Inches(width); th = Inches(0.4 * (len(rows) + 1))
            shape = s.shapes.add_table(len(rows) + 1, len(heads), Inches(0.8), Inches(y), tw, th)
            t = shape.table
            for j, h in enumerate(heads):
                c = t.cell(0, j); c.text = h
                for pr in c.text_frame.paragraphs:
                    for rr in pr.runs: rr.font.size = Pt(13); rr.font.bold = True
            for i, row in enumerate(rows):
                for j, v in enumerate(row):
                    c = t.cell(i + 1, j); c.text = str(v)
                    for pr in c.text_frame.paragraphs:
                        for rr in pr.runs: rr.font.size = Pt(13)
        if small:
            sb = s.shapes.add_textbox(Inches(0.6), Inches(6.8), Inches(12.1), Inches(0.5))
            sp = sb.text_frame.paragraphs[0]; sr = sp.add_run(); sr.text = small
            sr.font.size = Pt(10); sr.font.italic = True
        return s

    # 1 cover
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.7), Inches(2.5))
    tf = tb.text_frame
    r = tf.paragraphs[0].add_run(); r.text = "AZALEA HOSPICE & PALLIATIVE CARE"
    r.font.size = Pt(40); r.font.bold = True; r.font.color.rgb = NAVY
    p2 = tf.add_paragraph(); r2 = p2.add_run()
    r2.text = "Private Note Offering - $500,000  |  Tyler, Texas"
    r2.font.size = Pt(22)
    p3 = tf.add_paragraph(); r3 = p3.add_run()
    r3.text = "Confidential. For discussion with invited accredited investors only. Not an offer to sell securities."
    r3.font.size = Pt(12); r3.font.italic = True

    # 2 opportunity
    slide("Why hospice, why Tyler, why now",
          bullets=[
            "East Texas is one of the oldest regions in one of the fastest-aging states - hospice demand grows every year regardless of the economy.",
            "Texas is effectively not issuing new Medicaid hospice contracts, and CMS froze ALL new hospice Medicare enrollments nationwide in May 2026 (Texas named specifically).",
            "That means nobody can create what we already control: an active, ready-to-bill Medicare + Medicaid hospice license.",
            "We are not building a startup from a piece of paper - we are turning on a proven license with a proven team and a built-in patient census."])

    # 3 the asset
    slide("The asset: a license you cannot replicate",
          bullets=[
            "Medicare + Medicaid certified hospice provider number - proven, active, billable from day one.",
            "Medicaid certification lets us serve nursing-facility patients (room & board billing) - faster census growth and better margins than Medicare-only competitors.",
            "Comparable licenses: Medicare-only shells list at $225-350K; the only two Medicare+Medicaid licenses on the market listed at ~$400K - and supply is shrinking under the federal freeze.",
            "Purchase is fully financed (bank term loan, 6%, 3 years) - this raise is NOT buying the license; it funds growth and working capital."])

    # 4 team + census
    slide("Proven team, built-in census",
          bullets=[
            "Founder Geoff Schackmann: multi-hospice operator, 10+ years building and running East Texas hospice and home-care businesses.",
            "Full leadership team in seat from day one: Executive Director, Director of Nursing, Director of Sales, ADON/Intake - all with Tyler-market experience and equity in the company.",
            "A migrated patient panel and referral network from the team's prior operation - census plan: 24 patients by month 2, 34 by month 6, 40 by month 12.",
            "Anchor investor James E. Bullard has committed $195,000; founder capital brings committed equity to $250,000."])

    # 5 unit economics
    slide("The economics, per patient-day",
          kicker="Medicare pays a fixed per-diem for every enrolled patient, every day.",
          table=(["Metric", "Value"],
                 [["Net revenue per patient-day", "~$172"],
                  ["Direct cost per patient-day (staff + meds + equipment)", "~$75-80"],
                  ["Contribution per patient-day", "~$92-97 (55%+ margin)"],
                  ["Monthly net revenue at 40 patients", "~$205,000"],
                  ["EBITDA margin at month-12 run rate", "22.3%"]], 8.0),
          small="Source: company operating model, CMS FY2026 Smith County rates, validated against prior-operation actuals.")

    # 6 financials
    slide("Three-year financial plan",
          table=(["", "Year 1", "Year 2", "Year 3"],
                 [["Net patient revenue", "$1.93M", "$2.90M", "$3.48M"],
                  ["EBITDA (pre-debt-service)", "$263K", "$680K", "$976K"],
                  ["EBITDA margin", "13.6% (ramp)", "23.5%", "28.0%"],
                  ["Patient census (end of year)", "40", "50", "56"]], 9.0),
          bullets=None,
          small="Year 1 margin reflects the startup ramp; month-12 run rate is 22.3%. Full model (36-month, monthly) available in diligence.")

    # 7 the raise
    slide("The offering: $500,000 in promissory notes",
          table=(["Term", "Detail"],
                 [["Amount", "$500,000 total (minimum investment $50,000)"],
                  ["Interest", "10% per year, paid quarterly in cash"],
                  ["Maturity", "3 years; principal balloon at maturity"],
                  ["Prepayment", "Company may prepay anytime without penalty"],
                  ["Security", "Note obligations of the company; subordinate to the bank acquisition loan (disclosed in note documents)"],
                  ["Investor tax", "Simple 1099-INT interest income - no K-1s, no ownership complexity"]], 10.0),
          small="Private placement to accredited investors under Reg D. Definitive documents prepared by securities counsel.")

    # 8 use of funds
    slide("Use of funds",
          table=(["Use", "Amount"],
                 [["Working capital through the census ramp (payroll ahead of Medicare payment cycle)", "$250,000"],
                  ["Growth: referral development, clinical staffing ahead of census", "$150,000"],
                  ["Reserve (protects quarterly interest payments in any scenario)", "$100,000"]], 9.5),
          bullets=["Total capitalization: $250K committed equity + $500K notes + $500K bank license financing = $1.25M project.",
                   "The license itself is already financed - every raise dollar goes to operations and growth."])

    # 9 downside protection
    slide("How noteholders are protected",
          bullets=[
            "Coverage: quarterly interest is $12,500 - covered ~5x by EBITDA from the second quarter onward, and covered by raise proceeds before that.",
            "The license itself is a hard asset: comparable Medicare+Medicaid licenses list at $400K+, and the federal enrollment freeze makes them scarcer every month.",
            "Stress-tested: at 85% of planned census with 10% higher payroll, the model still services all obligations.",
            "Aligned operators: the leadership team holds equity that vests only if the business performs; the founder personally guarantees the bank debt.",
            "Quarterly investor reporting: census, revenue, EBITDA, cash - the same dashboard management runs the business on."])

    # 10 timeline
    slide("Timeline",
          bullets=[
            "License acquisition closing and first patients: within 30-60 days of funding.",
            "Census 24 by month 2 - the migrated panel arrives in the first weeks.",
            "Cash-flow positive operations: inside the first year (month-12 EBITDA run rate ~$45K/month).",
            "Note maturity in year 3: repaid from operating cash flow ($976K EBITDA year 3) or conventional refinance."])

    # 11 risks
    slide("Risks, stated plainly",
          bullets=[
            "Census ramp: if patients migrate slower than planned, margins compress - mitigated by the reserve, a working-capital facility, and management salary deferral commitments.",
            "Medicare payment timing: after a change of ownership, the first claims can be held 30-60 days - our cash plan assumes this.",
            "Regulatory: hospice is heavily regulated (surveys, cap limits, enrollment rules) - the team has operated under these rules for a decade.",
            "This is a private, illiquid investment - notes are not transferable and there is no public market."])

    # 12 ask
    slide("The ask",
          bullets=[
            "$500,000 in notes; $250,000 already committed by the founding investors on the equity side.",
            "Minimum participation $50,000; documents ready for review with counsel.",
            "Next step: a call or lunch with Geoff, then the data room (full model, license diligence, team agreements).",
            "Geoff Schackmann  |  geoff@azaleahospice.com  |  480-495-5474"],
          small="Confidential. Not an offer to sell or a solicitation. Any offer is made only through definitive offering documents to accredited investors.")

    path = os.path.join(OUT, "Azalea_Investor_Deck.pptx")
    prs.save(path); print("wrote", path)


# ============================ DOCX HELPERS ============================
def _doc():
    d = Document()
    st = d.styles["Normal"].font; st.name = "Calibri"; st.size = DPt(10.5)
    d.styles["Normal"].paragraph_format.space_after = DPt(5)
    for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(d.sections[0], m, DIn(0.8))
    return d

def _p(d, text, bold=False, italic=False, size=10.5, color=None):
    p = d.add_paragraph(); r = p.add_run(text)
    r.bold = bold; r.italic = italic; r.font.size = DPt(size)
    if color: r.font.color.rgb = color
    return p

def _h(d, text):
    p = d.add_paragraph(); r = p.add_run(text)
    r.bold = True; r.font.size = DPt(11); r.font.color.rgb = DNAVY
    p.paragraph_format.space_before = DPt(8)
    return p


# ============================ TEASER ============================
def teaser():
    d = _doc()
    _p(d, "AZALEA HOSPICE & PALLIATIVE CARE - Tyler, Texas", bold=True, size=14, color=DNAVY)
    _p(d, "Private note offering - the 60-second version", italic=True, size=10, color=DGREY)
    _p(d, "")
    _p(d, "What this is: a Tyler hospice launch built on an active Medicare + Medicaid provider number - "
          "a license Texas and CMS are no longer issuing. Proven operating team, patient census that "
          "arrives with the team, and the license purchase already financed by a bank.")
    _p(d, "The raise: $500,000 in promissory notes. 10% per year, paid quarterly in cash. 3-year maturity, "
          "prepayable anytime. Minimum $50,000. Simple 1099 interest - no K-1s, no ownership strings.")
    _p(d, "Where your money goes: working capital and growth, not the license (that's financed). The plan "
          "runs 24 patients by month 2, 34 by month 6, 40 by month 12 - at ~$172/day Medicare per-diem "
          "and 55%+ contribution margins.")
    _p(d, "Who's already in: $250,000 of founding capital committed, led by James E. Bullard.")
    _p(d, "Why the asset holds value: the only two comparable Medicare+Medicaid licenses publicly listed "
          "in Texas asked ~$400,000 - and the federal freeze on new hospice enrollments makes existing "
          "licenses scarcer every month.")
    _p(d, "The team: Geoff Schackmann (founder, 10+ years operating East Texas hospice/home care) with a "
          "full leadership bench - Executive Director, Director of Nursing, Director of Sales - all "
          "holding performance-vested equity.")
    _p(d, "Next step: call Geoff. 480-495-5474 / geoff@azaleahospice.com. Data room and note documents "
          "available for review with your counsel.")
    _p(d, "")
    _p(d, "Confidential. For invited accredited investors only. This summary is not an offer; any offer is "
          "made solely through definitive offering documents.", italic=True, size=8.5, color=DGREY)
    path = os.path.join(OUT, "Azalea_Teaser_One_Pager.docx")
    d.save(path); print("wrote", path)


# ============================ FAQ ============================
def faq():
    d = _doc()
    _p(d, "AZALEA HOSPICE - INVESTOR FAQ (Note Offering)", bold=True, size=13, color=DNAVY)
    _p(d, "Plain answers. Definitive terms live in the note documents your counsel will review.", italic=True, size=9, color=DGREY)
    qa = [
     ("What exactly am I buying?", "A promissory note from the company: 10% annual interest paid quarterly in cash, principal returned in 3 years (or earlier - we can prepay without penalty). You are a lender, not an owner."),
     ("Where does my interest come from?", "Operating cash flow. Quarterly interest on the full $500,000 is $12,500. By the second quarter the business generates several times that in EBITDA; before that, the raise proceeds themselves cover it. Coverage math is in the model in the data room."),
     ("What stands behind the note?", "The operating business and its license. The note is subordinate to the bank loan that financed the license purchase (about $500K, 3-year amortization) - that is disclosed plainly in the documents. The license itself is a scarce asset: comparable Medicare+Medicaid licenses list around $400K and can no longer be created."),
     ("What happens if the census ramps slower than planned?", "Three cushions, in order: the $100K reserve inside the raise, a working-capital facility, and management salary deferrals the operating agreement already contemplates. We stress-tested 85% census with 10% higher payroll - obligations are still serviced."),
     ("Why notes instead of equity?", "Cleaner for everyone. The company is an S corporation with a settled ownership group; notes avoid ownership complexity and give you a defined return and exit date. You get a 1099-INT, not a K-1."),
     ("Can I convert to equity later?", "Not automatically. If an equity round ever happens, noteholders will hear about it first, but this instrument is straight debt."),
     ("When and how do I get paid?", "Interest hits your account quarterly starting the first full quarter after closing. Principal comes back at month 36, or sooner if we prepay."),
     ("What reporting do I get?", "Quarterly: census, revenue, EBITDA, cash, and progress vs plan - the same one-page dashboard management uses. Annual financials prepared by our CPA."),
     ("Who else is in?", "Founding capital of $250,000 is committed - anchor investor James E. Bullard plus the founder. The leadership team holds equity that vests only on performance."),
     ("Who runs the company?", "Geoff Schackmann is the manager and largest owner, with a decade-plus operating hospice in East Texas. The clinical and sales leadership all come from the same market and book of relationships."),
     ("What are the real risks?", "Slower census, Medicare payment timing after the ownership change (first claims can be held 30-60 days - our plan assumes it), regulatory surveys, and illiquidity: there is no market for these notes and you should plan to hold to maturity."),
     ("Am I eligible to invest?", "This is a private placement for accredited investors (income $200K+/yr, $300K joint, or $1M+ net worth excluding your home). You'll complete a short questionnaire; if you refer patients to hospice professionally, tell us up front - healthcare law restricts investments from referral sources and our counsel must clear it."),
    ]
    for q, a in qa:
        _h(d, q); _p(d, a)
    _p(d, "")
    _p(d, "Not an offer to sell securities. Offer made only through definitive documents to accredited "
          "investors under Regulation D. Consult your own advisors.", italic=True, size=8.5, color=DGREY)
    path = os.path.join(OUT, "Azalea_Investor_FAQ.docx")
    d.save(path); print("wrote", path)


# ============================ DATA ROOM + QUESTIONNAIRE OUTLINE ============================
def dataroom():
    d = _doc()
    _p(d, "AZALEA HOSPICE - DATA ROOM CHECKLIST & COMPLIANCE OUTLINE", bold=True, size=13, color=DNAVY)
    _h(d, "A. Data room contents (assemble before first investor meeting)")
    for x in ["36-month financial model (Rev 3.00 dynamic proforma) + KPI dashboard",
              "License diligence: CCN/certification status, survey history, 36-month-rule position, Medicaid contract",
              "Purchase agreement + bank financing terms for the license",
              "Amended & Restated Operating Agreement + subsidiary OA",
              "Team: leadership bios, equity/vesting summary, key-person provisions",
              "Market: license comps exhibit, census/referral plan, prior-operation actuals",
              "Note documents (from securities counsel): note purchase agreement, form of note, subordination terms, risk factors",
              "Insurance certificates; entity good-standing certificates"]:
        _p(d, "  [ ]  " + x)
    _h(d, "B. Accredited investor questionnaire - OUTLINE for counsel (not a legal form)")
    for x in ["Identity, residence (US individual), investment amount",
              "Accreditation basis: income test / net-worth test / license-holder",
              "Sophistication & risk acknowledgments (illiquidity, subordination, loss of capital)",
              "Source-of-funds representation",
              "HEALTHCARE ITEM: does the investor (or spouse/practice) refer patients to hospice or certify terminal illness? If yes - route to healthcare regulatory counsel BEFORE acceptance (Anti-Kickback Statute)",
              "OIG/SAM exclusion representation",
              "Pre-existing-relationship note (who introduced, when) - supports the private-offering posture"]:
        _p(d, "  -  " + x)
    _h(d, "C. Compliance guardrails (counsel to confirm; this is not legal advice)")
    for x in ["Reg D 506(b): no general solicitation - no public advertising, no mass outreach, warm introductions only; accredited investors only as a practice",
              "File Form D within 15 days of first sale + Texas blue-sky notice",
              "AKS: no notes to physicians or others in a position to refer patients without written clearance from healthcare regulatory counsel; document the analysis",
              "Every recipient of the deck/teaser logged (name, date, who introduced) - keeps the offering private and provable",
              "Subordination to the bank license loan disclosed in every document that mentions security"]:
        _p(d, "  -  " + x)
    path = os.path.join(OUT, "Azalea_DataRoom_and_Compliance.docx")
    d.save(path); print("wrote", path)


# ============================ OUTREACH SCRIPTS ============================
def scripts():
    d = _doc()
    _p(d, "OUTREACH SCRIPTS - warm-path raise (Geoff's voice)", bold=True, size=13, color=DNAVY)
    _p(d, "Rule of the raise: nobody hears about this except through a personal introduction. No blasts, no ads, no posts.", italic=True, size=9.5)
    _h(d, "1. The connector ask (to someone who knows the prospect)")
    _p(d, '"I\'m putting together a small group of local investors for the hospice we\'re launching in Tyler. '
          'It\'s a note - 10% paid quarterly, 3 years - not equity, so it\'s simple. $250K is already committed. '
          'I\'m only talking to people who come recommended by someone I trust. Do you think [NAME] would want '
          'to hear about it? If so, could you make the intro - lunch or a call, whichever they prefer?"')
    _h(d, "2. First note after the introduction (email/text)")
    _p(d, '"[Connector] suggested we talk. Short version: I\'ve run hospice companies in East Texas for over '
          'ten years and I\'m launching one in Tyler on a Medicare-and-Medicaid license - the kind Texas isn\'t '
          'issuing anymore. The license purchase is already bank-financed and $250K of founding capital is '
          'committed. I\'m raising $500K in notes at 10%, paid quarterly, 3-year term. Happy to walk you '
          'through the numbers over lunch - no pressure either way. When works?"')
    _h(d, "3. The meeting close")
    _p(d, '"Here\'s what I\'d suggest: take the one-pager and the FAQ, have your CPA or attorney look at the '
          'note documents, and let\'s talk again Friday. Minimum is $50K. I\'d rather have ten local people '
          'who know the market than one stranger. Any question you\'ve got, ask me straight - you\'ll get a '
          'straight answer."')
    _h(d, "4. If they ask 'what if it goes wrong?'")
    _p(d, '"Fair question. Three things: we raised more than we need so quarterly interest is covered even '
          'before the business is; if the census comes slower, the leadership - including me - defers salary '
          'before your check is ever late; and worst case, the license itself is worth several hundred '
          'thousand dollars to the next buyer because nobody can create one anymore. You\'re also behind the '
          'bank, and I\'ll show you exactly what that means - I don\'t hide the order of payments."')
    _h(d, "5. If the prospect is a physician")
    _p(d, '"One thing I have to ask because healthcare law requires it: do you or your practice refer '
          'patients to hospice? If so, our lawyers have to look at it before we can even talk terms - it\'s '
          'an anti-kickback thing, it protects both of us. If you\'re retired or your specialty doesn\'t '
          'touch hospice referrals, we\'re fine to proceed."')
    path = os.path.join(OUT, "Azalea_Outreach_Scripts.docx")
    d.save(path); print("wrote", path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    deck(); teaser(); faq(); dataroom(); scripts()
