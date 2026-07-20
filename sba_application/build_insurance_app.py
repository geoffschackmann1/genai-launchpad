"""Fill Todd Plummer's PL/GL insurance application + reply email.

Named insured: Refuge Hospice, LLC dba Azalea Hospice & Palliative Care (EIN 92-1541610),
parent Tyler Hospice Hold, LLC. Effective 8/1/2026. Numbers basis: Rev 4.10 year one
(gross revenue ~$1.93M; W2 payroll $972,167; ~12 staff at launch to ~18-20 by month 12).
Fields are numbered sequentially through the form (mapped by page coordinates 7/16).
Output: sba_application/14_insurance_2026-07-16/
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docfmt

from docx import Document
from docx.shared import Pt, RGBColor
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject

SRC = "/root/.claude/uploads/7fe25b71-1087-5ae8-a835-ec8dc7454817/b4a4b196-apdmedhomehealthgeneralliabilityappplication144APP0724.pdf"
OUT = "sba_application/14_insurance_2026-07-16/"

TEXT = {
    # ---- page 1: Section I general
    "1": "Refuge Hospice, LLC dba Azalea Hospice & Palliative Care (parent: Tyler Hospice Hold, LLC)",
    "2": "13387 Hwy 69 N, Tyler, TX 75706",
    "3": "azaleahospice.com",
    "4": "2023 (confirming)",
    "5": "92-1541610",
    "13": "100% of membership interests being acquired by Tyler Hospice Hold, LLC (WY): 49% on ~8/1/2026, remaining 51% on 1/15/2027 per the CMS 36-month rule (42 CFR 424.550(b)). Operations under Tyler Hospice Hold management from 8/1/2026.",
    "21": "70",   # patient's home
    "23": "20",   # nursing home
    "24": "10",   # assisted living
    # ---- page 2
    "39": "1,933,000",   # projected next 12 months
    "40": "0",           # current YTD (if present)
    "41": "0",           # 1st year prior (minimal operations)
    "42": "0",           # 2nd year prior
    # staffing table Q15: employees / ICs / billable hours
    "50": "3",            # CNA employees
    "65": "0",            # LPN
    "89": "5",  "90": "2",   # RN employees (3 case mgrs + DON + ADON), RN ICs (PRN)
    "95": "1",  "96": "1",   # Social Worker employee + PRN IC
    "101": "5",              # Volunteers
    "104": "Chaplain (PRN contract); contracted physician Medical Director; 4 administrative/leadership employees (Executive Director, Director of Sales, Volunteer/Bereavement Coordinator, plus office support)",
    "105": "4", "106": "2",
    # ---- page 3
    "117": "100",  # % hospice care
    "131": "~400 visits/month at steady census (RN, aide, MSW, chaplain combined)",
    "132": "0",
    "136": "CHAP accredited - effective 1/8/2024, expires 1/8/2027. CLIA #45D2282700.",
    "137": "NHPCO / Texas & New Mexico Hospice Organization (membership in process)",
    "138": "QAPI program per Medicare CoP 418.58: quarterly IDG record review, infection control, incident/complaint tracking, HIS & CAHPS quality reporting; dedicated QAPI role added at census 40. Leadership team has 10+ years hospice operations experience.",
    # ---- page 4
    "146": "Quarterly (pre-employment, random, and for-cause)",
    "147": "Physicians, hospital discharge planners, nursing facilities, ALFs and community referrals. No remuneration for referrals.",
    "151": "[To be appointed - contracted hospice-experienced physician]; specialty: Hospice & Palliative Medicine",
    "154": "$1,000,000 / $3,000,000 (required by contract)",
    "156": "Directs the interdisciplinary group (IDG) per Medicare CoPs; clinical protocol oversight jointly with the Director of Nursing.",
    "159": "Contracted physicians (medical director) must carry $1M/$3M professional liability; PRN 1099 clinicians must evidence individual PL coverage.",
    "161": "Zero-tolerance abuse policy: two-person visit protocols where indicated, background/registry checks, mandatory reporting training at hire and annually; policy reviewed annually.",
    # ---- page 5
    "174": "NONE (startup)",
    "199": "N/A - requesting retroactive date = policy inception",
    # ---- page 6: GL locations
    "207": "Azalea Hospice & Palliative Care (administrative office)",
    "208": "13387 Hwy 69 N, Tyler, TX 75706",
    "209": "Admin office - no patient care",
    "210": "1,607",
    "211": "Refuge Hospice suite",
    "212": "8746 Wurzbach #201E San Antonio",
    "213": "Transitional admin suite",
    "214": "~1,000",
    "228": "100",
    "232": "100",
    # maintenance table: subbed janitorial + general maintenance (landlord/CAM per lease)
    "251": "100",
    "253": "100",
    # ---- page 7: autos
    "274": "8",
    "288": "0",
    "289": "Annually",
    # ---- page 9: additional insureds
    "323": "Fair Investments, Ltd., PO Box 689, Tyler, TX 75710",
    "324": "Landlord (per lease requirement)",
    # ---- page 13
    "Title": "Managing Member, Tyler Hospice Hold, LLC (Manager of applicant)",
}

# radio groups / checkboxes: name -> state
CHECKS = {
    "9": "/Yes",     # entity type: LLC
    "12": "/No",     # Q7 owned/controlled by another entity -> details given anyway in 13? -> set Yes
    "17": "/Yes",    # firm type: Hospice
    "35": "/No",     # Q10 services in affiliated facility
    "37": "/No",     # Q11 other businesses
    "43": "/No",     # Q13 pain management
    "45": "/No",     # Q14 opioid prescribing (physician prescribes; agency does not - see email note)
    "130": None,     # DO provide hospice services - leave "do not provide" UNCHECKED
    "133": "/N/A",   # 18c inpatient beds in NF - none
    "135": "/Yes",   # Q19 accredited
    "139": "/Yes",   # Q22 background checks
    "141": "/Yes", "142": "/Yes", "143": "/Yes",  # state/federal/sex-offender registry
    "145": "/Yes",   # Q23 drug screens
    "148": "/Yes",   # Q25 each patient has attending physician
    "150": "/Yes",   # Q26 medical director
    "152": "/Yes",   # 26b direct patient care (face-to-face/recert)
    "153": "/Yes",   # 26b.i carries malpractice
    "155": "/Yes",   # 26c supervisory duties
    "157": "/Yes",   # Q27 back-up procedures
    "158": "/Yes",   # Q28 ICs required to carry PL
    "160": "/Yes",   # Q29 abuse-prevention policy
    "162": "/No",    # cyber: standalone policy
    "169": "/Yes",   # antivirus/firewalls
    "170": "/Yes",   # encryption
    "171": "/Yes",   # password mgmt
    "172": "/Yes",   # HIPAA/HITECH compliant
    "173": "/No",    # breach history
    "200": "/No",    # currently insured under CGL
    "201": "/Yes",   # want GL quote
    "202": "/No",    # ever declined/cancelled
    "204": "/No",    # any claim ever
    "205": "/No",    # aware of circumstances
    "227": "/Lease", # loc 1
    "229": "/No",    # loc 1 other occupants
    "231": "/Lease", # loc 2
    "233": "/Yes",   # loc 2 other occupants (multi-tenant building)
    "247": "/No",    # full-time maintenance staff
    "259": "/Yes",   # written inspection procedures
    "260": "/Yes",   # records retained 5 yrs
    "261": "/No",    # construction planned
    # fire-life safety 5a-h (single-story historic house office; see email note)
    "265": "/No",    # sprinkler
    "266": "/Yes",   # two marked exits
    "267": "/Yes",   # smoke detectors
    "268": "/No",    # emergency electrical
    "269": "/No",    # heat sensors
    "270": "/No",    # fire escapes (single story - N/A)
    "271": "/Yes",   # posted evacuation procedures
    "272": "/Yes",   # fire extinguishers
    "273": "/Yes",   # non-owned auto: private passenger
    "281": "/Yes",   # usage: errands
    "282": "/Yes",   # usage: regular sales/service calls (staff visits to patient homes)
    "290": "/Yes",   # MVRs checked
    "292": "/Yes",   # state minimum limits required
    "293": "/No",    # drivers with violations
    "295": "/Yes",   # prohibit unlicensed/DUI drivers
    "296": "/No",    # owned/leased/hired autos
    "304": "/No",    # auto claims 5 yrs
    "305": "/No", "306": "/No", "307": "/No", "308": "/No",  # 15a-d food/rec/gym/pool
    "309": "/0",     # 15e daycare
    "310": "/No", "311": "/No", "312": "/No",               # 15f-h events/fundraising/alcohol
    "313": "/No", "314": "/No",                              # 15i trade shows, 15j construction
    "315": "/No",    # Q16 rent equipment to others
    "320": "/Yes",   # Q17 consents for advertising
    "321": "/No",    # Q18 structure/function claims
    "330": "/No",    # Q20 sell products
    # products 29-36: none
    "334": "/No", "335": "/No", "336": "/No", "337": "/No", "338": "/No",
    "339": "/No", "340": "/No", "341": "/No", "342": "/No", "343": "/No",
    "344": "/No", "345": "/No", "346": "/No", "347": "/No", "348": "/No", "349": "/No",
    "351": "/No",
    "353": "/No",
    "354": "/1",     # prior GL: NONE checkbox
    "355": "/Yes",   # NONE box (belt & suspenders per layout)
    "391": "/No",    # Q38 GL claims
    "392": "/None",  # Q39 aware of GL circumstances
    "394": "/None",  # Q40 product defects
}
CHECKS["12"] = "/Yes"  # owned/controlled: Yes (Tyler Hospice Hold acquiring; details in 13)


def fill():
    reader = PdfReader(SRC)
    writer = PdfWriter()
    writer.append(reader)
    fields = reader.get_fields()
    text_vals = {k: v for k, v in TEXT.items() if k in fields}
    missing = [k for k in TEXT if k not in fields]
    for page in writer.pages:
        writer.update_page_form_field_values(page, text_vals)
    set_cb = skip_cb = 0
    for page in writer.pages:
        for a in page.get("/Annots", []):
            obj = a.get_object()
            nm = obj.get("/T")
            parent = obj.get("/Parent")
            pname = parent.get_object().get("/T") if parent else None
            target = nm if nm in CHECKS else (pname if pname in CHECKS else None)
            if target is None or CHECKS[target] is None:
                continue
            want = CHECKS[target]
            ap = obj.get("/AP", {})
            states = list(ap.get("/N", {}).keys()) if ap else []
            holder = parent.get_object() if parent else obj
            if want in states:
                holder[NameObject("/V")] = NameObject(want)
                obj[NameObject("/AS")] = NameObject(want)
                set_cb += 1
            elif not parent and states == ["/Off"]:
                skip_cb += 1
    # radio groups whose kids carry states: set /V on the field object directly
    root_fields = writer._root_object["/AcroForm"]["/Fields"]
    for fref in root_fields:
        fo = fref.get_object()
        nm = fo.get("/T")
        if nm in CHECKS and CHECKS[nm] is not None and "/V" not in fo:
            fo[NameObject("/V")] = NameObject(CHECKS[nm])
            for kid in fo.get("/Kids", []):
                ko = kid.get_object()
                kstates = list(ko.get("/AP", {}).get("/N", {}).keys())
                ko[NameObject("/AS")] = NameObject(CHECKS[nm] if CHECKS[nm] in kstates else "/Off")
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        writer._root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)
    dst = OUT + "Azalea Insurance Application FILLED.pdf"
    with open(dst, "wb") as f:
        writer.write(f)
    print(f"wrote {dst} | text fields {len(text_vals)} (missing {missing}) | checks set {set_cb}")


def _p(doc, text, bold=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(11)
    return p


def email():
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading("Draft reply - Todd Plummer (insurance quote)", 0)
    _p(doc, "To: Todd Plummer | From: gs@h-care.us | Re: Insurance quote - Azalea Hospice", bold=True)
    for para in [
        "Todd,",
        "Thanks for turning this around. The completed application is attached and your questions are answered "
        "below. Quick context so the quote lands right: we are acquiring Refuge Hospice, LLC, a Texas licensed "
        "Medicare and Medicaid hospice, and operating it as Azalea Hospice & Palliative Care out of Tyler. We take "
        "over operations August 1, so that is the effective date we need.",
        "Full legal name is Refuge Hospice, LLC dba Azalea Hospice & Palliative Care. Parent company is Tyler "
        "Hospice Hold, LLC (Wyoming), which acquires 49% on August 1 and the remaining 51% on January 15, 2027 "
        "under the CMS 36 month ownership rule, with operations under our management from August 1. Owners of "
        "Tyler Hospice Hold: Geoff Schackmann 39.9%, James Bullard 19.5%, Silas Shelton 13.3%, Dana Davenport "
        "13.3%, Bradley Woodard 13.3%, and a reserved 0.7% pool.",
        "FEIN is 92-1541610 for Refuge Hospice, LLC. I'm still confirming the year the LLC was established with "
        "the sellers - the license itself was CMS certified 1/8/2024 and is CHAP accredited through 1/8/2027.",
        "Estimated annual gross revenue is $1,933,000 for the first 12 months as census ramps from about 12 "
        "patients up to 40 over the year. Headcount starts around 12 at launch and grows to 18-20 by month 12. "
        "Estimated annual W2 payroll for the first 12 months is $972,167, straight from the staffing model. On "
        "top of that we'll have contracted 1099s - a physician medical director (about $48K/yr), PRN nurses and "
        "aides (about $30K/yr), and a PRN chaplain.",
        "Here's the workers comp payroll breakdown by class for the first 12 months:",
    ]:
        _p(doc, para)

    t = doc.add_table(rows=1, cols=2)
    t.style = "Light Grid Accent 1"
    for i, htxt in enumerate(["Class", "Annual payroll"]):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(htxt)
        r.bold = True
        r.font.size = Pt(10)
    for cls, amt in [
        ("Registered nurses (case managers, on-call)", "$276,667"),
        ("Hospice aides / CNAs", "$148,000"),
        ("Medical social worker", "$70,000"),
        ("Clinical leadership - RNs (Director of Nursing, ADON/Intake)", "$170,000"),
        ("Clerical / administrative (Executive Director, Volunteer & Bereavement Coordinator)", "$187,500"),
        ("Outside marketing / community liaison (Director of Sales)", "$120,000"),
        ("Total", "$972,167"),
    ]:
        cells = t.add_row().cells
        cells[0].text = ""
        r0 = cells[0].paragraphs[0].add_run(cls)
        r0.font.size = Pt(10)
        cells[1].text = ""
        r1 = cells[1].paragraphs[0].add_run(amt)
        r1.font.size = Pt(10)
        if cls == "Total":
            r0.bold = True
            r1.bold = True

    for para in [
        "A few things worth flagging so there are no surprises. Care is delivered in patients' homes, nursing "
        "facilities and assisted living facilities - the office itself is administrative only, with no patient "
        "care on premises. The Tyler office is a leased single story historic house (1,607 sq ft), so the fire "
        "safety answers reflect that: smoke detectors, marked exits, extinguishers, and posted evacuation "
        "procedures are all in place, but there's no sprinkler system or emergency generator. Let me know if that "
        "needs anything on your end.",
        "The license's current service area is Texas Region 8 (San Antonio and the surrounding counties), and "
        "we're relocating the primary service area to East Texas (Tyler / Smith County), so please make sure the "
        "policy territory covers both during the transition. Staff drive their own vehicles to patient homes, so "
        "we'll need hired and non-owned auto coverage - there's no patient transport involved. History is clean: "
        "no prior claims and no prior professional liability policies for this operation, so we'd want the retro "
        "date set at policy inception. We'll also need the landlord, Fair Investments, Ltd., named as additional "
        "insured on the GL per the lease, and once the SBA loan closes our lender will need certificates with "
        "lender's loss payee wording.",
        "These numbers are first year projections for a ramping startup, so happy to true them up at audit. What "
        "else do you need from me to get quotes moving? I'd like coverage bound by August 1.",
        "Thanks,",
        "Geoff",
    ]:
        _p(doc, para)
    docfmt.finalize(doc, "Azalea Hospice - Insurance Application Reply")
    path = OUT + "Reply Email Todd Plummer DRAFT.docx"
    doc.save(path)
    print("wrote", path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    fill()
    email()
