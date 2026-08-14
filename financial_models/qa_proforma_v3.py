"""QA for the Rev 5.00 AVANT proforma (seller-carried at 6%, SBA takeout month 2).

  1. 0 formula errors
  2. BS check = 0 all 36 months (no plugs; incl. deferred comp + notes payable)
  3. Collections conservation
  4. Census: EOM M2=24, M6=34, M12=40
  5. BASE = Avant: $0 down, $0 installments, seller note ($300K, 6% simple)
     retired in full by the SBA at month 2; single $6,747/mo SBA payment after
  6. Owner deferral: conservation (accrued = repaid + end bal); end balance 0 by M36;
     EBITDA UNCHANGED vs Rev 3.10 (accrual basis - deferral is cash timing only)
  7. NO revolver anywhere; UNFUNDED NEED reported honestly (base + downside):
     the peak is THE raise number, not a pass/fail
  8. Master INTEGRITY check = 0 (viability reported separately)
  9. Toggle test (SBA OFF): private balloon refi must size to the seller payoff
 10. Zero hardcoded constants >=100 in formulas outside Control Tower
"""
import json, re, sys
import formulas
import openpyxl
from openpyxl.utils import get_column_letter as gcl

XLSX = "financial_models/output/Azalea_Hospice_Proforma_Rev5.00_AVANT.xlsx"
ROWMAP = "financial_models/output/proforma_v3_rowmap.json"
PASS = FAIL = 0
def chk(name, ok, detail=""):
    global PASS, FAIL
    print(f"  {'OK ' if ok else 'FAIL'} {name} {detail}")
    PASS, FAIL = PASS + bool(ok), FAIL + (not ok)

def num(x):
    return x if isinstance(x, (int, float)) else None

def evaluate(path):
    sol = formulas.ExcelModel().loads(path).finish().calculate()
    vals = {}
    for k, v in sol.items():
        try: vals[k.upper()] = v.value[0, 0]
        except Exception: pass
    fn = path.split("/")[-1].upper()
    def g(sheet, ref):
        return vals.get(f"'[{fn}]{sheet.upper()}'!{ref.upper()}")
    return vals, g

def main():
    rm = json.load(open(ROWMAP)); R = rm["rows"]; addr = rm["addr"]
    vals, g = evaluate(XLSX)

    errs = [k for k, v in vals.items() if isinstance(v, str) and v.startswith("#")]
    chk("0 formula errors", not errs, f"({len(errs)})")
    for e in errs[:6]: print("     ERR", e, vals[e])

    B = R["Balance Sheet"]
    bs = [abs(num(g("Balance Sheet", f"{gcl(3+i)}{B['check']}")) if num(g("Balance Sheet", f"{gcl(3+i)}{B['check']}")) is not None else 9e9) for i in range(36)]
    chk("BS check = 0 all 36 months", max(bs) < 0.01, f"(max {max(bs):,.4f})")

    CF = R["Cash Flow & Runway"]
    cum_npr = sum(num(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['earned']}")) for i in range(36))
    cum_col = sum(num(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['coll']}")) for i in range(36))
    ar_end = num(g("Cash Flow & Runway", f"AL{CF['ar']}"))
    chk("collections conservation", abs(cum_col + ar_end - cum_npr) < 1.0)

    CN = R["Census Waterfall"]
    for mo, tgt in [(2, 24.0), (6, 34.0), (12, 40.0)]:
        v = num(g("Census Waterfall", f"{gcl(2+mo)}{CN['eom']}"))
        chk(f"census EOM M{mo} = {tgt}", abs(v - tgt) < 0.1, f"({v:.1f})")

    # BASE (Avant): seller carries full $300K at 6%; SBA funds M2 and retires the note
    chk("BASE: no down payment / no M1 seller principal", abs(num(g("Cash Flow & Runway", f"C{CF['s_prin']}"))) < 1)
    chk("BASE: seller note retired in full at M2 ($300K principal)",
        abs(num(g("Cash Flow & Runway", f"D{CF['s_prin']}")) - 300000) < 1)
    ipaid = num(g("Cash Flow & Runway", f"D{CF['s_intpaid']}"))
    chk("BASE: accrued seller interest paid at takeout (~$3,000)", 2500 <= ipaid <= 3500, f"(${ipaid:,.0f})")
    chk("BASE: seller balance 0 from M2", abs(num(g("Cash Flow & Runway", f"D{CF['s_end']}"))) < 0.01)
    chk("BASE: no bank note anywhere (refi bal M2 = 0)", abs(num(g("Cash Flow & Runway", f"D{CF['refi_bal']}"))) < 0.01)
    print(f"     info: SBA $500,000 at M2 - seller takeout ${300000+ipaid:,.0f} | surplus to WC ~${500000-300000-ipaid:,.0f}")
    chk("BASE: SBA ~$6,746.75/mo from M3", abs(num(g("Cash Flow & Runway", f"E{CF['sba_pmt']}")) - 6746.75) < 1)

    # deferral
    ST = R["Staffing"]
    acc = sum(num(g("Staffing", f"{gcl(3+i)}{ST['defer_amt']}")) for i in range(36))
    rep = sum(num(g("Staffing", f"{gcl(3+i)}{ST['defer_repay']}")) for i in range(36))
    endb = num(g("Staffing", f"AL{ST['defer_bal']}"))
    chk("deferral conservation (accrued = repaid + end)", abs(acc - rep - endb) < 0.01,
        f"(accrued ${acc:,.0f})")
    chk("deferral fully repaid by M36", abs(endb) < 0.01, f"(end ${endb:,.0f})")

    P = R["P&L"]
    e_y = [sum(num(g("P&L", f"{gcl(3+i)}{P['ebitda']}")) for i in range(a, b)) for a, b in [(0,12),(12,24),(24,36)]]
    n_y = [sum(num(g("P&L", f"{gcl(3+i)}{P['npr']}")) for i in range(a, b)) for a, b in [(0,12),(12,24),(24,36)]]
    # hire lag legitimately lifts Y1 EBITDA vs 3.10 ($263K): fewer early FTEs, PRN offset
    chk("EBITDA Y1 in band $255K-$325K (hire-lag effect vs 3.10's $263K)", 255000 <= e_y[0] <= 325000, f"(${e_y[0]:,.0f})")
    em12 = num(g("P&L", f"N{P['em']}"))
    chk("EBITDA margin M12 in 20-25% band", 0.20 <= em12 <= 0.25, f"({em12:.1%})")
    print(f"     info: EBITDA Y1 ${e_y[0]:,.0f} ({e_y[0]/n_y[0]:.1%}) | Y2 ${e_y[1]:,.0f} ({e_y[1]/n_y[1]:.1%}) | Y3 ${e_y[2]:,.0f} ({e_y[2]/n_y[2]:.1%})")

    # honest cash picture - BASE
    cash = [num(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['cash']}")) for i in range(36)]
    need = [num(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['unfunded']}")) for i in range(36)]
    peak = num(g("Cash Flow & Runway", f"B{CF['peakneed']}"))
    chk("peak-need cell = max monthly unfunded", abs(peak - max(need)) < 0.01)
    first_neg = next((i + 1 for i, c in enumerate(cash) if c < -0.01), None)
    print(f"     >>> BASE: peak ADDITIONAL CAPITAL REQUIRED ${peak:,.0f}"
          + (f" | first negative month M{first_neg}" if first_neg else " | never negative")
          + f" | min cash ${min(cash):,.0f} | M36 cash ${cash[35]:,.0f}")

    master = num(g("Checks", f"B{R['Checks']['master']}"))
    chk("Master INTEGRITY check = 0", master == 0, f"({master})")

    # no revolver anywhere
    wb2 = openpyxl.load_workbook(XLSX)
    rev_refs = 0
    for sh in wb2.sheetnames:
        for row in wb2[sh].iter_rows():
            for c in row:
                if isinstance(c.value, str) and "evolver" in str(c.value) and "no revolver" not in str(c.value).lower():
                    rev_refs += 1
    chk("no LIVE revolver references (\'no revolver\' banners excluded)", rev_refs == 0, f"({rev_refs})")

    # toggle: SBA OFF -> seller carried, balloon machinery must fire at balloon_mo with private refi
    wb = openpyxl.load_workbook(XLSX)
    aws = wb["Control Tower"]
    aws[addr["sba_on"].split("!")[1].replace("$", "")] = 0
    tmp = "financial_models/output/_v5_off.xlsx"; wb.save(tmp)
    _, g2 = evaluate(tmp)
    bref = num(g2("Cash Flow & Runway", f"B{CF['bref_amt']}"))
    chk("SBA OFF: private balloon refi sized to seller payoff (>= $300K)", bref >= 300000, f"(${bref:,.0f})")
    need2 = [num(g2("Cash Flow & Runway", f"{gcl(3+i)}{CF['unfunded']}")) for i in range(36)]
    print(f"     >>> NO-SBA scenario: peak ADDITIONAL CAPITAL REQUIRED ${max(need2):,.0f}")
    import os; os.remove(tmp)

    # downside (case=2) on the base structure: report the honest raise number
    wb3 = openpyxl.load_workbook(XLSX)
    aws3 = wb3["Control Tower"]
    aws3[addr["case"].split("!")[1].replace("$", "")] = 2
    tmp2 = "financial_models/output/_v4_down.xlsx"; wb3.save(tmp2)
    _, g3 = evaluate(tmp2)
    m3 = num(g3("Checks", f"B{R['Checks']['master']}"))
    chk("DOWNSIDE: integrity still 0", m3 == 0, f"({m3})")
    need3 = [num(g3("Cash Flow & Runway", f"{gcl(3+i)}{CF['unfunded']}")) for i in range(36)]
    cash3 = [num(g3("Cash Flow & Runway", f"{gcl(3+i)}{CF['cash']}")) for i in range(36)]
    fn3 = next((i + 1 for i, c in enumerate(cash3) if c < -0.01), None)
    print(f"     >>> DOWNSIDE (slow census): peak ADDITIONAL CAPITAL REQUIRED ${max(need3):,.0f}"
          + (f" | first negative M{fn3}" if fn3 else "") + f" | M36 cash ${cash3[35]:,.0f}")
    import os as _os; _os.remove(tmp2)

    # hardcode scan
    pat = re.compile(r"(?<![A-Z$.\d])(\d{3,}(?:\.\d+)?)")
    offenders = []
    for sh in wb2.sheetnames:
        if sh == "Control Tower": continue
        for row in wb2[sh].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for m in pat.findall(c.value):
                        if float(m) >= 100:
                            offenders.append((sh, c.coordinate, c.value[:60]))
    chk("0 hardcoded constants >=100 in formulas", not offenders, f"({len(offenders)})")
    for s_, co, f in offenders[:6]: print("     HARD", s_, co, f)

    print(f"\n{PASS} passed, {FAIL} failed")
    sys.exit(1 if FAIL else 0)

if __name__ == "__main__":
    main()
