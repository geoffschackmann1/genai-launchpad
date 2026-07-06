"""QA for the Rev 3.xx dynamic proforma (PE-grade pass).

  1. 0 formula errors
  2. BS check = 0 all 36 months (no plugs)
  3. Collections conservation
  4. Census: EOM M2=24, M6=34, M12=40 (user spec 7/8)
  5. BASE = $500K/6%/36 note: seller paid off Sept ($375K), no Jan balloon, $15,211/mo from Oct
  6. Revolver: cash never < -0.01; revolver never > commitment; no draw while cash > floor
  7. Checks-tab MASTER = 0
  8. EBITDA margin M12 in 20-25% band (pre-debt-service, contingency below the line)
  9. Toggle test (refi OFF): Jan balloon = $275K principal; balloon-refi services it
 10. Zero hardcoded constants >=100 in formulas outside Control Tower
"""
import json, re, sys
import formulas
import openpyxl
from openpyxl.utils import get_column_letter as gcl

XLSX = "financial_models/output/Azalea_Hospice_Proforma_Rev3.00_DYNAMIC.xlsx"
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

    chk("BASE: seller paid off Sept ($375K)", abs(num(g("Cash Flow & Runway", f"E{CF['s_prin']}")) - 375000) < 1)
    chk("BASE: no January balloon", abs(num(g("Cash Flow & Runway", f"I{CF['s_prin']}"))) < 1)
    chk("BASE: $15,211/mo from Oct", abs(num(g("Cash Flow & Runway", f"F{CF['refi_pmt']}")) - 15210.97) < 1)

    cash = [num(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['cash']}")) for i in range(36)]
    rev = [num(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['rev_bal']}")) for i in range(36)]
    floor = num(g("Control Tower", addr["floor"].split("!")[1].replace("$", "")))
    lim = num(g("Control Tower", addr["rev_limit"].split("!")[1].replace("$", "")))
    chk("revolver: cash never negative", min(cash) > -0.01, f"(min {min(cash):,.0f})")
    chk("revolver: never over commitment", max(rev) <= lim + 0.01, f"(max {max(rev):,.0f})")
    excl = all(not (rv > 0.01 and c > floor + 1) for rv, c in zip(rev, cash))
    chk("revolver: no balance while cash above floor", excl)
    print(f"     info: peak revolver ${max(rev):,.0f} | M12 cash ${cash[11]:,.0f} | M36 cash ${cash[35]:,.0f}")

    master = num(g("Checks", f"B{R['Checks']['master']}"))
    chk("Checks tab MASTER = 0", master == 0, f"({master})")

    P = R["P&L"]
    em12 = num(g("P&L", f"N{P['em']}"))
    chk("EBITDA margin M12 in 20-25% band", 0.20 <= em12 <= 0.25, f"({em12:.1%})")
    e_y = [sum(num(g("P&L", f"{gcl(3+i)}{P['ebitda']}")) for i in range(a, b)) for a, b in [(0,12),(12,24),(24,36)]]
    n_y = [sum(num(g("P&L", f"{gcl(3+i)}{P['npr']}")) for i in range(a, b)) for a, b in [(0,12),(12,24),(24,36)]]
    print(f"     info: EBITDA Y1 ${e_y[0]:,.0f} ({e_y[0]/n_y[0]:.1%}) | Y2 ${e_y[1]:,.0f} ({e_y[1]/n_y[1]:.1%}) | Y3 ${e_y[2]:,.0f} ({e_y[2]/n_y[2]:.1%})")

    # toggle test: refi OFF -> balloon path
    wb = openpyxl.load_workbook(XLSX)
    aws = wb["Control Tower"]
    aws[addr["refi_on"].split("!")[1].replace("$", "")] = 0
    tmp = "financial_models/output/_v3_off.xlsx"; wb.save(tmp)
    _, g2 = evaluate(tmp)
    chk("refi OFF: Jan balloon principal = $275,000", abs(num(g2("Cash Flow & Runway", f"I{CF['s_prin']}")) - 275000) < 1)
    chk("refi OFF: seller end-Dec = $275,000", abs(num(g2("Cash Flow & Runway", f"H{CF['s_end']}")) - 275000) < 1)
    import os; os.remove(tmp)

    # hardcode scan
    wb2 = openpyxl.load_workbook(XLSX)
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
