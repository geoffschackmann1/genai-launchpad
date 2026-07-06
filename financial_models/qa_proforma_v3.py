"""QA for the Rev 3.00 dynamic proforma.

  1. 0 formula errors (formulas evaluation)
  2. Balance-sheet check row = 0 in all 36 months (no plugs)
  3. Collections conservation: cum collected + ending AR == cum NPR
  4. Seller schedule: end-Dec balance $275,000; Jan balloon principal $275,000
  5. Refi toggle: seller paid off in Sept; no Jan balloon; $15,211/mo from Oct
  6. Zero hardcoded numeric constants >100 embedded in formulas (the Rev 2.00 sin)
  7. Sanity: M6 gross margin 45-65%; census EOM M6 = 35
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
    PASS, FAIL = PASS + ok, FAIL + (not ok)

def evaluate(path):
    sol = formulas.ExcelModel().loads(path).finish().calculate()
    vals = {}
    for k, v in sol.items():
        try: vals[k.upper()] = v.value[0, 0]
        except Exception: pass
    fn = path.split("/")[-1].upper()
    def g(sheet, ref):
        v = vals.get(f"'[{fn}]{sheet.upper()}'!{ref.upper()}")
        try: return float(v)
        except Exception: return v
    return vals, g

def main():
    rm = json.load(open(ROWMAP)); R = rm["rows"]; addr = rm["addr"]
    vals, g = evaluate(XLSX)

    errs = [k for k, v in vals.items() if isinstance(v, str) and v.startswith("#")]
    chk("0 formula errors", not errs, f"({len(errs)})")
    for e in errs[:6]: print("     ERR", e, vals[e])

    B = R["Balance Sheet"]
    def val0(x):
        return abs(x) if isinstance(x, (int, float)) else 9e9
    checks = [val0(g("Balance Sheet", f"{gcl(3+i)}{B['check']}")) for i in range(36)]
    chk("balance sheet check = 0 all 36 months (no plugs)", max(checks) < 1.0, f"(max {max(checks):,.2f})")

    CF = R["Cash Flow & Runway"]
    cum_npr = sum(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['earned']}") for i in range(36))
    cum_col = sum(g("Cash Flow & Runway", f"{gcl(3+i)}{CF['coll']}") for i in range(36))
    ar_end = g("Cash Flow & Runway", f"AL{CF['ar']}")
    chk("collections conservation (cum coll + AR = cum NPR)", abs(cum_col + ar_end - cum_npr) < 1.0,
        f"({cum_col:,.0f}+{ar_end:,.0f} vs {cum_npr:,.0f})")

    chk("seller balance end-Dec = $275,000", abs(g("Cash Flow & Runway", f"H{CF['s_end']}") - 275000) < 1,
        f"({g('Cash Flow & Runway', f'H{CF[chr(39)+chr(39)] if False else chr(115)+chr(95)+chr(101)+chr(110)+chr(100)}'):,.0f})" if False else "")
    chk("Jan balloon principal = $275,000", abs(g("Cash Flow & Runway", f"I{CF['s_prin']}") - 275000) < 1)
    chk("census EOM M6 = 35.0", abs(g("Census Waterfall", f"H{R['Census Waterfall']['eom']}") - 35.0) < 0.1)
    gm6 = g("P&L", f"H{R['P&L']['gm']}")
    chk("M6 gross margin 45-65%", 0.45 <= gm6 <= 0.65, f"({gm6:.1%})")
    cash = [g("Cash Flow & Runway", f"{gcl(3+i)}{CF['cash']}") for i in range(36)]
    print(f"     info: min cash ${min(cash):,.0f} (M{cash.index(min(cash))+1}) | M12 ${cash[11]:,.0f} | M36 ${cash[35]:,.0f}")

    # refi-ON scenario
    wb = openpyxl.load_workbook(XLSX)
    aws = wb["Control Tower"]
    aws[addr["refi_on"].split("!")[1].replace("$", "")] = 1
    tmp = "financial_models/output/_v3_refi.xlsx"; wb.save(tmp)
    _, g2 = evaluate(tmp)
    chk("refi ON: seller principal Sept = $375,000", abs(g2("Cash Flow & Runway", f"E{CF['s_prin']}") - 375000) < 1)
    chk("refi ON: no Jan balloon", abs(g2("Cash Flow & Runway", f"I{CF['s_prin']}")) < 1)
    chk("refi ON: $15,211/mo from Oct", abs(g2("Cash Flow & Runway", f"F{CF['refi_pmt']}") - 15210.97) < 1)
    import os; os.remove(tmp)

    # hardcode hunt (the Rev 2.00 sin): constants >=100 inside formulas, excluding
    # cell refs, PMT month counts and the month-index comparisons
    wb2 = openpyxl.load_workbook(XLSX)
    pat = re.compile(r"(?<![A-Z$.\d])(\d{3,}(?:\.\d+)?)")
    offenders = []
    for sh in wb2.sheetnames:
        if sh == "Control Tower": continue
        for row in wb2[sh].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for mnum in pat.findall(c.value):
                        if float(mnum) >= 100 and mnum not in ("100",):
                            offenders.append((sh, c.coordinate, c.value[:60]))
    chk("0 hardcoded constants >=100 in formulas (outside Control Tower)", not offenders, f"({len(offenders)})")
    for s, co, f in offenders[:6]: print("     HARD", s, co, f)

    print(f"\n{PASS} passed, {FAIL} failed")
    sys.exit(1 if FAIL else 0)

if __name__ == "__main__":
    main()
