"""PPEO stress test on the Rev 4.10 proforma.

Palmetto's 3/20/2026 reactivation letter places Refuge under a provisional period of
enhanced oversight (42 CFR 424.527) with medical review effective on the FIRST claim.
Prepayment review can delay Medicare payment 60-120+ days. The model's existing
"CHOW hold (months)" input is exactly this lever: service month 1 is held `hold`
extra months beyond the normal +1 collection lag, month 2 held `hold-1`, etc. -
held claims release together when review clears (the realistic prepayment-review
shape). Base case already assumes hold = 1.

Scenarios (run from repo root: python3 -m financial_models.ppeo_stress):
  A. BASE (hold=1)                      - reference
  B. PPEO ~90 days  (hold=3)           - first cash lands month 5
  C. PPEO ~120+ days (hold=4)          - first cash lands month 6
  D. PPEO ~120+ days + slow census     - hold=4 on the downside census case

Reports per scenario: peak ADDITIONAL FUNDING REQUIREMENT, min cash, first
negative month, months negative, peak A/R (cash stuck in review), M36 cash.
"""
import json
import formulas
import openpyxl
from openpyxl.utils import get_column_letter as gcl

XLSX = "financial_models/output/Azalea_Hospice_Proforma_Rev4.10_DYNAMIC.xlsx"
ROWMAP = "financial_models/output/proforma_v3_rowmap.json"


def evaluate(path):
    sol = formulas.ExcelModel().loads(path).finish().calculate()
    vals = {}
    for k, v in sol.items():
        try:
            vals[k.upper()] = v.value[0, 0]
        except Exception:
            pass
    fn = path.split("/")[-1].upper()

    def g(sheet, ref):
        return vals.get(f"'[{fn}]{sheet.upper()}'!{ref.upper()}")
    return g


def run(name, overrides):
    """overrides: dict of Control Tower cell -> value (e.g. {'B37': 3})."""
    wb = openpyxl.load_workbook(XLSX)
    ct = wb["Control Tower"]
    for cell, val in overrides.items():
        ct[cell] = val
    tmp = "financial_models/output/_ppeo_tmp.xlsx"
    wb.save(tmp)
    g = evaluate(tmp)

    rm = json.load(open(ROWMAP))
    CF = rm["rows"]["Cash Flow & Runway"]
    CK = rm["rows"]["Checks"]

    cash = [g("Cash Flow & Runway", f"{gcl(3+i)}{CF['cash']}") for i in range(36)]
    need = [g("Cash Flow & Runway", f"{gcl(3+i)}{CF['unfunded']}") for i in range(36)]
    ar = [g("Cash Flow & Runway", f"{gcl(3+i)}{CF['ar']}") for i in range(36)]
    master = g("Checks", f"B{CK['master']}")

    peak_need = max(need)
    min_cash = min(cash)
    first_neg = next((i + 1 for i, c in enumerate(cash) if c < -0.01), None)
    n_neg = sum(1 for c in cash if c < -0.01)
    peak_ar = max(ar)
    peak_ar_mo = ar.index(peak_ar) + 1

    print(f"\n=== {name} ===")
    print(f"  integrity check:                 {master}")
    print(f"  peak ADDITIONAL FUNDING REQ'D:   ${peak_need:>10,.0f}"
          + (f"  (month {need.index(peak_need)+1})" if peak_need > 0 else ""))
    print(f"  minimum cash:                    ${min_cash:>10,.0f}")
    print(f"  first negative month:            {'M'+str(first_neg) if first_neg else 'never'}")
    print(f"  months cash below zero:          {n_neg}")
    print(f"  peak A/R (cash stuck in review): ${peak_ar:>10,.0f}  (month {peak_ar_mo})")
    print(f"  month-36 cash:                   ${cash[35]:>10,.0f}")

    import os
    os.remove(tmp)
    return peak_need, min_cash


if __name__ == "__main__":
    # Control Tower cells (from rowmap addr): case B33; hld case-inputs B37/C37/D37
    run("A. BASE (hold = 1 month)", {})
    run("B. PPEO ~90 days (hold = 3, base census)", {"B37": 3})
    run("C. PPEO ~120+ days (hold = 4, base census)", {"B37": 4})
    run("D. PPEO ~120+ days + SLOW CENSUS (case 2, hold = 4)", {"B33": 2, "C37": 4})
    print("\nNote: 'additional funding required' is on top of the full modeled stack "
          "($250K equity on deposit, Sept bank note, Jan SBA). Cash actually available "
          "going forward is ~$215K of the $250K ($35K already spent on startup costs), "
          "so real-world need runs modestly higher than modeled where the model's "
          "startup budget understates spend to date.")
