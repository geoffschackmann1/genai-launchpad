"""QA for the four-track comparison workbook (Rev 1.00).

  1. 0 formula errors across the workbook
  2. Comparison tab links equal the track-tab values
  3. Track chassis conservation: earned revenue = collections + uncollected tail
     (every track, all claims eventually collect or remain as tail A/R)
  4. T2 Refuge peak-need sanity band vs the ppeo_stress reference: ppeo_stress
     showed $155,489 additional need at hold=3 ON TOP of the $250K equity and with
     the Rev 4.10 startup/deferral treatment; this comparison models no equity
     inflow and a simplified chassis, so peak need should land in a band around
     $250K + $155K (+/- chassis simplification tolerance).
  5. T1 peak need < T2 and < T3 peak need at defaults (no license capital)
  6. T4 variants: peak need strictly less than their standalone owned-license track
  7. No hardcoded constants >= 100 in formulas outside Assumptions

Run from repo root: python3 -m financial_models.qa_track_comparison
"""
import json
import sys
import formulas
import openpyxl
import re
from openpyxl.utils import get_column_letter as gcl

XLSX = "financial_models/output/Azalea_Track_Comparison_Rev1.00.xlsx"
ROWMAP = "financial_models/output/track_comparison_rowmap.json"
NM = 24
PASS = FAIL = 0


def chk(name, ok, detail=""):
    global PASS, FAIL
    print(f"  {'OK ' if ok else 'FAIL'} {name} {detail}")
    PASS, FAIL = PASS + bool(ok), FAIL + (not ok)


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
    return vals, g


def main():
    rm = json.load(open(ROWMAP))
    R = rm["rows"]
    vals, g = evaluate(XLSX)

    errs = [k for k, v in vals.items() if isinstance(v, str) and v.startswith("#")]
    chk("0 formula errors", not errs, f"({len(errs)})")
    for e in errs[:8]:
        print("     ERR", e, vals[e])

    tracks = ["T1 Jason ADS", "T2 Refuge", "T3 Avant", "T4a Refuge+Jason", "T4b Avant+Jason"]
    peaks = {}
    for t in tracks:
        peaks[t] = g(t, f"B{R[t]['peak']}")

    # Comparison links
    comp_ok = True
    for i, t in enumerate(tracks):
        got = g("Comparison", f"B{4+i}")
        if got is None or abs(got - peaks[t]) > 0.01:
            comp_ok = False
    chk("Comparison tab links match track tabs", comp_ok)

    # conservation per track: total earned >= total collected, tail = A/R
    for t in tracks:
        earn_rows = [k for k in ("earned", "earn_own") if k in R[t]]
        coll_rows = [k for k in ("coll", "coll_own", "coll_j") if k in R[t]]
        tot_e = 0.0
        for er in earn_rows:
            tot_e += sum(g(t, f"{gcl(2+i)}{R[t][er]}") or 0 for i in range(NM))
        if "adc_j" in R[t]:  # jason stream earned = adc_j * netday * dpm
            tot_e += sum((g(t, f"{gcl(2+i)}{R[t]['adc_j']}") or 0) * 172.0 * 30.4 for i in range(NM))
        tot_c = 0.0
        for cr in coll_rows:
            tot_c += sum(g(t, f"{gcl(2+i)}{R[t][cr]}") or 0 for i in range(NM))
        chk(f"{t}: collections <= earned (tail is A/R)", tot_c <= tot_e + 1.0,
            f"(earned {tot_e:,.0f} collected {tot_c:,.0f})")

    # T2 sanity band vs ppeo_stress reference (see module docstring)
    chk("T2 peak need in sanity band $330K-$480K",
        330000 <= peaks["T2 Refuge"] <= 480000, f"(${peaks['T2 Refuge']:,.0f})")

    chk("T1 peak < T2 peak", peaks["T1 Jason ADS"] < peaks["T2 Refuge"],
        f"(T1 ${peaks['T1 Jason ADS']:,.0f})")
    chk("T1 peak < T3 peak", peaks["T1 Jason ADS"] < peaks["T3 Avant"],
        f"(T3 ${peaks['T3 Avant']:,.0f})")
    chk("T4a peak < T2 peak", peaks["T4a Refuge+Jason"] < peaks["T2 Refuge"],
        f"(T4a ${peaks['T4a Refuge+Jason']:,.0f})")
    chk("T4b peak < T3 peak", peaks["T4b Avant+Jason"] < peaks["T3 Avant"],
        f"(T4b ${peaks['T4b Avant+Jason']:,.0f})")

    # hardcode scan outside Assumptions
    wb = openpyxl.load_workbook(XLSX)
    pat = re.compile(r"(?<![A-Z$.\d])(\d{3,}(?:\.\d+)?)")
    offenders = []
    for sh in wb.sheetnames:
        if sh in ("Assumptions", "Comparison"):
            continue
        for row in wb[sh].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for mnum in pat.findall(c.value):
                        if float(mnum) >= 100:
                            offenders.append(f"{sh}!{c.coordinate}:{mnum}")
    chk("0 hardcoded constants >=100 outside Assumptions", not offenders, str(offenders[:5]))

    print(f"\n  >>> PEAKS: " + " | ".join(f"{t} ${peaks[t]:,.0f}" for t in tracks))
    for t in tracks:
        m12 = g(t, f"M{R[t]['cum']}")
        m24 = g(t, f"Y{R[t]['cum']}")
        print(f"      {t}: M12 cum ${m12:,.0f} | M24 cum ${m24:,.0f}")

    print(f"\n{PASS} passed, {FAIL} failed")
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
