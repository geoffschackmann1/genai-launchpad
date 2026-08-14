# Azalea Hospice — Three Purpose-Built Financial Workbooks

Three audience-specific Excel workbooks generated from **one shared, formula-driven
engine**, all reconciled to Paloma Tyler's real 2025 actuals.

| Workbook | File | Audience |
|---|---|---|
| **A — SBA Loan Package** | `output/Azalea_SBA_Loan_Package.xlsx` | SBA lender / underwriter |
| **B — Investor Model** | `output/Azalea_Investor_Model.xlsx` | Equity investor / partner |
| **C — Operations Dashboard** | `output/Azalea_Operations_Dashboard.xlsx` | Owner-operator / clinical leadership |

## Deal framing
Azalea is a **CHOW relaunch** of **Hickory Hospice** — Azalea acquires Hickory's
existing Medicare-certified provider number for **$300K, seller-financed at 6% over
36 months**. This lets Azalea bill from day 1 (no fresh 855A enrollment delay) in
**San Antonio** (Hickory's existing market) and operate an **alternative-delivery
site for Tyler / East Texas**. Paloma's clinicians join Azalea and Paloma's
patients migrate onto Azalea's provider number over months 1–2. Revenue ramps with
that migration (M1 ADC 12 → M2 19.8 → M3 22 steady), and labor ramps via a
staggered-hire schedule. Paloma's proven ~$118K/mo, ~22 ADC book is the benchmark
the model reconciles to.

CHAP/ACHC accreditation transfers with the CHOW. Capital stack:
- **$500K SBA 7(a)** @ 11.5%, 10-yr amortization
- **$300K Hickory seller note** @ 6%, 36-month amortization (retires at end of Year 3)
- **$250K owner equity**

## How to regenerate
```bash
pip install openpyxl
python financial_models/build_all.py        # builds all three into output/
```

## QA / verification
The model is validated by a pure-Python reference calculation (`engine/model.py :: compute`)
plus a formula evaluator (`pip install formulas`):
```bash
python financial_models/qa.py        # recalc engine vs Python truth (every line Δ≈0; BS balances to 0)
python financial_models/qa_all.py    # scan all 3 workbooks for formula errors + spot-check audience tabs
```
Latest run: **0 formula errors** across ~2,600 evaluated cells per workbook; balance
sheet check row = 0.000000 every period; model net rate = **$181.30/PD** vs actual
$181.00; steady-state net revenue reconciles to Paloma's **$118K** Apr–Jun average.

## Architecture
```
financial_models/
  source_data/        # the 5 uploaded source files (P&L, 3 payrolls, prior engine)
  engine/
    styles.py         # color palette, number formats, cell-note helper
    model.py          # assumptions (single source of truth) + Python reference calc + period structure
    build.py          # Book helper + shared engine-tab writers (Inputs, Revenue, Staffing,
                      #   Operating Budget, Debt, Cash Flow & BS, Actuals vs Model)
    workbooks.py      # audience-specific tabs + the three assemble_* functions
  build_all.py        # orchestrator -> output/*.xlsx
  qa.py, qa_all.py    # verification harnesses
  output/             # the three deliverables
```

### Shared engine tabs (identical in all three)
`Inputs` → `Revenue Model` → `Staffing & Payroll` → `Operating Budget` (P&L) →
`Debt Schedule` → `Cash Flow & BS` (3-statement) → `Actuals vs Model`.

Time horizon: **Year 1 monthly (M1–M12) + Years 2–3 quarterly (8 quarters)** = 20 periods,
each in the same spreadsheet column across every tab so cross-sheet links line up.

### Build rules honored
- **No hardcoded business numbers in formulas** — every rate/%/$ lives in a labeled
  `Inputs` cell and is referenced (named ranges created for each).
- **Color code:** blue = input · black = formula · green = cross-sheet link.
- **Cell notes** cite `Source: Paloma P&L 2025` / `Source: Paloma payrolls Apr–May 2025`
  on every actuals-derived assumption (100+ per workbook).
- Number formats: parentheses for negatives, `-` for zeros, `$#,##0`, `0.0%`, `0.0x`.
- **Balance sheet balances through real AR/AP/LOC, never a plug** (check row = 0).
- Every workbook reconciles to the historical actuals (`Actuals vs Model` tab).

## Scenarios (drive the whole model from `Inputs`)
- **Capture rate** (default 100% → steady-state ADC 22). Set 80% → ADC ~18, revenue −20%.
- **Census scenario** (1 = Base flat / conservative; 2 = Upside referral growth → ADC ~60 by Y3).

## Flagged assumptions to confirm (blue `CONFIRM` cells on `Inputs`)
1. **Capital:** $500K SBA @ 11.5% / 10-yr + $300K Hickory seller note @ 6% / 36-mo +
   $250K owner equity. Combined Y1/Y2/Y3 DSCR = **1.68× / 1.56× / 1.45×** (seller
   note retires end of Y3 → coverage jumps in Y4+).
2. **Startup one-time costs:** CHOW filing, TX licensure, EMR setup, supply stock,
   legal/formation, contingency. (Accreditation removed — transfers with CHOW.)
3. **Owner comp:** Brad/Silas salaries = full comp (no separate draw).
4. **Rent & marketing:** $3,000/mo each (Paloma actuals were $0 — embedded/rent-free).

Other carried-forward notes: **Medical Director** modeled at $4,000/mo (Paloma
actual ran $1,000/mo — flagged). **855A cash gap is eliminated** by the CHOW —
Hickory's existing provider number transfers; Azalea bills from day 1. The
**Medicare cap** monitor uses inherited mid-episode day-counts.

## Hickory license amortization
- **License intangible** ($300K) amortized **15 years** straight-line → ~$20K/yr
  non-cash, sits below EBITDA (no DSCR impact).
- **Seller note** principal + interest are real cash outflows; included in combined
  debt service and DSCR.
