# Model Reconciliation — which workbook is authoritative for what
**Date:** 2026-07-08 (final pass of the financial-modeling QA loop)

Three generations of models now live in this repo. They intentionally differ; this note says
which to use for which purpose, and why their numbers don't match.

## 1. `Azalea_Hospice_Proforma_Rev3.10_DYNAMIC.xlsx`  ← AUTHORITATIVE (planning / investor / SBA)
Built by `build_proforma_v3.py`; QA `qa_proforma_v3.py` (24/24). **Rev 3.10 (2026-07-16): BASE CASE
is now the JANUARY TAKEOUT** per the executed MIPA (7/14/2026) — $31,250/mo Sep–Dec (equity-funded),
~$258,489.46 balloon retired by the SBA/bank note at the 51% transfer 1/15/2027, modeled at 6%/36
($7,957/mo from Feb). The Sept $500K/6%/36 refi is the UPSIDE toggle (refi_on=1) — a September SBA
close would require seller guaranties under partial-CHOW rules. Base peak revolver $182,753;
**downside (slow census) now needs a ~$375K facility vs the $250K base commitment — disclosed gap**
(mitigants: census recovery, owner deferral, September takeout).
- **Basis:** granular COA P&L (account codes), post-audit costs (benefits **22%**, workers comp
  **3%**, leadership at Control-Tower salaries: Silas $150K, Dana $120K, Brad $120K, Rhonda $60K),
  CMS rate build (high/low RHC split, ~$172 blended net), monthly ×36.
- **Census:** EOM **24 @M2 / 34 @M6 / 40 @M12** → 50 @M24 → 56 @M36 (user spec 7/8).
- **Debt base:** $500K / 6% / 36-mo note funds Sept ($15,211/mo); seller terms + SBA as toggles.
- **PE conventions:** revolver balancing facility ($250K), Checks tab (master flag), KPI
  Dashboard (TTM DSCR + covenant headroom, DSO, cap cushion), scenario switch, S&U, EBITDA bridge,
  zero hardcodes (scanned).
- **Headlines:** EBITDA Y1 $263K (13.6%) / Y2 $680K (23.5%) / Y3 $976K (28.0%); M12 margin 22.3%;
  base peak revolver $78.7K; **downside case maxes the $250K revolver and stays drawn** (needs
  census recovery / equity / larger facility — disclosed).

## 2. `Azalea_Refuge_Operating_Model.xlsx`  ← weekly cash operations (Jul–Sep window)
Built by `build_operating_model.py`; QA `qa_operating_model.py` (15/15).
- **Basis:** leaner cost set from the Paloma actuals era (benefits 18.27%, $181/day net,
  engine roster salaries), 13-week weekly grain + monthly ×36; semi-monthly payroll timing and
  vendor bill calendar — the only model with intra-month cash precision.
- **Census:** the earlier 7/6 plan (27.4 @M2 / 35 @M6). **One census generation behind Rev 3.00.**
- **Use it for:** week-by-week cash management in the July–September danger zone; update its
  admissions row to actuals as they land. For any *decision* number, prefer Rev 3.00.

## 3. `engine/` workbooks (SBA/Investor/Ops packages)  ← tied to the paused SBA-package track
Built by `engine/build_all`; last structured for SBA $450K + bank $300K + equity $195K (the
John Hart submission set, ~30 documents in `sba_application/`).
- **Stale vs. current reality:** census and deal economics predate Refuge/7-8 census spec.
  **If the SBA track resumes, regenerate this set from the Rev 3.00 assumptions first** —
  otherwise the application contradicts the operating plan.

## Why the EBITDA figures differ (largest deltas)
| Driver | Operating model | Rev 3.00 proforma |
|---|---|---|
| Net rate | $181/PD (Paloma actual) | ~$172/PD (CMS build, 15% high-rate) |
| Benefits / WC | 18.27% / in-load | 22% / +3% |
| Leadership comp | $360K/yr (3 owners) | $487.5K/yr (4 + vol/bereavement) |
| Month-1 revenue | weekly front-load (ADC ~19.8) | monthly grain (ADC ~10–13) |
| Contingency | none | 5% of NPR (below EBITDA, cash-deducted) |

Rev 3.00 is deliberately the more conservative, lender-defensible basis.

## Open items carried out of the loop
1. Confirm CMS FY2026 Smith County rates and the 15% high-rate share against the actual rate
   letter (inputs on Control Tower).
2. Commit the $250K revolver (or equivalent) before close — the downside case depends on it.
3. Entity naming: uploaded docs referenced "Tyler Hospice OpCo, LLC"; the session-locked entity is
   Tyler Hospice Hold, LLC (WY) → Refuge sub. Align before anything goes external.
4. If the John Hart SBA track resumes: regenerate the engine workbooks + narratives from Rev 3.00.
