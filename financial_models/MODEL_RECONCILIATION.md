# Model Reconciliation — which workbook is authoritative for what
**Date:** 2026-07-08 (final pass of the financial-modeling QA loop)

Three generations of models now live in this repo. They intentionally differ; this note says
which to use for which purpose, and why their numbers don't match.

## 1. `Azalea_Hospice_Proforma_Rev4.10_DYNAMIC.xlsx`  ← AUTHORITATIVE (real-cash + Path A)
Built by `build_proforma_v3.py`; QA `qa_proforma_v3.py` (20/20). **Rev 4.00 (2026-07-16):
NO revolver / no invented capital.** BASE = Sept bank refi $500K/6%/36 (confirmed Bullard
bank relationship) pays seller $375K; owner deferral (Silas/Dana/Brad accrue to break-even
ADC 18, repaid M7; $42.5K accrued); 1-mo clinical hire lag with day-1 core 1 RN + 1 CNA;
shortfalls SHOWN as UNFUNDED NEED. Headlines:
**Rev 4.10 adds Path A:** SBA 7(a) $500K funds January (51% transfer) and REFINANCES the
bank note (~$461,676 payoff; ~$38K surplus to cash); debt service drops $15,211 → $6,747/mo.
- **BASE: additional capital required $0** — never negative; min cash $3,354 (tight!, M3);
  M36 cash $1,297,392. EBITDA Y1 $299K (15.5%) / Y2 $696K / Y3 $985K.
- **NO-BANK-REFI (SBA covers the Jan balloon): raise $82,328.**
- **DOWNSIDE (slow census): raise $30,254** (first negative M3) — the Jan SBA refi rescues
  the downside. Honest raise target for full coverage: **~$85K**; comfort target ~$100K.
Toggles: refi_on (bank refi), seller_defer (installments roll to balloon), note_amt
(investor notes 10% IO qtrly — default $0, set only when checks clear), sba_on.

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
