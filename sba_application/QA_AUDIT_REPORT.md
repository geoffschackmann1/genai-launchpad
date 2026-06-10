# QA Audit — Azalea SBA Package (2026-06-10)

Comprehensive adversarial QA of the three Excel workbooks and all 30 generated
documents against the source-of-truth structure (SBA $555K / equity $195K /
Bullard 19.5% / 26% injection / license paid at close, no seller note).

## Source-of-truth numbers (all verified to tie out)
- Project $750,000 = SBA $555,000 (74.0%) + equity $195,000 (26.0%)
- Uses: seller $300,000 + startup $63,000 + equipment $15,000 + WC reserve $372,000 = $750,000
- SBA debt service $93,637/yr ($7,803.05/mo); SBA interest $62,202 / $58,741 / $54,551
- EBITDA $302,522 / $679,524 / $1,007,520; Net income $199,308 / $574,951 / $903,362
- DSCR 3.23x / 7.26x / 10.76x; global 7.08x; min cash $262,558
- Cap table A&L 39.9 / Bullard 19.5 / Silas-Dana-Brad 13.3x3 / reserved 0.7 = 100.0%

## Model checks — ALL PASS
- 0 formula errors in all three workbooks (4,300+ cells each).
- Balance sheet identity (Assets − (L+E)) = 0.000000 across all 21 periods, all three workbooks.
- Sources − Uses = 0.
- DSCR, net income, EBITDA, opening cash, min cash recompute to the figures above.
- Investor Position tab: Bullard's 19.5% pro-rata return — IRR 83.6%, MOIC 5.82x, 41% cash-on-cash.
- Engine tabs byte-for-byte identical across all three workbooks.

## Bugs found and FIXED (documents)
All were stale figures left from the $500K/$250K → $555K/$195K re-cuts. None were in the
models (which are formula-driven); all were hand-authored narrative/table values.

1. **Business Plan P&L — interest row did not reconcile.** Showed old $500K interest
   ($56,038 / $52,604 / $48,754); the net-income row in the SAME table implied $62,202 /
   $58,741 / $54,551. Fixed to the model's interest; P&L now reconciles to the penny.
   (Also fixed the identical row in the Reconciliation memo.)

2. **DSCR Sensitivity memo — every stress ratio computed on the OLD $93,637→$84,357 debt
   service** while the header correctly showed $93,637. All 8 scenarios overstated. Recomputed:
   S1 2.07→1.87x, S2 0.56→0.50x, S3 3.13→2.82x, S4 2.83→2.55x, S5 2.56→2.30x, S6 3.07→2.77x,
   C1 2.22→2.00x, **C2 1.10→0.99x** (moderate triple-downside now shown below the floor /
   ~breakeven, was previously near-floor). Narrative updated.

3. **Business Plan Year-2 stress table did not match the model's Stress Tests tab.** Hand-authored
   EBITDA/DSCR drifted from the engine. Aligned to the model exactly: Census -20% 2.10→1.99x
   ($186,766), Wage/cost +10% 5.45→5.35x ($501,098), Combined 0.30→**0.09x** ($8,340). Narrative
   and risk-table references updated.

4. **Reconciliation memo comparison row** showed old current-model net income ($205K/$581K/$909K),
   contradicting the detailed P&L below it. Fixed to $199K/$575K/$903K.

5. **Stale singletons:** Business Plan tile "SBA $500K + Equity $250K" → $555K/$195K; "combined
   DSCR" → "DSCR"; "33% equity injection" → "26%"; Interview worksheet "$250K" → "$195K" and the
   obsolete "Seller note — confirm $300,000, 36 months" → "$300,000 paid in full at close (no note)";
   Submission Readiness "$250K on deposit" → "$195K"; three residual 3-tranche / "June and July 2026"
   references → consolidated 2-tranche ($100K received + $95K by 7/31/2026).

## Final state
- Stale-value sweep across all 30 generated docs: **0 hits.**
- Every DSCR figure in every doc reconciles to base-case or a correctly-computed stress.
- Arithmetic identities (sources=uses, ratios, cap-table=100%, interest reconciliation, monthly×12
  = annual debt service): **all pass.**
- The two AS_PROVIDED OA uploads are intentionally left at their original (pre-amendment) figures.
