# SBA Loan Package model - re-cut to seller-paid-at-close (startup, $750K)

`Azalea_Hospice_SBA_Loan_Package.xlsx` was re-cut from the prior dual-debt structure
($1.05M project with a $300K retained seller note) to the structure confirmed with John Hart:
a startup CHOW with the seller paid in full at close and no retained note.

## What changed (four input-driven edits; all formulas left live)
1. **Inputs!B139** (opening cash): now subtracts the license cost - `= equity + loan - startup - capex - license`.
   Opening cash $672,000 -> **$372,000** (the $300K license is now paid in cash at close).
2. **Sources & Uses!B6** (seller-note source): set to **0** - the note is no longer a funding source.
3. **Debt Schedule!C18** (seller-note beginning balance): set to **0** - zeroes the entire note
   schedule, so combined debt service becomes SBA-only.
4. **Cash Flow & BS!B34** (seller-note liability at open): set to **0** - no note on the balance sheet.

Everything else (revenue, staffing, EBITDA, the license intangible asset and its 15-yr amortization,
the SBA loan) is unchanged. Descriptive header text on Sources & Uses, Lender Summary, and Debt
Schedule was updated to say "seller paid at close."

## Verified recalculated outputs (computed independently with the `formulas` engine)
| Metric | Value |
|---|---|
| Total sources / uses | $750,000 / $750,000 (check = 0) |
| Opening cash | $372,000 |
| SBA debt service (only debt) | $84,357 / yr |
| DSCR | Y1 **3.59x** / Y2 **8.06x** / Y3 **11.94x** |
| Global 3-yr DSCR | **7.86x** |
| EBITDA | $302,522 / $679,524 / $1,007,520 |
| Net income | $205,472 / $580,772 / $908,768 |
| **Minimum cash (any period)** | **$264,104** - well above the $25,000 floor |
| Balance-sheet check (open and M12) | 0 (balances) |

The generated package documents (Business Plan Rev 5.00, credit memo, etc.) match these to within
rounding (<0.1% on net income, from minor working-capital-line interest).

## IMPORTANT - two caveats
1. **Open in Excel once before sending.** This environment could not run a spreadsheet engine to
   refresh the file's *cached* cell values, so the workbook is set to **recalculate automatically on
   open** (fullCalcOnLoad). Excel and Google Sheets will recompute and show the values above the
   moment it opens. To be safe, open it once and save before forwarding to John, so the saved values
   are baked in.
2. **The other two workbooks still show the OLD structure.** `Azalea_Hospice_Investor_Package.xlsx`
   and `Azalea_Hospice_Operations_Dashboard.xlsx` were NOT re-cut - they still reflect the $1.05M /
   $300K-seller-note model. If either goes to John or an investor, it needs the same four edits.
