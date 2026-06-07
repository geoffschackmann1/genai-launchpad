# SBA Loan Package model - re-cut to $555K SBA / $195K equity (Bullard 19.5%)

`Azalea_Hospice_SBA_Loan_Package.xlsx` reflects the current structure: a startup CHOW with the
seller paid in full at close (no retained note), funded by a $555,000 SBA loan and a $195,000 cash
equity injection from James Bullard (19.5% passive minority interest).

## Input cells that define the structure (Inputs sheet)
| Cell | Meaning | Value |
|---|---|---|
| B110 | SBA 7(a) loan principal | **$555,000** |
| B113 | Equity injection (Bullard) | **$195,000** |
| B120 | Hickory license / acquisition cost | $300,000 (paid at close; intangible amortized 15 yr) |
| B139 | Opening cash = equity + loan - startup - capex - license | **$372,000** |
| Debt Schedule C18 | Seller-note beginning balance | **0** (no note) |
| Sources & Uses B6 | Seller-note source | **0** |
| Cash Flow & BS B34 | Seller-note liability at open | **0** |

## Verified recalculated outputs (computed independently with the `formulas` engine)
| Metric | Value |
|---|---|
| Total sources / uses | $750,000 / $750,000 (check = 0) |
| SBA loan / equity | $555,000 (74%) / $195,000 (26%) |
| Opening cash | $372,000 |
| SBA debt service (only debt) | $93,637 / yr ($7,803.05/mo) |
| DSCR | Y1 **3.23x** / Y2 **7.26x** / Y3 **10.76x** |
| Global 3-yr DSCR | **7.08x** |
| EBITDA | $302,522 / $679,524 / $1,007,520 |
| Net income | $199,308 / $574,951 / $903,362 |
| **Minimum cash (any period)** | **$262,558** - well above the $25,000 floor |
| Balance-sheet check (open and M12) | 0 (balances) |

The package documents (Business Plan Rev 5.00, credit memo, equity memo, etc.) match these figures.

## Notes
- **Equity injection is 26% of project** - above the 10% SOP 50 10 8 minimum, but **below the ~30%
  some startup lenders (including the LOC lender John mentioned) prefer.** Still a strong, low-leverage
  deal given DSCR 3.23x+; John can place it with lenders comfortable at this level.
- **Open in Excel once before sending.** The workbook recalculates automatically on open
  (fullCalcOnLoad) - this environment can't run a spreadsheet engine to bake cached values. Open and
  save once before forwarding.
- **The other two workbooks** (Investor Package, Operations Dashboard) still reflect the old
  $1.05M / dual-debt structure and have NOT been re-cut.
