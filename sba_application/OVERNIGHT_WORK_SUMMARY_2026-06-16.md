# Overnight Work Summary — 2026-06-16

Completed the three tasks requested: (1) the operating agreement, scrutinized and finalized;
(2) the SBA financial model; (3) the full SBA loan package for John Hart. Everything builds
clean, ties out to the dollar, and is committed and pushed to `claude/happy-hopper-z2ubs`.

## Final, locked structure
- **Borrower:** Tyler Hospice Hold, LLC — a **Wyoming** LLC (EIN 41-4966640), foreign-qualified in
  Texas — which owns 100% of **Hickory Hospice LLC** (Texas; dba **Azalea Hospice & Palliative Care**;
  keeps its Medicare provider number via the CHOW).
- **Financing (total project $945,000):** SBA 7(a) **$450,000** (15-yr, ~10.5%, startup + working
  capital) + **bank acquisition loan $300,000** (6%, 3-yr, Jim Bullard personally guarantees,
  subordinate to SBA) + equity **$195,000** (Bullard, 19.5%; $180K designated as the SBA injection).
- **Guarantors of the SBA loan:** Geoff Schackmann (PG, controlling Manager); Adeline & Lilah, LLC
  (entity, 39.9% owner); Hickory Hospice LLC (corporate guaranty + lien). Bullard guarantees only the
  bank loan, stays <20% / no SBA PG.
- **Coverage:** combined debt service $169,211/yr (Yrs 1-3); **DSCR 1.79x / 4.02x / 5.95x** (global
  3.92x); min cash ~$445K; LOC never drawn. Base case is solid; ramp-year downside is tighter because
  the bank loan is front-loaded on a 3-yr amortization (flagged for the lender).

## 1) Operating Agreement — drafted, attorney-reviewed, finalized
- **`08_entity_documents/Tyler_Hospice_Hold_LLC_AMENDED_and_RESTATED_OA.docx`** — full A&R OA
  (Wyoming, partnership tax): restated cap table 39.9/19.5/13.3×3/0.7; Article V vesting (entire 13.3%
  at risk; 4.9% time-vested base + 8.4% dual-trigger earn-up; good/bad-leaver call); Bullard 19.5%
  floor; SBA carve-outs; Texas Shootout deferred while SBA debt outstanding; financing/closing
  coordination; S-corporation removed throughout.
- **`08_entity_documents/Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx`** — single-member OA for the
  operating subsidiary (disregarded entity; authorizes the SBA corporate guaranty + lien;
  subordinates the bank loan).
- **`08_entity_documents/OA_RESTATEMENT_REVIEW.md`** — the attorney-perspective review: issues found
  and fixed (superseded incorporation-by-reference restated in full; capital-vs-profits-interest
  characterization corrected; unexecuted-original savings clause; A&L second-member consent), plus
  residual items needing your/counsel/CPA input.
- The interim **Amendment No. 1** is marked **SUPERSEDED**; the A&R OA is the operative document.

## 2) Financial model — rebuilt and QA-clean
- `financial_models/engine/model.py` updated to the new structure; all three workbooks regenerated.
- QA harness: **0 formula errors** (4,380+ cells/workbook), Sources − Uses = 0, global DSCR 3.92x,
  balance-sheet identity holds. Note: the *blended* equity-return tab shows a very high MOIC because it
  divides the whole company's value by the small cash injection (sweat-equity cap table) — that is an
  investor-package artifact, not a lender figure; Bullard's own return tab is ~5-6x.

## 3) SBA loan package — rebuilt to the new structure and cross-checked
All ~30 documents regenerated and verified: Company Profile, Use of Funds + Equity Injection memo,
Business Plan Rev 5.00, Reconciliation/Break-even memo, **Business Debt Schedule (now lists both the
SBA and the bank loan)**, Lender Credit Memo, DSCR Sensitivity/Stress memo, Cover Letter + Q&A,
Form 1919, PFS 413, closing checklist, submission cover/TOC, and the master tracker. Verification:
clean build, **0 stale tokens**, **0 stale entity names**, numbers tie out across model ↔ narratives ↔ OA.
Also fixed a hallucinated seller name ("Greiner" → Gleason/Lozano).

## What still needs YOU / counsel / lender (the real open items)
1. **Bank acquisition loan** — secure the $300K term sheet, the intercreditor/subordination agreement,
   and **SBA-lender approval of the piggyback** (B1). Confirm Bullard's PG of it.
2. **MIPA amendment** — change the Buyer to "Tyler Hospice Hold, LLC, a Wyoming LLC" (it currently
   names a Texas entity) (B2). Confirm the exact charter name against the WY Certificate of Formation.
3. **Execute** the A&R OA + Hickory single-member OA (B3).
4. **Personal items** — Geoff's IDs/tax returns/PFS values; Dana Davenport's street address; CPA
   confirmation of capital-vs-profits-interest treatment for the operator grants.
5. **Lender confirmations** — NAICS 621610; seller PPP/EIDL check; COVID questionnaire.

Full detail and document locations are in `00_CHECKLIST_and_STATUS.md` and
`COUNSEL_QA_REVIEW_2026-06-16.md`.
