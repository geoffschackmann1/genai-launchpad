# Amended & Restated OA — Attorney-Perspective Review & Build Notes
**Date:** 2026-06-16

Two operating agreements were drafted and self-reviewed as counsel would, then finalized:

1. **`Tyler_Hospice_Hold_LLC_AMENDED_and_RESTATED_OA.docx`** — full A&R OA for the holding
   company / SBA borrower (Wyoming), built by `build_oa_restatement.py`.
2. **`Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx`** — single-member OA for the operating
   subsidiary (Texas, dba Azalea), built by `build_hickory_oa.py`, adopted at the CHOW closing.

## What the A&R OA integrates (vs. the provided v7)
- Entity conformed to **Tyler Hospice Hold, LLC** (Wyoming, EIN 41-4966640), foreign-qualified in TX.
- **Partnership** tax treatment; the S-corporation election and every S-corp construct removed
  (definition 2.23, §9.1(e), §11.2, Article XII certification, §§8.5/8.11 eligibility conditions,
  §4.5 single-class language).
- Restated **cap table** 39.9 / 19.5 / 13.3×3 / 0.7 (Exhibit A).
- **Article V** rewritten: full 13.3% at risk; 4.9% time-vested Initial Base + 8.4% dual-trigger
  Earn-Up; vesting from the Effective Date; good/bad-leaver call at FMV / lower-of-cost-or-FMV;
  SBA-subordinated payment; §83(b) requirement retained.
- Bullard floor **19.5%**; SBA/Bank-loan **carve-out** from Reserved Matters; **Texas Shootout
  deferred** while the SBA Loan is outstanding; unvested interests **vote**.
- New **§1.7 / §3.16** financing & closing coordination; **§6.2(d)** debt carve-out; Hickory
  corporate-guaranty mechanics referenced; MIPA-amendment acknowledgment.

## Issues found in self-review and FIXED
1. **Superseded incorporation-by-reference.** The first draft pulled §§3.7–3.13 "by reference" from
   the Original Agreement, which the restatement supersedes in full — internally inconsistent.
   → Fully **restated** §§3.6–3.13 (additional capital, delinquency, contribution loans, dilution,
   appraised value, no interest on capital, withdrawal limits).
2. **Tax characterization.** The Restricted Interests carry pari-passu economic rights (including
   liquidation, §4.5), which makes them **capital interests**, not profits interests — the original
   "profits-interest / Rev. Proc. 93-27" claim was inconsistent. → §5.10 now states they are
   expected to be **capital interests** with nominal FMV fixed by the §83(b) election, and
   authorizes the Manager to implement a **liquidation hurdle** if the CPA instead elects profits-
   interest treatment.
3. **Unexecuted-original savings clause.** The v7 was an "EXECUTION COPY — PENDING FORMATION."
   → Recital A now states that if the Original Agreement was never executed, this Agreement is the
   **initial** Operating Agreement.
4. **A&L second-member consent.** Adeline & Lilah is 50/50 Geoff / Mary Elizabeth Burcham.
   → Added a **consent of the second member** (and spousal acknowledgment) signature block, rather
   than a duplicate "authorized representative."
5. Removed vestigial S-corp "single class" phrasing in §5.1.

## Residual items requiring human / counsel / CPA input before execution
- **Exact charter name** — confirm against the Wyoming Certificate of Formation (this draft uses
  "Tyler Hospice Hold, LLC"; the v7 title said "Tyler Hospice HoldCo L.L.C."). The MIPA must be
  amended to the confirmed name (its buyer is currently a Texas LLC).
- **Dana L. Davenport's physical street address** (required for the §11.11 notice block).
- **Capital vs. profits interest** — CPA to confirm treatment and, if profits interests are desired,
  set the liquidation hurdle value per Equity Grantee.
- **Mary Elizabeth Burcham's address** for the consent block.
- **SBA/Bank loan documents** — the corporate guaranty, security agreements, and the
  intercreditor/subordination agreement are referenced but are separate lender-form documents to be
  papered at closing.
- Counsel to confirm Wyoming governing law + Texas (Smith County) arbitration venue remains intended
  for a TX-operating subsidiary structure.
