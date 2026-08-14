# Azalea Hospice — SBA 7(a) Application — Master Tracker

**Borrower:** Tyler Hospice Hold, LLC (dba Azalea Hospice & Palliative Care), a Wyoming LLC
foreign-qualified in Texas | EIN 41-4966640
**Target/operating co:** Avant Hospice, LLC (TX) | CMS CCN 741798 | Medicare-certified,
never billed | acquired via a single change of ownership (CHOW) — 100% at closing
**Sellers:** Kimberly Carlisle (with David Groom on the seller side)
**Loan:** $500,000 SBA 7(a) (Rev 5.00 AVANT) | **Project:** $750,000 | **Equity:** $250,000 (33%)
**Lender contact:** John Hart (SourceFunding)

> Start with **`13_checklist_response_2026-07-16/`** for the current lender submission package
> (all active documents there are Avant-based and pass `qa_package.py` 41/41).
> Legend: **READY** = drafted/complete · **PRE-FILL** = mostly done, add personal data ·
> **NEEDS YOU** = borrower must provide · **PROVIDED** = source doc on file · **REF** = reference only.

---

## Current structure (authoritative — Rev 5.00 AVANT)

- **Acquisition:** $300,000 purchase price for 100% of Avant Hospice, LLC's membership
  interests, transferring in a single step at CHOW approval. No split closing, no 36-month
  staging: the license is fully in the borrower's name from day one, which satisfies the
  lender rule that SBA cannot fund until the license is under the borrower's name.
- **Payment structure:** nothing paid at closing. The seller carries the full $300,000 at 6%
  interest, **fully deferred** — no payment due until Medicare claims are actually paying
  (i.e., until the prepayment-review gate clears). Prepayable without penalty.
- **Financing sequence:** the SBA 7(a) loan — $500,000 at ~10.5%, 10-year term — funds at
  **month 2**, after the license is 100% in the borrower's name, and retires the seller note
  in full (~$303,000 including accrued interest). The remaining ~$197,000 of proceeds goes to
  ramp working capital. There is **no interim/bridge bank financing anywhere in the structure**.
  Single debt thereafter: ~$6,747/month.
- **Prepayment review (PPEO), modeled head-on:** Avant has never billed, so its first claims
  enter enhanced-oversight prepayment review (42 CFR 424.527). The model carries a 2-month
  base-case payment hold (4 months in the downside). Because the seller note is fully deferred
  until that gate clears, no seller payment is due while Medicare cash is held.
- **No revolver / assumed facility.** The financial model carries real-cash discipline: if
  projected cash goes negative it's reported as an explicit "additional funding requirement,"
  never plugged with a fictional credit line. (Base case never goes negative; the no-SBA
  fallback scenario's peak requirement is only $7,608.)
- **Equity injection:** $250,000 total — 33% of the $750,000 project, well above SBA's 10%
  minimum. James Bullard $100,000 (wired 5/7/2026) + $95,000 (pending); Geoff Schackmann
  $55,000 (pending).
- **Cap table (Tyler Hospice Hold, LLC — unchanged across target switches):** Geoff Schackmann
  39.9% (held directly as an individual, Manager) / James Bullard 19.5% / Silas Shelton, Dana
  Davenport, Bradley Woodard 13.3% each / reserved pool 0.7%. **Operating-agreement documents
  in `08_entity_documents/` still reference Hickory Hospice as the operating subsidiary and
  need to be reconciled to Avant Hospice, LLC before execution — see Open Items below.**
- **Financials (Rev 5.00 AVANT proforma, QA 22/22):** EBITDA $299,195 / $695,981 / $985,429
  (Years 1-3); net income $130,630 / $472,185 / $733,029; EBITDA margin 15.5% / 24.1% / 28.2%
  (month-12 margin 22.3%); debt-service coverage 3.7x / 8.6x / 12.2x; minimum projected cash
  **$192,358**; month-36 cash **$1,473,592**; downside (slow census) peak additional capital
  $99,356 with month-36 cash still $444,066.

## License and regulatory snapshot

- Avant Hospice, LLC: CMS CCN 741798, Medicare-certified, **never billed** (clean claims
  history; no ADR/audit tail). Medicare-only — no Medicaid contract, so no nursing-facility
  room-and-board pass-through; the SNF referral channel is built on facility relationships
  instead.
- **Palmetto GBA tie-in / CHOW processing** is the key regulatory timing risk: the sellers'
  tie-in has a long unresolved history (escalated through two legislators and HHSC as of
  June-July 2026). The financing sequence tolerates this — the seller deferral means nothing
  in the deal breaks while CMS processes.
- **CMS hospice enrollment moratorium** in effect nationwide since 5/13/2026. The sellers read
  the CHOW as unaffected; **that favorable read is the sellers', not counsel's — get Avant's
  36-month/moratorium position confirmed in writing by healthcare regulatory counsel** (open
  item below).
- **Prepayment review:** first claims from a never-billed number enter enhanced oversight
  (42 CFR 424.527), typically 60-120+ day payment delays. Modeled at 2 months base / 4 downside.

---

## Status at a glance

| # | Item | Status | Where it lives |
|---|------|--------|-----------------|
| 1 | Lender checklist response (PDFs + memos, Avant) | **DRAFTED** | `13_checklist_response_2026-07-16/` |
| 2 | Use of Proceeds workbook (John Hart's template format) | **DRAFTED** | `13_checklist_response_2026-07-16/UOP Azalea Avant Rev3.00.xlsx` |
| 3 | Business Plan (Rev 7.00 AVANT, current) | **DRAFTED** | `03_business_plan/Azalea SBA Business Plan Rev7.00 AVANT.docx` |
| 4 | Financial model (Rev 5.00 AVANT proforma, 22/22 QA) | **COMPLETE** | `financial_models/output/Azalea_Hospice_Proforma_Rev5.00_AVANT.xlsx` |
| 5 | Package audit harness (content + formatting, 41/41) | **COMPLETE** | `sba_application/qa_package.py` |
| 6 | Insurance application (PL/GL) + reply to Todd Plummer | **DRAFTED** (Refuge-era, sent) | `14_insurance_2026-07-16/` |
| 7 | Workers' comp underwriting answers (BerkleyNet / Pie) | **SENT** | reply thread, not filed |
| 8 | Refuge FMV valuation package for William (healthfmv.com) | **HISTORICAL** (Refuge track) | `15_valuation_2026-07-20/` |
| 9 | Interim operating plan (pre-CHOW critical path) | **HISTORICAL** (Refuge track) | `16_refuge_operational_transition/` |
| 10 | Investor raise package ($500K raise) | **DRAFTED** | `12_investor_raise/` |
| 11 | Operating agreements (holdco A&R + single-member opco) | **NEEDS RECONCILIATION** — still name Hickory Hospice | `08_entity_documents/` |
| 12 | Personal forms (413 PFS, 912, cash flow, resumes) | **PRE-FILL** — values still apply to Geoff personally | `05_*`, `06_*`, `07_*` |
| — | Lease (Tyler office, 13387 Hwy 69 N) | **EXECUTED** — reviewed for tenant-entity/term flags | `10_supporting_documents/office_lease/` |

---

## Open items

1. **Avant purchase agreement.** The uploaded agreement is unsigned; the sellers have agreed
   in principle to full deferral until PPEO clears — get that term (and the 6% rate, month-2
   prepayable takeout) into the executed document.
2. **Counsel confirmation of Avant's regulatory position:** 36-month rule and moratorium
   treatment of the CHOW, in writing, plus the Palmetto tie-in status. The sellers' favorable
   moratorium read is not counsel's.
3. **Reconcile the operating agreements to Avant Hospice, LLC.** `08_entity_documents/` still
   contains the Amended & Restated holdco OA and a single-member OA drafted for "Hickory
   Hospice LLC" as the operating subsidiary. The cap table and holdco structure carry over
   unchanged; every reference to the acquired entity needs to become Avant Hospice, LLC.
4. **Jim Bullard note → equity conversion.** SBA SOP 50 10 8 counts only true equity toward
   the injection; the convertible note must convert (or be re-papered) before the injection is
   documented to the lender.
5. **CPA / S-corp election (Form 2553)** — confirm filing status with the CPA (predates the
   target switches and still applies unchanged).
6. **Personal items from Geoff** (unchanged by the target switch): joint PFS (SBA 413),
   personal cash flow, passports/citizenship proof, source-of-injection statements, 2025
   return or extension copy — the outstanding half of John Hart's 11-item checklist.

---

## Historical note — prior targets (retained for reference only)

The application originally targeted **Hickory Hospice, LLC** ($300,000, Medicare-only), which
failed its state site survey; the target then moved to **Refuge Hospice, LLC** (San Antonio;
$500,000, dual-certified, phased 49/51 closing around the CMS 36-month window) as of
7/16/2026. On 8/14/2026 the target moved to **Avant Hospice, LLC** ($300,000, Medicare-only,
single 100% CHOW closing, fully deferred seller note) on capital-efficiency grounds: the
Refuge structure's interim bank note plus phased purchase payments could not absorb a
prepayment-review scenario (the PPEO stress test showed a $155K-$315K gap), while the Avant
structure's base case never goes cash-negative (minimum cash $192,358). Refuge-era documents
(`15_valuation_2026-07-20/`, `16_refuge_operational_transition/`, `17_mipa_negotiation_2026-07-22/`,
and the sent drafts 8/9/11 inside `13_checklist_response_2026-07-16/`) and Hickory-era files
(`08_entity_documents/Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx`,
`10_supporting_documents/purchase_agreement_seller_note/`) are left in place as historical
record but are not part of the active package — do not send them to the lender.
