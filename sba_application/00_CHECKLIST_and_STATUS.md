# Azalea Hospice — SBA 7(a) Application — Master Tracker

**Borrower:** Tyler Hospice Hold, LLC (dba Azalea Hospice & Palliative Care), a Wyoming LLC
foreign-qualified in Texas | EIN 41-4966640
**Target/operating co:** Refuge Hospice, LLC (San Antonio, TX) | NPI 1174237572 | CCN/PTAN A91679
| being acquired via a phased change of ownership (CHOW)
**Sellers:** Jorge Resendiz and Dennis Hendrix, through DHJR Limited Partnership
**Loan:** $500,000 SBA 7(a) (Rev 4.10) | **Project:** $750,000 | **Equity:** $250,000 (33%)
**Lender contact:** John Hart (SourceFunding)

> Start with **`13_checklist_response_2026-07-16/`** for the current lender submission package,
> **`15_valuation_2026-07-20/`** for the FMV valuation package sent to William at healthfmv.com,
> and **`16_refuge_operational_transition/refuge-operating-plan.md`** for the interim
> operating plan covering the period before the formal CHOW closes.
> Legend: **READY** = drafted/complete · **PRE-FILL** = mostly done, add personal data ·
> **NEEDS YOU** = borrower must provide · **PROVIDED** = source doc on file · **REF** = reference only.

---

## Current structure (authoritative — Rev 4.10)

- **Acquisition:** $500,000 purchase price for 100% of Refuge Hospice, LLC's membership
  interests, structured in two steps to respect the CMS 36-month rule (Refuge's CCN was
  certified 1/8/2024, so the window closes 1/8/2027): **49%** transfers at closing
  (~8/1/2026), the remaining **51%** transfers **1/15/2027**, the first date past that window.
- **Payment structure:** $125,000 down at closing; $31,250/month September-December 2026;
  a final payment of approximately $258,489 due 1/15/2027; 6% interest on the unpaid balance;
  security interest plus personal guaranty; prepayable at any time without penalty.
- **Financing sequence (Path A):** an interim bank note of $500,000 at 6%, 36-month
  amortization (via Jim Bullard's Texas bank relationship) funds in September 2026 and pays
  off the seller balance. The SBA 7(a) loan — $500,000 at ~10.5%, 10-year term — funds and
  refinances that bank note in January 2027, concurrent with the 51% ownership transfer.
- **No revolver / assumed facility.** The financial model carries real-cash discipline: if
  projected cash goes negative it's reported as an explicit "additional funding requirement,"
  never plugged with a fictional credit line.
- **Equity injection:** $250,000 total — 33% of the $750,000 project, well above SBA's 10%
  minimum. James Bullard $100,000 (wired 5/7/2026) + $95,000 (pending); Geoff Schackmann
  $55,000 (pending).
- **Cap table (Tyler Hospice Hold, LLC — unchanged across the target switch from Hickory to
  Refuge):** Geoff Schackmann 39.9% (held directly as an individual, Manager) / James Bullard
  19.5% / Silas Shelton, Dana Davenport, Bradley Woodard 13.3% each / reserved pool 0.7%.
  **Operating-agreement documents in `08_entity_documents/` still reference Hickory Hospice as
  the operating subsidiary and need to be reconciled to Refuge Hospice, LLC before execution —
  see Open Items below.**
- **Financials (Rev 4.10 proforma):** EBITDA $299,195 / $695,981 / $985,429 (Years 1-3); net
  income $128,491 / $457,463 / $718,153; EBITDA margin ~15.5% ramping to 22.3% blended Year 1,
  24.1% Year 2, 28.2% Year 3; debt-service coverage 3.8x / 8.6x / 12.2x (8.2x aggregate); minimum
  projected cash $3,354; month-36 cash $1,297,392.

## License and regulatory snapshot

- CCN/PTAN A91679, NPI 1174237572, PECOS enrollment active.
- CHAP accredited 1/8/2024 - 1/8/2027; CLIA #45D2282700.
- Texas HHSC Medicaid Contract No. HHS000004700408.
- **PTAN reactivation:** CMS approval letter dated 3/20/2026 confirms the billing number was
  reactivated after a deactivation period; reactivation expires **September 2026** absent proof
  of active billing (drives the interim operating plan's July-admission timeline).
- **CMS hospice enrollment moratorium** in effect nationwide since 5/13/2026 (initial 6-month
  term). Blocks new Medicare enrollments and majority-ownership changes that fall inside the
  36-month window; does not appear to block the 1/15/2027 transfer, which lands outside that
  window — recheck with counsel closer to year-end.
- **Interim operating arrangement:** sellers agreed on a 7/21/2026 call to let Geoff operate
  under Refuge's license via a management arrangement ahead of the formal CHOW. Full critical
  path (admit a patient by 7/31 → submit a claim in August → receive payment in September) is
  in `16_refuge_operational_transition/refuge-operating-plan.md`, which also flags the
  de-facto-control-before-CHOW-approval risk as the item requiring counsel sign-off first.

---

## Status at a glance

| # | Item | Status | Where it lives |
|---|------|--------|-----------------|
| 1 | Lender checklist response (PDFs + memos) | **DRAFTED** | `13_checklist_response_2026-07-16/` |
| 2 | Use of Proceeds workbook (John Hart's template format) | **DRAFTED** | `13_checklist_response_2026-07-16/UOP Azalea Refuge Rev2.00.xlsx` |
| 3 | Business Plan (Rev 6.00, current) | **DRAFTED** | `03_business_plan/Azalea SBA Business Plan Rev6.00.docx` |
| 4 | Financial model (Rev 4.10 proforma, 23/23 QA) | **COMPLETE** | `financial_models/output/Azalea_Hospice_Proforma_Rev4.10_DYNAMIC.xlsx` |
| 5 | Package audit harness (content + formatting, 40/40) | **COMPLETE** | `sba_application/qa_package.py` |
| 6 | Insurance application (PL/GL) + reply to Todd Plummer | **DRAFTED** | `14_insurance_2026-07-16/` |
| 7 | Workers' comp underwriting answers (BerkleyNet / Pie) | **SENT** | reply thread, not filed |
| 8 | Refuge FMV valuation package for William (healthfmv.com) | **DRAFTED / DELIVERED** | `15_valuation_2026-07-20/` |
| 9 | Interim operating plan (pre-CHOW critical path) | **DRAFTED / DELIVERED** | `16_refuge_operational_transition/refuge-operating-plan.md` |
| 10 | Investor raise package ($500K raise) | **DRAFTED** | `12_investor_raise/` |
| 11 | Operating agreements (holdco A&R + single-member opco) | **NEEDS RECONCILIATION** — still name Hickory Hospice | `08_entity_documents/` |
| 12 | Personal forms (413 PFS, 912, cash flow, resumes) | **PRE-FILL** — Hickory-era drafts, values still apply to Geoff personally | `05_*`, `06_*`, `07_*` |
| — | Lease (Tyler office, 13387 Hwy 69 N) | **EXECUTED** — reviewed for tenant-entity/term flags | `10_supporting_documents/office_lease/` |

---

## Open items

1. **Reconcile the operating agreements to Refuge Hospice, LLC.** `08_entity_documents/`
   still contains the Amended & Restated holdco OA and a single-member OA drafted for
   "Hickory Hospice LLC" as the operating subsidiary. The cap table and holdco structure carry
   over unchanged (see above), but every reference to Hickory as the acquired entity needs to
   become Refuge Hospice, LLC, and — per the interim operating plan — a management services
   agreement needs drafting for the pre-CHOW period, reviewed by healthcare regulatory counsel.
2. **MIPA still unsigned.** Terms are settled (see Current structure above) but nothing is
   executed. Lock the protections listed in `16_refuge_operational_transition/refuge-operating-plan.md`
   §3 before operating under the sellers' license.
3. **855A / DDE / TULIP sequencing** for the interim operating period — see the operating plan's
   critical-path table; gated on counsel sign-off first.
4. **CPA / S-corp election (Form 2553)** — the Company elects S-corp treatment under IRC §1361;
   confirm filing status with the CPA (this item predates the Hickory-to-Refuge switch and still
   applies unchanged).
5. **Personal items from Geoff** (unchanged by the target switch): SSN/DOB/citizenship/tax
   returns/credit report for the SBA 912/413 forms; PFS dollar values; personal cash-flow values.

---

## Historical note — Hickory Hospice (abandoned target, retained for reference only)

The application originally targeted **Hickory Hospice, LLC**, a different Texas hospice license,
with a $945,000 project ($450K SBA + $300K subordinate bank loan + $195K equity at 19.5%
Bullard-only injection). Hickory failed its site survey and the target changed to **Refuge
Hospice, LLC** as of 7/16/2026. None of the Hickory-specific deal terms, financing structure, or
dollar figures apply to the current application. Files that still carry the Hickory name
(`08_entity_documents/Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx`,
`10_supporting_documents/purchase_agreement_seller_note/` Hickory MIPA and lease drafts) are left
in place as historical record but are not part of the active package — do not send them to the
lender. No reference to an entity named "Avant Hospice" was found anywhere in this repository;
if that name refers to something else, flag it and it can be tracked down and addressed
separately.
