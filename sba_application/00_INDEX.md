# Azalea Hospice — SBA 7(a) Application — File Index

**Borrower:** Tyler Hospice Hold, LLC (dba Azalea Hospice & Palliative Care) · EIN 41-4966640
**Target/operating co:** Avant Hospice, LLC (TX) · CMS CCN 741798 · Medicare-certified, never
billed · acquired via a single change of ownership (CHOW) — 100% at closing
**Loan:** $500,000 SBA 7(a) (Rev 5.00 AVANT) · **Project:** $750,000 · **Equity:** $250,000 (33%)
**Lender contact:** John Hart (SourceFunding)

> Start with **`00_CHECKLIST_and_STATUS.md`** (master tracker) for current status and open items.
> The active, current-deal folder is **`13_checklist_response_2026-07-16/`** (lender submission
> package — all active documents there are Avant-based). `14_insurance_2026-07-16/` (insurance),
> `15_valuation_2026-07-20/` (Refuge FMV package), `16_refuge_operational_transition/`, and
> `17_mipa_negotiation_2026-07-22/` are Refuge-era: sent or superseded, retained as record.
> Legend: **READY** = drafted/complete · **PRE-FILL** = mostly done, add personal data ·
> **NEEDS YOU** = borrower must provide · **PROVIDED** = source doc on file · **REF** = reference only ·
> **HISTORICAL** = superseded (Hickory- or Refuge-era) material, not part of the active package.

---

## Root
| File | Purpose | Status |
|---|---|---|
| `00_CHECKLIST_and_STATUS.md` | Master tracker — current structure, status of every item, locked decisions, open items | READY |
| `00_INDEX.md` | This file — maps every file to its purpose | READY |
| `qa_package.py` | Automated content + formatting audit for the Avant lender package (41/41; Refuge-era sent drafts exempted as history) | READY |
| `docfmt.py` | Shared document-formatting helper (fonts, footers, TOC fields) used by all `build_*.py` generators | REF |
| `SUBMISSION_COVER_SHEET_and_TOC.docx`, `SUBMISSION_READINESS_CHECKLIST.docx` | Hickory-era submission materials | HISTORICAL |
| `INFORMATION_NEEDED_FROM_YOU.docx`, `INTERVIEW_for_remaining_drafts.docx` | Hickory-era Q&A worksheets | HISTORICAL |
| `QA_AUDIT_REPORT.md`, `COUNSEL_QA_REVIEW_2026-06-16.md`, `OVERNIGHT_WORK_SUMMARY_2026-06-16.md`, `DISCUSSION_BRIEF_for_John_Hart_lens.md` | Hickory-era QA/strategy notes | HISTORICAL |
| `build_documents.py`, `build_hickory_oa.py`, `build_oa_restatement.py`, `build_refuge_loi.py` | Hickory-era or superseded generator scripts | HISTORICAL |
| `build_business_plan_rev6.py`, `build_checklist_response.py`, `build_uop_workbook.py` | Current, active generator scripts for the Avant package (business plan Rev 7.00, checklist response, UOP workbook) | REF |
| `build_uop_reply.py`, `build_insurance_app.py`, `build_investor_raise.py`, `build_prospect_research.py`, `build_valuation_package.py`, `build_mipa_reply.py`, `build_track_memo.py` | Refuge-era generator scripts (sent/superseded deliverables) | HISTORICAL |

## 00_source_forms/ — blank forms + lender material (as provided)
| File | Purpose | Status |
|---|---|---|
| `1_SBA_7a_Loan_Checklist*.docx` | Lender's master checklist (incl. John Hart's 5/19 version) | REF |
| `2…7a_*_FORM.pdf` | Blank SBA forms (Company Profile, Use of Funds, Business Plan guide, Debt Schedule, 912, 413, Cash Flow) | REF |
| `8_Management_Resume_TEMPLATE.docx` | Resume template | REF |
| `DRAFT_reply_to_John_Hart.md`, `JOHN_HART_OPEN_ITEMS_and_DECISIONS.md`, `John_Hart_email_thread…pdf` | Lender correspondence + open items (Hickory-era; superseded by `13_checklist_response_2026-07-16/`) | HISTORICAL |

## 01_company_profile/ through 07_personal_cash_flow/ — personal & entity forms
These folders hold SBA personal forms (Company Profile, Form 1919, Form 413 PFS, Form 912,
management resumes, personal cash flow). The *values* for Geoff personally (income, assets,
citizenship, etc.) apply regardless of the target switches, but the drafted narrative text in
these files references earlier targets and structures — **needs a pass to swap in Avant
Hospice, LLC and the Rev 5.00 figures** before these go to the lender. Tracked as an open item
in `00_CHECKLIST_and_STATUS.md`. (The filled Company Profile in
`13_checklist_response_2026-07-16/` is already Avant-based.)

## 08_entity_documents/
| File | Purpose | Status |
|---|---|---|
| `Tyler_Hospice_Hold_LLC_AMENDED_and_RESTATED_OA.docx` | Holdco operating agreement — cap table (39.9/19.5/13.3×3/0.7) is current and unchanged, but the operating-subsidiary references still name Hickory | **NEEDS RECONCILIATION** to Avant |
| `Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx` | Single-member OA drafted for Hickory as the acquired opco | HISTORICAL — needs an Avant Hospice, LLC equivalent drafted |
| `OA_Amendment_No_1_DRAFT.docx`, `Tyler_Hospice_HoldCo_OA_v7*_AS_PROVIDED.docx` | Prior OA versions, superseded by the A&R OA | HISTORICAL |
| `OA_RECONCILIATION_MEMO.md`, `OA_RESTATEMENT_REVIEW.md` | Notes from reconciling the OA to the (then-Hickory) SBA structure | REF |
| `Bullard_Subscription_and_Capital_Contribution_DRAFT.docx`, `Bullard_Investor_Attestation_*` | Bullard's $195K-of-the-current-$250K subscription and passive-investor attestation — dollar figures need a pass for the current $250K/33% structure; note the note→equity conversion open item | PRE-FILL |

## 09_financial_model/ and financial_models/
| File | Purpose | Status |
|---|---|---|
| `financial_models/output/Azalea_Hospice_Proforma_Rev5.00_AVANT.xlsx` | **Current financial model** — Avant structure (deferred $300K seller note, month-2 SBA takeout), full engine, cover sheet, 22/22 QA clean | READY |
| `financial_models/build_proforma_v3.py` (run as `python3 -m financial_models.build_proforma_v3`), `qa_proforma_v3.py`, `format_workbook.py` | Generator, QA harness, and post-build formatter for the current model | REF |
| `financial_models/MODEL_RECONCILIATION.md` | Explains why different model generations' figures differ | REF |
| `financial_models/output/Azalea_Hospice_Proforma_Rev4.10_DYNAMIC.xlsx` | Refuge-era Rev 4.10 model | HISTORICAL |
| `financial_models/output/Azalea_Track_Comparison_Rev1.00.xlsx`, `ppeo_stress.py`, `build_track_comparison.py`, `build_track_memo.py` outputs | Four-track cash-flow comparison and PPEO stress analyses that drove the Avant pivot | REF |
| `09_financial_model/*.xlsx`, `MODEL_RECUT_NOTE.md` | Hickory-era workbooks | HISTORICAL |

## 10_supporting_documents/
| File | Purpose | Status |
|---|---|---|
| `office_lease/` | Executed Tyler office lease (13387 Hwy 69 N, 1,607 sf) — reviewed for tenant-entity/term flags | PROVIDED |
| `purchase_agreement_seller_note/` | Hickory-era MIPA drafts and conflicts memo | HISTORICAL — the current (unsigned) Avant purchase agreement terms are tracked in `00_CHECKLIST_and_STATUS.md`; the executed Avant agreement is not yet filed here |
| `personal_financial_statement_support/` | Geoff's PFS asset support (VistaRiver sale/note) | PROVIDED |
| `equity_injection_evidence/` | Bullard source-of-funds evidence — needs a pass for the current $250K structure | PRE-FILL |
| `bank_brokerage_statements/`, `credit_drivers_license/`, `tax_returns/` | Drop-zones for borrower documents | NEEDS YOU |

## 11_lender_credit_memo/
Hickory-era credit memo, DSCR stress test, cover letter, and closing checklist — dollar figures
and structure are superseded. **HISTORICAL** — rebuild against the Rev 5.00 model before use,
or supersede into `13_checklist_response_2026-07-16/`.

## 12_investor_raise/
$500,000 investor raise materials (deck, teaser, FAQ, data room, prospect research) — built
against the Refuge structure; figures need an Avant pass before reuse. HISTORICAL/PRE-FILL.

## 13_checklist_response_2026-07-16/ — current lender submission package (AVANT)
Filled PDF forms (Company Profile, Use of Proceeds, Business Debt Schedule), the checklist
status memo, cover email, Avant addendum, assumptions narrative, the financing-timing reply to
John Hart, the UOP workbook (`UOP Azalea Avant Rev3.00.xlsx`, John Hart's template format), the
Rev 5.00 projections copy, and `AUDIT REPORT.docx` from `qa_package.py`. Documents 8, 9, and 11
are Refuge-era sent drafts retained as history (exempted from the stale scan). **READY — this
is the authoritative current package.**

## 14_insurance_2026-07-16/
PL/GL insurance application (Todd Plummer) and reply email, including the year-one W2 payroll
table by class. Workers' comp underwriting answers (BerkleyNet, Pie) were sent by email but not
yet filed here. Refuge-era (sent); insurance will need re-papering to Avant at binding. SENT/REF.

## 15_valuation_2026-07-20/
Preliminary Refuge Hospice information package and cover email for William at healthfmv.com
(FMV valuation engagement). Refuge-era. HISTORICAL / DELIVERED.

## 16_refuge_operational_transition/
Refuge interim operating plan (task/owner/dependency/deadline, de-facto-control flag) and the
earlier category-level checklist. Refuge-era. HISTORICAL / DELIVERED.

## 17_mipa_negotiation_2026-07-22/
Refuge MIPA negotiation record — Rev 1.02 threaded-reply redline, Bender correspondence, and
transmittal email. Refuge-era. HISTORICAL / DELIVERED.

## 18-20 — analysis folders (drove the Avant pivot)
`18_*` interim management agreement (Refuge; signature-ready final), `19_*` Jason
alternate-delivery-site proposal, `20_track_comparison_2026-07-28/` four-track cash-flow
comparison memo. The track comparison and PPEO stress test are the analyses that motivated the
Avant target: the Refuge stack could not absorb prepayment review, while Avant's fully
deferred seller note plus month-2 SBA takeout keeps the base case cash-positive throughout. REF.

---

### Not in this folder (historical — Hickory Hospice, in Google Drive "9. Hickory Hospice LLC")
Hickory's TX formation and EIN documents (Certificate of Formation, Certificate of Filing, CP-575
EIN 86-2886807) are retained in Google Drive for record-keeping but are not part of the active
Avant Hospice application.
