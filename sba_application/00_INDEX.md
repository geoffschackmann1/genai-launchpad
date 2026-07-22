# Azalea Hospice — SBA 7(a) Application — File Index

**Borrower:** Tyler Hospice Hold, LLC (dba Azalea Hospice & Palliative Care) · EIN 41-4966640
**Target/operating co:** Refuge Hospice, LLC (San Antonio, TX) · NPI 1174237572 · CCN/PTAN A91679
· being acquired via a phased change of ownership (CHOW)
**Loan:** $500,000 SBA 7(a) (Rev 4.10) · **Project:** $750,000 · **Equity:** $250,000 (33%)
**Lender contact:** John Hart (SourceFunding)

> Start with **`00_CHECKLIST_and_STATUS.md`** (master tracker) for current status and open items.
> The active, current-deal folders are **`13_checklist_response_2026-07-16/`** (lender
> submission package), **`14_insurance_2026-07-16/`** (PL/GL/WC insurance), **`15_valuation_2026-07-20/`**
> (FMV valuation package), and **`16_refuge_operational_transition/`** (interim operating plan).
> Legend: **READY** = drafted/complete · **PRE-FILL** = mostly done, add personal data ·
> **NEEDS YOU** = borrower must provide · **PROVIDED** = source doc on file · **REF** = reference only ·
> **HISTORICAL** = superseded Hickory-era material, not part of the active package.

---

## Root
| File | Purpose | Status |
|---|---|---|
| `00_CHECKLIST_and_STATUS.md` | Master tracker — current structure, status of every item, locked decisions, open items | READY |
| `00_INDEX.md` | This file — maps every file to its purpose | READY |
| `qa_package.py` | Automated content + formatting audit for the checklist-response and insurance packages (40/40) | READY |
| `docfmt.py` | Shared document-formatting helper (fonts, footers, TOC fields) used by all `build_*.py` generators | REF |
| `SUBMISSION_COVER_SHEET_and_TOC.docx`, `SUBMISSION_READINESS_CHECKLIST.docx` | Hickory-era submission materials | HISTORICAL |
| `INFORMATION_NEEDED_FROM_YOU.docx`, `INTERVIEW_for_remaining_drafts.docx` | Hickory-era Q&A worksheets | HISTORICAL |
| `QA_AUDIT_REPORT.md`, `COUNSEL_QA_REVIEW_2026-06-16.md`, `OVERNIGHT_WORK_SUMMARY_2026-06-16.md`, `DISCUSSION_BRIEF_for_John_Hart_lens.md` | Hickory-era QA/strategy notes | HISTORICAL |
| `build_documents.py`, `build_hickory_oa.py`, `build_oa_restatement.py`, `build_refuge_loi.py` | Hickory-era or superseded generator scripts | HISTORICAL |
| `build_business_plan_rev6.py`, `build_checklist_response.py`, `build_uop_workbook.py`, `build_uop_reply.py`, `build_insurance_app.py`, `build_investor_raise.py`, `build_prospect_research.py`, `build_valuation_package.py` | Current, active generator scripts for the Refuge-era package | REF |

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
citizenship, etc.) still apply regardless of the Hickory-to-Refuge target switch, but the
drafted narrative text in these files references Hickory Hospice and the old $945K structure —
**needs a pass to swap in Refuge Hospice, LLC and the Rev 4.10 figures** before these go to the
lender. Tracked as an open item in `00_CHECKLIST_and_STATUS.md`.

## 08_entity_documents/
| File | Purpose | Status |
|---|---|---|
| `Tyler_Hospice_Hold_LLC_AMENDED_and_RESTATED_OA.docx` | Holdco operating agreement — cap table (39.9/19.5/13.3×3/0.7) is current and unchanged, but the operating-subsidiary references still name Hickory | **NEEDS RECONCILIATION** to Refuge |
| `Hickory_Hospice_LLC_SINGLE_MEMBER_OA.docx` | Single-member OA drafted for Hickory as the acquired opco | HISTORICAL — needs a Refuge Hospice, LLC equivalent drafted |
| `OA_Amendment_No_1_DRAFT.docx`, `Tyler_Hospice_HoldCo_OA_v7*_AS_PROVIDED.docx` | Prior OA versions, superseded by the A&R OA | HISTORICAL |
| `OA_RECONCILIATION_MEMO.md`, `OA_RESTATEMENT_REVIEW.md` | Notes from reconciling the OA to the (then-Hickory) SBA structure | REF |
| `Bullard_Subscription_and_Capital_Contribution_DRAFT.docx`, `Bullard_Investor_Attestation_*` | Bullard's $195K-of-the-current-$250K subscription and passive-investor attestation — dollar figures need a pass for the current $250K/33% structure | PRE-FILL |

## 09_financial_model/ and financial_models/
| File | Purpose | Status |
|---|---|---|
| `financial_models/output/Azalea_Hospice_Proforma_Rev4.10_DYNAMIC.xlsx` | **Current financial model** — full engine, cover sheet, 23/23 QA clean | READY |
| `financial_models/qa_proforma_v3.py`, `financial_models/format_workbook.py` | QA harness and post-build formatter for the current model | REF |
| `financial_models/MODEL_RECONCILIATION.md` | Explains why different model generations' figures differ | REF |
| `09_financial_model/*.xlsx`, `MODEL_RECUT_NOTE.md` | Hickory-era workbooks | HISTORICAL |

## 10_supporting_documents/
| File | Purpose | Status |
|---|---|---|
| `office_lease/` | Executed Tyler office lease (13387 Hwy 69 N, 1,607 sf) — reviewed for tenant-entity/term flags | PROVIDED |
| `purchase_agreement_seller_note/` | Hickory-era MIPA drafts and conflicts memo | HISTORICAL — current MIPA terms are in `00_CHECKLIST_and_STATUS.md` and `15_valuation_2026-07-20/`; the Refuge MIPA itself is still unsigned and not yet filed here |
| `personal_financial_statement_support/` | Geoff's PFS asset support (VistaRiver sale/note) | PROVIDED |
| `equity_injection_evidence/` | Bullard source-of-funds evidence — needs a pass for the current $250K structure | PRE-FILL |
| `bank_brokerage_statements/`, `credit_drivers_license/`, `tax_returns/` | Drop-zones for borrower documents | NEEDS YOU |

## 11_lender_credit_memo/
Hickory-era credit memo, DSCR stress test, cover letter, and closing checklist — dollar figures
and structure are superseded by Rev 4.10. **HISTORICAL** — needs a rebuild against the current
model before use, or supersedes into `13_checklist_response_2026-07-16/`.

## 12_investor_raise/
$500,000 investor raise materials (deck, teaser, FAQ, data room, prospect research) — built
against the current Refuge structure. READY.

## 13_checklist_response_2026-07-16/ — current lender submission package
Filled PDF forms (Company Profile, Use of Proceeds, Business Debt Schedule), the checklist
memo, cover email, Refuge addendum, assumptions narrative, the UOP workbook (John Hart's
template format), and `AUDIT REPORT.docx` from `qa_package.py`. **READY — this is the
authoritative current package.**

## 14_insurance_2026-07-16/
PL/GL insurance application (Todd Plummer) and reply email, including the year-one W2 payroll
table by class. Workers' comp underwriting answers (BerkleyNet, Pie) were sent by email but not
yet filed here. READY.

## 15_valuation_2026-07-20/
Preliminary Refuge Hospice information package and cover email for William at healthfmv.com
(FMV valuation engagement) — license/certification detail, negotiated purchase terms, market
comps, and the PTAN-reactivation and zero-census/near-zero-financials flags. READY / DELIVERED.

## 16_refuge_operational_transition/
`refuge-operating-plan.md` — the current, authoritative interim operating plan (task/owner/
dependency/deadline, working back from the July 31 patient admission), including the
de-facto-control regulatory flag. `REFUGE_TRANSITION_PLAN.md` is the earlier category-level
checklist it supersedes; kept for reference. READY / DELIVERED.

---

### Not in this folder (historical — Hickory Hospice, in Google Drive "9. Hickory Hospice LLC")
Hickory's TX formation and EIN documents (Certificate of Formation, Certificate of Filing, CP-575
EIN 86-2886807) are retained in Google Drive for record-keeping but are not part of the active
Refuge Hospice application.

No reference to an entity named "Avant Hospice" was found in this repository. If that refers to
something not yet captured here, flag it so it can be tracked down and folded into this index.
