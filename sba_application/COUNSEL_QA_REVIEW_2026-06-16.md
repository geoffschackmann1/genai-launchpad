# Counsel-Style QA Review — Tyler Hospice Hold, LLC / Azalea SBA 7(a) Package
**Date:** 2026-06-16  **Prepared by:** (engagement-style review at client request)

> **Not legal advice.** This memo is organized the way deal counsel would triage the file:
> it identifies issues, risk, and the decisions needed. You should have your retained
> attorney and SBA lender confirm the conclusions before anything is signed or submitted.
> Scope: the entire `sba_application/` package, the provided OA v7, the Hickory MIPA
> (Rev 2.00), and the financial model — read against the structure you confirmed on 6/16.

---

## A. Structure as you confirmed it (6/16)

Two-tier; **no operating company between them**:

```
Adeline & Lilah, LLC (AZ)  39.9% ─┐
James E. Bullard           19.5% ─┤
Shelton / Davenport / Woodard      ├─►  TYLER HOSPICE HOLD, LLC  ──100%──►  HICKORY HOSPICE LLC
   13.3% each (Restricted)        ─┤      (holding co. + SBA borrower)        (operating sub;
Reserved pool               0.7% ─┘                                           keeps Medicare CCN
                                                                              via CHOW; dba "Azalea
                                                                              Hospice & Palliative Care")
```

- **Acquisition mechanics:** purchase of **100% of Hickory's membership interests** (CHOW), so
  Hickory survives as a wholly-owned subsidiary and **retains its Medicare/Medicaid provider
  numbers** — that is the entire point of the CHOW and is correctly reflected in the MIPA.
- **Financing as now described:**
  - **SBA 7(a)** = the main facility (working capital + startup).
  - **$300K acquisition financing** = a **bank term loan from Jim Bullard's bank, 6% / 3 yr**,
    refinancing the MIPA's interest-free 12×$25K seller installments.
  - **Equity** = Bullard's $195K cash (19.5%).

This two-tier model **matches what the package already describes** (Hickory as "wholly owned
operating subsidiary"). So the structure itself is sound; the problems are (1) **which entity /
state / name is real**, (2) **four inconsistent capital structures across the file**, (3) **the
SBA mechanics of a holdco borrower + a concurrent non-SBA bank loan**, and (4) **a federal-
application integrity item** flagged by the lender.

---

## B. BLOCKERS — resolve before signing or submitting anything

### B1. Entity name + state knot (the single most important cleanup)
Three different identities are in play for the **same** holding company:

| Source | Name | State |
|---|---|---|
| Hickory MIPA (Rev 2.00, binding) | **Tyler Hospice Hold, LLC** | **Texas** |
| Operating Agreement v7 | **Tyler Hospice HoldCo L.L.C.** | **Wyoming** |
| SBA forms / EIN 41-4966640 / closing checklist | Tyler Hospice Hold LLC | **Wyoming** (formed 03/18/2026) |

The borrower, the MIPA buyer, the OA party, and the SBA Form 1919 borrower **must be the exact
same legal name and state**. A federal loan application and a binding purchase agreement that name
different entities/states is a first-pass underwriting reject and a diligence red flag.

**I need from you:** the **actual Certificate of Formation** — is Tyler Hospice Hold, LLC a **Texas**
or **Wyoming** entity? — its **exact charter name** (with/without "Co", comma, "L.L.C." vs "LLC"),
and **which entity holds EIN 41-4966640**. Everything else conforms to that answer.

### B2. Four conflicting capital / seller-financing structures
The file currently contains four different versions of how the deal is funded:

| Version | SBA loan | Equity | $300K Hickory payment | Source |
|---|---|---|---|---|
| **"Locked" master tracker** | $555K | $195K | **Paid in full at close, no note** | `00_CHECKLIST_and_STATUS.md` |
| **MIPA (binding)** | n/a | n/a | **12 × $25K, interest-free, over 12 mo** | MIPA Rev 2.00 §1.02 |
| **Financial model defaults** | $555K @ 11.5% / 120 mo | $195K | License paid at close (note flag off) | engine `model.py` |
| **Your instructions (6/15 + 6/16)** | $600K @ ~10.5% / 180 mo | $180K used | **Bank loan $300K / 6% / 3 yr** | this session |

These must collapse to **one** structure. As counsel the controlling fact is: **the MIPA is the
binding document and it says 12 × $25K interest-free.** Everything else (model, credit memo, DSCR,
runway) currently describes a *different* obligation. Either (a) the MIPA is honored and the
application is rebuilt on 12×$25K (Year-1 DSCR ≈ 0.77x — a problem), or (b) the MIPA is amended to
match the real plan (pay-at-close, or the bank-refi), or (c) the bank refinance closes
**concurrently** so the seller is paid and the company owes the bank $300K/6%/3yr from day one.
Until this is settled, **no DSCR or runway number in the package is reliable** (see §E table).

### B3. Holdco is the borrower, but the cash flow and license live in Hickory
The borrower (Tyler Hospice Hold, LLC) is a **pure holding company** — no operations, no Medicare
billing, no employees. All revenue, the provider number, the receivables, and the assets sit in
**Hickory / Azalea**. SBA lenders will not lend to an equity holdco on an unsecured look-through;
expect them to require:
- **Hickory (the operating company) to be a co-borrower or full corporate guarantor**, and
- a **lien on Hickory's assets and receivables** (and a pledge of the Hickory membership interests).

This is the SBA EPC/OC pattern (here a holdco/opco variant). The package does **not** currently
name Hickory as co-borrower/guarantor anywhere. **Decision needed:** co-borrower vs. corporate
guarantor for Hickory (recommend confirming with John Hart early — it drives the loan documents
and the resolutions).

### B4. Concurrent non-SBA bank loan = "piggyback" financing (SBA disclosure + lien priority)
A **second lender (Jim's bank) making a $300K term loan at the same time** as the SBA 7(a),
secured against the same business, is *piggyback financing* under SBA rules. Consequences:
- The **SBA lender almost always requires first lien**; the bank's $300K would have to be
  **subordinated** or secured only by a specific carve-out, with an **intercreditor/subordination
  agreement**.
- The piggyback loan **must be disclosed to and approved by** the SBA lender; undisclosed
  concurrent debt is a serious problem.
- If the SBA loan is *only* working capital while a separate bank funds the **acquisition**, the
  SBA lender may ask why the CHOW isn't in the SBA facility. This is allowable but needs a clean
  explanation.
- **Bullard's role:** if the bank is lending partly on Jim's relationship and he **guarantees**
  that bank loan, that is permissible, but note it sits slightly against the "Bullard is passive
  and guarantees nothing" posture in the SBA file — keep the narratives consistent and disclose
  the related-party relationship (his depository bank). Confirm Jim has **no ownership** in the
  bank (if he does, that's a conflict to disclose).

The OA Amendment §5.2 already carves "Hickory seller financing… and any commercial loan that
refinances or replaces it" out of Bullard's Reserved-Matters consent — so the *governance* side is
handled, but the *SBA lien/disclosure* side is not yet addressed in the credit package.

### B5. Expansion-vs-startup and affiliate accuracy (integrity-critical)
The lender (John Hart) flagged that Geoff told him both "**no affiliated companies**" and that he
has "**10+ years of hospice ownership**," and pointedly asked what happened to the prior business.
The package resolves this as: **startup CHOW**, VistaRiver interest **sold Aug 2025** (passive note
only), **no current operating affiliate**. That is internally consistent **only if** Geoff truly
owns no other operating business today. This must be **exactly right** — affiliate disclosure and
the startup-vs-expansion framing go on a federal application (18 U.S.C. §1001 exposure for
misstatement). **Confirm:** Geoff currently owns/controls no operating business other than passive
interests already disclosed → startup is correct and the "no affiliate" answer stands.

---

## C. HIGH priority

- **C1. Two operating agreements needed.** We have been amending the **holdco** OA (Tyler Hospice
  Hold). At close you also need a **single-member OA for Hickory** (holdco as sole member,
  manager-managed) — the MIPA already provides for the old managers' resignation. The holdco OA
  must be retitled to the exact charter name/state from B1 and finalized as the **Amended &
  Restated OA** you chose.
- **C2. Tax.** Holdco = **partnership** (multi-member; the Amendment already moves it off the
  invalid S-corp election — correct). Hickory, once 100%-owned by the holdco, is a **single-member
  LLC = disregarded entity**; its results flow up into the holdco partnership return. Confirm no
  separate election is made for Hickory and that this is the intended treatment.
- **C3. Equity-injection trace.** SBA requires the $195K to be **in the borrower** and traceable.
  Confirm the Mercury account ****1275 is **Tyler Hospice Hold, LLC's** account (the borrower), not
  Hickory's or Geoff's, and that the source-of-funds chain (Bullard's statements → wire → borrower)
  is complete for both tranches.
- **C4. Runway / DSCR must be rebuilt on the chosen B2 structure.** Current published DSCR
  (3.23x / 7.26x / 10.76x) assumes **paid-at-close, no note**. If the real structure is the bank
  refi, the correct Year-1 figure is ~**1.60x** (still fine); if the MIPA stands unrefinanced,
  Year-1 is ~**0.77x** (a lender problem). The credit memo, business plan, and stress tests all
  cite the 3.23x line and would need updating once B2 is locked.
- **C5. Guaranty look-through.** Through the holdco: A&L is a **39.9% indirect owner of the OC →
  must guarantee** (entity guaranty). Geoff is **19.95% indirect** (just under 20%) **but is the
  controlling Manager**, so the SBA lender will require **his personal guaranty as the controlling
  principal** regardless — which the package already provides (Geoff sole PG). Document the control
  rationale so the "<20% but still guarantees" point is not questioned. Bullard at 19.5% indirect
  stays under 20% (no SBA PG required) — but see B4 re: the separate bank loan.

---

## D. MEDIUM / cleanup

- **D1. "License" terminology.** You are buying **100% of Hickory's membership interests**, not a
  standalone "license." The model/docs label the $300K a "Hickory Medicare license" and a "license
  note / license intangible." For accuracy (and lender clarity) describe it as **purchase of 100%
  membership interests of Hickory (CHOW; provider number conveys with the entity)**. The $300K
  intangible amortization treatment should be revisited with the CPA (equity purchase vs. asset/
  intangible step-up has different tax mechanics).
- **D2. Seller PPP/EIDL.** John's checklist requires confirming whether **Gleason/Lozano** (Hickory
  sellers) have any PPP/EIDL/stimulus loans — needed for a CHOW.
- **D3. COVID questionnaire** (10 items on John's template) — still to draft.
- **D4. NAICS 621610** — confirm with lender (used throughout).
- **D5. Dana Davenport** address/email still missing (OA signature block + roster).
- **D6.** The two `*_AS_PROVIDED*` OA files intentionally retain old (60/25/4.9) figures — fine as
  historical uploads, but make sure the lender receives the **Amended & Restated** version as the
  operative document, not the as-provided one.
- **D7.** OA cross-references and stale worked examples (§2.20 / §3.10 / §3.14 / §6.2 parenthetical /
  §2.19 pool / §9.1(d)) were **already corrected** this session.

---

## E. Runway under each candidate structure (model recomputed 6/16)

EBITDA is identical across all three (pre-debt): Y1 ≈ $302,522.

| Structure | Opening cash | Y1 debt service | **Y1 DSCR** | Min cash (36 mo) | LOC drawn |
|---|---|---|---|---|---|
| 1. Paid at close, no note (locked) | $372,000 | $93,637 | **3.23x** | $262,558 | $0 |
| 2. MIPA as signed (12×$25K, 0%) | $672,000 | $393,637 | **0.77x** | $400,917 | $0 |
| 3. Bank refi $300K/6%/3yr + SBA $600K | $702,000 | $189,108 | **1.60x** | $576,646 | $0 |

**Runway conclusion:** on pure liquidity you have ample runway in **every** scenario — the LOC is
never touched and minimum cash stays between ~$263K and ~$577K. The decision driver is **Year-1
DSCR for lender approval**, not cash survival. Financing the acquisition (2 or 3) actually leaves
*more* cash on hand than paying at close (1), at the cost of weaker Year-1 coverage optics.

---

## F. Decisions / facts I need to converge the file

1. **Entity (B1):** Texas or Wyoming? Exact charter name? Which entity owns EIN 41-4966640?
2. **Capital structure (B2):** pick ONE — (a) pay seller at close, no note; (b) honor the MIPA's
   12×$25K; or (c) bank refi $300K/6%/3yr concurrent at close. And **SBA loan size** ($500K / $555K
   / $600K) and **injection** ($180K / $195K).
3. **Hickory as co-borrower or corporate guarantor (B3)?** (Recommend confirming with John Hart.)
4. **Piggyback bank loan (B4):** is Jim's bank loan **at close** or a **later refinance**? Will Jim
   **guarantee** it? Does Jim have any **ownership** in that bank?
5. **Affiliate/startup (B5):** confirm Geoff owns no current operating business → startup framing
   stands.
6. **Bullard:** stay 19.5% no-PG, or move to a second guarantor if a lender requires it?

Once you lock items 1–2 (and ideally 3–4), I can rebuild the OA (A&R, correct entity), the MIPA
alignment note, the financial model, and the entire SBA narrative on a single consistent structure
in one pass.

---

## G. RESOLVED — decisions locked 2026-06-16

| # | Decision | Answer |
|---|---|---|
| 1 | Entity | **Tyler Hospice Hold, LLC — Wyoming** (EIN 41-4966640), foreign-qualified in TX. **MIPA must be amended** (its TX buyer is the error). |
| 2 | Capital structure | **SBA $450,000** (startup + working capital) + **bank loan $300,000 / 6% / 3 yr** (acquisition) + **equity $195,000** |
| 3 | Hickory | **Corporate guarantor** of the SBA loan (lien on its assets/receivables) |
| 4 | Bank loan | At close; **Jim Bullard personally guarantees it**; subordinate to SBA; no ownership in the bank |
| 5 | Eligibility | **Startup, no affiliate** (VistaRiver sold 8/2025) |
| 6 | Bullard | **19.5%, no SBA personal guaranty** (guarantees only the bank loan) |
| — | Injection | Bullard **$195,000 total** (cap table 19.5%); **$180,000 designated** as the SBA injection; $15,000 extra WC |

### Authoritative numbers (model rebuilt + workbooks regenerated 6/16)
- **Sources $945,000** = SBA $450,000 (47.6%) + bank $300,000 (31.7%) + equity $195,000 (20.6%)
- **Uses $945,000** = acquisition $300,000 + startup $63,000 + equipment $15,000 + WC reserve **$567,000**
- **Debt service (Yrs 1-3) $169,211/yr** = SBA $59,692 + bank $109,519 (after Yr 3, SBA only ~$59,692)
- **DSCR 1.79x / 4.02x / 5.95x; global 3.92x**; min cash **$444,962**; LOC never drawn

### Counsel caveats carried forward
- **Downside coverage is thinner** than the old paid-at-close structure: the $300K bank loan amortizes
  over only 3 years (~$109.5K/yr front-loaded). A ~20% census shock pushes **Year-1 DSCR below 1.0x**
  (about 0.29x), recovering by Year 2 (~1.18x). Base case (1.79x) is approvable and liquidity holds
  (the $567K reserve keeps the LOC undrawn), but expect lender stress-test scrutiny on the ramp year.
  A longer bank-loan amortization, or a smaller WC reserve, would ease the optics.
- **Piggyback bank loan** requires SBA-lender approval + subordination/intercreditor (SBA first lien);
  disclose Jim's guaranty of it.
- **MIPA amendment** (buyer -> Wyoming entity) and **Hickory corporate guaranty + intercreditor** still
  to be papered. **Two OAs** to finalize: holdco A&R (Wyoming) + a single-member OA for Hickory.
