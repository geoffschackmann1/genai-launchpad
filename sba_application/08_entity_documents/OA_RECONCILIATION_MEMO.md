# OA v7 vs SBA-Compliant Structure - Reconciliation Memo

**Status:** Material conflicts identified. Direction needed before further drafting.

The Operating Agreement (`Tyler_Hospice_HoldCo_OA_v7_AS_PROVIDED.docx`, title-page version 5,
March 2026) is well-drafted, but it was structured **before** the SBA-compliance work we've done
together over the last two days. Several provisions in the OA directly contradict the structure
documented in the rest of this package, and two of them are structural enough to potentially
invalidate the S-Corporation tax election or trigger SBA underwriting flags.

---

## Conflict Map

### 1. CAP TABLE — direct contradiction

| Position                  | OA v7 | SBA package |
|---------------------------|-------|-------------|
| Adeline & Lilah, LLC      | 60.0% | 39.9%       |
| James E. Bullard          | 25.0% | 19.9%       |
| Silas R. Shelton          | 4.9%  | 13.3%       |
| Dana L. Davenport         | 4.9%  | 13.3%       |
| Bradley G. Woodard        | 4.9%  | 13.3%       |
| Unissued / reserved       | 0.3%  | 0.3%        |

**Why this matters:** at **25%, Bullard is a 20%+ owner under SBA rules**, which triggers
mandatory PFS (Form 413), Personal History (Form 912), three years of tax returns, and an
unconditional personal guaranty from him. The whole reason we structured him at 19.9% in the
SBA package was to avoid that — you confirmed Bullard does not want to provide those personal
disclosures, and the equity-injection rules of SOP 50 10 8 still let his cash count as
qualifying injection at <20%.

**Reconciliation options:**
- (a) Amend the OA cap table to 39.9 / 19.9 / 13.3 × 3 / 0.3 and keep the SBA structure as
  documented.
- (b) Keep the OA at 60 / 25 / 4.9 × 3 and re-do the SBA package to require Bullard's PFS,
  Form 912, tax returns, and personal guaranty.

### 2. BULLARD'S CONTROL RIGHTS — SBA underwriting risk

The OA gives Bullard:
- Voting rights on all **Reserved Matters** (§6.2): sale of assets, merger, additional
  issuances, debt > $500K, amendments, dissolution, change of control
- **Anti-dilution protection with a hard 20% floor** (§§6.8, 12.3) — no new investor can
  reduce him below 20% without his written consent
- **Pro-rata participation right** on any new issuance (§12.6)
- **Texas Shootout** buy-sell to break deadlock with Geoff (§8.10)

**Why this matters:** SOP 50 10 8 specifically polices the practice of "an investor taking
less than 20% to dodge a guaranty while using a side agreement to control the business." The
combination of (i) voting on Reserved Matters, (ii) anti-dilution to 20%, and (iii) Texas
Shootout deadlock-buyout could be read by an SBA credit analyst as effective control —
exactly the eligibility flag we drafted Bullard's no-control attestation to avoid.

**Reconciliation options:**
- (a) Strip Bullard's voting rights on Reserved Matters; recharacterize his rights as
  information / observer only; remove anti-dilution and 20% floor. He becomes truly passive.
- (b) Keep his rights and accept that he's a controlling member (combine with cap-table
  option 1(b) above — full SBA disclosures).

### 3. S-CORP ELECTION vs ADELINE & LILAH OWNERSHIP — potential structural defect

The OA elects **S-Corporation tax treatment** under IRC §1361 and lists Adeline & Lilah, LLC
as a 60% Member. The OA recital describes A&L as **"wholly owned and controlled by Schackmann."**

Last night you confirmed Mary Elizabeth Burcham is a **50% member** of Adeline & Lilah, LLC
(consistent with the VistaRiver MIPA signature page where she signs as 50% owner of A&L).

**Why this matters:** Under IRC §1361, an S-Corporation's shareholders may only be (i) U.S.
individuals, (ii) single-member LLCs treated as disregarded entities, (iii) certain trusts,
or (iv) certain estates. **A multi-member LLC is taxed as a partnership and cannot be an
S-Corp shareholder.** If A&L has two members (Geoff and Mary Elizabeth), it is a partnership
for federal tax purposes — and admitting it as a Member of an S-Corp would **automatically
terminate the S-Corp election** under §1362(d)(2).

**Reconciliation options:**
- (a) Restructure A&L as a single-member LLC owned 100% by Geoff (Mary Elizabeth removed as
  member — would require buying her out, gifting, or restructuring her interest as a
  beneficial-only / community-property interest under Arizona law without separate membership).
- (b) Drop the S-Corp election and let Tyler Hospice HoldCo be taxed as a partnership (default
  multi-member LLC treatment). This is simpler and avoids the eligibility issue. Most lenders
  don't care which way you go.
- (c) Hold Tyler Hospice HoldCo membership directly in Geoff's individual name (rather than
  through A&L) — preserves S-Corp eligibility but undoes the holding-entity structure.

**This is the most urgent issue.** Until it's resolved, the S-Corp election in the OA is
arguably invalid as of the Effective Date.

### 4. ACQUISITION TARGET — naming update

The OA §3.14(a) describes the acquisition target as **"Healing Hands Palliative Hospice INC."**
You've confirmed the actual target is **Hickory Hospice LLC** (current owners Tracy Gleason
and Ann Lozano). The OA needs the target name updated wherever it appears (Section 3.14, the
Section 1.2 business description, and the HCSSA refund clause in Section 3.15).

### 5. OPERATOR-MEMBER VESTING — different schemes

The OA (§§5.1-5.5) uses **milestone-based forfeiture**: 25% lapses at Breakeven, 50% lapses
at 3 consecutive Profitable months, 25% lapses at 12 consecutive Profitable months. Each
operator gets a 4.9% Restricted Interest tied to §83(b) election.

You confirmed (last message) you want **4-year time vesting with a 1-year cliff** instead.

**These are quite different philosophies.** The OA's milestone approach actually has some
appeal — it ties vesting to the business hitting financial goals (which is what you really
care about), and the §83(b) framework is already drafted. The 4/1 cliff is what most
operating agreements use and is what the lender will recognize without explanation. You could
go either way. Worth a short conversation before we lock it.

---

## Smaller items captured from the OA (good to update package with)

- **Geoff's address:** 4602 E Cheery Lynn Rd, Phoenix, AZ 85018 (you wrote "Cherryland Road" —
  that's not a Phoenix street name; "Cheery Lynn Road" is what's in the OA and is a real
  street near 44th/Camelback. Confirming this is the right one.)
- **Bullard's address:** 13910 Indiana Ave, Suite 300, Lubbock, TX 79423
- **Silas's address:** 1503 Lake Park Circle, Hideaway, TX 75771
- **Bradley's address:** 421 W Cumberland Rd, Apt 403, Tyler, TX 75703
- **Dana's address:** missing in the OA (placeholder)
- **Geoff's OA-era email:** geoff@adelineandlilah.com (will use geoff@azaleahospice.com for
  the SBA file per your direction)
- **Bullard's accredited-investor representation** (§9.2) is already in the OA — useful for
  the Subscription Agreement framing.
- **OIG / SAM exclusion representations** (§9.1(d)) are already covered — useful for
  affiliate/eligibility memo.
- **Non-compete on operator-members:** 75-mile radius around Tyler, 24 months post-departure
  (§6.4) — STRONG protection for the deal; good to highlight to the lender.
- **HCSSA refund clause** (§3.15): if acquisition isn't completed in 12 months, Bullard can
  demand return of "Undeployed HCSSA Funds." Lender will want to see this and may ask whether
  it's structured to fund/secure during the SBA loan term.

---

## Recommended path forward

1. **Decide cap table** (60/25 with Bullard guaranteeing, or 39.9/19.9 with him passive).
   Recommendation: keep the SBA-compliant 39.9/19.9 unless Bullard is genuinely willing to
   personally guaranty and disclose. Confirm with Bullard first.
2. **Decide S-Corp vs partnership tax**. Recommendation: switch to partnership taxation
   (default LLC). Simpler, avoids the multi-member A&L issue, and no SBA lender cares about
   the difference. Re-electing S-Corp later if it makes sense is straightforward.
3. **Decide on operator vesting** (milestone-based per OA, or 4/1 cliff per your message).
   Recommendation: keep the OA's milestone vesting — it's well-drafted, ties to performance,
   and §83(b) is already structured.
4. **Update acquisition target language** to Hickory Hospice LLC throughout the OA.
5. **Strip or scale back Bullard's voting / anti-dilution rights** so the no-control story for
   SBA is clean. Re-paper his rights as observer/information only.

Once you confirm direction on 1-3, I can produce an **Amendment and Restatement of the
Operating Agreement** that resolves all five items in one document, plus a revised Exhibit A,
and an updated Bullard Subscription Agreement to replace the language in OA §11.5.

---

## Separate timing issue (raised earlier)

You mentioned **Hickory target closing date: June 23, 2026.** We previously locked
**SBA close after the July equity tranche** so the full $250K is on deposit at SBA
disbursement (SOP 50 10 8). These can't both happen — in a CHOW the SBA loan typically funds
at or near the acquisition close so working-capital proceeds are in the account.

Options:
- (a) Push Hickory close to late July / early August (after Tranche 3).
- (b) Accelerate Bullard tranches 2 and 3 to land before June 23.
- (c) Negotiate a disbursement holdback with the lender — SBA loan closes at Hickory close
  but working-capital proceeds release as remaining tranches land.

Recommendation: (b) if Bullard can move money up; otherwise (c).
