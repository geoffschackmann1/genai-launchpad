"""Debt-service + runway overlay for the user's scenario:
   - SBA 7(a): $600,000 over 15 years (180 mo) at current typical SBA rate
   - Hickory license: $300,000 seller note at 6% over 3 years (36 mo)
   - Equity injection: $180,000 of Jim Bullard's $195,000
Recomputes the existing model's 36-month ramp with these terms.
"""
import financial_models.engine.model as m

SBA_P, SBA_N, EQUITY = 600_000, 180, 180_000


def make_assumptions(rate):
    orig = m.assumptions_dict.__wrapped__ if hasattr(m.assumptions_dict, "__wrapped__") else _ORIG
    def f():
        a = _ORIG()
        a["sba_principal"] = SBA_P
        a["sba_term_mo"] = SBA_N
        a["sba_rate"] = rate
        a["equity"] = EQUITY
        return a
    return f


_ORIG = m.assumptions_dict


def pmt(P, annual, n):
    r = annual / 12
    return P * r / (1 - (1 + r) ** -n)


def run(rate):
    m.LICENSE_PAID_AT_CLOSE = False           # license = $300K seller note
    m.assumptions_dict = make_assumptions(rate)
    return m.compute(scenario=1)              # base census path


def labels():
    out = []
    for p in m.PERIODS:
        if p["year"] == 1:
            out.append("M%d" % p["month_index"])
        else:
            q = ((p["month_index"] - 1) % 12) // 3 + 1
            out.append("Y%dQ%d" % (p["year"], q))
    return out


def main():
    print("=== LOAN TERMS ===")
    for rr in (0.095, 0.105, 0.115):
        mp = pmt(SBA_P, rr, SBA_N)
        print("SBA $600,000 / 180mo @ %.1f%%:  monthly $%,.0f   annual $%,.0f"
              .replace("%,", "{:,}").format(rr * 100, mp, mp * 12)
              if False else
              "SBA $600,000 / 180mo @ {:.1f}%:  monthly ${:,.0f}   annual ${:,.0f}".format(rr*100, mp, mp*12))
    lp = pmt(300_000, 0.06, 36)
    print("License $300,000 / 36mo @ 6.0%:  monthly ${:,.0f}   annual ${:,.0f}".format(lp, lp * 12))

    R, a = run(0.105)
    sc = R["_scalars"]
    startup = sc["startup_total"]
    beg = a["equity"] + a["sba_principal"] - startup - a["capex"]
    print("\n=== OPENING CASH (base rate 10.5%, license as seller note) ===")
    print("  equity ${:,.0f} + SBA ${:,.0f} - startup ${:,.0f} - capex ${:,.0f} - license $0 (note)"
          .format(a["equity"], a["sba_principal"], startup, a["capex"]))
    print("  = OPENING CASH ${:,.0f}   (min-cash floor ${:,.0f}; LOC backstops shortfalls)"
          .format(beg, a["min_cash"]))

    labs = labels()
    print("\n{:>6} {:>10} {:>9} {:>6} {:>10} {:>9}".format(
        "PER", "EBITDA", "DEBTSVC", "DSCR", "ENDCASH", "LOC_BAL"))
    minc, minlab, first = 1e18, "", None
    loc_total = 0.0
    for i, p in enumerate(m.PERIODS):
        e, ds = R["ebitda"][i], R["ds"][i]
        dscr = e / ds if ds else 0
        ec, loc = R["end_cash"][i], R["loc_bal"][i]
        loc_total += R["loc_draw"][i]
        if ec < minc:
            minc, minlab = ec, labs[i]
        if first is None and dscr >= 1.25 and p["month_index"] >= 3:
            first = labs[i]
        print("{:>6} {:>10,.0f} {:>9,.0f} {:>6.2f} {:>10,.0f} {:>9,.0f}".format(
            labs[i], e, ds, dscr, ec, loc))

    print("\nMin cash trough: ${:,.0f} at {}".format(minc, minlab))
    print("Total LOC drawn over 36 mo: ${:,.0f}".format(loc_total))
    print("First sustained DSCR >= 1.25x: {}".format(first))

    def yr(y):
        e = sum(R["ebitda"][i] for i, p in enumerate(m.PERIODS) if p["year"] == y)
        ds = sum(R["ds"][i] for i, p in enumerate(m.PERIODS) if p["year"] == y)
        return e, ds
    print("\n=== ANNUAL DSCR ===")
    for y in (1, 2, 3):
        e, ds = yr(y)
        print("Year {}: EBITDA ${:,.0f} | Debt service ${:,.0f} | DSCR {:.2f}x".format(y, e, ds, e / ds))

    print("\n=== SBA RATE SENSITIVITY (Year-2 DSCR & min-cash trough) ===")
    for rr in (0.095, 0.105, 0.115, 0.125):
        R2, a2 = run(rr)
        e2 = sum(R2["ebitda"][i] for i, p in enumerate(m.PERIODS) if p["year"] == 2)
        ds2 = sum(R2["ds"][i] for i, p in enumerate(m.PERIODS) if p["year"] == 2)
        mc = min(R2["end_cash"])
        loc = sum(R2["loc_draw"])
        print("  @ {:.1f}%: Y2 DSCR {:.2f}x | min cash ${:,.0f} | total LOC draw ${:,.0f}"
              .format(rr * 100, e2 / ds2, mc, loc))


if __name__ == "__main__":
    main()
