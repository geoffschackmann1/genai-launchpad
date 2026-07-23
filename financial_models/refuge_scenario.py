"""Refuge Hospice acquisition affordability solver.

Base case: NO SBA loan; $250,000 total liquid capital; startup/capex funded
elsewhere; Paloma panel migrates at close (engine base census, Tyler rates).
Seller terms modeled: Down at close + monthly payments (with optional 2-month
holiday) + January (month-6) balloon paid by Geoff's personal refinance.
Deferred balance accrues 6% simple; from month 7 the company services the refi
note (default 9% APR, 120-month amortization) - flagged assumption.
"""
import financial_models.engine.model as m

PRICE = 500_000
CASH = 250_000
SELLER_RATE = 0.06
REFI_RATE = 0.09
REFI_TERM = 120

_O = m.assumptions_dict
def _base():
    a = _O()
    a["sba_principal"] = 0.0000001
    a["equity"] = CASH
    a["capex"] = 0
    a["license_cost"] = 0
    return a


def engine_cfo():
    m.assumptions_dict = _base
    m.LICENSE_PAID_AT_CLOSE = False
    R, a = m.compute(scenario=1)
    return [R["cfo"][i] for i in range(12)], R, a


def pmt(P, annual, n):
    r = annual / 12
    return P * r / (1 - (1 + r) ** -n)


def run_structure(cfo, down, monthly, holiday=0, label=""):
    """Return (min_cash, min_month, balloon, interest_paid, path)."""
    cash = CASH - down
    bal = PRICE - down
    interest_accrued = 0.0
    path = []
    balloon = None
    refi_ds = 0.0
    minc, minm = cash, 0
    for i in range(12):
        mno = i + 1
        cash += cfo[i]
        interest_accrued += bal * SELLER_RATE / 12
        pay = 0.0
        if mno <= 6 and mno > holiday and monthly > 0:
            pay = min(monthly, bal)
            bal -= pay
            cash -= pay
        if mno == 6:
            balloon = bal + interest_accrued
            refi_ds = pmt(balloon, REFI_RATE, REFI_TERM) if balloon > 0 else 0
            bal = 0.0
        if mno > 6:
            cash -= refi_ds
        if cash < minc:
            minc, minm = cash, mno
        path.append(cash)
    return dict(label=label, down=down, monthly=monthly, holiday=holiday,
                min_cash=minc, min_month=minm, balloon=balloon,
                interest=interest_accrued, refi_ds=refi_ds,
                end_cash=path[-1], path=path)


def main():
    cfo, R, a = engine_cfo()
    print("Engine CFO (12 mo):", " ".join(f"{c/1000:,.0f}k" for c in cfo))
    print()
    structures = [
        run_structure(cfo, 245_000, 0, 0, "Seller ask: 49% down, balloon Jan"),
        run_structure(cfo, 106_000, 0, 0, "Max down, no monthlies"),
        run_structure(cfo, 100_000, 10_000, 0, "100K down + 10K/mo M1-6"),
        run_structure(cfo, 75_000, 15_000, 0, "75K down + 15K/mo M1-6"),
        run_structure(cfo, 100_000, 25_000, 2, "100K down + 25K/mo M3-6 (2-mo holiday)"),
        run_structure(cfo, 125_000, 25_000, 2, "125K down + 25K/mo M3-6 (2-mo holiday)"),
        run_structure(cfo, 150_000, 20_000, 2, "150K down + 20K/mo M3-6 (2-mo holiday)"),
    ]
    hdr = f"{'Structure':44} {'Down':>8} {'Paid pre-balloon':>16} {'Balloon(Jan)':>13} {'MinCash':>9} {'@M':>3} {'EndCash':>9} {'RefiDS/mo':>10}"
    print(hdr); print("-" * len(hdr))
    for s in structures:
        paid = s["down"] + (min(6 - s["holiday"], 6)) * s["monthly"]
        print(f"{s['label']:44} {s['down']:>8,.0f} {paid:>16,.0f} {s['balloon']:>13,.0f} "
              f"{s['min_cash']:>9,.0f} {s['min_month']:>3} {s['end_cash']:>9,.0f} {s['refi_ds']:>10,.0f}")
    print()
    print("Floor guidance: >=50k comfortable | 25-50k tight | <25k unsafe (no LOC in base case)")


if __name__ == "__main__":
    main()
