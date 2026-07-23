"""Refuge acquisition - operating-cost-refined base + sensitivity grid.

Interview-locked refinements vs the engine defaults:
  - G&A: rent $3,000 -> $2,000; credit-card fees ($1,004) cut  => -$2,004/mo
  - Health insurance: 90-day waiting period per hire ($550/head/mo saved until covered)
  - Med director ($4K/mo) contracted -> excluded from payroll inflation
Sensitivity axes: census capture (100/85/70/55%) x payroll inflation (+0/10/15%),
plus combined downside; mitigation lever = owner salary deferral (40% of $30K/mo, M1-6).
Cash: 1-month collection lag (expected) and 45-day AR (conservative) both reported.
Floor: $25K safe / <$25K flagged / <$0 fail.
Offer overlay: $100K down + 2-mo holiday + $25K x4 (M3-6) + Jan balloon (refi).
"""
import financial_models.engine.model as m
from financial_models.refuge_scenario import PRICE, CASH, SELLER_RATE, REFI_RATE, REFI_TERM, pmt

GA_SAVE = 2004.0            # rent -1000, credit-card fees -1004
HEALTH_PM = 550.0
HEALTH_WAIT = 3             # months (90-day waiting period)
OWNER_DEFER = 12_000.0      # 40% of the $30K/mo owner pool, months 1-6
N = 12

_O = m.assumptions_dict
def _base():
    a = _O()
    a["sba_principal"] = 0.0000001
    a["equity"] = CASH
    a["capex"] = 0
    a["license_cost"] = 0
    return a


def _health_uncovered(month):
    """Heads hired but still inside the 90-day health waiting period in `month` (1-12)."""
    n = 0
    for (_, _, sal, start, grp, _o) in m.ROSTER:
        if start <= month < start + HEALTH_WAIT:
            n += 1
    return n


def scenario(capture=1.0, payroll_mult=1.0, defer=False):
    """Return monthly dict of P&L + cash paths under a stress scenario."""
    m.assumptions_dict = _base
    m.LICENSE_PAID_AT_CLOSE = False
    R, a = m.compute(capture_rate=capture, scenario=1)
    net = [R["net"][i] for i in range(N)]
    out = dict(net=net, adc=[R["adc"][i] for i in range(N)])
    lic_am = PRICE / 15 / 12
    ebitda, opex = [], []
    for i in range(N):
        mo = i + 1
        med = R["med_dir"][i]
        labor = (R["ft_direct"][i] + R["prn"][i] + R["burden_d"][i] + R["health_d"][i]
                 + R["ft_indirect"][i] + R["rhonda"][i] + R["burden_i"][i] + R["health_i"][i])
        labor -= HEALTH_PM * _health_uncovered(mo)              # health waiting period
        labor *= payroll_mult                                    # payroll stress
        if defer and mo <= 6:
            labor = max(0.0, labor - OWNER_DEFER)                # owner deferral lever
        ga = R["ga_fixed"][i] - GA_SAVE                          # G&A refinements
        other = R["cogs_patient"][i] + R["qr"][i] + R["billing"][i]
        ox = labor + med + ga + other
        opex.append(ox)
        ebitda.append(net[i] - ox)
    out["ebitda"], out["opex"] = ebitda, opex

    # ---- offer overlay: 100K down, holiday 2, 25K x4, Jan balloon (refi) ----
    down, monthly, holiday = 100_000, 25_000, 2
    bal, acc = PRICE - down, 0.0
    seller_pay, seller_int = [], []
    balloon = 0.0
    for i in range(N):
        mo = i + 1
        if mo <= 6:
            si = bal * SELLER_RATE / 12; acc += si; seller_int.append(si)
            p = monthly if (holiday < mo <= 6) else 0.0
            bal -= p; seller_pay.append(p)
            if mo == 6:
                balloon = bal + acc; bal = 0.0
        else:
            seller_int.append(0.0); seller_pay.append(0.0)
    refi_p = pmt(balloon, REFI_RATE, REFI_TERM)
    out["balloon"], out["seller_pay"] = balloon, seller_pay

    # taxes / NI (display-consistent)
    rbal = balloon; refi_int = []
    for i in range(N):
        if i + 1 <= 6:
            refi_int.append(0.0)
        else:
            ri = rbal * REFI_RATE / 12; rbal -= (refi_p - ri); refi_int.append(ri)
    ni = []
    for i in range(N):
        pt = ebitda[i] - lic_am - seller_int[i] - refi_int[i]
        tx = net[i] * a["tx_tax"] if pt > 0 else 0.0
        ni.append(pt - tx)
    out["ni"] = ni

    # ---- cash path A: 1-month collection lag (expected) ----
    cashA, beg = [], CASH - down
    for i in range(N):
        coll = net[i - 1] if i >= 1 else 0.0
        c = beg + coll - opex[i] - seller_pay[i]
        if i + 1 > 6:
            c -= refi_p
        cashA.append(c); beg = c
    # ---- cash path B: 45-day AR (conservative) ----
    ar = [R["ar"][i] for i in range(N)]; ap = [R["ap"][i] for i in range(N)]
    d_ar = [ar[0]] + [ar[i] - ar[i - 1] for i in range(1, N)]
    d_ap = [ap[0]] + [ap[i] - ap[i - 1] for i in range(1, N)]
    # scale AR delta if capture changed (ar from engine already reflects capture)
    cashB, beg = [], CASH - down
    rbal = balloon
    for i in range(N):
        cfo = ni[i] + lic_am + seller_int[i] - d_ar[i] + d_ap[i]
        pr = 0.0
        if i + 1 > 6:
            ri = rbal * REFI_RATE / 12; pr = refi_p - ri; rbal -= pr
        c = beg + cfo - seller_pay[i] - pr
        cashB.append(c); beg = c
    out["cash_lag1"], out["cash_ar45"] = cashA, cashB
    out["min1"], out["min1_m"] = min(cashA), cashA.index(min(cashA)) + 1
    out["min45"], out["min45_m"] = min(cashB), cashB.index(min(cashB)) + 1
    return out


def breakeven_adc(payroll_mult=1.0):
    """Steady-state (M6 economics) ADC where EBITDA = 0, via capture bisection."""
    lo, hi = 0.1, 1.5
    for _ in range(40):
        mid = (lo + hi) / 2
        e = scenario(capture=mid, payroll_mult=payroll_mult)["ebitda"][5]
        if e > 0: hi = mid
        else: lo = mid
    s = scenario(capture=(lo + hi) / 2, payroll_mult=payroll_mult)
    return s["adc"][5]


def max_seller_monthly(capture=1.0, payroll_mult=1.0, defer=False, floor=25_000):
    """Max $X/mo (M3-6, after $100K down) keeping 1-mo-lag min cash >= floor."""
    lo, hi = 0.0, 80_000.0
    def minc(mon):
        # inline: rerun with custom monthly
        s = scenario(capture, payroll_mult, defer)
        # rebuild cash with custom monthly
        down, holiday = 100_000, 2
        bal, acc = PRICE - down, 0.0
        pays, balloon = [], 0.0
        for i in range(N):
            mo = i + 1
            if mo <= 6:
                acc += bal * SELLER_RATE / 12
                p = mon if (holiday < mo <= 6) else 0.0
                bal -= p; pays.append(p)
                if mo == 6: balloon = bal + acc; bal = 0.0
            else: pays.append(0.0)
        refi_p = pmt(balloon, REFI_RATE, REFI_TERM) if balloon > 0 else 0.0
        cash, beg = [], CASH - down
        for i in range(N):
            coll = s["net"][i - 1] if i >= 1 else 0.0
            c = beg + coll - s["opex"][i] - pays[i] - (refi_p if i + 1 > 6 else 0.0)
            cash.append(c); beg = c
        return min(cash)
    for _ in range(30):
        mid = (lo + hi) / 2
        if minc(mid) >= floor: lo = mid
        else: hi = mid
    return lo


GRID = [
    ("Base (100% capture)",            1.00, 1.00, False),
    ("Census 85%",                     0.85, 1.00, False),
    ("Census 70%",                     0.70, 1.00, False),
    ("Census 55%",                     0.55, 1.00, False),
    ("Payroll +10%",                   1.00, 1.10, False),
    ("Payroll +15%",                   1.00, 1.15, False),
    ("Census 85% + payroll +10%",      0.85, 1.10, False),
    ("COMBINED: census 70% + pay +10%", 0.70, 1.10, False),
    ("Combined + owner deferral",      0.70, 1.10, True),
]


def main():
    print("Refined base (G&A -$2,004/mo; health 90-day wait):")
    b = scenario()
    print("  M1-M6 EBITDA:", " ".join(f"{b['ebitda'][i]/1000:,.1f}k" for i in range(6)))
    print(f"  min cash (1-mo lag) ${b['min1']:,.0f} @M{b['min1_m']} | (45-day AR) ${b['min45']:,.0f} @M{b['min45_m']}")
    print()
    hdr = f"{'Scenario':34}{'6moEBITDA':>10}{'Min$ lag1':>11}{'@M':>3}{'Min$ ar45':>11}{'Verdict':>9}"
    print(hdr); print("-" * len(hdr))
    for label, c, p, dfr in GRID:
        s = scenario(c, p, dfr)
        e6 = sum(s["ebitda"][:6])
        v = "FAIL" if s["min1"] < 0 else ("flag" if s["min1"] < 25_000 else "safe")
        print(f"{label:34}{e6:>10,.0f}{s['min1']:>11,.0f}{s['min1_m']:>3}{s['min45']:>11,.0f}{v:>9}")
    print()
    print(f"Breakeven ADC (steady state): {breakeven_adc():.1f} | at payroll +10%: {breakeven_adc(1.10):.1f}")
    for label, c, p, dfr in [("Base", 1.0, 1.0, False), ("Census 85%", 0.85, 1.0, False),
                             ("Census 70%", 0.70, 1.0, False), ("Combined 70%/+10%", 0.70, 1.10, False),
                             ("Combined + deferral", 0.70, 1.10, True)]:
        mx = max_seller_monthly(c, p, dfr)
        print(f"Max seller monthly (floor $25K) - {label}: ${mx:,.0f}")


if __name__ == "__main__":
    main()
