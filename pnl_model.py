"""Franchise Model (admins only): a theoretical new centre's monthly P&L from opening, with and
without a campus manager, and the month its profit pays back the investment.

Members ramp from 0 to RAMP_MEMBERS at RAMP_WEEKS, then grow by GROWTH a month up to a ceiling
(or stay flat at the ceiling). Revenue = members x price per member x 4.2 weeks. Tutor wages
use the % for that many members (members bands); marketing is a % of revenue; rent, agency
fees and OPEX are fixed $. Royalties are the higher of the minimum or 8% of revenue, nothing
during the first relief weeks. The manager is hired the first month members reach the
trigger and stays. Profit for payback is EBITDA, less the owner's wage when that is on.
"""
import pandas as pd

import pnl
import pnl_forecast as pf

MANAGER = pnl.MANAGER
LINES = pnl.ROW_ORDER[1:]                       # cost lines, condensed P&L order
PCT = {l: f"{l} %" for l in LINES}
GRID = ["Members", "AMV $ / week"] + [PCT[l] for l in LINES]
OWNER = "Owner's wage"
FIXED = ["Rent", "Marketing - Agency Fees", "OPEX"]
WITH, WITHOUT = "With manager", "Without manager"

DEFAULTS = dict(target_aud=110000.0, payback_months=42, ramp_members=50.0, ramp_weeks=12.0,
                growth=7.0, manager_trigger=80.0, owner_on=False, owner_wage=60000.0,
                owner_stops=False, relief_weeks=18.0, royalty_min=2266.0, royalty_rise=3.0,
                royalty_fy=2026, price_rise=0.0, wage_rise=0.0, flat=False)

# Each country's royalties and investment target. royalty_min is in the country's own
# currency for the financial year starting in July of royalty_fy; target_aud is in AUD.
# A country not listed uses Australia's figures as a placeholder (flagged on the page).
COUNTRY_DEFAULTS = {
    "Australia": dict(royalty_min=2266.0, royalty_rise=3.0, royalty_fy=2026, relief_weeks=18.0,
                      target_aud=110000.0),
}
COUNTRY_KEYS = ["royalty_min", "royalty_rise", "royalty_fy", "relief_weeks", "target_aud"]


def settings_for(country):
    """(defaults for this country, True when they are Australia's placeholder figures)."""
    own = COUNTRY_DEFAULTS.get(country)
    return {**DEFAULTS, **(own or COUNTRY_DEFAULTS["Australia"])}, own is None


# ── Typical figures from actual data ─────────────────────────────────────────
def country_typical(cent, country):
    """Typical (median) figures for a country's centres, established ones where there are
    any. cent: centre_table rows (monthly averages). Returns a dict, or None if none."""
    c = cent[(cent["Country"].astype(str) == str(country)) & (cent["Total Revenue"] > 0)]
    if c.empty:
        return None
    est, _ = pnl.split_by_age(c)
    base = est if len(est) else c
    mem = base[base["Members"].fillna(0) > 0]
    payers = base[base[MANAGER] > 0]
    rates = pnl.tutor_benchmarks(base)
    return {
        "centres": len(base), "established": len(est) > 0,
        "members": float(mem["Members"].median()) if len(mem) else 0.0,
        "amv": float((mem["Total Revenue"] / mem["Members"]).median() / pnl.WEEKS_PER_MONTH)
        if len(mem) else 0.0,
        "marketing_pct": float((base["Marketing"] / base["Total Revenue"]).median()),
        "fixed": {l: float(base[l].median()) for l in FIXED},
        "manager": float(payers[MANAGER].median()) if len(payers) else 0.0,
        "manager_centres": len(payers),
        "rates": dict(zip(rates.index, rates["Use %"])),
        "fx": float(base["FX"].median()) if "FX" in base and base["FX"].notna().any() else 1.0,
        "peers": base,
    }


# ── The model ─────────────────────────────────────────────────────────────────
def months_from(start, n):
    return [pf.add_months(start, i) for i in range(n)]


def horizon(start, payback_months):
    """Months shown: from opening to the end of the financial year after the payback target."""
    end_fy = pf.fy_start(pf.add_months(start, payback_months - 1)) + 1
    n = 0
    while pf.fy_start(pf.add_months(start, n)) <= end_fy:
        n += 1
    return n


def ramp(month, s, ceiling):
    """Members in model month 1, 2, … (middle of the month)."""
    if s["flat"]:
        return ceiling
    weeks = (month - 0.5) * pnl.WEEKS_PER_MONTH
    if weeks <= s["ramp_weeks"]:
        m = s["ramp_members"] * weeks / s["ramp_weeks"]
    else:
        m = s["ramp_members"] + s["growth"] * (weeks - s["ramp_weeks"]) / pnl.WEEKS_PER_MONTH
    return min(m, max(ceiling, s["ramp_members"])) if ceiling else m


def _rises(period, start, pct):
    """1 + pct% for each July after opening (compounded)."""
    return (1 + pct / 100) ** (pf.fy_start(period) - pf.fy_start(start))


def royalty_min(period, s):
    return s["royalty_min"] * (1 + s["royalty_rise"] / 100) ** (pf.fy_start(period) -
                                                               s["royalty_fy"])


def relief_share(month, weeks):
    """Share of model month `month` inside the first `weeks` weeks (royalty relief)."""
    done = (month - 1) * pnl.WEEKS_PER_MONTH
    return min(max(weeks - done, 0.0), pnl.WEEKS_PER_MONTH) / pnl.WEEKS_PER_MONTH


def run(s, typ, start, n, manager=True, grid=None, base=None):
    """Monthly P&L for n months from `start`. grid/base: the adjustment grid as changed and
    before changes (any cell that differs overrides the model for that month)."""
    rows, hired, cum = [], False, 0.0
    for i, p in enumerate(months_from(start, n), start=1):
        def changed(col):
            return grid is not None and p in grid.index and \
                abs(float(grid.at[p, col]) - float(base.at[p, col])) > 1e-6
        mem = float(grid.at[p, "Members"]) if changed("Members") else \
            ramp(i, s, typ["members"])
        amv = float(grid.at[p, "AMV $ / week"]) if changed("AMV $ / week") else \
            typ["amv"] * _rises(p, start, s["price_rise"])
        rev = mem * amv * pnl.WEEKS_PER_MONTH
        wage = _rises(p, start, s["wage_rise"])
        hired = manager and (hired or mem >= s["manager_trigger"])
        r = {"Month": p, "Model month": i, "Members": mem, "AMV": amv, "Total Revenue": rev,
             pnl.TUTOR: pnl.rate_for(mem, typ["rates"]) * rev,
             "Marketing": typ["marketing_pct"] * rev,
             MANAGER: typ["manager"] * wage if hired else 0.0,
             "Royalties": (1 - relief_share(i, s["relief_weeks"])) *
             pnl.royalty(rev, royalty_min(p, s))}
        r.update({l: typ["fixed"][l] for l in FIXED})
        for l in LINES:
            if changed(PCT[l]) and (l != MANAGER or manager):
                r[l] = float(grid.at[p, PCT[l]]) / 100 * rev
        r["EBITDA"] = rev - sum(r[l] for l in LINES)
        r[pnl.GROSS_MARGIN] = rev - r[pnl.TUTOR]
        r[OWNER] = s["owner_wage"] / 12 * wage if s["owner_on"] and not \
            (s["owner_stops"] and hired) else 0.0
        r["Profit"] = r["EBITDA"] - r[OWNER]
        cum += r["Profit"]
        r["Cumulative"] = cum
        rows.append(r)
    return pd.DataFrame(rows).set_index("Month")


def grid_from(table):
    """The adjustment grid (members, price and each line's % of revenue) for these months."""
    g = pd.DataFrame(index=table.index)
    g["Members"] = table["Members"]
    g["AMV $ / week"] = table["AMV"]
    rev = table["Total Revenue"].where(table["Total Revenue"] > 0)
    for l in LINES:
        g[PCT[l]] = (table[l] / rev * 100).fillna(0.0)
    return g


def payback_month(table, target):
    """First model month cumulative profit reaches the target (None if never)."""
    hit = table[table["Cumulative"] >= target]
    return int(hit["Model month"].iloc[0]) if len(hit) else None


def needed(s, typ, start, target, manager, what):
    """Members ceiling ('members') or price per member ('amv') needed to pay back within
    s['payback_months'], everything else as set. None if not reachable (or not needed)."""
    def ok(v):
        t = run(s, {**typ, what: v}, start, s["payback_months"], manager)
        return payback_month(t, target) is not None
    lo, hi = 0.0, max(typ[what], 1.0)
    while not ok(hi):
        hi *= 2
        if hi > 1e5:
            return None
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if ok(mid) else (mid, hi)
    return hi


def fy_totals(table):
    """One row per financial year: months, amounts added up and cumulative profit at the end."""
    t = table.assign(FY=[pf.fy_start(p) for p in table.index])
    amounts = ["Total Revenue", pnl.GROSS_MARGIN, "EBITDA", OWNER, "Profit"]
    out = t.groupby("FY")[amounts].sum()
    out["Months"] = t.groupby("FY").size()
    out["Members (end)"] = t.groupby("FY")["Members"].last()
    out["Cumulative"] = t.groupby("FY")["Cumulative"].last()
    return out


# ── Actual against the model ─────────────────────────────────────────────────
AGE_BANDS = [(1, 6, "Months 1–6"), (7, 12, "Months 7–12"), (13, 24, "Year 2"),
             (25, 36, "Year 3"), (37, 48, "Year 4"), (49, 10 ** 6, "Year 5+")]


def band_of(month):
    return next(lbl for lo, hi, lbl in AGE_BANDS if lo <= month <= hi)


def actual_by_age(lm, members, weeks_old, country_of, country, as_at, young_weeks=18):
    """Each centre-month of the country's P&L with the centre's age then (model month).
    lm: pnl.location_months; members: (Location, Period) series; weeks_old: {location:
    weeks old at as_at}. Centres young_weeks old or younger at that month are left out."""
    d = lm.reset_index()
    d = d[d["Location"].map(country_of) == country]
    d["Weeks"] = [weeks_old.get(l, float("nan")) -
                  (as_at - (pd.Period(p, "M").end_time)).days / 7
                  for l, p in zip(d["Location"], d["Period"])]
    d = d[d["Weeks"] > young_weeks]
    d["Model month"] = (d["Weeks"] / pnl.WEEKS_PER_MONTH).apply(lambda w: int(w) + 1)
    d["Members"] = [members.get((l, p)) if len(members) else None
                    for l, p in zip(d["Location"], d["Period"])]
    return d


def compare(actual, model):
    """Typical actual (median) against the model's average, by centre age band."""
    rows = []
    a = actual.assign(Band=actual["Model month"].map(band_of))
    m = model.assign(Band=model["Model month"].map(band_of))
    for _, _, band in AGE_BANDS:
        g, mm = a[a["Band"] == band], m[m["Band"] == band]
        if g.empty and mm.empty:
            continue
        rev = g["Total Revenue"]
        pct = lambda col, df: float((df[col] / df["Total Revenue"]).median() * 100) \
            if len(df) and (df["Total Revenue"] > 0).any() else None
        mpct = lambda col: float(mm[col].sum() / mm["Total Revenue"].sum() * 100) \
            if len(mm) and mm["Total Revenue"].sum() > 0 else None
        rows.append({"Centre age": band, "Centres": g["Location"].nunique(),
                     "Members (actual)": g["Members"].median() if len(g) else None,
                     "Members (model)": mm["Members"].mean() if len(mm) else None,
                     "Revenue / month (actual)": rev.median() if len(g) else None,
                     "Revenue / month (model)": mm["Total Revenue"].mean() if len(mm) else None,
                     "Gross margin % (actual)": pct(pnl.GROSS_MARGIN, g),
                     "Gross margin % (model)": mpct(pnl.GROSS_MARGIN),
                     "EBITDA % (actual)": pct("EBITDA", g),
                     "EBITDA % (model)": mpct("EBITDA")})
    return pd.DataFrame(rows)
