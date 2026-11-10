"""Franchise Model (admins only): a theoretical new centre's monthly P&L from opening, with and
without a campus manager, and the month its profit pays back the investment.

Members are worked out month by month from opening: about 17.5 a month (50 members over 12
weeks) until 50, then +7 a month with no limit. In a term-holiday month members change by that
month's holiday % instead. Revenue = members x price per member x 4.2 weeks. Tutor wages
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
                royalty_fy=2026, price_rise=0.0, wage_rise=0.0)
MONTHS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
FALLBACK_HOLIDAYS = ["01", "04", "07", "10", "12"]   # when the School Holidays tab has none
HOLIDAY_STATES = {"Australia": ["NSW", "VIC", "QLD", "WA", "SA"], "New Zealand": ["NZ"]}

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


def next_members(end, period, s, reached=False):
    """Members at the end of `period`, from `end` at the end of the month before. reached:
    the members at week 12 have been reached before (from then on it's +growth a month)."""
    pct = s.get("holiday", {}).get(period[5:7], 0.0)
    if pct:
        return max(end * (1 + pct / 100), 0.0)
    if not reached and end < s["ramp_members"]:
        step = s["ramp_members"] * pnl.WEEKS_PER_MONTH / s["ramp_weeks"]
        return min(end + step, s["ramp_members"])
    return end + s["growth"]


# ── Term-holiday months and their typical change ─────────────────────────────
def holiday_months(holidays, country):
    """Calendar months ("01".."12") with term holidays for the country: a month counts when
    most of the country's state-years in the School Holidays tab have a holiday week in it.
    Returns (months, True when the fallback list was used)."""
    states = HOLIDAY_STATES.get(country, [])
    if holidays is None or holidays.empty or not states or "State" not in holidays:
        return FALLBACK_HOLIDAYS, True
    h = holidays[holidays["State"].isin(states)]
    seen = {}
    for _, r in h.iterrows():
        y, a, b = int(r["Year"]), int(r["Start_Week"]), int(r["End_Week"])
        weeks = [(y, w) for w in range(a, b + 1)] if b >= a else \
            [(y, w) for w in range(a, 53)] + [(y + 1, w) for w in range(1, b + 1)]
        for yy, w in weeks:
            try:
                m = pd.Timestamp.fromisocalendar(yy, w, 4).strftime("%m")
            except ValueError:
                continue
            seen.setdefault((r["State"], y), set()).add(m)
    if not seen:
        return FALLBACK_HOLIDAYS, True
    return [m for m in MONTHS if sum(m in v for v in seen.values()) / len(seen) > 0.5], False


def holiday_changes(members, locations, months, min_members=pf.NETWORK_MIN_MEMBERS):
    """{month: typical (median) % change in members} for these calendar months, from these
    centres' history (centres with fewer than min_members the month before left out), and
    the number of centres behind it. members: (Location, Period) series."""
    parts, used = [], set()
    for loc, ser in members.groupby(level=0):
        if loc not in locations:
            continue
        ser = ser.droplevel(0)
        ch = pf._changes(ser)
        prev = ser.reindex([pf.add_months(p, -1) for p in ch.index]).values
        ch = ch[(prev >= min_members) & (ch.abs() <= 0.5)]
        if len(ch):
            parts.append(ch)
            used.add(loc)
    if not parts:
        return {}, 0
    allc = pd.concat(parts)
    by = allc.groupby(allc.index.str[5:7])
    return {m: float(by.get_group(m).median() * 100) for m in months
            if m in by.groups and len(by.get_group(m)) >= pf.NETWORK_MIN_OBS}, len(used)


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
    rows, hired, cum, end, reached = [], False, 0.0, 0.0, False
    for i, p in enumerate(months_from(start, n), start=1):
        def changed(col):
            return grid is not None and p in grid.index and \
                abs(float(grid.at[p, col]) - float(base.at[p, col])) > 1e-6
        new = float(grid.at[p, "Members"]) if changed("Members") else \
            next_members(end, p, s, reached)            # members at the end of the month
        mem, end = (end + new) / 2, new                 # revenue uses the month's average
        reached = reached or end >= s["ramp_members"]
        amv = float(grid.at[p, "AMV $ / week"]) if changed("AMV $ / week") else \
            typ["amv"] * _rises(p, start, s["price_rise"])
        rev = mem * amv * pnl.WEEKS_PER_MONTH
        wage = _rises(p, start, s["wage_rise"])
        hired = manager and (hired or mem >= s["manager_trigger"])
        r = {"Month": p, "Model month": i, "Members": mem, "Members end": end, "AMV": amv,
             "Total Revenue": rev,
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
    g["Members"] = table["Members end"]              # end of month, as in the table
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
    """New members a month ('growth') or price per member ('amv') needed to pay back within
    s['payback_months'], everything else as set. None if not reachable."""
    def ok(v):
        s2, t2 = ({**s, what: v}, typ) if what in s else (s, {**typ, what: v})
        t = run(s2, t2, start, s["payback_months"], manager)
        return payback_month(t, target) is not None
    lo, hi = 0.0, max(s[what] if what in s else typ[what], 1.0)
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
    out["Members (end)"] = t.groupby("FY")["Members end"].last()
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
