"""Franchise Model (admins only): a theoretical new centre's monthly P&L over its first 5 years
(60 months), with and without a campus manager, and the month its profit pays back the
investment.

Members on the books grow about 14 a month (40 members over 12 weeks) until 40. After that,
each month new sign-ups join (10 a month on average) and a share of members leave (8.5% a
month), so a centre levels off where leavers equal sign-ups (about 118). The best case grows to
yearly targets instead (80 by the end of Year 1, then +30, +20, +10, +5). In a term-holiday
month sign-ups drop by the holiday drop % (85%) and the other months make up for it, so a
year still averages 10 a month; an extra share of members suspends for that month, paying the
suspension fee ($5 a week) instead of the full price, and comes back the next month.
Price: the first 50 members (foundation members) pay tier prices ($50 for the first 20, $60
for the next 20, $65 for the next 10) for as long as they stay (8.5% leave a month); everyone
else pays the standard price (the state RRP for the typical package mix). These are list
prices including GST: the model takes off GST and applies the collected % (95% to start,
for every member). Prices rise each July from the first July at least
12 months after opening; rent, OPEX and wages rise each July after opening. No rent is
paid in the rent-free months at the start (2).
Revenue = tuition (active members x price x 4.2 weeks, billed from the billing week) +
suspension fees + joining fees ($100 for each sign-up, from week 1). Marketing is advertising a day (presale amount during the presale campaign, then the usual amount); the
agency fee is the presale fee, then the typical fee but at least the minimum. Tutor wages use
the % for that many members (members bands) of total revenue; rent, agency fees and OPEX are
fixed $. Royalties are the higher of the minimum or 8% of total revenue, nothing during the
first relief weeks. The manager is hired the first month members reach the trigger and stays.
Profit for payback is EBITDA, less the owner's wage when that is on.
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

DEFAULTS = dict(target_aud=110000.0, payback_months=42, ramp_members=40.0, ramp_weeks=12.0,
                leave_pct=8.5, growth_mode="churn", targets=(80.0, 40.0, 30.0, 15.0, 10.0),
                worst_signups=8.0, worst_leave=10.0,
                founder_sizes=(20.0, 20.0, 10.0), founder_prices=(50.0, 60.0, 65.0),
                founder_leave=4.0, standard_price=70.0, gst=10.0, collected=95.0,
                manager_trigger=80.0, manager_salary=32000.0, super_pct=12.0,
                owner_on=False, owner_wage=60000.0,
                owner_stops=False, relief_weeks=8.0, royalty_min=2266.0, royalty_rise=3.0,
                rent=3936.0, opex=2066.0,      # Australian established centres' medians
                royalty_fy=2026, price_rise=3.0, wage_rise=3.0, cost_rise=3.0, billing_week=9.0,
                rent_free_months=2.0,
                # tutor wages + super % of revenue at these member numbers (straight lines
                # between them, flat beyond the ends)
                tutor_curve=((10.0, 50.0), (30.0, 40.0), (50.0, 36.0), (80.0, 33.0),
                             (100.0, 30.0), (125.0, 27.0), (150.0, 25.0)),
                presale_weeks=12.0, presale_agency=1200.0, presale_ads_day=100.0,
                agency_min=250.0, ads_day=30.0, signups=10.0, joining_fee=100.0,
                suspension_fee=5.0, holiday_drop=85.0)
MODEL_MONTHS = 60                                    # 5 years from opening
MONTHS = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
FALLBACK_HOLIDAYS = ["01", "04", "07", "10", "12"]   # when the School Holidays tab has none
# Package mix when there is no Membership Composition history (HQ's ideal composition).
FALLBACK_MIX = {("Group", 1): 15.0, ("Group", 2): 80.0, ("Group", 3): 1.0, ("1:1", 1): 4.0}
# Extra % of members suspended in holiday months when there is no suspension history (typical
# Australian centre, Jul 2022 – Nov 2026, above the usual ~9%).
FALLBACK_SUSPENDED = {"01": 17.0, "04": 9.0, "07": 5.0, "10": 2.5, "12": 15.0}
HOLIDAY_STATES = {"Australia": ["NSW", "VIC", "QLD", "WA", "SA"], "New Zealand": ["NZ"]}

# Each country's royalties and investment target. royalty_min is in the country's own
# currency for the financial year starting in July of royalty_fy; target_aud is in AUD.
# A country not listed uses Australia's figures as a placeholder (flagged on the page).
COUNTRY_DEFAULTS = {
    "Australia": dict(royalty_min=2266.0, royalty_rise=3.0, royalty_fy=2026, relief_weeks=8.0,
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


def holiday_keep(period, s):
    """Share of the usual new members who join this month: fewer in a holiday month."""
    return 1 - s["holiday_drop"] / 100 if period[5:7] in s.get("holiday_months", ()) else 1.0


def signup_factor(period, s):
    """Sign-ups this month as a share of the average month. Holiday months get fewer and the
    other months make up for it, so a year still adds up to 12 average months (the average
    from actual data already includes the holidays)."""
    lost = len(set(s.get("holiday_months", ()))) * s["holiday_drop"] / 100
    return (12 / (12 - lost) if lost < 12 else 1.0) * holiday_keep(period, s)


def year_target(month, s):
    """Best case: members on the books at the end of the model year that holds `month`."""
    y = min((month - 1) // 12, len(s["targets"]) - 1)
    return float(sum(s["targets"][:y + 1]))


def month_members(books, period, s, reached=False, month=1):
    """(members on the books at the end of `period`, new sign-ups that month), from `books` at
    the end of the month before. reached: the week-12 members have been reached before.
    month: model month (1 = opening)."""
    keep = holiday_keep(period, s)
    if not reached and books < s["ramp_members"]:
        step = s["ramp_members"] * pnl.WEEKS_PER_MONTH / s["ramp_weeks"] * keep
        end = min(books + step, s["ramp_members"])
        return end, end - books                          # everyone joining is a sign-up
    leavers = s["leave_pct"] / 100 * books
    if s.get("growth_mode") == "targets":                # grow to the year's target
        left = 12 - (month - 1) % 12
        add = max(year_target(month, s) - books, 0.0) / left * (keep if left > 1 else 1.0)
        return books + add, add + leavers                # sign-ups also replace leavers
    joins = s["signups"] * signup_factor(period, s)
    return max(books + joins - leavers, 0.0), joins


def founders_after(remaining, joined, joins, s, leaving):
    """Foundation members left in each price tier at the end of the month. remaining: per tier
    at the start; joined: members who had ever joined before this month; joins: sign-ups
    this month (the first ones to join fill the tiers in order). leaving: founders leave this
    month (after the first 12 weeks)."""
    out, lo = [], 0.0
    stay = 1 - s["founder_leave"] / 100 if leaving else 1.0
    for r, size in zip(remaining, s["founder_sizes"]):
        hi = lo + size
        out.append(r * stay + max(0.0, min(joined + joins, hi) - max(joined, lo)))
        lo = hi
    return out


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


def suspension_extra(wm, locations, months, min_active=20):
    """{month: extra % of members suspended} in these holiday months, above the usual level
    (the middle of the other months), from weekly membership history, and the number of
    centres behind it. Each week takes the typical (median) centre; centres with fewer than
    min_active active members that week are left out."""
    need = {"Success Tutoring - Business name", "Date", "# Active members", "# Suspended members"}
    if wm is None or wm.empty or not need <= set(wm.columns):
        return {}, 0
    d = wm[wm["Success Tutoring - Business name"].isin(locations)]
    act = pd.to_numeric(d["# Active members"], errors="coerce")
    sus = pd.to_numeric(d["# Suspended members"], errors="coerce").fillna(0)
    d = d.assign(share=sus / (act + sus) * 100)[act >= min_active].dropna(subset=["share"])
    if d.empty:
        return {}, 0
    weekly = d.groupby("Date")["share"].median()
    by_month = weekly.groupby(weekly.index.strftime("%m")).mean()
    usual = by_month.drop([m for m in months if m in by_month.index]).median()
    if pd.isna(usual):
        return {}, 0
    return ({m: round(max(float(by_month[m]) - float(usual), 0.0), 1) for m in months
             if m in by_month.index},
            d["Success Tutoring - Business name"].nunique())


def _rises(period, start, pct):
    """1 + pct% for each July after opening (compounded)."""
    return (1 + pct / 100) ** (pf.fy_start(period) - pf.fy_start(start))


def price_rises(period, start, pct):
    """1 + pct% for each July from the first July at least 12 months after opening."""
    first = pf.add_months(start, 12)
    n = pf.fy_start(period) - pf.fy_start(first) + (1 if first[5:7] == "07" else 0)
    return (1 + pct / 100) ** max(n, 0)


def collected_share(s):
    """What a $1 list price (GST included) brings in: less GST, times the collected %."""
    return 1 / (1 + s["gst"] / 100) * s["collected"] / 100


def joining_fee_banked(s):
    """Joining fee per sign-up without GST (the fee is set including GST)."""
    return s["joining_fee"] / (1 + s["gst"] / 100)


def list_price(founders, books, s):
    """Average list price a week across members on the books: foundation members at their
    tier price, everyone else at the standard price."""
    if books <= 0:
        return s["founder_prices"][0] if s["founder_prices"] else s["standard_price"]
    f = sum(founders)
    scale = min(1.0, books / f) if f else 1.0
    paid = sum(r * scale * pr for r, pr in zip(founders, s["founder_prices"]))
    return (paid + max(books - f * scale, 0.0) * s["standard_price"]) / books


def plan_mix(comp, country="Australia"):
    """{(type, sessions a week): % of members} from Membership Composition uploads (prepared),
    all weeks, freeze prices left out, and the number of centres behind it. Falls back to
    HQ's ideal composition."""
    import membership_composition as mc
    cc = mc.country_code(country)
    if comp is None or comp.empty:
        return dict(FALLBACK_MIX), 0
    c = comp[(comp["Country"] == cc) & ~comp["Freeze"]]
    if c.empty:
        return dict(FALLBACK_MIX), 0
    g = c.groupby(["Type", "Sessions/week"])["Members"].sum()
    return {k: float(v / g.sum() * 100) for k, v in g.items()}, int(c["Location"].nunique())


def rrp_mix_price(rrp, mix):
    """List price a week for the package mix at this RRP (GST included)."""
    import membership_composition as mc
    if not rrp:
        return None
    parts = [(w, mc.rrp_price(rrp, kind, int(n))) for (kind, n), w in mix.items()]
    parts = [(w, v) for w, v in parts if v is not None]
    tot = sum(w for w, _ in parts)
    return sum(w * v for w, v in parts) / tot if tot else None


def collected_pct(peers, price_of_region, s):
    """Collected % (the typical centre): what centres bank per member a week, without joining
    fees, against their state's RRP for the package mix less GST. price_of_region(region)
    gives the RRP mix price. Returns (%, centres) or (None, 0)."""
    vals = []
    for _, r in peers.iterrows():
        mem, rev = r.get("Members"), r.get("Total Revenue")
        rrp = price_of_region(r.get("Region"))
        if not mem or pd.isna(mem) or mem <= 0 or not rrp or not rev:
            continue
        banked = rev / mem / pnl.WEEKS_PER_MONTH - \
            s["signups"] * joining_fee_banked(s) / (mem * pnl.WEEKS_PER_MONTH)
        vals.append(banked / (rrp / (1 + s["gst"] / 100)) * 100)
    return (float(pd.Series(vals).median()), len(vals)) if vals else (None, 0)


def manager_monthly(s):
    """Manager cost a month: salary a year plus super."""
    return s["manager_salary"] * (1 + s["super_pct"] / 100) / 12


def tutor_rate(members, s):
    """Tutor wages + super as a share of revenue (0–1) for this many members: the tutor curve,
    blended in a straight line between its points and flat beyond the first and last."""
    pts = sorted(s["tutor_curve"])
    if not pts:
        return 0.0
    if members <= pts[0][0]:
        return pts[0][1] / 100
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        if members <= x1:
            return (y0 + (y1 - y0) * (members - x0) / (x1 - x0)) / 100 if x1 > x0 else y1 / 100
    return pts[-1][1] / 100


def rent_free_share(month, s):
    """Share of model month `month` that is rent free (the first rent-free months)."""
    return min(max(s.get("rent_free_months", 0.0) - (month - 1), 0.0), 1.0)


def royalty_min(period, s):
    return s["royalty_min"] * (1 + s["royalty_rise"] / 100) ** (pf.fy_start(period) -
                                                               s["royalty_fy"])


def weeks_share(month, start_week, end_week):
    """Share of model month `month` (weeks (month-1)*4.2 to month*4.2 since opening) that falls
    between start_week and end_week weeks since opening."""
    lo, hi = (month - 1) * pnl.WEEKS_PER_MONTH, month * pnl.WEEKS_PER_MONTH
    return max(0.0, min(hi, end_week) - max(lo, start_week)) / pnl.WEEKS_PER_MONTH


def billed_share(month, s):
    """Share of the month billed: billing starts at the beginning of billing week."""
    return weeks_share(month, s["billing_week"] - 1, float("inf"))


def presale_share(month, s):
    return weeks_share(month, 0.0, s["presale_weeks"])


def relief_share(month, weeks):
    """Share of model month `month` inside the first `weeks` weeks (royalty relief)."""
    done = (month - 1) * pnl.WEEKS_PER_MONTH
    return min(max(weeks - done, 0.0), pnl.WEEKS_PER_MONTH) / pnl.WEEKS_PER_MONTH


def run(s, typ, start, n, manager=True, grid=None, base=None):
    """Monthly P&L for n months from `start`. grid/base: the adjustment grid as changed and
    before changes (any cell that differs overrides the model for that month)."""
    rows, hired, cum, books, reached = [], False, 0.0, 0.0, False
    founders, joined = [0.0] * len(s["founder_sizes"]), 0.0
    for i, p in enumerate(months_from(start, n), start=1):
        def changed(col):
            return grid is not None and p in grid.index and \
                abs(float(grid.at[p, col]) - float(base.at[p, col])) > 1e-6
        new, joins = month_members(books, p, s, reached, i)   # on the books, end of month
        if changed("Members"):
            new = float(grid.at[p, "Members"])
            if not reached:
                joins = max(new - books, 0.0)
        founders = founders_after(founders, joined, joins, s, leaving=reached)
        joined += joins
        prev, was_reached = books, reached
        avg, books = (books + new) / 2, new                   # revenue uses the month's average
        reached = reached or books >= s["ramp_members"]
        sus_pct = s.get("holiday", {}).get(p[5:7], 0.0) / 100   # extra suspended this month
        sus = avg * sus_pct
        mem = avg - sus                                       # active (paying full price)
        listed = list_price(founders, books, s)               # GST included
        amv = float(grid.at[p, "AMV $ / week"]) if changed("AMV $ / week") else \
            listed * collected_share(s) * price_rises(p, start, s["price_rise"])
        bill = pnl.WEEKS_PER_MONTH * billed_share(i, s)       # billed from billing week
        tuition, sus_fees = mem * amv * bill, sus * s["suspension_fee"] * bill
        joining = joins * joining_fee_banked(s)               # from week 1, presale too
        rev = tuition + sus_fees + joining
        pre = presale_share(i, s)
        wage = _rises(p, start, s["wage_rise"])
        hired = manager and (hired or avg >= s["manager_trigger"])
        r = {"Month": p, "Model month": i, "Members": mem, "Members end": books * (1 - sus_pct),
             "On books": books, "Suspended": sus, "Sign-ups": joins,
             "Leavers": max(prev + joins - books, 0.0) if was_reached else 0.0,
             "Founders": min(sum(founders), books), "List price": listed, "AMV": amv,
             "Tuition": tuition, "Suspension fees": sus_fees, "Joining fees": joining,
             "Total Revenue": rev,
             # % of total revenue, but none before sessions start (presale)
             pnl.TUTOR: tutor_rate(mem, s) * rev if bill > 0 else 0.0,
             # advertising a day: the presale amount, then the usual amount
             "Marketing": (s["presale_ads_day"] * pre + s["ads_day"] * (1 - pre)) * 7 *
             pnl.WEEKS_PER_MONTH,
             MANAGER: typ["manager"] * wage if hired else 0.0,
             "Royalties due": pnl.royalty(rev, royalty_min(p, s)),   # with no relief
             "Royalties": (1 - relief_share(i, s["relief_weeks"])) *
             pnl.royalty(rev, royalty_min(p, s))}
        r.update({l: typ["fixed"][l] for l in FIXED})
        for l in ("Rent", "OPEX"):                            # rise each July after opening
            r[l] = typ["fixed"][l] * _rises(p, start, s["cost_rise"])
        r["Rent"] *= 1 - rent_free_share(i, s)                # rent-free months at the start
        r["Marketing - Agency Fees"] = s["presale_agency"] * pre + max(       # then at least
            typ["fixed"]["Marketing - Agency Fees"], s["agency_min"]) * (1 - pre)  # the minimum
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
    g["Members"] = table["On books"]                 # members on the books, end of month
    g["AMV $ / week"] = table["AMV"]
    rev = table["Total Revenue"].where(table["Total Revenue"] > 0)
    for l in LINES:
        g[PCT[l]] = (table[l] / rev * 100).fillna(0.0)
    return g


def payback_month(table, target):
    """First model month cumulative profit reaches the target (None if never)."""
    hit = table[table["Cumulative"] >= target]
    return int(hit["Model month"].iloc[0]) if len(hit) else None


def _with(s, what, v):
    """Settings with `what` set to v. 'price': the standard price, with the foundation tier
    prices moved in proportion."""
    if what == "price":
        k = v / s["standard_price"] if s["standard_price"] else 1.0
        return {**s, "standard_price": v,
                "founder_prices": tuple(x * k for x in s["founder_prices"])}
    return {**s, what: v}


def needed(s, typ, start, target, manager, what):
    """Sign-ups a month ('signups') or standard list price ('price', foundation prices moved in
    proportion) needed to pay back within s['payback_months'], everything else as set.
    None if not reachable."""
    def ok(v):
        t = run(_with(s, what, v), typ, start, s["payback_months"], manager)
        return payback_month(t, target) is not None
    lo, hi = 0.0, max(s["standard_price"] if what == "price" else s[what], 1.0)
    while not ok(hi):
        hi *= 2
        if hi > 1e5:
            return None
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if ok(mid) else (mid, hi)
    return hi


YEAR_AMOUNTS = ["Total Revenue", pnl.GROSS_MARGIN, "Royalties", "EBITDA", OWNER, "Profit"]


def year_totals(table, fy=False):
    """One row per year: Year 1–5 from opening (months 1–12, 13–24 …), or financial years
    (fy=True). Amounts added up, members on the books and cumulative profit at the end."""
    key = [pf.fy_start(p) for p in table.index] if fy else \
        ((table["Model month"] - 1) // 12 + 1).tolist()
    g = table.assign(Year=key).groupby("Year")
    out = g[YEAR_AMOUNTS].sum()
    out["Months"] = g.size()
    out["Members (end)"] = g["On books"].last()
    out["Cumulative"] = g["Cumulative"].last()
    return out


def fy_totals(table):
    return year_totals(table, fy=True)


def breakeven_month(table):
    """First model month with a profit (above $0), or None."""
    hit = table[table["Profit"] > 0]
    return int(hit["Model month"].iloc[0]) if len(hit) else None


def peak_cash(table):
    """(the most the centre is behind, as a positive $, and the model month it happens): the
    lowest point of profit added up since opening. (0, None) if it is never behind."""
    low = table["Cumulative"].min()
    if low >= 0:
        return 0.0, None
    return float(-low), int(table.loc[table["Cumulative"].idxmin(), "Model month"])


def royalty_start(table):
    """Best month to start paying royalties: the first model month from which profit after
    paying full royalties stays at or above $0 every month to the end. None if never."""
    after = table["Profit"] + table["Royalties"] - table["Royalties due"]
    ok = (after >= -0.005)[::-1].cummin()[::-1]          # True from here to the end
    hit = table[ok]
    return int(hit["Model month"].iloc[0]) if len(hit) else None


def summary(table, target):
    """The headline results for one scenario."""
    cash, cash_month = peak_cash(table)
    return {"Payback month": payback_month(table, target),
            "Break-even month": breakeven_month(table),
            "Peak cash needed": cash, "Peak cash month": cash_month,
            "Royalty start month": royalty_start(table),
            "5-year profit": float(table["Profit"].sum())}


def _week_month(weeks):
    """Model month that holds the end of this many weeks since opening."""
    return max(int(-(-weeks // pnl.WEEKS_PER_MONTH)), 1)


def milestones(table, s, target):
    """The road to profitability for one scenario: [(model month or None, label)] in order,
    events that don't happen within the months shown last (None)."""
    def first(mask):
        hit = table[mask]
        return int(hit["Model month"].iloc[0]) if len(hit) else None
    cash, cash_month = peak_cash(table)
    events = [
        (int((s["billing_week"] - 1) // pnl.WEEKS_PER_MONTH) + 1,      # start of that week
         "Sessions and billing start"),
        (_week_month(s["presale_weeks"]), "Presale campaign ends"),
        (first(table["On books"] >= s["ramp_members"] - 1e-6),
         f"{s['ramp_members']:g} members"),
        (_week_month(s["relief_weeks"]) if s["relief_weeks"] else None, "Royalty relief ends"),
        (int(s.get("rent_free_months", 0.0)) + 1 if s.get("rent_free_months") else None,
         "Rent starts"),
        (breakeven_month(table), "First month in profit"),
        (cash_month, "Lowest cash point"),
        (royalty_start(table), "Can afford royalties every month"),
        (first(table[MANAGER] > 0), "Manager hired"),
        (payback_month(table, target), "Investment paid back"),
    ]
    if not (table[MANAGER] > 0).any():
        events = [e for e in events if e[1] != "Manager hired"]
    return sorted(events, key=lambda e: (e[0] is None, e[0] or 0))


def opening_months(s, typ, start, target, count=12, n=MODEL_MONTHS):
    """Payback month and 5-year profit for each of `count` opening months from `start`,
    both scenarios."""
    rows = []
    for o in months_from(start, count):
        r = {"Opening month": o}
        for name in (WITHOUT, WITH):
            t = run(s, typ, o, n, name == WITH)
            r[f"Payback ({name.lower()})"] = payback_month(t, target)
            r[f"5-year profit ({name.lower()})"] = float(t["Profit"].sum())
        rows.append(r)
    return pd.DataFrame(rows)


def _results(label, s, typ, start, target, n, extra=None):
    r = {"Case": label, **(extra or {})}
    for name in (WITHOUT, WITH):
        t = run(s, typ, start, n, name == WITH)
        r[f"Payback ({name.lower()})"] = payback_month(t, target)
        r[f"5-year profit ({name.lower()})"] = float(t["Profit"].sum())
        r[f"Members after 5 years ({name.lower()})"] = float(t["On books"].iloc[-1])
    return r


def cases(s, typ, start, target, n=MODEL_MONTHS):
    """Worst case (fewer sign-ups, more leaving: like an underperforming centre), base case
    (as set) and best case (growing to the yearly targets)."""
    worst = {**s, "growth_mode": "churn", "signups": s["worst_signups"],
             "leave_pct": s["worst_leave"]}
    best = {**s, "growth_mode": "targets"}
    t = s["targets"]
    return pd.DataFrame([
        _results("Worst case", worst, typ, start, target, n,
                 {"Growth": f"{s['worst_signups']:g} sign-ups a month, "
                            f"{s['worst_leave']:g}% leave"}),
        _results("Base case", s, typ, start, target, n,
                 {"Growth": f"{s['signups']:g} sign-ups a month, {s['leave_pct']:g}% leave"}),
        _results("Best case", best, typ, start, target, n,
                 {"Growth": f"{t[0]:g} by the end of Year 1, then "
                            + ", ".join(f"+{x:g}" for x in t[1:])})])


def price_test(s, typ, start, target, step=5.0, n=MODEL_MONTHS):
    """Foundation prices all moved down or up by `step` (list $ a week), everything else as
    set: payback and 5-year profit."""
    rows = []
    for label, k in ((f"${step:g} lower", -1), ("As set", 0), (f"${step:g} higher", 1)):
        s2 = {**s, "founder_prices": tuple(max(x + k * step, 0.0) for x in s["founder_prices"])}
        rows.append(_results(label, s2, typ, start, target, n,
                             {"Foundation prices": " / ".join(f"${x:g}"
                                                              for x in s2["founder_prices"])}))
    return pd.DataFrame(rows)


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
