"""Rolling membership trend: each centre's growth over the last few weeks against the
network (or its state), grouped into five categories."""
import pandas as pd

LOC = "Success Tutoring - Business name"
WINDOW = 4                       # weeks in each window: last 4 vs the 4 before
MIN_WEEKS_TRADING = 18           # newer centres are listed separately
SMALL_CENTRE = 20                # fewer members than this: % swings a lot
SPARK_WEEKS = 12

# (name, lower bound in points vs average, status)
CATEGORIES = [
    ("Significantly above average trend", 5, "Green"),
    ("Above average trend", 2, "Green"),
    ("About average", -2, "Orange"),
    ("Below average trend", -5, "Red"),
    ("Significantly below average trend", None, "Red"),
]
CATEGORY_ORDER = [c[0] for c in CATEGORIES]
STATUS = {name: status for name, _, status in CATEGORIES}


def category(diff):
    """Category for a centre's trend vs average, in percentage points. The bounds are
    'more than +5', '+2 to +5', '−2 to +2', '−5 to −2', 'below −5'."""
    if diff > 5:
        return CATEGORY_ORDER[0]
    if diff > 2:
        return CATEGORY_ORDER[1]
    if diff >= -2:
        return CATEGORY_ORDER[2]
    if diff >= -5:
        return CATEGORY_ORDER[3]
    return CATEGORY_ORDER[4]


def _window_trend(weekly, dates, end_i):
    """Per centre: average members over the WINDOW weeks ending at dates[end_i], the WINDOW
    weeks before, and the % change. weekly: Location × Date table of active members."""
    last = dates[end_i - WINDOW + 1:end_i + 1]
    prev = dates[end_i - 2 * WINDOW + 1:end_i - WINDOW + 1]
    a_last = weekly[last].mean(axis=1)
    a_prev = weekly[prev].mean(axis=1)
    pct = (a_last / a_prev.where(a_prev > 0) - 1) * 100
    return a_last, a_prev, pct


def trend_table(full, as_at, basis="All centres"):
    """One row per centre for the week ending `as_at`.

    full: Weekly Membership with Location, Date, '# Active members' and State.
    basis: 'All centres' or 'State' (the average trend the centre is compared with).
    Returns (table, too_new, summary). The average uses every established centre in the
    network (or state), whatever filters the report shows.
    """
    d = full[full["Date"] <= as_at]
    weekly = d.pivot_table(index=LOC, columns="Date", values="# Active members",
                           aggfunc="sum")
    dates = sorted(weekly.columns)
    state = d.groupby(LOC)["State"].last()
    if len(dates) < 2 * WINDOW:
        return pd.DataFrame(), pd.DataFrame(), {"weeks": len(dates)}
    weeks_trading = (weekly > 0).sum(axis=1)
    now = weekly[dates[-1]]
    trading = now.fillna(0) > 0
    established = trading & (weeks_trading >= MIN_WEEKS_TRADING)

    def compare(end_i):
        a_last, a_prev, pct = _window_trend(weekly, dates, end_i)
        est = established & a_prev.notna() & (a_prev > 0)
        if basis == "State":
            grp = state[est]
            avg = (a_last[est].groupby(grp).sum() / a_prev[est].groupby(grp).sum() - 1) * 100
            avg_for = state.map(avg)
        else:
            net = (a_last[est].sum() / a_prev[est].sum() - 1) * 100 if est.any() else None
            avg_for = pd.Series(net, index=weekly.index)
        return pct, avg_for, est

    pct, avg_for, est = compare(len(dates) - 1)
    diff = pct - avg_for
    rows = pd.DataFrame({
        "Location": weekly.index, "State": state.reindex(weekly.index).values,
        "Trend %": pct.values, "Average trend %": avg_for.values, "vs average": diff.values,
        "Members now": now.values,
        "Change in members": (now - weekly[dates[-1 - WINDOW]]).values,
    }).set_index("Location")
    rows = rows[est.reindex(rows.index).fillna(False) & rows["vs average"].notna()]
    rows["Category"] = rows["vs average"].map(category)
    rows["Status"] = rows["Category"].map(STATUS)
    rows["Small centre"] = rows["Members now"] < SMALL_CENTRE
    rows["Last 12 weeks"] = [weekly.loc[l, dates[-SPARK_WEEKS:]].fillna(0).tolist()
                             for l in rows.index]

    # Moved group since the previous week.
    rank = {c: i for i, c in enumerate(CATEGORY_ORDER)}
    rows["Moved"] = ""
    if len(dates) >= 2 * WINDOW + 1:
        p_pct, p_avg, p_est = compare(len(dates) - 2)
        p_cat = (p_pct - p_avg).dropna().map(category)
        for loc in rows.index:
            if loc in p_cat.index and p_est.get(loc, False):
                before, after = rank[p_cat[loc]], rank[rows.at[loc, "Category"]]
                rows.at[loc, "Moved"] = "▲ up" if after < before else \
                    "▼ down" if after > before else ""

    new = trading & ~established
    too_new = pd.DataFrame({"Weeks trading": weeks_trading[new], "Members now": now[new]})
    net_last, net_prev, _ = _window_trend(weekly[established], dates, len(dates) - 1)
    summary = {
        "weeks": len(dates),
        "network trend": (net_last.sum() / net_prev.sum() - 1) * 100
        if net_prev.sum() > 0 else None,
        "counts": rows["Category"].value_counts().reindex(CATEGORY_ORDER).fillna(0)
        .astype(int).to_dict(),
        "window": (dates[-2 * WINDOW], dates[-1]),
    }
    return rows, too_new, summary
