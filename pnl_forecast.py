"""P&L Centre Forecast: the calculations behind the admin-only 12-month forecast.

Revenue = members × average membership value × 4.2 weeks, as in the break-even. Each cost
line of the condensed P&L is set as a % of revenue; a line left at its default keeps its
own behaviour (tutor wages and marketing move with revenue, rent, manager wages, agency
fees and OPEX stay at their recent $ amount, royalties are the higher of the minimum or 8%).
"""
import pandas as pd

import pnl

RECENT = 3                      # months averaged for the starting point
DAMP = 0.9                      # each month the recent trend counts this much less
MAX_TREND = 0.10                # recent growth is capped at ±10% a month
OWN_PATTERN_MONTHS = 24         # months of membership history before a centre's own pattern
NETWORK_MIN_MEMBERS = 30        # centres this small are left out of the network pattern
NETWORK_MIN_OBS = 5             # fewer changes than this for a month: no seasonal effect

VARIABLE_LINES = ["Tutor Wages + Super", "Marketing"]
FIXED_LINES = ["Rent", "Wages - Manager/Admin + Super", "Marketing - Agency Fees", "OPEX"]
COST_LINES = pnl.ROW_ORDER[1:]                    # condensed P&L order, revenue left out
PCT = {line: f"{line} %" for line in COST_LINES}  # column names in the adjustment grid
INPUTS = ["Members", "AMV $ / week"] + [PCT[l] for l in COST_LINES]

ACTUAL, FORECAST, ESTIMATE, MISSING = "Actual", "Forecast", "Estimate ⚠", "No P&L ⚠"


# ── Months ────────────────────────────────────────────────────────────────────
def add_months(period, n):
    return (pd.Period(period, "M") + n).strftime("%Y-%m")


def fy_start(period):
    """Financial year (July to June) a month falls in, as the year it starts."""
    y, m = int(period[:4]), int(period[5:7])
    return y if m >= 7 else y - 1


def fy_months(start_year):
    return [add_months(f"{start_year}-07", i) for i in range(12)]


def fy_label(start_year):
    return f"FY{start_year}–{str(start_year + 1)[2:]} (Jul {start_year} – Jun {start_year + 1})"


def window(last_actual, view, fy=None):
    """Months shown: a financial year, or the 12 months after the last P&L."""
    if view == "Next 12 months":
        return [add_months(last_actual, i + 1) for i in range(12)]
    return fy_months(fy if fy is not None else fy_start(add_months(last_actual, 1)))


# ── Members ───────────────────────────────────────────────────────────────────
def _changes(series):
    """Month-on-month % change of a members series indexed by Period (consecutive only)."""
    s = series.sort_index()
    out = {}
    for prev, cur in zip(s.index[:-1], s.index[1:]):
        if add_months(prev, 1) == cur and s[prev] > 0:
            out[cur] = s[cur] / s[prev] - 1
    return pd.Series(out, dtype=float)


def _detrend(changes):
    """Typical change for each calendar month, less the typical change across all months."""
    if changes.empty:
        return {}, {}
    by_month = changes.groupby(changes.index.str[5:7])
    overall = changes.median()
    return {m: g.median() - overall for m, g in by_month}, by_month.size().to_dict()


def network_pattern(members, min_members=NETWORK_MIN_MEMBERS, min_obs=NETWORK_MIN_OBS):
    """{"01".."12": seasonal change} from every centre's membership history.
    members: average active members indexed by (location, Period)."""
    parts = []
    for _, s in members.groupby(level=0):
        s = s.droplevel(0)
        ch = _changes(s)
        prev = s.reindex([add_months(p, -1) for p in ch.index]).values
        parts.append(ch[prev >= min_members])
    allc = pd.concat(parts) if parts else pd.Series(dtype=float)
    allc = allc[allc.abs() <= 0.5]                  # opening months and data errors
    if allc.empty:
        return {}
    pattern, counts = _detrend(allc)
    return {m: v for m, v in pattern.items() if counts[m] >= min_obs}


def own_pattern(series, min_months=OWN_PATTERN_MONTHS):
    """The centre's own seasonal pattern, once it has two years of membership history."""
    if len(series) < min_months:
        return None
    pattern, _ = _detrend(_changes(series))
    return pattern


def forecast_members(series, months, pattern, damp=DAMP, cap=MAX_TREND):
    """{period: members} for each month. Months with membership data use it; later months
    carry on from the last known month: seasonal change + the recent trend, which fades.
    series: the centre's average active members indexed by Period."""
    s = series.sort_index()
    s = s[s > 0]
    if s.empty:
        return {p: None for p in months}
    ch = _changes(s).tail(RECENT)
    trend = (ch - [pattern.get(p[5:7], 0.0) for p in ch.index]).mean() if len(ch) else 0.0
    trend = max(-cap, min(cap, 0.0 if pd.isna(trend) else trend))
    last = s.index[-1]
    value, out = s.iloc[-1], {}
    step, p = 0, last
    end = max(months)
    while p < end:
        p = add_months(p, 1)
        step += 1
        value = value * (1 + pattern.get(p[5:7], 0.0) + trend * damp ** step)
        out[p] = value
    return {m: (s[m] if m in s.index else out.get(m)) for m in months}


# ── Starting point ────────────────────────────────────────────────────────────
def starting_point(data, n=RECENT):
    """Averages over the centre's last n P&L months: AMV, the variable lines' % of revenue,
    the fixed lines' $ per month and its royalties.
    data: the centre's months (Period index, condensed rows and Members)."""
    recent = data.sort_index().tail(n)
    rev = recent["Total Revenue"].sum()
    mem = recent["Members"].where(recent["Members"] > 0)
    with_members = recent[mem.notna()]
    amv = (with_members["Total Revenue"] / with_members["Members"]).mean() / pnl.WEEKS_PER_MONTH \
        if len(with_members) else None
    return {
        "months": list(recent.index),
        "amv": None if amv is None or pd.isna(amv) else float(amv),
        "pct": {l: (recent[l].sum() / rev if rev > 0 else 0.0) for l in VARIABLE_LINES},
        "fixed": {l: float(recent[l].mean()) for l in FIXED_LINES},
        "royalties": float(recent["Royalties"].mean()),
    }


def royalty_for(revenue, period, royalty_min, relief_until=None):
    """Higher of the minimum or 8% of revenue; nothing while royalty relief lasts."""
    if relief_until and period <= relief_until:
        return 0.0
    return pnl.royalty(revenue, royalty_min)


def defaults(months, members, start, royalty_min, relief_until=None, rise_from=None,
             rise_pct=0.0):
    """The adjustment grid before any changes: one row per forecast month."""
    rows = []
    for p in months:
        amv = start["amv"] or 0.0
        if rise_from and p >= rise_from:
            amv *= 1 + rise_pct / 100
        mem = members.get(p) or 0.0
        rev = mem * amv * pnl.WEEKS_PER_MONTH
        row = {"Month": p, "Members": float(mem), "AMV $ / week": amv}
        for l in VARIABLE_LINES:
            row[PCT[l]] = start["pct"][l] * 100
        for l in FIXED_LINES:
            row[PCT[l]] = start["fixed"][l] / rev * 100 if rev > 0 else 0.0
        roy = royalty_for(rev, p, royalty_min, relief_until)
        row[PCT["Royalties"]] = roy / rev * 100 if rev > 0 else 0.0
        rows.append(row)
    return pd.DataFrame(rows, columns=["Month"] + INPUTS).set_index("Month")


def changed(edited, base, col, p, tol=1e-6):
    return abs(float(edited.at[p, col]) - float(base.at[p, col])) > tol


def compute(edited, base, start, royalty_min, relief_until=None):
    """Each forecast month's condensed P&L from the grid. A line left at its default keeps
    its own behaviour (see the module docstring); a changed line uses its % of revenue."""
    rows = {}
    for p in edited.index:
        mem, amv = float(edited.at[p, "Members"] or 0), float(edited.at[p, "AMV $ / week"] or 0)
        rev = mem * amv * pnl.WEEKS_PER_MONTH
        col = {"Members": mem, "AMV": amv, "Total Revenue": rev}
        for l in COST_LINES:
            pct = float(edited.at[p, PCT[l]] or 0) / 100
            if l in VARIABLE_LINES or changed(edited, base, PCT[l], p):
                col[l] = pct * rev
            elif l in FIXED_LINES:
                col[l] = start["fixed"][l]
            else:
                col[l] = royalty_for(rev, p, royalty_min, relief_until)
        col["Royalty changed"] = changed(edited, base, PCT["Royalties"], p)
        rows[p] = col
    out = pd.DataFrame.from_dict(rows, orient="index")
    if out.empty:
        return out
    out[pnl.GROSS_MARGIN] = out["Total Revenue"] - out[pnl.TUTOR]
    out["EBITDA"] = out["Total Revenue"] - out[COST_LINES].sum(axis=1)
    for p in out.index:
        out.loc[p, "BE members"] = break_even_members(out.loc[p], royalty_min, relief_until, p)
    return out


def break_even_members(col, royalty_min, relief_until=None, period=None):
    """Members needed for EBITDA of 0 at this month's AMV and cost percentages."""
    rev, mem = col["Total Revenue"], col["Members"]
    if rev <= 0 or mem <= 0:
        return None
    fixed = sum(col[l] for l in FIXED_LINES)
    variable = sum(col[l] for l in VARIABLE_LINES)
    if col.get("Royalty changed"):
        rmin, rate = 0.0, col["Royalties"] / rev
    elif relief_until and period and period <= relief_until:
        rmin, rate = 0.0, 0.0
    else:
        rmin, rate = royalty_min, pnl.ROYALTY_RATE
    _, members = pnl.break_even(rev, variable, fixed, rmin, mem, rate)
    return members


# ── Putting it together ───────────────────────────────────────────────────────
def months_table(actual, forecast, months, seen_months, royalty_min=pnl.DEFAULT_ROYALTY_MIN):
    """Every month shown, in order, with how it was worked out (the Basis column).
    actual: the centre's P&L months; forecast: compute()'s result; seen_months: calendar
    months ("01".."12") the centre has P&L for (others are flagged as estimates)."""
    rows = []
    for p in months:
        if p in actual.index:
            r = actual.loc[p].to_dict()
            r["AMV"] = r["Total Revenue"] / r["Members"] / pnl.WEEKS_PER_MONTH \
                if r.get("Members") and r["Members"] > 0 else None
            r["Basis"] = ACTUAL
            relief = p if r["Royalties"] == 0 else None         # $0 royalties: relief
            r["BE members"] = break_even_members(r, royalty_min, relief, p) \
                if r["AMV"] else None
        elif p in forecast.index:
            r = forecast.loc[p].to_dict()
            r["Basis"] = FORECAST if p[5:7] in seen_months else ESTIMATE
        else:
            r = {"Basis": MISSING}
        r["Month"] = p
        rows.append(r)
    return pd.DataFrame(rows).set_index("Month")


def totals(table):
    """Headline figures for the months shown (months without a P&L left out)."""
    t = table[table["Basis"] != MISSING]
    rev = t["Total Revenue"].sum()
    ebitda = t["EBITDA"].sum()
    return {"Revenue": rev, "EBITDA": ebitda,
            "EBITDA %": ebitda / rev * 100 if rev > 0 else None,
            "End members": t["Members"].dropna().iloc[-1] if t["Members"].notna().any() else None,
            "Months below 0": int((t["EBITDA"] < 0).sum()),
            "Forecast months": int(t["Basis"].isin([FORECAST, ESTIMATE]).sum()),
            "Estimate months": int((t["Basis"] == ESTIMATE).sum()),
            "Months": len(t)}


def benchmark_values(table, country=None):
    """Per-month averages of the forecast months, ready for pnl.compare_lines."""
    f = table[table["Basis"].isin([FORECAST, ESTIMATE])]
    if f.empty:
        return None
    v = pnl.with_gross_margin(f)[pnl.ROW_ORDER + [pnl.GROSS_MARGIN, "EBITDA"]].mean()
    v["Members"] = f["Members"].mean()
    v["Country"] = country
    return v
