"""P&L Share: the calculations behind the partner-friendly one-centre report.

Benchmarks use every centre, but nothing here returns another centre's name, and every
benchmark needs at least MIN_PEERS similar centres, so no single centre's figures can be
worked out from what is shown:
- the normal range is the middle half (25th–75th percentile), never lowest-to-highest;
- a line is only benchmarked when most similar centres actually use that account;
- sensitive lines (rent, royalties, manager wages, interest, CRM software) are never
  benchmarked.
"""
import re

import pandas as pd

import membership_trend
import pnl

MIN_PEERS = 5                   # fewer similar centres than this: no benchmark
MOST_PEERS = 0.5                # a line is benchmarked when more than half the peers use it
YOUNG_WEEKS = 18                # centres this young are left out of the peer groups

# Sensitive lines: shown as the centre's own figure only, never against other centres.
# royalties, rent, manager wages + super, interest, CRM software
SENSITIVE_CODES = {467, 469, 471, 478, 437, 486}
SENSITIVE_ROWS = {"Rent", "Royalties", pnl.MANAGER, "Interest"}
_SENSITIVE_NAME = re.compile(r"royalt|\brent\b|manager|interest\s+expense|\bcrm\b", re.I)

NOT_BENCHMARKED = "Not benchmarked"
NO_BENCHMARK = "No benchmark"
NOT_ENOUGH = "Not enough similar centres to compare"

# Softer wording for partners. Costs: higher is worse.
COST_WORDS = {"Very high": "Worth a look", "High": "A little above typical",
              "Typical": "In line with similar centres", "Low": "Below typical",
              "Very low": "Well below typical"}
GAIN_WORDS = {"Very high": "Well above typical", "High": "Above typical",
              "Typical": "In line with similar centres", "Low": "A little below typical",
              "Very low": "Worth a look"}

PLAIN = {"Tutor Wages + Super": "tutor wages and super", "Marketing": "marketing",
         "Marketing - Agency Fees": "marketing agency fees",
         "Average Membership Value": "your average membership value"}


# ── Sensitive lines ───────────────────────────────────────────────────────────
def sensitive_accounts(df_pnl):
    """Account names that are sensitive: by account code in the data, plus any name that
    looks like rent, royalties, manager wages, interest expense or CRM software."""
    names = set()
    if df_pnl is not None and not df_pnl.empty:
        codes = pd.to_numeric(df_pnl["Account code"], errors="coerce")
        names = set(df_pnl.loc[codes.isin(SENSITIVE_CODES), "Account"].astype(str))
        names |= {a for a in df_pnl["Account"].astype(str).unique() if _SENSITIVE_NAME.search(a)}
    return names


def is_sensitive(line, sensitive_names=()):
    return line in SENSITIVE_ROWS or line in sensitive_names or \
        bool(_SENSITIVE_NAME.search(str(line)))


# ── Similar centres ───────────────────────────────────────────────────────────
def _stage_word(age):
    return "developing" if pnl.is_developing(age) else "established"


def describe(stage, desc):
    """'developing' + 'Australian centres with 40–139 members'
    -> 'developing Australian centres with 40–139 members'."""
    if stage is None:
        return desc[0].lower() + desc[1:] if desc.startswith("Centres") else desc
    if desc.startswith("all "):
        return f"all {stage} {desc[4:]}"
    if desc.startswith("Centres"):
        return f"{stage} c{desc[1:]}"
    return f"{stage} {desc}"


def share_peers(bench, age, country, members, exclude, young=(), min_n=MIN_PEERS):
    """Similar centres with at least min_n in the group. Tries the same stage (developing
    or established) first, widening the size band and then the country as pnl.peer_set
    does; with too few, drops the stage. Young centres (YOUNG_WEEKS or younger) are left
    out. Returns (peers, description) or (empty, None) when there aren't enough."""
    pool = bench[~bench.index.isin(young)] if len(young) else bench
    established, developing = pnl.split_by_age(pool)
    stage = _stage_word(age)
    same = developing if stage == "developing" else established
    for p, word in ((same, stage), (pool, None)):
        peers, desc = pnl.peer_set(p, country, members, exclude, min_n)
        if len(peers[peers["Total Revenue"] > 0]) >= min_n:
            return peers, describe(word, desc)
    return pool.iloc[0:0], None


# ── Comparison table ──────────────────────────────────────────────────────────
def soft_status(line, status):
    words = GAIN_WORDS if line in pnl.HIGHER_IS_BETTER else COST_WORDS
    return words.get(status, status)


def benchmark_lines(values, peers, lines, sensitive_names=(), min_n=MIN_PEERS):
    """Every line against similar centres, with the share report's rules applied.

    Returns one row per line: Line, Actual $, Actual %, Typical %, Range low %,
    Range high % (middle half), Centres, Users (peers with a non-zero amount), Gap $
    (what the difference is worth a month, positive = costs more), Status (internal
    word), Benchmarked (bool) and Note ('' or why it isn't benchmarked). Lines that
    aren't benchmarked keep only their own $ and %."""
    peers = peers if peers is not None else pd.DataFrame()
    cmp = pnl.compare_lines(values, peers, lines) if len(peers) else pd.DataFrame()
    if cmp.empty:
        rev = values.get("Total Revenue", 0)
        cmp = pd.DataFrame([{"Line": l, "Actual $": values.get(l, 0.0),
                             "Actual %": values.get(l, 0.0) / rev * 100 if rev > 0 else None}
                            for l in lines if l in values])
    cmp = cmp[~cmp["Line"].astype(str).str.startswith("↳")].reset_index(drop=True)
    pp = peers[peers["Total Revenue"] > 0] if len(peers) else peers
    blank = ["Typical %", "Range low %", "Range high %", "Typical $", "Range low $",
             "Range high $", "Gap $", "Status"]
    for c in blank + ["Centres"]:
        if c not in cmp:
            cmp[c] = None
    users, ok, notes = [], [], []
    for i, r in cmp.iterrows():
        line = r["Line"]
        n_users = int((pp[line].fillna(0) != 0).sum()) if line in pp else None
        if is_sensitive(line, sensitive_names):
            note = NOT_BENCHMARKED
        elif len(pp) < min_n or (pd.notna(r["Centres"]) and r["Centres"] < min_n):
            note = NOT_ENOUGH
        elif n_users is not None and n_users <= len(pp) * MOST_PEERS:
            note = NO_BENCHMARK
        else:
            note = ""
        users.append(n_users)
        ok.append(note == "")
        notes.append(note)
    cmp["Users"], cmp["Benchmarked"], cmp["Note"] = users, ok, notes
    cmp.loc[~cmp["Benchmarked"], blank] = None
    cmp.loc[~cmp["Benchmarked"], "Centres"] = None
    return cmp


def out_of_step(table):
    """Benchmarked lines outside the normal range: (costing more, costing less)."""
    t = table[table["Benchmarked"] & table["Status"].isin(["High", "Very high", "Low",
                                                           "Very low"])]
    t = t[t["Line"] != "Average Membership Value"]
    bad = pd.Series([pnl.is_bad(l, s) for l, s in zip(t["Line"], t["Status"])],
                    index=t.index, dtype=bool)
    more, less = t[bad], t[~bad]
    order = lambda d: d.assign(_g=d["Gap $"].abs()).sort_values("_g", ascending=False) \
        .drop(columns="_g")
    return order(more), order(less)


def top_opportunities(table, n=3):
    """The n benchmarked lines worth the most a month if they matched similar centres,
    in plain English. Sensitive and unbenchmarked lines never appear."""
    t = table[table["Benchmarked"] & (table["Gap $"].fillna(0) > 0)]
    t = t[t["Line"] != "EBITDA"].sort_values("Gap $", ascending=False).head(n)
    out = []
    for _, r in t.iterrows():
        name = PLAIN.get(r["Line"], r["Line"].lower())
        verb = "earn" if r["Line"] in pnl.HIGHER_IS_BETTER else "keep"
        out.append({"Line": r["Line"], "Gap $": r["Gap $"],
                    "Text": f"If {name} matched similar centres, you'd {verb} about "
                            f"${r['Gap $']:,.0f} more a month."})
    return out


def opportunity_lines(grouped, accounts, account_rows):
    """Candidates for the top opportunities: the grouped lines (tutor wages, marketing),
    each OPEX account on its own, and the average membership value. account_rows:
    {account name: summary row} so OPEX accounts can be picked out."""
    g = grouped[grouped["Line"].isin(["Tutor Wages + Super", "Marketing",
                                      "Marketing - Agency Fees", "Average Membership Value"])]
    a = accounts[accounts["Line"].map(lambda l: account_rows.get(l) == "OPEX")]
    return pd.concat([g, a], ignore_index=True)


# ── Quarter bands ─────────────────────────────────────────────────────────────
BANDS = ["Top quarter", "Above the middle", "Below the middle", "Bottom quarter"]


def quarter_band(value, peer_values, higher_is_better=True, min_n=MIN_PEERS):
    """Which quarter of similar centres a value falls in (never an exact rank).
    None when there are fewer than min_n peer values."""
    p = pd.Series(peer_values, dtype=float).dropna()
    if value is None or pd.isna(value) or len(p) < min_n:
        return None
    q1, med, q3 = p.quantile(0.25), p.median(), p.quantile(0.75)
    if not higher_is_better:
        value, q1, med, q3 = -value, -q3, -med, -q1
    if value > q3:
        return BANDS[0]
    if value > med:
        return BANDS[1]
    if value >= q1:
        return BANDS[2]
    return BANDS[3]


def centre_bands(values, peers, min_n=MIN_PEERS):
    """{measure: band} for EBITDA %, gross margin % and revenue per member (same country)."""
    pp = peers[peers["Total Revenue"] > 0] if len(peers) else peers
    rev = values.get("Total Revenue", 0)
    out = {}
    if rev <= 0 or len(pp) < min_n:
        return out
    out["EBITDA % of revenue"] = quarter_band(
        values["EBITDA"] / rev * 100, pp["EBITDA"] / pp["Total Revenue"] * 100, True, min_n)
    out["Gross margin (after tutor wages)"] = quarter_band(
        (rev - values["Tutor Wages + Super"]) / rev * 100,
        (pp["Total Revenue"] - pp["Tutor Wages + Super"]) / pp["Total Revenue"] * 100,
        True, min_n)
    mem = values.get("Members")
    if mem and not pd.isna(mem) and mem > 0 and "Members" in pp:
        same = pp[(pp["Members"].fillna(0) > 0)]
        if isinstance(values.get("Country"), str) and "Country" in same:
            same = same[same["Country"].astype(str) == values["Country"]]
        out["Revenue per member"] = quarter_band(rev / mem, same["Total Revenue"] /
                                                 same["Members"], True, min_n)
    return {k: v for k, v in out.items() if v}


# ── 12-month trend ────────────────────────────────────────────────────────────
def typical_by_month(lm, members, peer_locs, periods, country=None, country_of=None,
                     min_n=MIN_PEERS):
    """Typical (median) similar centre each month: revenue ($, same country only),
    revenue per member, EBITDA %, gross margin % and OPEX %. A month with fewer than min_n of these centres
    reporting is left blank. lm: pnl.location_months; members: (Location, Period) series;
    country_of: {location: country}."""
    rows = []
    sub = lm[lm.index.get_level_values("Location").isin(peer_locs)]
    for p in periods:
        m = sub[sub.index.get_level_values("Period") == p].droplevel("Period")
        m = m[m["Total Revenue"] > 0]
        row = {"Period": p, "Revenue": None, "Revenue per member": None, "EBITDA %": None,
               "Gross margin %": None, "OPEX %": None, "Centres": len(m)}
        if len(m) >= min_n:
            row["EBITDA %"] = float((m["EBITDA"] / m["Total Revenue"] * 100).median())
            gm = m["Total Revenue"] - m[pnl.TUTOR]
            row["Gross margin %"] = float((gm / m["Total Revenue"] * 100).median())
            row["OPEX %"] = float((m["OPEX"] / m["Total Revenue"] * 100).median())
        if country is not None and country_of is not None:
            m = m[[str(country_of.get(l, "")) == str(country) for l in m.index]]
        if len(m) >= min_n:
            row["Revenue"] = float(m["Total Revenue"].median())
            if members is not None and not members.empty:
                mem = pd.Series({l: members.get((l, p)) for l in m.index}, dtype=float)
                per = (m["Total Revenue"] / mem.where(mem > 0)).dropna()
                if len(per) >= min_n:
                    row["Revenue per member"] = float(per.median())
        rows.append(row)
    return pd.DataFrame(rows).set_index("Period") if rows else pd.DataFrame()


# ── Members trend ─────────────────────────────────────────────────────────────
def members_trend(full_wm, loc, as_at):
    """This centre's 4-week membership trend against the network's (membership_trend).
    Returns a dict, or None if there isn't enough history (or the centre is too new)."""
    if full_wm is None or full_wm.empty or "Date" not in full_wm:
        return None
    d = full_wm.copy()
    if "State" not in d:
        d["State"] = ""
    d["# Active members"] = pd.to_numeric(d["# Active members"], errors="coerce").fillna(0)
    rows, too_new, summary = membership_trend.trend_table(d, pd.Timestamp(as_at))
    if loc in too_new.index:
        return {"too_new": True, "weeks": int(too_new.at[loc, "Weeks trading"])}
    if rows.empty or loc not in rows.index:
        return None
    r = rows.loc[loc]
    return {"too_new": False, "trend": float(r["Trend %"]),
            "average": float(r["Average trend %"]), "vs": float(r["vs average"]),
            "category": r["Category"], "status": r["Status"],
            "members": float(r["Members now"]), "change": float(r["Change in members"]),
            "spark": list(r["Last 12 weeks"]), "small": bool(r["Small centre"]),
            "window": summary.get("window")}


# ── Check before sharing ──────────────────────────────────────────────────────
MAJOR_LINES = ["Tutor Wages + Super", "Rent", pnl.MANAGER, "Marketing", "OPEX"]


def checks(values, peers, accounts=None, missed=(), interest_income=0.0, young=False,
           min_n=MIN_PEERS):
    """Odd things the person sharing should look at first (shown on screen, never printed).
    Royalties at 0 are not flagged: royalty relief is normal."""
    out = []
    pp = peers[peers["Total Revenue"] > 0] if peers is not None and len(peers) else None
    if pp is None or len(pp) < min_n:
        out.append(f"Fewer than {min_n} similar centres reported for this period, so the "
                   "comparisons are left out.")
    elif pnl.small_sample(len(pp)):
        out.append(f"The comparisons rest on {len(pp)} similar centres (a small sample): "
                   "treat them as a guide.")
    if pp is not None and len(pp):
        for line in MAJOR_LINES:
            if line in values and line in pp and abs(values.get(line, 0) or 0) < 0.5 and \
                    (pp[line].fillna(0) != 0).sum() > len(pp) * MOST_PEERS:
                out.append(f"{line} is $0 this period, but most similar centres have an "
                           "amount. Check the P&L before sharing.")
    if accounts is not None and len(accounts):
        neg = [l for l, v in zip(accounts["Line"], accounts["Actual $"])
               if l != "Average Membership Value" and v is not None and v < -0.5]
        if neg:
            out.append("Negative costs (credits): " + ", ".join(neg) + ".")
    if missed:
        out.append("No P&L for " + ", ".join(missed) + ": the average uses the months "
                   "reported.")
    if interest_income and interest_income > 0.5:
        out.append(f"Revenue includes ${interest_income:,.0f} of interest income.")
    mem = values.get("Members")
    if mem is None or pd.isna(mem) or mem <= 0:
        out.append("No members figures for this period, so per-member numbers are missing.")
    if values.get("Total Revenue", 0) < pnl.LOW_REVENUE:
        out.append("Revenue is very low this period, so percentages may look odd.")
    if young:
        out.append(f"The centre is {YOUNG_WEEKS} weeks old or younger: scheduling "
                   "comparisons are left out.")
    return out
