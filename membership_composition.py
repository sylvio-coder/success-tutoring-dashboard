"""Membership Composition: each centre's plans and prices against RRP, from the portal's
weekly Membership Composition export. Calculations only; the page is composition_report.py.

Rules agreed with HQ:
- RRP is national for Australia, New Zealand and Canada; the US is priced by region.
- Prices under $10 a week (a freeze fee or discretionary amount) are left out of the plan
  rows but kept in totals, ARPM and composition, with a note.
- Prices are exported per week for every country (the US was monthly until 22 Nov 2026;
  a group price over $100 a session is flagged on upload as probably monthly).
- 4+ sessions a week: RRP = the 3x rate per session x sessions. 1:1: RRP per session.
- Ideal composition: 15% 1x, 80% 2x, 1% 3+ and 4% 1:1 of the centre's members, at RRP or
  the centre's own price where that is higher.
"""
import csv
import io
import math
import statistics

import pandas as pd

import ingest

TAB = "Membership Composition"
HEADERS = ["Week ending", "Location", "Country", "Type", "Sessions/week", "Members",
           "Weekly price", "Export name"]
RRP_TAB = "RRP"
RRP_HEADERS = ["Country", "Region", "1x", "2x", "3x", "Foundational 1x", "Foundational 2x",
               "Foundational 3x", "1:1 per session", "Effective from", "Notes"]
CRM = "Composition"            # alias scope: names matched for these uploads only
LOC = "Success Tutoring - Business name"

FREEZE_BELOW = 10              # weekly price under this = freeze fee / discretionary
LOOKS_MONTHLY = 100            # a group price per session above this is probably monthly
REGIONAL = {"US", "CA"}        # countries priced by region (others: national RRP only)
TARGET_MIX = {"1×": 15, "2×": 80, "3+": 1, "1:1": 4}
IDEAL_GROUPS = [("Group", "1×", 1, 1, 1), ("Group", "2×", 2, 2, 2),
                ("Group", "3+", 3, 3, 3), ("1:1", "1:1", 1, 99, 1)]    # type, plan, min, max, sessions
NOT_RECOMMENDED_FROM = 4       # group plans with this many sessions a week or more
NOT_RECOMMENDED = "4×+ (not recommended)"


def not_recommended(kind, sessions):
    """Group plans with 4+ sessions a week: outside the recommended plans, so they are
    kept at today's price (no RRP uplift) and left out of the ideal mix and % of RRP."""
    return kind == "Group" and sessions >= NOT_RECOMMENDED_FROM
# Scheduling: the target students per study hub rises with members, in straight lines
# between these points (CRM: young centres ~3.1-3.4; the best centres over 150 members
# 5.5-6). A hub holds at most 6. 1:1 sessions need a tutor each.
FILL_POINTS = [(20, 3.0), (40, 3.5), (120, 5.0), (150, 5.5), (180, 6.0)]
TUTOR_HOUR = {"AU": 37.90}     # tutor cost an hour incl. super ($33.85 + 12%); one per hub
CURRENCY = {"AU": "AUD", "NZ": "NZD", "US": "USD", "CA": "CAD", "GB": "GBP"}
COUNTRY_CODE = {"australia": "AU", "new zealand": "NZ", "united states": "US", "usa": "US",
                "united states of america": "US", "canada": "CA", "united kingdom": "GB",
                "uk": "GB"}
EXPORT_HEADERS = {"date", "centre", "country", "row", "type", "no. members",
                  "sessions/week", "membership price"}


class CompositionError(ValueError):
    pass


# ── RRP ───────────────────────────────────────────────────────────────────────
def _rrp(country, region, mkt, fnd, one, notes=""):
    return dict(zip(RRP_HEADERS, [country, region, *mkt, *fnd, one, "2026-07-01", notes]))


# Weekly prices from 1 July 2026 (market benchmark, foundational) and 1:1 per session.
DEFAULT_RRP = [
    _rrp("AU", "All", (42, 74, 100), (36, 63, 85), 90, "National RRP"),
    _rrp("NZ", "All", (47, 82, 110), (40, 70, 94), 100, "National RRP"),
    _rrp("CA", "All", (51, 89, 121), (43, 76, 103), 108, "National RRP"),
    _rrp("US", "All", (47, 82, 110), (40, 70, 94), 100, "National US RRP; billed monthly"),
    _rrp("US", "Texas", (49, 85, 115), (42, 72, 98), 103, "Austin & Houston"),
]


def country_code(value):
    v = str(value or "").strip()
    return v.upper() if v.upper() in CURRENCY else COUNTRY_CODE.get(v.lower(), v.upper())


def _region_key(region):
    return str(region or "").strip().lower()


def rrp_table(df_rrp=None):
    """{(country, region): {mkt, one, label}} from the RRP tab (or the defaults)."""
    rows = DEFAULT_RRP if df_rrp is None or df_rrp.empty else df_rrp.to_dict("records")
    out = {}
    for r in rows:
        try:
            num = lambda h: float(str(r.get(h, "")).replace("$", "").replace(",", ""))
            cc = country_code(r["Country"])
            out[(cc, _region_key(r["Region"]))] = dict(
                mkt=tuple(num(f"{n}x") for n in (1, 2, 3)), one=num("1:1 per session"),
                label=f"{cc} national" if _region_key(r["Region"]) == "all"
                else f"{cc} · {r['Region']}")
        except (ValueError, KeyError):
            continue
    return out


def rrp_for(table, country, region=None):
    """National RRP; the US (and Canada, once regions are added) by region if listed."""
    cc = country_code(country)
    if cc in REGIONAL and table.get((cc, _region_key(region))):
        return table[(cc, _region_key(region))]
    return table.get((cc, "all"))


def rrp_price(rrp, kind, sessions):
    """Weekly RRP for a plan. 4+ a week: the 3x rate per session x sessions."""
    if rrp is None or not sessions:
        return None
    if kind == "1:1":
        return rrp["one"] * sessions
    return rrp["mkt"][sessions - 1] if sessions <= 3 else rrp["mkt"][2] / 3 * sessions


def plan_label(kind, sessions):
    return "1:1" if kind == "1:1" else f"Group {int(sessions)}×"


# ── Reading the export ────────────────────────────────────────────────────────
def _num(v):
    try:
        return float(str(v).replace("$", "").replace(",", "").strip())
    except ValueError:
        return None


def read_export(data):
    """The portal CSV (bytes or text) -> {week_end, rows, empty}."""
    text = data.decode("utf-8-sig") if isinstance(data, bytes) else data
    reader = csv.DictReader(io.StringIO(text))
    fields = {ingest.norm_header(h): h for h in (reader.fieldnames or [])}
    missing = EXPORT_HEADERS - set(fields)
    if missing:
        raise CompositionError("This doesn't look like the Membership Composition export "
                               f"(missing: {', '.join(sorted(missing))}).")
    get = lambda r, h: (r.get(fields[h]) or "").strip()
    rows, centres, dates = [], set(), set()
    for r in reader:
        centre = get(r, "centre")
        if not centre:
            continue
        centres.add(centre)
        if get(r, "date"):
            dates.add(get(r, "date"))
        if get(r, "row").lower() != "plan":
            continue
        members, sessions, price = (_num(get(r, h)) for h in
                                    ("no. members", "sessions/week", "membership price"))
        if not members or not sessions or price is None:
            continue
        cc = country_code(get(r, "country"))
        rows.append({"Centre": centre, "Country": cc,
                     "Type": "1:1" if get(r, "type").replace(" ", "") == "1:1" else "Group",
                     "Sessions/week": int(sessions), "Members": int(members),
                     "Weekly price": round(price, 2)})
    if not centres:
        raise CompositionError("The file has no centres in it.")
    week = None
    if dates:
        last = sorted(dates)[-1]
        week = pd.to_datetime(last, format="%Y-%m-%d", errors="coerce")
        if pd.isna(week):
            week = pd.to_datetime(last, dayfirst=True, errors="coerce")
    df = pd.DataFrame(rows, columns=["Centre", "Country", "Type", "Sessions/week", "Members",
                                     "Weekly price"])
    return {"week_end": None if week is None or pd.isna(week) else week.date(),
            "rows": df, "empty": sorted(centres - set(df["Centre"]))}


def looks_monthly(rows):
    """Group plans priced above LOOKS_MONTHLY a session: the export was probably monthly
    (the old US export). Rows from read_export."""
    if rows is None or rows.empty:
        return rows
    per = rows["Weekly price"] / rows["Sessions/week"]
    return rows[(rows["Type"] == "Group") & (per > LOOKS_MONTHLY)]


def match_centres(names, master, aliases):
    """{export name: master location or None}, plus suggestions for the unknown ones."""
    matcher = ingest.LocationMatcher(master["Location"].tolist(), aliases)
    found = {n: matcher.resolve(n, crm=CRM) for n in names}
    return found, {n: matcher.suggestions(n) for n, loc in found.items() if not loc}


def records(rows, locations, week_end):
    """Sheet rows for the matched centres (export name -> location)."""
    label = "'" + week_end.isoformat()           # kept as text, so it reads back the same
    return [{"Week ending": label, "Location": locations[r["Centre"]], "Country": r["Country"],
             "Type": r["Type"], "Sessions/week": r["Sessions/week"], "Members": r["Members"],
             "Weekly price": r["Weekly price"], "Export name": r["Centre"]}
            for _, r in rows.iterrows() if locations.get(r["Centre"])]


def prepare(df):
    """The sheet tab, typed, with a Freeze column (price under $10)."""
    cols = ["Week ending", "Location", "Country", "Type", "Sessions/week", "Members",
            "Weekly price", "Freeze"]
    if df is None or df.empty or not {"Week ending", "Location"} <= set(df.columns):
        return pd.DataFrame(columns=cols)
    out = df.copy()
    out["Week ending"] = pd.to_datetime(out["Week ending"].astype(str).str.lstrip("'"),
                                        format="%Y-%m-%d", errors="coerce")
    for c in ["Sessions/week", "Members", "Weekly price"]:
        out[c] = pd.to_numeric(out[c].astype(str).str.replace("[$,]", "", regex=True),
                               errors="coerce")
    out = out.dropna(subset=["Week ending", "Sessions/week", "Members", "Weekly price"])
    out = out[out["Sessions/week"] > 0]
    out["Sessions/week"] = out["Sessions/week"].astype(int)
    out["Members"] = out["Members"].astype(int)
    out["Country"] = out["Country"].map(country_code)
    out["Freeze"] = out["Weekly price"] < FREEZE_BELOW
    return out[cols]


# ── Medians ───────────────────────────────────────────────────────────────────
def plan_medians(week):
    """Median price members pay for each plan (type + sessions) in each country, from
    every centre in the week: {(country, type, sessions): (median, centres)}."""
    out = {}
    ok = week[~week["Freeze"]]
    for k, g in ok.groupby(["Country", "Type", "Sessions/week"]):
        prices = [p for p, m in zip(g["Weekly price"], g["Members"]) for _ in range(int(m))]
        out[k] = (statistics.median(prices), g["Location"].nunique())
    return out


def _variance(price, ref):
    if ref is None or not ref:
        return None, None
    return price - ref, (price / ref - 1) * 100


# ── Table 1: current composition ──────────────────────────────────────────────
def current_table(plans, rrp, medians):
    """Plan rows (prices under $10 left out), totals over every member, the composition
    by plan and the under-$10 note."""
    n = int(plans["Members"].sum())
    value = float((plans["Members"] * plans["Weekly price"]).sum())
    sessions = float((plans["Members"] * plans["Sessions/week"]).sum())
    rows = []
    show = plans[~plans["Freeze"]].sort_values(["Type", "Sessions/week", "Weekly price"],
                                               key=lambda c: c.map({"Group": 0, "1:1": 1})
                                               if c.name == "Type" else c)
    for _, p in show.iterrows():
        cc, kind, s, price = p["Country"], p["Type"], int(p["Sessions/week"]), p["Weekly price"]
        r = rrp_price(rrp, kind, s)
        med, cnt = medians.get((cc, kind, s), (None, 0))
        rv, rp = _variance(price, r)
        mv, mp = _variance(price, med) if cnt > 1 else (None, None)
        rows.append({"Type": kind, "Members": int(p["Members"]), "Sessions/week": s,
                     "Total sessions": p["Members"] * s,
                     "Not recommended": not_recommended(kind, s),
                     "Membership price": price, "Total value": p["Members"] * price,
                     "Revenue per session": price / s, "RRP": r, "RRP variance": rv,
                     "RRP variance %": rp, "Median": med if cnt > 1 else None,
                     "Median variance": mv, "Median variance %": mp})
    comp = (plans.assign(Plan=[plan_label(t, s) for t, s in
                               zip(plans["Type"], plans["Sessions/week"])],
                         order=[(t == "1:1", s) for t, s in
                                zip(plans["Type"], plans["Sessions/week"])])
            .assign(Value=plans["Members"] * plans["Weekly price"],
                    Sess=plans["Members"] * plans["Sessions/week"])
            .groupby("Plan").agg(Members=("Members", "sum"), order=("order", "min"),
                                 Sess=("Sess", "sum"), **{"Total value": ("Value", "sum")})
            .sort_values("order").reset_index())
    members = comp["Members"].where(comp["Members"] > 0)
    comp["% of members"] = comp["Members"] / n * 100 if n else 0
    comp["Total sessions"] = comp["Sess"]                   # members × sessions a week
    comp["Average price"] = comp["Total value"] / members
    comp["Revenue per session"] = comp["Total value"] / comp["Sess"].where(comp["Sess"] > 0)
    comp["Sessions per member"] = comp["Sess"] / members
    frozen = plans[plans["Freeze"]]
    return {"rows": pd.DataFrame(rows), "members": n, "total": value,
            "per_session": value / sessions if sessions else None,
            "arpm": value / n if n else None, "sessions": sessions,
            "composition": comp[["Plan", "Members", "% of members", "Total sessions",
                                 "Sessions per member",
                                 "Average price", "Total value", "Revenue per session"]],
            "freeze_members": int(frozen["Members"].sum()),
            "freeze_value": float((frozen["Members"] * frozen["Weekly price"]).sum())}


# ── Table 2: ideal composition ────────────────────────────────────────────────
def ideal_table(plans, rrp, target=None):
    """The centre's members spread to the target mix, each at RRP or the centre's own
    (higher) average price per session. Composition: current %, ideal %, variance."""
    target = target or TARGET_MIX
    n = int(plans["Members"].sum())
    current_value = float((plans["Members"] * plans["Weekly price"]).sum())
    nr = plans[[not_recommended(k, s) for k, s in zip(plans["Type"], plans["Sessions/week"])]]
    n_rec = n - int(nr["Members"].sum())             # only these are spread to the ideal mix
    rows, comp = [], []
    for kind, plan, lo, hi, s in IDEAL_GROUPS:
        g = plans[(plans["Type"] == kind) & plans["Sessions/week"].between(lo, hi)]
        priced = g[~g["Freeze"]]
        own = None
        if priced["Members"].sum():
            own = (priced["Members"] * priced["Weekly price"]).sum() / \
                (priced["Members"] * priced["Sessions/week"]).sum() * s
        r = rrp_price(rrp, kind, s)
        price = max(x for x in (r, own) if x is not None) if (r or own) else None
        m = n_rec * target[plan] / 100
        rows.append({"Type": kind, "Plan": plan, "Members": m, "Sessions/week": s,
                     "Total sessions": m * s,
                     "Membership price": price, "Total value": m * price if price else None,
                     "Revenue per session": price / s if price else None,
                     "Own price used": bool(own and r and own > r + 0.005),
                     "Not recommended": False})
        cur = g["Members"].sum() / n * 100 if n else 0
        ideal_pct = target[plan] * n_rec / n if n else target[plan]
        comp.append({"Plan": f"Group {plan}" if kind == "Group" else plan,
                     "Current %": cur, "Ideal %": ideal_pct, "Variance": cur - ideal_pct})
    if len(nr):                                      # kept as they are, at today's price
        m_nr = float(nr["Members"].sum())
        v_nr = float((nr["Members"] * nr["Weekly price"]).sum())
        s_nr = float((nr["Members"] * nr["Sessions/week"]).sum())
        rows.append({"Type": "Group", "Plan": NOT_RECOMMENDED, "Members": m_nr,
                     "Sessions/week": s_nr / m_nr, "Total sessions": s_nr,
                     "Membership price": v_nr / m_nr, "Total value": v_nr,
                     "Revenue per session": v_nr / s_nr if s_nr else None,
                     "Own price used": False, "Not recommended": True})
        cur = m_nr / n * 100 if n else 0
        comp.append({"Plan": f"Group {NOT_RECOMMENDED}", "Current %": cur, "Ideal %": cur,
                     "Variance": 0.0})
    t = pd.DataFrame(rows)
    total = float(t["Total value"].sum())
    sessions = float(t["Total sessions"].sum())
    return {"rows": t, "members": n, "total": total, "sessions": sessions,
            "per_session": total / sessions if sessions else None,
            "arpm": total / n if n else None, "vs_today": total - current_value,
            "composition": pd.DataFrame(comp)}


# ── Scheduling ────────────────────────────────────────────────────────────────
def target_fill(members):
    """Target students per study hub for a centre this size (1 decimal)."""
    pts = FILL_POINTS
    if members <= pts[0][0]:
        return pts[0][1]
    for (m0, f0), (m1, f1) in zip(pts, pts[1:]):
        if members <= m1:
            return math.floor((f0 + (f1 - f0) * (members - m0) / (m1 - m0)) * 10 + 0.5 + 1e-9) / 10
    return pts[-1][1]


def scheduling(members, group_sessions, one_sessions, value, country, fill=None):
    """Hubs needed a week at the target fill, tutor cost, and the margin after tutor cost.
    Tutor cost and margin are None where no tutor rate is set for the country."""
    fill = fill or target_fill(members)
    hubs = group_sessions / fill + one_sessions
    rate = TUTOR_HOUR.get(country)
    cost = hubs * rate if rate else None
    margin = value - cost if cost is not None else None
    return {"Target fill": fill, "Hubs needed a week": hubs, "Tutor cost": cost,
            "Tutor cost %": cost / value * 100 if cost is not None and value else None,
            "Margin a week": margin,
            "Margin per member a week": margin / members if margin is not None and members
            else None}


# ── Trends ────────────────────────────────────────────────────────────────────
MIX = ["1×", "2×", "3+", "1:1"]


def _bucket(kind, sessions):
    return "1:1" if kind == "1:1" else "1×" if sessions <= 1 else "2×" if sessions == 2 \
        else "3+"


def weekly_totals(comp):
    """One row per (week, centre): members, sessions, 1:1 sessions, total value, margin a
    week after tutor cost at the centre's target fill, and members on each plan type."""
    if comp is None or comp.empty:
        return pd.DataFrame()
    c = comp.assign(Value=comp["Members"] * comp["Weekly price"],
                    Sess=comp["Members"] * comp["Sessions/week"],
                    Bucket=[_bucket(k, s) for k, s in zip(comp["Type"], comp["Sessions/week"])])
    c["One"] = c["Sess"].where(c["Type"] == "1:1", 0)
    keys = ["Week ending", "Location", "Country"]
    out = c.groupby(keys).agg(Members=("Members", "sum"), Sessions=("Sess", "sum"),
                              One=("One", "sum"), Value=("Value", "sum")).reset_index()
    mix = c.pivot_table(index=keys, columns="Bucket", values="Members", aggfunc="sum",
                        fill_value=0).reindex(columns=MIX, fill_value=0).reset_index()
    out = out.merge(mix, on=keys)
    out["Margin"] = [scheduling(n, s - o, o, v, cc)["Margin a week"] for n, s, o, v, cc in
                     zip(out["Members"], out["Sessions"], out["One"], out["Value"],
                         out["Country"])]
    return out


def trend_measures(t):
    """Per-row measures from (summed) weekly totals: members, sessions, sessions per member,
    ARPM, revenue per session, margin per member and the plan mix in %."""
    n = t["Members"].where(t["Members"] > 0)
    out = pd.DataFrame({"Members": t["Members"], "Total sessions/week": t["Sessions"],
                        "Sessions per member": t["Sessions"] / n, "ARPM": t["Value"] / n,
                        "Revenue per session": t["Value"] / t["Sessions"].where(
                            t["Sessions"] > 0),
                        "Margin per member a week": t["Margin"] / n}, index=t.index)
    for k in MIX:
        out[k] = t[k] / n * 100
    return out


# ── Measured: study hubs run (CRM weekly Revenue export) ───────────────────────
CHECK_GAP = 0.20     # members vs CRM active students differ by more than this: data check


def hubs_run(revenue, week_end):
    """{location: (sessions run, active students)} for the week ending week_end, from the
    Revenue tab (Date = week ending; 'Total Sessions'; '# Active Students')."""
    if revenue is None or revenue.empty or "Date" not in revenue or LOC not in revenue:
        return {}
    r = revenue[pd.to_datetime(revenue["Date"]).dt.normalize() ==
                pd.Timestamp(week_end).normalize()]
    return {loc: (float(s), float(a)) for loc, s, a in
            zip(r[LOC], r["Total Sessions"], r["# Active Students"])}


def measured(row, run):
    """Hubs run, students per hub and extra hubs vs the target, for one group_table row.
    run: (sessions run, active students) or None. Measured figures are None (and Check
    says why) where the Revenue row is missing or its students don't match the members."""
    out = {"Hubs run a week": None, "Students per hub": None, "Extra hubs a week": None,
           "Check": ""}
    if not run:
        out["Check"] = "No sessions-run data"
        return out
    hubs, active = run
    members, one = row["Members"], row.get("1:1 sessions", 0) or 0
    if active and abs(members - active) / active > CHECK_GAP:
        out["Check"] = f"Members {members:.0f} vs CRM students {active:.0f}"
        return out
    out["Hubs run a week"] = hubs
    if hubs > one:
        out["Students per hub"] = (row["Total sessions/week"] - one) / (hubs - one)
    out["Extra hubs a week"] = hubs - row["Hubs needed a week"]
    return out


def add_measured(t, revenue, week_end):
    """group_table plus the measured columns for that week."""
    if t is None or t.empty:
        return t
    runs = hubs_run(revenue, week_end)
    m = pd.DataFrame([measured(r, runs.get(r["Location"])) for _, r in t.iterrows()],
                     index=t.index)
    return pd.concat([t, m], axis=1)


# ── Today vs ideal: where the money is ───────────────────────────────────────
def where_the_money_is(plans, rrp, ideal, country, run=None):
    """Today vs the ideal composition, after tutor cost, and the gain split into price, mix
    and scheduling (weekly, in the centre's currency).

    Fill: the target for the centre's size, or the centre's measured students per hub where
    that is higher (a centre already beating the target isn't asked to do worse).
    Today's tutor cost uses the hubs actually run where measured, else the hubs needed.
    - price: today's members, each plan priced at least at RRP (prices under $10 kept);
    - mix: moving to the ideal composition, less the change in tutor cost;
    - scheduling: hubs run above those needed at that fill (never negative).
    Tutor cost and margins are None where the country has no tutor rate."""
    n = int(plans["Members"].sum())
    one_t = plans["Type"] == "1:1"
    sess = float((plans["Members"] * plans["Sessions/week"]).sum())
    one = float((plans.loc[one_t, "Members"] * plans.loc[one_t, "Sessions/week"]).sum())
    v0 = float((plans["Members"] * plans["Weekly price"]).sum())
    v1 = 0.0
    for _, p in plans.iterrows():
        r = rrp_price(rrp, p["Type"], int(p["Sessions/week"])) \
            if not p["Freeze"] and not not_recommended(p["Type"], p["Sessions/week"]) else None
        v1 += p["Members"] * max(p["Weekly price"], r or 0)
    rows = ideal["rows"]
    v2 = float(ideal["total"])
    sess2 = float(rows["Total sessions"].sum())
    one2 = float(rows.loc[rows["Type"] == "1:1", "Total sessions"].sum())
    target = target_fill(n)
    hubs_run = run[0] if run else None
    measured = (sess - one) / (hubs_run - one) if hubs_run and hubs_run > one else None
    fill = max(target, measured or 0)
    need0 = (sess - one) / fill + one if fill else 0.0
    need2 = (sess2 - one2) / fill + one2 if fill else 0.0
    hubs0 = hubs_run if hubs_run else need0
    rate = TUTOR_HOUR.get(country)
    cost0 = hubs0 * rate if rate else None
    cost2 = need2 * rate if rate else None

    def side(v, s, hubs, cost):
        return {"Revenue per session": v / s if s else None,
                "Sessions per member": s / n if n else None, "ARPM": v / n if n else None,
                "Hubs a week": hubs, "Tutor cost %": cost / v * 100 if cost and v else None,
                "Margin a week": v - cost if cost is not None else None,
                "Margin per member a week": (v - cost) / n if cost is not None and n
                else None, "Total value": v}
    mix = (v2 - v1) - ((need2 - need0) * rate if rate else 0)
    use_mix = mix > 0.5
    gains = [("Price", v1 - v0,
              "Bring every plan up to at least RRP (today's members and plans)."),
             ("Mix", mix if use_mix else 0.0,
              ("Move members towards the ideal mix"
               + (", after the change in tutor time." if rate else "."))
              if use_mix else "Today's mix already earns more than the ideal mix: keep it.")]
    if hubs_run is not None and rate:
        extra = max(hubs_run - need0, 0)
        gains.append(("Scheduling", extra * rate,
                      f"Run about {need0:.0f} hubs a week instead of {hubs_run:.0f} "
                      f"({fill:.1f} students a hub)." if extra >= 0.5 else
                      f"Already at or under the hubs needed at {fill:.1f} students a hub."))
    # Plans: members on 4x+ group plans moved onto 2x at the ideal 2x price.
    nrs = _nr_summary(plans, n)
    p2 = rows.loc[rows["Plan"] == "2×", "Membership price"]
    p2 = float(p2.iloc[0]) if len(p2) and pd.notna(p2.iloc[0]) else None
    plans_gain = None
    if nrs["members"] and p2:
        fees = nrs["members"] * p2 - nrs["value"]
        saved = nrs["sessions"] - 2 * nrs["members"]          # sessions a week saved
        freed = saved / fill if fill else 0.0                  # group hubs a week freed
        net = fees + (freed * rate if rate else 0)
        plans_gain = {"net": net, "fees": fees, "sessions": saved, "hubs": freed}
        gains.append(("Plans", max(net, 0.0),
                      f"Move the {nrs['members']} members on {NOT_RECOMMENDED_FROM}×+ plans onto "
                      f"2× at ${p2:,.2f}: {saved:.0f} fewer sessions (~{freed:.0f} hubs freed), "
                      f"fees {'+' if fees >= 0 else '−'}${abs(fees):,.0f}"
                      + (f", tutor cost −${freed * rate:,.0f}" if rate else "")
                      + ("." if net > 0.5 else f". Net {'+' if net >= 0 else '−'}"
                         f"${abs(net):,.0f}: worth it for the capacity, not the fees.")))
    use_plans = bool(plans_gain and plans_gain["net"] > 0.5)
    if use_mix:
        bv, bs, bh = v2, sess2, need2
    else:
        bv, bs, bh = v1, sess, need0
    if use_plans:
        bv += plans_gain["fees"]
        bs -= plans_gain["sessions"]
        bh -= plans_gain["hubs"]
    best = side(bv, bs, bh, bh * rate if rate else None)
    # Theoretical: every member on the recommended mix, at the ideal prices and target fill.
    shares = {"1×": (1, 0.15), "2×": (2, 0.80), "3+": (3, 0.01), "1:1": (1, 0.04)}
    shares = {k: (s_, TARGET_MIX[k] / 100) for k, (s_, _) in shares.items()}
    price = dict(zip(rows["Plan"], rows["Membership price"]))
    tv = sum(n * sh * float(price.get(k) or 0) for k, (_, sh) in shares.items())
    ts = sum(n * sh * s_ for k, (s_, sh) in shares.items())
    t_one = n * shares["1:1"][1]
    th = (ts - t_one) / target + t_one if target else 0.0
    theory = side(tv, ts, th, th * rate if rate else None)
    return {"today": side(v0, sess, hubs0, cost0), "ideal": best, "use_mix": use_mix,
            "use_plans": use_plans, "plans": plans_gain, "theory": theory,
            "gains": sorted(gains, key=lambda g: -g[1]), "fill": fill, "target": target,
            "measured_fill": measured, "hubs_needed": need0, "hubs_run": hubs_run,
            "rate": rate, "not_recommended": nrs,
            "total_sessions": sess, "total_value": v0}


def _nr_summary(plans, n):
    nr = plans[[not_recommended(k, s) for k, s in zip(plans["Type"], plans["Sessions/week"])]]
    m = int(nr["Members"].sum())
    return {"members": m, "share": m / n * 100 if n else 0,
            "value": float((nr["Members"] * nr["Weekly price"]).sum()),
            "sessions": float((nr["Members"] * nr["Sessions/week"]).sum())}


# ── Group overview ────────────────────────────────────────────────────────────
def group_table(week, vl, table):
    """One row per centre: members, ARPM, revenue per session, % of RRP and ARPM vs the
    country median (None where the country has one centre). Lowest % of RRP first."""
    region = {}
    if vl is not None and not vl.empty and LOC in vl.columns and "Region" in vl.columns:
        region = dict(zip(vl[LOC], vl["Region"]))
    rows = []
    for loc, g in week.groupby("Location"):
        cc = g["Country"].iloc[0]
        rrp = rrp_for(table, cc, region.get(loc))
        n = int(g["Members"].sum())
        value = float((g["Members"] * g["Weekly price"]).sum())
        sessions = float((g["Members"] * g["Sessions/week"]).sum())
        nr = [not_recommended(k, s) for k, s in zip(g["Type"], g["Sessions/week"])]
        priced = g[~g["Freeze"] & ~pd.Series(nr, index=g.index)]
        at = sum(m * rrp_price(rrp, t, s) for m, t, s in
                 zip(priced["Members"], priced["Type"], priced["Sessions/week"])) if rrp else 0
        paid = float((priced["Members"] * priced["Weekly price"]).sum())
        one = g["Type"] == "1:1"
        one_sess = float((g.loc[one, "Members"] * g.loc[one, "Sessions/week"]).sum())
        rows.append({"Location": loc, "Centre": loc.removeprefix(ingest.PREFIX),
                     "Country": cc, "Currency": CURRENCY.get(cc, ""),
                     "RRP used": rrp["label"] if rrp else "No RRP set", "Members": n,
                     "Total sessions/week": sessions, "1:1 sessions": one_sess,
                     "Sessions per member": sessions / n if n else None,
                     "ARPM": value / n if n else None,
                     "Revenue per session": value / sessions if sessions else None,
                     **scheduling(n, sessions - one_sess, one_sess, value, cc),
                     "% of RRP": paid / at * 100 if at else None,
                     "Members on 4×+": int(g.loc[nr, "Members"].sum()) if any(nr) else 0})
    t = pd.DataFrame(rows)
    if t.empty:
        return t
    med = t.groupby("Country")["ARPM"].transform("median")
    count = t.groupby("Country")["Centre"].transform("count")
    t["ARPM vs country median"] = (t["ARPM"] - med).where(count > 1)
    return t.sort_values("% of RRP", na_position="last").reset_index(drop=True)
