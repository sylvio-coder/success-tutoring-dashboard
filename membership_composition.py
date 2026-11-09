"""Membership Composition: each centre's plans and prices against RRP, from the portal's
weekly Membership Composition export. Calculations only; the page is composition_report.py.

Rules agreed with HQ:
- RRP is national for Australia, New Zealand and Canada; the US is priced by region.
- Prices under $10 a week (a freeze fee or discretionary amount) are left out of the plan
  rows but kept in totals, ARPM and composition, with a note.
- US prices are exported per month; weekly = monthly / 4 (as in the RRP guide).
- 4+ sessions a week: RRP = the 3x rate per session x sessions. 1:1: RRP per session.
- Ideal composition: 15% 1x, 80% 2x, 1% 3+ and 4% 1:1 of the centre's members, at RRP or
  the centre's own price where that is higher.
"""
import csv
import io
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
MONTHLY_BILLED = {"US"}        # exported per month; weekly = monthly / 4
REGIONAL = {"US", "CA"}        # countries priced by region (others: national RRP only)
TARGET_MIX = {"1×": 15, "2×": 80, "3+": 1, "1:1": 4}
IDEAL_GROUPS = [("Group", "1×", 1, 1, 1), ("Group", "2×", 2, 2, 2),
                ("Group", "3+", 3, 99, 3), ("1:1", "1:1", 1, 99, 1)]   # type, plan, min, max, sessions
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
                     "Weekly price": round(price / 4 if cc in MONTHLY_BILLED else price, 2)})
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
                     "Membership price": price, "Total value": p["Members"] * price,
                     "Revenue per session": price / s, "RRP": r, "RRP variance": rv,
                     "RRP variance %": rp, "Median": med if cnt > 1 else None,
                     "Median variance": mv, "Median variance %": mp})
    comp = (plans.assign(Plan=[plan_label(t, s) for t, s in
                               zip(plans["Type"], plans["Sessions/week"])],
                         order=[(t == "1:1", s) for t, s in
                                zip(plans["Type"], plans["Sessions/week"])])
            .groupby("Plan").agg(Members=("Members", "sum"), order=("order", "min"))
            .sort_values("order").reset_index())
    comp["% of members"] = comp["Members"] / n * 100 if n else 0
    frozen = plans[plans["Freeze"]]
    return {"rows": pd.DataFrame(rows), "members": n, "total": value,
            "per_session": value / sessions if sessions else None,
            "arpm": value / n if n else None,
            "composition": comp[["Plan", "Members", "% of members"]],
            "freeze_members": int(frozen["Members"].sum()),
            "freeze_value": float((frozen["Members"] * frozen["Weekly price"]).sum())}


# ── Table 2: ideal composition ────────────────────────────────────────────────
def ideal_table(plans, rrp, target=None):
    """The centre's members spread to the target mix, each at RRP or the centre's own
    (higher) average price per session. Composition: current %, ideal %, variance."""
    target = target or TARGET_MIX
    n = int(plans["Members"].sum())
    current_value = float((plans["Members"] * plans["Weekly price"]).sum())
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
        m = n * target[plan] / 100
        rows.append({"Type": kind, "Plan": plan, "Members": m, "Sessions/week": s,
                     "Membership price": price, "Total value": m * price if price else None,
                     "Revenue per session": price / s if price else None,
                     "Own price used": bool(own and r and own > r + 0.005)})
        cur = g["Members"].sum() / n * 100 if n else 0
        comp.append({"Plan": f"Group {plan}" if kind == "Group" else plan,
                     "Current %": cur, "Ideal %": target[plan],
                     "Variance": cur - target[plan]})
    t = pd.DataFrame(rows)
    total = float(t["Total value"].sum())
    sessions = float((t["Members"] * t["Sessions/week"]).sum())
    return {"rows": t, "members": n, "total": total,
            "per_session": total / sessions if sessions else None,
            "arpm": total / n if n else None, "vs_today": total - current_value,
            "composition": pd.DataFrame(comp)}


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
        priced = g[~g["Freeze"]]
        at = sum(m * rrp_price(rrp, t, s) for m, t, s in
                 zip(priced["Members"], priced["Type"], priced["Sessions/week"])) if rrp else 0
        paid = float((priced["Members"] * priced["Weekly price"]).sum())
        rows.append({"Location": loc, "Centre": loc.removeprefix(ingest.PREFIX),
                     "Country": cc, "Currency": CURRENCY.get(cc, ""),
                     "RRP used": rrp["label"] if rrp else "No RRP set", "Members": n,
                     "ARPM": value / n if n else None,
                     "Revenue per session": value / sessions if sessions else None,
                     "% of RRP": paid / at * 100 if at else None})
    t = pd.DataFrame(rows)
    if t.empty:
        return t
    med = t.groupby("Country")["ARPM"].transform("median")
    count = t.groupby("Country")["Centre"].transform("count")
    t["ARPM vs country median"] = (t["ARPM"] - med).where(count > 1)
    return t.sort_values("% of RRP", na_position="last").reset_index(drop=True)
