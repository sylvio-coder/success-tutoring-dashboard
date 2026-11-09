"""Membership Composition: each centre's plan mix (1x, 2x, 3x, 4+x a week, 1:1) and
pricing against its regional RRP, from the portal's weekly Membership Composition
export. Calculations only; the report page is in composition_report.py.

Rules agreed with HQ:
- Prices under $10 a week are a freeze fee or a discretionary amount: they count in
  the mix but not in prices or RRP figures, and are listed to check.
- US prices are exported per month; weekly = monthly / 4 (as in the RRP guide).
- 4+ sessions a week: RRP = the 3x rate per session x sessions; flagged to phase out.
- 1:1: RRP per session x sessions.
- Pathway to RRP: 1x members stay on 1x (more 2x needs new members); 1:1 and 4+ are
  left as they are; foundational pricing ends 12 months after opening; other prices
  rise each 1 July by at most 5%, never above RRP.
"""
import csv
import io
import math
from datetime import date

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
REVIEW_MONTH = 7               # price review each 1 July...
REVIEW_CAP = 0.05              # ...by no more than 5%
FOUNDATIONAL_MONTHS = 12       # foundational pricing ends 12 months after opening
YOUNG_WEEKS = 18               # centres this young or younger are shown for information
OUTLIER_SD = 2                 # flag centres 2 standard deviations from the median
TOP_GAPS = 5                   # always flag the 5 biggest RRP gaps
SMALL_SAMPLE = 8
NEW_MEMBER_WEEKS = 8           # recent sign-up rate = average new members over 8 weeks
MIX = ["1×", "2×", "3×", "4+×", "1:1"]
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


# Market benchmark and foundational (pre-launch) weekly prices from 1 July 2026.
# 1:1 is $90 a session in Australia, scaled elsewhere by each region's 2x RRP.
# Region "All" applies where no row matches the centre's state or region.
DEFAULT_RRP = [
    _rrp("AU", "All", (42, 74, 100), (36, 63, 85), 90, "National RRP (QLD/SA)"),
    _rrp("AU", "New South Wales", (46, 81, 110), (39, 69, 93), 99, "Premium"),
    _rrp("AU", "Australian Capital Territory", (46, 81, 110), (39, 69, 93), 99, "Premium"),
    _rrp("AU", "Victoria", (44, 77, 104), (37, 65, 88), 94, "Mid-premium"),
    _rrp("AU", "Western Australia", (44, 77, 104), (37, 65, 88), 94, "Mid-premium"),
    _rrp("AU", "Queensland", (42, 74, 100), (36, 63, 85), 90, "Mid-market"),
    _rrp("AU", "South Australia", (42, 74, 100), (36, 63, 85), 90, "Mid-market"),
    _rrp("AU", "Tasmania", (40, 71, 96), (34, 60, 82), 86, "Value"),
    _rrp("AU", "Northern Territory", (40, 71, 96), (34, 60, 82), 86, "Value"),
    _rrp("NZ", "All", (47, 82, 110), (40, 70, 94), 100,
         "Auckland partners may add $3-5 at their discretion"),
    _rrp("CA", "All", (51, 89, 121), (43, 76, 103), 108, "Foundational: 15% below RRP"),
    _rrp("US", "All", (47, 82, 110), (40, 70, 94), 100, "National US RRP; billed monthly"),
    _rrp("US", "Texas", (49, 85, 115), (42, 72, 98), 103, "Austin & Houston"),
]


def country_code(value):
    v = str(value or "").strip()
    return v.upper() if v.upper() in CURRENCY else COUNTRY_CODE.get(v.lower(), v.upper())


def _region_key(region):
    return ingest.state_abbr(region) or str(region or "").strip().lower()


def rrp_table(df_rrp=None):
    """{(country, region key): {mkt, fnd, one, label}} from the RRP tab (or defaults)."""
    rows = DEFAULT_RRP if df_rrp is None or df_rrp.empty else df_rrp.to_dict("records")
    out = {}
    for r in rows:
        try:
            num = lambda h: float(str(r.get(h, "")).replace("$", "").replace(",", ""))
            entry = dict(mkt=tuple(num(f"{n}x") for n in (1, 2, 3)),
                         fnd=tuple(num(f"Foundational {n}x") for n in (1, 2, 3)),
                         one=num("1:1 per session"),
                         label=f"{country_code(r['Country'])} · {r['Region']}")
        except (ValueError, KeyError):
            continue
        out[(country_code(r["Country"]), _region_key(r["Region"]))] = entry
    return out


def rrp_for(table, country, region):
    """The RRP row for this centre: its region if listed, else the country's 'All'."""
    cc = country_code(country)
    return table.get((cc, _region_key(region))) or table.get((cc, "all"))


def rrp_price(rrp, kind, sessions, foundational=False):
    """Weekly RRP for a plan. 4+ a week: the 3x rate per session x sessions."""
    if rrp is None or not sessions:
        return None
    if kind == "1:1":
        # No foundational 1:1 price: discounted like the 2x plan (AU logic).
        scale = rrp["fnd"][1] / rrp["mkt"][1] if foundational and rrp["mkt"][1] else 1
        return rrp["one"] * sessions * scale
    tier = rrp["fnd"] if foundational else rrp["mkt"]
    return tier[sessions - 1] if sessions <= 3 else tier[2] / 3 * sessions


def mix_of(kind, sessions):
    if kind == "1:1":
        return "1:1"
    return {1: "1×", 2: "2×", 3: "3×"}.get(int(sessions), "4+×")


# ── Reading the export ────────────────────────────────────────────────────────
def _num(v):
    try:
        return float(str(v).replace("$", "").replace(",", "").strip())
    except ValueError:
        return None


def read_export(data):
    """The portal CSV (bytes or text) -> {week_end, rows, empty}.
    rows: one per plan with a weekly price; empty: centres listed with no members."""
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
        if not members or sessions is None or price is None:
            continue
        cc = country_code(get(r, "country"))
        weekly = price / 4 if cc in MONTHLY_BILLED else price
        rows.append({"Centre": centre, "Country": cc,
                     "Type": "1:1" if get(r, "type").replace(" ", "") == "1:1" else "Group",
                     "Sessions/week": int(sessions), "Members": int(members),
                     "Weekly price": round(weekly, 2)})
    if not centres:
        raise CompositionError("The file has no centres in it.")
    week = pd.to_datetime(sorted(dates)[-1], dayfirst=False, errors="coerce") if dates else None
    if week is not None and pd.isna(week):
        week = pd.to_datetime(sorted(dates)[-1], dayfirst=True, errors="coerce")
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
    """The sheet tab, typed, with Mix and Freeze columns."""
    cols = ["Week ending", "Location", "Country", "Type", "Sessions/week", "Members",
            "Weekly price", "Mix", "Freeze"]
    if df is None or df.empty or not {"Week ending", "Location"} <= set(df.columns):
        return pd.DataFrame(columns=cols)
    out = df.copy()
    out["Week ending"] = pd.to_datetime(out["Week ending"].astype(str).str.lstrip("'"),
                                        format="%Y-%m-%d", errors="coerce")
    for c in ["Sessions/week", "Members", "Weekly price"]:
        out[c] = pd.to_numeric(out[c].astype(str).str.replace("[$,]", "", regex=True),
                               errors="coerce")
    out = out.dropna(subset=["Week ending", "Sessions/week", "Members", "Weekly price"])
    out["Sessions/week"] = out["Sessions/week"].astype(int)
    out["Country"] = out["Country"].map(country_code)
    out["Mix"] = [mix_of(t, s) for t, s in zip(out["Type"], out["Sessions/week"])]
    out["Freeze"] = out["Weekly price"] < FREEZE_BELOW
    return out[cols]


# ── One centre, one week ──────────────────────────────────────────────────────
def centre_metrics(plans, rrp):
    """Mix, prices and RRP figures for one centre's plans in one week."""
    n = int(plans["Members"].sum())
    m = {"Members": n}
    for k in MIX:
        sub = plans[plans["Mix"] == k]
        cnt = int(sub["Members"].sum())
        m[f"n {k}"] = cnt
        m[f"% {k}"] = 100 * cnt / n if n else 0.0
        priced = sub[~sub["Freeze"]]
        pm = priced["Members"].sum()
        m[f"Avg {k}"] = (priced["Members"] * priced["Weekly price"]).sum() / pm if pm else None
    priced = plans[~plans["Freeze"]]
    pm = int(priced["Members"].sum())
    actual = float((priced["Members"] * priced["Weekly price"]).sum())
    m["Freeze"] = int(plans.loc[plans["Freeze"], "Members"].sum())
    m["ARPM"] = actual / pm if pm else None
    if rrp is None or not pm:
        m.update({"% of RRP": None, "Below RRP": None, "Below foundational": None,
                  "RRP gap": None, "At RRP value": None})
        return m
    at_rrp = below = below_f = gap = 0.0
    for _, p in priced.iterrows():
        r = rrp_price(rrp, p["Type"], p["Sessions/week"])
        f = rrp_price(rrp, p["Type"], p["Sessions/week"], foundational=True)
        at_rrp += p["Members"] * r
        if p["Weekly price"] < r - 0.005:
            below += p["Members"]
            gap += p["Members"] * (r - p["Weekly price"])
        if p["Weekly price"] < f - 0.005:
            below_f += p["Members"]
    m.update({"% of RRP": 100 * actual / at_rrp if at_rrp else None,
              "Below RRP": int(below), "% below RRP": 100 * below / pm,
              "Below foundational": int(below_f), "RRP gap": gap, "At RRP value": at_rrp})
    return m


def _vl_info(vl):
    if vl is None or vl.empty or LOC not in vl.columns:
        return {}
    return {r[LOC]: r for r in vl.to_dict("records")}


def summary(comp, vl, table, age_group=None, fx=None, week=None):
    """One row per centre for a week (default: each centre's latest week)."""
    if comp.empty:
        return pd.DataFrame()
    info = _vl_info(vl)
    fx = fx or {}
    rows = []
    for loc, g in comp.groupby("Location"):
        wk = week if week is not None else g["Week ending"].max()
        plans = g[g["Week ending"] == wk]
        if plans.empty:
            continue
        cc = plans["Country"].iloc[0]
        v = info.get(loc, {})
        rrp = rrp_for(table, cc, v.get("Region"))
        weeks = _num(v.get("Weeks old"))
        months = _num(v.get("Age (Months)"))
        m = centre_metrics(plans, rrp)
        cur = CURRENCY.get(cc, "")
        rate = 1.0 if cur == "AUD" else fx.get(cur)
        m.update({"Location": loc, "Centre": loc.removeprefix(ingest.PREFIX),
                  "Country": cc, "Currency": cur, "Region": v.get("Region", ""),
                  "RRP region": rrp["label"] if rrp else "No RRP set",
                  "Week ending": wk, "Weeks old": weeks, "Age (Months)": months,
                  "Age group": age_group(months) if age_group and months is not None
                  else "Unknown",
                  "Young": weeks is not None and weeks <= YOUNG_WEEKS,
                  "RRP gap AUD": None if m["RRP gap"] is None or rate is None
                  else m["RRP gap"] * rate})
        rows.append(m)
    return pd.DataFrame(rows)


def history(comp, loc, rrp):
    """Week-by-week mix % and average price for one centre."""
    g = comp[comp["Location"] == loc]
    rows = []
    for wk, plans in g.groupby("Week ending"):
        m = centre_metrics(plans, rrp)
        m["Week ending"] = wk
        rows.append(m)
    return pd.DataFrame(rows).sort_values("Week ending") if rows else pd.DataFrame()


# ── Benchmarks and flags ──────────────────────────────────────────────────────
BENCH_METRICS = ["% 1×", "% 2×", "% of RRP", "% below RRP"]


def benchmarks(summ):
    """Median of each % measure by age group, from centres older than 18 weeks
    (or every centre in the group when none is). {age group: {metric: median, 'n': count}}"""
    out = {}
    if summ.empty:
        return out
    for grp, g in summ.groupby("Age group"):
        est = g[~g["Young"]]
        use = est if len(est) else g
        out[grp] = {k: use[k].dropna().median() for k in BENCH_METRICS if k in use}
        out[grp]["n"] = len(use)
    return out


def flags(summ):
    """Why each centre stands out: {location: [reasons]}.
    - Outside 2 standard deviations of the median (the bad side) on % of RRP, % 2x,
      % 1x or % below RRP, among centres older than 18 weeks.
    - One of the 5 biggest weekly RRP gaps (in AUD)."""
    out = {}
    est = summ[~summ["Young"]] if not summ.empty else summ
    worse = {"% of RRP": -1, "% 2×": -1, "% 1×": 1, "% below RRP": 1}
    names = {"% of RRP": "% of RRP", "% 2×": "share on 2×", "% 1×": "share on 1×",
             "% below RRP": "share below RRP"}
    for k, sign in worse.items():
        vals = est[k].dropna() if k in est else pd.Series(dtype=float)
        if len(vals) < 3:
            continue
        med, sd = vals.median(), vals.std()
        if not sd or math.isnan(sd):
            continue
        for _, r in est.iterrows():
            v = r.get(k)
            if v is not None and not pd.isna(v) and sign * (v - med) > OUTLIER_SD * sd:
                out.setdefault(r["Location"], []).append(
                    f"{names[k]} {v:.0f}% vs median {med:.0f}% (beyond {OUTLIER_SD} SD)")
    gaps = summ.dropna(subset=["RRP gap AUD"]) if not summ.empty else summ
    gaps = gaps[gaps["RRP gap AUD"] > 0].nlargest(TOP_GAPS, "RRP gap AUD") \
        if not gaps.empty else gaps
    for i, (_, r) in enumerate(gaps.iterrows(), 1):
        out.setdefault(r["Location"], []).append(f"#{i} biggest RRP gap")
    return out


# ── Pathway to RRP ────────────────────────────────────────────────────────────
def next_july(today):
    y = today.year + (1 if today >= date(today.year, REVIEW_MONTH, 1) else 0)
    return date(y, REVIEW_MONTH, 1)


def _add_months(d, months):
    y, m = divmod(d.month - 1 + int(months), 12)
    return date(d.year + y, m + 1, min(d.day, 28))


def pathway(plans, rrp, age_months=None, new_per_week=None, target_2x=None, today=None,
            max_reviews=10):
    """Steps to bring this centre's prices to RRP, and new 2x members to lift its mix.
    Only group plans of 1-3 sessions change; 1:1, 4+ and freeze/discretionary rows stay
    as they are. Returns {steps, projection, new_2x, weeks_to_2x, target_2x}."""
    today = today or date.today()
    if rrp is None:
        return {"steps": [], "projection": [], "new_2x": 0, "weeks_to_2x": None,
                "target_2x": target_2x}
    p = plans[~plans["Freeze"]].copy()
    p["RRP"] = [rrp_price(rrp, t, s) for t, s in zip(p["Type"], p["Sessions/week"])]
    p["Fnd"] = [rrp_price(rrp, t, s, True) for t, s in zip(p["Type"], p["Sessions/week"])]
    p["Moves"] = (p["Type"] == "Group") & (p["Sessions/week"] <= 3)
    value = lambda: float((p["Members"] * p["Price"]).sum())
    p["Price"] = p["Weekly price"]
    members = int(p["Members"].sum())
    at_rrp = float((p["Members"] * p["RRP"]).sum())
    proj = [{"When": "Today", "Date": today, "ARPM": value() / members if members else None,
             "% of RRP": 100 * value() / at_rrp if at_rrp else None}]
    steps = [{"Step": "Price every new member at RRP", "When": "From now",
              "Members": None, "Weekly gain": None,
              "Detail": f"2× at {rrp['mkt'][1]:,.0f}, 1× at {rrp['mkt'][0]:,.0f}, "
                        f"3× at {rrp['mkt'][2]:,.0f} a week"}]

    # Foundational pricing ends 12 months after opening (not capped at 5%).
    if age_months is not None and age_months < FOUNDATIONAL_MONTHS:
        end = _add_months(today, FOUNDATIONAL_MONTHS - age_months)
        f = p["Moves"] & (p["Price"] <= p["Fnd"] + 0.005) & (p["Price"] < p["RRP"] - 0.005)
        if f.any():
            before = value()
            p.loc[f, "Price"] = p.loc[f, "RRP"]
            steps.append({"Step": "End foundational pricing", "When": end.strftime("%b %Y"),
                          "Members": int(p.loc[f, "Members"].sum()),
                          "Weekly gain": value() - before,
                          "Detail": "Members on foundational rates move to RRP 12 months "
                                    "after opening"})
            proj.append({"When": "Foundational ends", "Date": end,
                         "ARPM": value() / members, "% of RRP": 100 * value() / at_rrp})

    # Each 1 July: +5% at most, never above RRP.
    review = next_july(today)
    for _ in range(max_reviews):
        low = p["Moves"] & (p["Price"] < p["RRP"] - 0.005)
        if not low.any():
            break
        before = value()
        p.loc[low, "Price"] = (p.loc[low, "Price"] * (1 + REVIEW_CAP)).clip(
            upper=p.loc[low, "RRP"])
        steps.append({"Step": f"July price review (up to {REVIEW_CAP:.0%})",
                      "When": review.strftime("%b %Y"), "Members": int(p.loc[low, "Members"].sum()),
                      "Weekly gain": value() - before,
                      "Detail": "Below-RRP group plans up by 5%, capped at RRP"})
        proj.append({"When": f"After {review.strftime('%b %Y')} review", "Date": review,
                     "ARPM": value() / members, "% of RRP": 100 * value() / at_rrp})
        review = date(review.year + 1, REVIEW_MONTH, 1)

    # More 2x needs new members (1x members rarely change once onboarded).
    n_all = int(plans["Members"].sum())
    two = int(plans.loc[plans["Mix"] == "2×", "Members"].sum())
    new_2x, weeks = 0, None
    if target_2x and n_all and two / n_all * 100 < target_2x < 100:
        t = target_2x / 100
        new_2x = math.ceil(round((t * n_all - two) / (1 - t), 6))
        weeks = math.ceil(new_2x / new_per_week) if new_per_week else None
        steps.append({"Step": "Recruit new members onto 2×",
                      "When": f"~{weeks} weeks at your recent sign-up rate" if weeks
                      else "Ongoing",
                      "Members": new_2x, "Weekly gain": new_2x * rrp["mkt"][1],
                      "Detail": f"{new_2x} new 2× members at RRP lifts the 2× share from "
                                f"{100 * two / n_all:.0f}% to {target_2x:.0f}%"})
    left = p[p["Moves"] & (p["Price"] < p["RRP"] - 0.005)]
    return {"steps": steps, "projection": proj, "new_2x": new_2x, "weeks_to_2x": weeks,
            "target_2x": target_2x, "not_moved": int(p.loc[~p["Moves"], "Members"].sum()),
            "still_below": int(left["Members"].sum())}


# ── Plain-English commentary (rules, no AI) ───────────────────────────────────
def _money(v, cur=""):
    return f"${v:,.0f}" + (f" {cur}" if cur and cur != "AUD" else "")


def commentary(row, bench, path=None):
    """A few sentences for one centre, from its numbers and its age group's medians."""
    out = []
    cur = row.get("Currency", "")
    b = bench.get(row.get("Age group"), {})
    p2, p1 = row.get("% 2×"), row.get("% 1×")
    if b.get("% 2×") is not None and p2 is not None:
        diff = p2 - b["% 2×"]
        word = "above" if diff >= 0 else "below"
        pts = f"{abs(diff):.0f} point" + ("" if round(abs(diff)) == 1 else "s")
        out.append(f"{p2:.0f}% of members are on 2× a week, {pts} {word} "
                   f"the median for centres of the same age ({b['% 2×']:.0f}%).")
    if p1 is not None and b.get("% 1×") is not None and p1 > b["% 1×"] + 10:
        out.append(f"The 1× share ({p1:.0f}%) is high. These members are unlikely to move "
                   "to 2×, so growing the 2× share depends on new enrolments.")
    if row.get("% of RRP") is not None:
        pct = row["% of RRP"]
        if pct >= 99.5:
            out.append(f"Pricing is at or above RRP ({pct:.0f}% of RRP).")
        else:
            out.append(f"Members pay {pct:.0f}% of RRP on average; {row['Below RRP']} of them "
                       f"are below RRP, a gap of {_money(row['RRP gap'], cur)} a week "
                       f"({_money(row['RRP gap'] * 52, cur)} a year).")
        if row.get("Below foundational"):
            out.append(f"{row['Below foundational']} member(s) pay less than even the "
                       "foundational (pre-launch) price.")
    worst = None
    for k in ["1×", "2×", "3×"]:
        avg = row.get(f"Avg {k}")
        r = row.get(f"RRP {k}")
        if avg and r and avg < r - 0.5 and row.get(f"n {k}"):
            gap = (r - avg) * row[f"n {k}"]
            if worst is None or gap > worst[0]:
                worst = (gap, k, avg, r)
    if worst:
        out.append(f"The biggest single gap is the {worst[1]} plan: average "
                   f"{_money(worst[2], cur)} against an RRP of {_money(worst[3], cur)}.")
    if row.get("n 4+×"):
        out.append(f"{row['n 4+×']} member(s) are on 4 or more sessions a week. These plans "
                   "should be phased out; they are left unchanged in the pathway.")
    if row.get("Freeze"):
        out.append(f"{row['Freeze']} member(s) pay under ${FREEZE_BELOW} a week (likely a "
                   "freeze fee or a discretionary amount). Check these.")
    if path and path.get("steps"):
        gain = sum(s["Weekly gain"] or 0 for s in path["steps"]
                   if s["Step"] != "Recruit new members onto 2×")
        if gain:
            last = path["projection"][-1]
            out.append(f"Following the pathway lifts existing members by about "
                       f"{_money(gain, cur)} a week, reaching {last['% of RRP']:.0f}% of RRP "
                       f"by {last['Date'].strftime('%b %Y')}.")
        if path.get("new_2x"):
            out.append(f"About {path['new_2x']} new 2× members would bring the 2× share to "
                       f"{path['target_2x']:.0f}%"
                       + (f", roughly {path['weeks_to_2x']} weeks at the recent sign-up rate."
                          if path.get("weeks_to_2x") else "."))
    return out


def with_rrp_columns(row, rrp):
    """Adds 'RRP 1×' ... for the commentary's plan comparison."""
    row = dict(row)
    if rrp:
        for i, k in enumerate(["1×", "2×", "3×"]):
            row[f"RRP {k}"] = rrp["mkt"][i]
    return row


def new_members_per_week(df_wm, loc, weeks=NEW_MEMBER_WEEKS):
    """Average new members a week over the last 8 weeks (Weekly Membership)."""
    if df_wm is None or df_wm.empty or "# New members" not in df_wm.columns:
        return None
    g = df_wm[df_wm[LOC] == loc].sort_values("Date").tail(weeks)
    if g.empty:
        return None
    v = pd.to_numeric(g["# New members"], errors="coerce").mean()
    return float(v) if v and not pd.isna(v) else None


def active_members(df_wm, loc):
    if df_wm is None or df_wm.empty or "# Active members" not in df_wm.columns:
        return None
    g = df_wm[df_wm[LOC] == loc]
    if g.empty:
        return None
    latest = g[g["Date"] == g["Date"].max()]
    return int(pd.to_numeric(latest["# Active members"], errors="coerce").sum())


def latest_fx(df_pnl):
    """{currency: rate to AUD} from the latest P&L month that has it."""
    if df_pnl is None or df_pnl.empty or "FX rate to AUD" not in df_pnl.columns:
        return {}
    d = df_pnl[["Period", "Currency", "FX rate to AUD"]].copy()
    d["FX rate to AUD"] = pd.to_numeric(d["FX rate to AUD"], errors="coerce")
    d = d.dropna().sort_values("Period")
    return {c: float(g["FX rate to AUD"].iloc[-1]) for c, g in d.groupby("Currency")}
