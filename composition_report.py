"""Membership Composition report (admins only for now): a ranked group overview of every
centre's plan mix and pricing against median and RRP, and a deep dive into one centre
with trend lines, gaps, a pathway to RRP and rule-based commentary."""
import html

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import membership_composition as mc

RED, GREY, GREEN, TEAL, TEXT = "#c8322f", "#9aa7ab", "#2e8540", "#1f9a9a", "#1c2a30"
MIX_COLOURS = {"1×": "#e08a3c", "2×": GREEN, "3×": TEAL, "4+×": RED, "1:1": GREY}
VIEWS = ["Group overview", "Centre deep dive"]

HOW = f"""
- **Source:** the portal's Membership Composition export, uploaded weekly by an admin
  (⚙ Composition Upload). Trend lines start from the first upload and grow week by week.
- **Mix:** share of members on 1×, 2×, 3× or 4+ sessions a week (group plans), and on 1:1.
- **Prices** are weekly, in each centre's own currency. US prices are billed monthly and
  divided by 4. Prices under ${mc.FREEZE_BELOW} a week (a freeze fee or discretionary amount)
  count in the mix but are left out of prices and RRP figures; they are listed to check.
- **RRP:** HQ's market benchmark from 1 July 2026 for the centre's state or country (the
  'RRP' tab in the Sheet). 1:1 = RRP per session × sessions. 4+ a week = the 3× rate per
  session × sessions; these plans should be phased out.
- **% of RRP** = what members pay ÷ what the same members would pay at RRP.
  **RRP gap** = the shortfall of members paying below RRP (members above RRP don't offset it).
  **Below foundational** = paying less than even the pre-launch price.
- **Median** = the middle centre of the same age group, from centres older than
  {mc.YOUNG_WEEKS} weeks. ⚠ = fewer than {mc.SMALL_SAMPLE} centres in the group.
- **Ranking** is by RRP gap per week in AUD (latest P&L exchange rates); centres
  {mc.YOUNG_WEEKS} weeks old or younger are shown separately, for information.
- **Flagged** = beyond {mc.OUTLIER_SD} standard deviations from the median on % of RRP,
  share on 2×, share on 1× or share below RRP, or one of the {mc.TOP_GAPS} biggest RRP gaps.
- **Pathway to RRP:** new members at RRP from now; foundational pricing ends
  {mc.FOUNDATIONAL_MONTHS} months after opening (members on foundational rates move to RRP);
  other below-RRP group plans rise each 1 July by at most {mc.REVIEW_CAP:.0%}, never above
  RRP. 1× members are assumed to stay on 1×, so a higher 2× share needs new members. 1:1 and
  4+ plans are left as they are. Weeks to reach the 2× target use the centre's average new
  members a week over the last {mc.NEW_MEMBER_WEEKS} weeks (Weekly Membership).
- **Commentary** is written by fixed rules from the numbers above (no AI), so it is the same
  every time and costs nothing to produce.
"""


def _section(title, caption=None):
    st.markdown(f'<div class="section-header">{html.escape(title)}</div>',
                unsafe_allow_html=True)
    if caption:
        st.caption(caption)


def _card(label, value, sub=""):
    return (f'<div class="metric-card"><div class="metric-label">{html.escape(label)}</div>'
            f'<div class="metric-value">{html.escape(value)}</div>'
            f'<div class="metric-delta" style="color:#5d6d73">{html.escape(sub)}</div></div>')


def _cards(items):
    cols = st.columns(len(items))
    for c, it in zip(cols, items):
        c.markdown(_card(*it), unsafe_allow_html=True)


def _rrp_colour(v):
    if v is None or pd.isna(v):
        return ""
    c = GREEN if v >= 99.5 else GREY if v >= 90 else RED
    return f"color:{c};font-weight:600"


def _pct(v):
    return "–" if v is None or pd.isna(v) else f"{v:.0f}%"


def _money(v, cur=""):
    if v is None or pd.isna(v):
        return "–"
    return f"${v:,.0f}" + (f" {cur}" if cur and cur != "AUD" else "")


# ── Group overview ────────────────────────────────────────────────────────────
TABLE_COLS = ["Rank", "Centre", "Country", "Age group", "Members", "% 1×", "% 2×",
              "2× vs median", "% 3×", "% 4+×", "% 1:1", "ARPM", "% of RRP", "Below RRP",
              "Below foundational", "RRP gap /wk", "RRP gap /wk AUD", "Flag"]


def _table(s, bench, fl):
    t = s.copy()
    t["2× vs median"] = [r["% 2×"] - bench.get(r["Age group"], {}).get("% 2×", float("nan"))
                         for _, r in t.iterrows()]
    t["RRP gap /wk"] = t["RRP gap"]
    t["RRP gap /wk AUD"] = t["RRP gap AUD"]
    t["Flag"] = ["⚑ " + "; ".join(fl[loc]) if loc in fl else "" for loc in t["Location"]]
    t = t.sort_values("RRP gap AUD", ascending=False, na_position="last")
    t["Rank"] = range(1, len(t) + 1)
    return t[TABLE_COLS]


def _show_table(t, key):
    fmt = {c: st.column_config.NumberColumn(c, format="%.0f%%")
           for c in ["% 1×", "% 2×", "% 3×", "% 4+×", "% 1:1", "% of RRP"]}
    fmt["2× vs median"] = st.column_config.NumberColumn("2× vs median", format="%+.0f pts")
    fmt["ARPM"] = st.column_config.NumberColumn("ARPM", format="%.2f")
    fmt["RRP gap /wk"] = st.column_config.NumberColumn("RRP gap /wk", format="%,.0f")
    fmt["RRP gap /wk AUD"] = st.column_config.NumberColumn("RRP gap /wk AUD", format="$%,.0f")
    styled = t.style.map(_rrp_colour, subset=["% of RRP"])
    st.dataframe(styled, hide_index=True, width="stretch", column_config=fmt, key=key)


def _rrp_chart(s, helpers):
    d = s.dropna(subset=["% of RRP"]).sort_values("% of RRP")
    if d.empty:
        return
    colours = [GREEN if v >= 99.5 else GREY if v >= 90 else RED for v in d["% of RRP"]]
    fig = go.Figure(go.Bar(x=d["% of RRP"], y=d["Centre"], orientation="h",
                           marker_color=colours, text=[f"{v:.0f}%" for v in d["% of RRP"]],
                           textposition="outside", hovertemplate="%{y}: %{x:.0f}% of RRP"
                           "<extra></extra>"))
    fig.add_vline(x=100, line_dash="dash", line_color=TEXT, annotation_text="RRP")
    med = d["% of RRP"].median()
    fig.add_vline(x=med, line_dash="dot", line_color=TEAL,
                  annotation_text=f"Median {med:.0f}%", annotation_position="bottom right")
    fig.update_layout(**helpers["std_layout"]("", "", max(300, 26 * len(d) + 80)))
    fig.update_xaxes(title="% of RRP", ticksuffix="%")
    helpers["show_chart"](fig)


def _mix_chart(s, helpers):
    d = s.sort_values("% 2×")
    fig = go.Figure()
    for k in mc.MIX:
        fig.add_bar(x=d[f"% {k}"], y=d["Centre"], orientation="h", name=k,
                    marker_color=MIX_COLOURS[k],
                    hovertemplate=f"%{{y}}: %{{x:.0f}}% on {k}<extra></extra>")
    fig.update_layout(**helpers["std_layout"]("", "", max(300, 26 * len(d) + 80)),
                      barmode="stack")
    fig.update_xaxes(title="Share of members", ticksuffix="%", range=[0, 100])
    helpers["show_chart"](fig)


def _bench_table(bench):
    rows = [{"Age group": g, "Centres": int(b["n"]) if b.get("n") else 0,
             "Median % 1×": b.get("% 1×"), "Median % 2×": b.get("% 2×"),
             "Median % of RRP": b.get("% of RRP"),
             "Sample": "⚠ small sample" if mc.SMALL_SAMPLE > b.get("n", 0) else ""}
            for g, b in bench.items()]
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch",
                 column_config={c: st.column_config.NumberColumn(c, format="%.0f%%")
                                for c in ["Median % 1×", "Median % 2×", "Median % of RRP"]})


def _group(s, bench, fl, helpers):
    countries = sorted(s["Country"].unique())
    pick = st.selectbox("Country", ["All countries"] + countries, key="comp_country")
    if pick != "All countries":
        s = s[s["Country"] == pick]
    est, young = s[~s["Young"]], s[s["Young"]]
    members = int(s["Members"].sum())
    two = s["n 2×"].sum() / members * 100 if members else 0
    gap = s["RRP gap AUD"].sum(min_count=1)
    missing_fx = s["RRP gap AUD"].isna() & s["RRP gap"].notna()
    _cards([("Centres", f"{len(s)}", f"{len(young)} aged {mc.YOUNG_WEEKS} weeks or less"),
            ("Members", f"{members:,}", "in this week's export"),
            ("On 2× a week", f"{two:.0f}%", "of all members"),
            ("Median % of RRP", _pct(est["% of RRP"].median() if len(est)
                                     else s["% of RRP"].median()), "centres over 18 weeks"),
            ("RRP gap", _money(gap) + " /wk" if not pd.isna(gap) else "–",
             f"{_money(gap * 52)} a year (AUD)" if not pd.isna(gap) else "")])
    if missing_fx.any():
        st.caption("No exchange rate yet for " + ", ".join(sorted(s.loc[missing_fx, "Currency"]
                   .unique())) + " (from the P&L tab); those gaps are left out of AUD totals.")

    _section("Biggest gaps and opportunities",
             f"Centres beyond {mc.OUTLIER_SD} standard deviations from the median, and the "
             f"{mc.TOP_GAPS} biggest RRP gaps.")
    shown = [loc for loc in s["Location"] if loc in fl]
    if not shown:
        st.success("No centre stands out this week.")
    for loc in sorted(shown, key=lambda l: -(s.loc[s["Location"] == l, "RRP gap AUD"]
                                             .fillna(0).iloc[0])):
        r = s[s["Location"] == loc].iloc[0]
        st.markdown(f"**{html.escape(r['Centre'])}** — {html.escape('; '.join(fl[loc]))}. "
                    f"{_pct(r['% of RRP'])} of RRP, {_pct(r['% 2×'])} on 2×, gap "
                    f"{_money(r['RRP gap'], r['Currency'])} a week.")

    _section(f"Ranking — centres older than {mc.YOUNG_WEEKS} weeks",
             "Ranked by RRP gap per week (AUD). % of RRP: green at RRP, grey 90–99%, red "
             "under 90%.")
    if est.empty:
        st.info(f"No centre older than {mc.YOUNG_WEEKS} weeks in this selection.")
    else:
        _show_table(_table(est, bench, fl), "comp_rank")
    if not young.empty:
        _section(f"Centres {mc.YOUNG_WEEKS} weeks old or younger (for information)")
        _show_table(_table(young, bench, fl), "comp_young")

    _section("% of RRP by centre", "Dashed line = RRP; dotted line = median.")
    _rrp_chart(s, helpers)
    _section("Membership mix by centre", "Sorted by share on 2× (the target plan).")
    _mix_chart(s, helpers)
    _section("Medians by age group", "Used for the comparisons above (every centre counts).")
    _bench_table(bench)


# ── Centre deep dive ──────────────────────────────────────────────────────────
def _trend_charts(hist, rrp, cur, helpers):
    if hist.empty:
        return
    n = len(hist)
    mode = "lines+markers" if n > 1 else "markers"
    fig = go.Figure()
    for k in mc.MIX:
        if hist[f"n {k}"].sum():
            fig.add_scatter(x=hist["Week ending"], y=hist[f"% {k}"], name=k, mode=mode,
                            line=dict(color=MIX_COLOURS[k], width=2.5), marker=dict(size=8))
    fig.update_layout(**helpers["std_layout"]("Share of members by plan", "% of members", 360))
    fig.update_yaxes(ticksuffix="%", rangemode="tozero")
    helpers["show_chart"](fig)

    fig = go.Figure()
    for i, k in enumerate(mc.MIX):
        col = f"Avg {k}"
        if hist[col].notna().any():
            fig.add_scatter(x=hist["Week ending"], y=hist[col], name=f"{k} average",
                            mode=mode, line=dict(color=MIX_COLOURS[k], width=2.5),
                            marker=dict(size=8))
            if rrp and i < 3:
                fig.add_hline(y=rrp["mkt"][i], line_dash="dot", line_color=MIX_COLOURS[k],
                              annotation_text=f"{k} RRP {rrp['mkt'][i]:,.0f}",
                              annotation_position="top left")
    fig.update_layout(**helpers["std_layout"](f"Average weekly price by plan ({cur})",
                                              cur, 380))
    helpers["show_chart"](fig)
    if n == 1:
        st.caption("Only one week uploaded so far: the trend lines grow as weekly exports "
                   "are added.")


def _gaps_table(plans, rrp, cur):
    rows = []
    for k in mc.MIX:
        sub = plans[(plans["Mix"] == k) & ~plans["Freeze"]]
        if sub.empty:
            continue
        n = int(sub["Members"].sum())
        avg = (sub["Members"] * sub["Weekly price"]).sum() / n
        at = sum(m * mc.rrp_price(rrp, t, s) for m, t, s in
                 zip(sub["Members"], sub["Type"], sub["Sessions/week"])) if rrp else None
        gap = sum(m * max(0, mc.rrp_price(rrp, t, s) - p) for m, t, s, p in
                  zip(sub["Members"], sub["Type"], sub["Sessions/week"], sub["Weekly price"])) \
            if rrp else None
        below = int(sum(m for m, t, s, p in zip(sub["Members"], sub["Type"], sub["Sessions/week"],
                                                 sub["Weekly price"])
                        if rrp and p < mc.rrp_price(rrp, t, s) - 0.005))
        rows.append({"Plan": k, "Members": n, "Average price": avg,
                     "RRP (average)": at / n if at else None,
                     "Below RRP": below, "Gap /wk": gap, "Gap /yr": gap * 52 if gap else gap,
                     "Note": "Phase out" if k == "4+×" else ""})
    t = pd.DataFrame(rows)
    if t.empty:
        return
    t = t.sort_values("Gap /wk", ascending=False, na_position="last")
    st.dataframe(t, hide_index=True, width="stretch", column_config={
        c: st.column_config.NumberColumn(f"{c} ({cur})", format="%,.2f" if "price" in c
                                         or "RRP" in c else "%,.0f")
        for c in ["Average price", "RRP (average)", "Gap /wk", "Gap /yr"]})


def _pathway(plans, rrp, row, bench, df_wm, cur, helpers):
    med = bench.get(row["Age group"], {}).get("% 2×")
    default = float(round(max(med or 0, row["% 2×"]))) if med is not None else \
        float(round(row["% 2×"]))
    target = st.number_input("Target share on 2× (%)", 0.0, 99.0, min(default, 99.0), 5.0,
                             key=f"comp_target_{row['Location']}",
                             help="Starts at the median for centres of the same age, or the "
                                  "centre's own share if that's higher.")
    per_week = mc.new_members_per_week(df_wm, row["Location"])
    path = mc.pathway(plans, rrp, row.get("Age (Months)"), per_week, target)
    if not path["steps"]:
        st.info("No RRP is set for this centre's country, so there's no pathway yet.")
        return path
    steps = pd.DataFrame(path["steps"])
    st.dataframe(steps[["Step", "When", "Members", "Weekly gain", "Detail"]], hide_index=True,
                 width="stretch", column_config={
                     "Weekly gain": st.column_config.NumberColumn(f"Weekly gain ({cur})",
                                                                  format="%,.0f"),
                     "Members": st.column_config.NumberColumn(format="%d")})
    notes = ["1× members stay on 1×; more 2× comes from new members"]
    if path.get("not_moved"):
        notes.append(f"{path['not_moved']} member(s) on 1:1 or 4+ a week are left as they are")
    if path.get("still_below"):
        notes.append(f"{path['still_below']} member(s) still below RRP after "
                     f"{len([s for s in path['steps'] if s['Step'].startswith('July')])} reviews")
    if per_week is None:
        notes.append("no recent sign-up rate in Weekly Membership, so no time estimate")
    st.caption(" · ".join(notes))
    proj = pd.DataFrame(path["projection"])
    if len(proj) > 1:
        fig = go.Figure(go.Scatter(x=proj["Date"], y=proj["% of RRP"], mode="lines+markers",
                                   line=dict(color=TEAL, width=3), marker=dict(size=9),
                                   text=proj["When"], hovertemplate="%{text}: %{y:.0f}% of RRP"
                                   "<extra></extra>"))
        fig.add_hline(y=100, line_dash="dash", line_color=TEXT, annotation_text="RRP")
        fig.update_layout(**helpers["std_layout"]("Existing members: % of RRP on the pathway",
                                                  "% of RRP", 320))
        fig.update_yaxes(ticksuffix="%")
        helpers["show_chart"](fig)
    return path


def _centre(comp, s, bench, fl, table, vl, df_wm, helpers):
    names = dict(sorted(zip(s["Centre"], s["Location"]), key=lambda kv: kv[0].lower()))
    pick = st.selectbox("Centre", list(names), key="comp_centre")
    loc = names[pick]
    row = s[s["Location"] == loc].iloc[0].to_dict()
    cur = row["Currency"]
    info = vl[vl[mc.LOC] == loc].iloc[0].to_dict() if mc.LOC in vl.columns and \
        (vl[mc.LOC] == loc).any() else {}
    rrp = mc.rrp_for(table, row["Country"], info.get("Region"))
    plans = comp[(comp["Location"] == loc) & (comp["Week ending"] == row["Week ending"])]
    b = bench.get(row["Age group"], {})
    active = mc.active_members(df_wm, loc)
    st.caption(f"{row['Age group']} · RRP region {row['RRP region']}"
               + (f" · {rrp['mkt'][0]:,.0f} / {rrp['mkt'][1]:,.0f} / {rrp['mkt'][2]:,.0f} "
                  f"a week for 1× / 2× / 3× ({cur})" if rrp else "")
               + (" · ⚠ small comparison group" if mc.SMALL_SAMPLE > b.get("n", 0) else ""))
    _cards([("Members", f"{row['Members']}",
             f"{row['Members'] / active:.0%} of {active} active members" if active else
             "in the export"),
            ("On 2× a week", _pct(row["% 2×"]),
             f"median {_pct(b.get('% 2×'))}" if b else ""),
            ("ARPM", _money(row["ARPM"], cur), "average weekly price"),
            ("% of RRP", _pct(row["% of RRP"]), f"median {_pct(b.get('% of RRP'))}"),
            ("RRP gap", _money(row["RRP gap"], cur) + " /wk",
             f"{_money((row['RRP gap'] or 0) * 52, cur)} a year")])
    if loc in fl:
        st.warning("⚑ " + "; ".join(fl[loc]))

    path = None
    with st.container():
        _section("What stands out")
        box = st.empty()

    _section("Trends", "Each week's upload adds a point.")
    _trend_charts(mc.history(comp, loc, rrp), rrp, cur, helpers)

    _section("Gaps by plan", f"Weekly, {cur}. Members paying above RRP don't offset the gap.")
    _gaps_table(plans, rrp, cur)

    _section("Recommended pathway to RRP")
    path = _pathway(plans, rrp, row, bench, df_wm, cur, helpers)

    text = mc.commentary(mc.with_rrp_columns(row, rrp), bench, path)
    box.markdown("\n".join(f"- {html.escape(t)}" for t in text) or "Nothing stands out.")

    odd = plans[plans["Freeze"]]
    with st.expander(f"All plans this week ({len(plans)})"
                     + (f" · {len(odd)} under ${mc.FREEZE_BELOW} to check" if len(odd) else "")):
        view = plans[["Type", "Sessions/week", "Members", "Weekly price"]].copy()
        view["RRP"] = [mc.rrp_price(rrp, t, n) for t, n in zip(view["Type"], view["Sessions/week"])]
        view["Check"] = ["Under $10: freeze fee or discretionary?" if f else
                         ("4+ a week: phase out" if n >= 4 and t == "Group" else "")
                         for f, n, t in zip(plans["Freeze"], view["Sessions/week"], view["Type"])]
        st.dataframe(view.sort_values(["Type", "Sessions/week", "Weekly price"]),
                     hide_index=True, width="stretch",
                     column_config={"Weekly price": st.column_config.NumberColumn(format="%.2f"),
                                    "RRP": st.column_config.NumberColumn(format="%.2f")})


# ── Page ──────────────────────────────────────────────────────────────────────
def report_composition(df_comp, df_wm, vl, df_rrp, df_pnl, helpers):
    st.markdown('<div class="report-title">Membership Composition</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Each centre\'s mix of 1×, 2×, 3× and 4+ a week '
                'plans and its prices, against the median for its age and HQ\'s RRP.</div>',
                unsafe_allow_html=True)
    if st.session_state.get("access_level") != "admin":
        st.error("This report is for admins only.")
        return
    comp = mc.prepare(df_comp)
    if comp.empty:
        st.info("No Membership Composition data yet. An admin can add a week on the "
                "⚙ Composition Upload page.")
        return
    vl = vl if vl is not None else pd.DataFrame()
    table = mc.rrp_table(df_rrp)
    weeks = sorted({d.strftime("%Y-%m-%d") for d in comp["Week ending"]}, reverse=True)
    with st.container(border=True):
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b = st.columns([2, 1])
        view = a.radio("View", VIEWS, horizontal=True, key="comp_view",
                       label_visibility="collapsed")
        week = b.selectbox("Week ending", weeks, key="comp_week_pick",
                           format_func=lambda d: pd.Timestamp(d).strftime("%d/%m/%Y"))
    week = pd.Timestamp(week)
    s = mc.summary(comp, vl, table, helpers.get("age_group"), mc.latest_fx(df_pnl),
                   week=week)
    if s.empty:
        st.info("No centres in that week.")
        return
    bench = mc.benchmarks(s)
    fl = mc.flags(s)
    if view == VIEWS[0]:
        _group(s, bench, fl, helpers)
    else:
        _centre(comp, s, bench, fl, table, vl, df_wm, helpers)
    with st.expander("How this is worked out"):
        st.markdown(HOW)
