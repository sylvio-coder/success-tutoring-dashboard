"""Membership Composition report (admins only): a group overview of every centre's
ARPM, revenue per session and % of RRP, and a centre view with its current composition
and the ideal composition to achieve RRP, in the portal's table layout."""
import html
import math
import re

import pandas as pd
import streamlit as st

import membership_composition as mc

VIEWS = ["Group overview", "Centre deep dive"]
RED, GREEN = "#c8322f", "#2e8540"

CSS = """<style>
.mc-card{border:1px solid #e3e8ea;border-radius:10px;background:#fff;padding:14px 16px 6px;
 margin-bottom:14px}
.mc-head{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px}
.mc-head b{font-size:16px;color:#1c2a30}.mc-head span{color:#5d6d73;font-size:13px}
.mc{border-collapse:collapse;width:100%;font-size:13px}
.mc th{text-align:right;color:#5d6d73;font-weight:600;font-size:11px;letter-spacing:.04em;
 text-transform:uppercase;background:#f5f7f8;padding:9px 10px;border-bottom:1px solid #e3e8ea}
.mc th:first-child,.mc td:first-child{text-align:left}
.mc td{text-align:right;padding:9px 10px;border-bottom:1px solid #eef2f3;color:#1c2a30}
.mc td.t{color:#5d6d73}
.mc tr.total td{font-weight:700;border-top:1.5px solid #1c2a30}
.mc tr.arpm td{font-weight:700}
.mc .neg{color:#c8322f}.mc .pos{color:#2e8540}
.mc-note{color:#5d6d73;font-size:12px;margin:6px 0 4px}
.mc-sub{font-size:13px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;
 color:#1f9a9a;margin:2px 0 6px}
.mc-nr{color:#8a5a00;font-weight:600}
.mc-gains{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;
 margin:6px 0 12px}
.mc-gain{background:#fff;border:1px solid #dde5e6;border-left:5px solid var(--c);
 border-radius:10px;padding:12px 14px}
.mc-gain .k{font-size:12px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;
 color:var(--c)}
.mc-gain .v{font-size:26px;font-weight:700;color:#1c2a30;margin:2px 0}
.mc-gain .v small{font-size:13px;font-weight:600;color:#5d6d73}
.mc-gain .w{font-size:13.5px;color:#1c2a30;line-height:1.35}
.mc-gain.zero .v{color:#8a9a9f}
.mc-total{font-size:15px;color:#1c2a30;margin:0 0 4px}
</style>"""

HOW = f"""
- **Source:** the portal's Membership Composition export, uploaded weekly by an admin
  (⚙ Composition Upload). Prices are weekly, in the centre's own currency (the CRM exports
  every country, including the US, at weekly prices).
- **RRP:** HQ's national RRP from 1 July 2026 for Australia, New Zealand and Canada; the US by
  region (e.g. Texas = Austin & Houston). Kept in the Sheet's 'RRP' tab.
  1:1 = RRP per session × sessions. 4 or more a week = the 3× rate per session × sessions.
- **RRP variance** = membership price − RRP for that plan.
- **Variance from median** = membership price − the median price members pay for the same
  plan (type and sessions a week) across every centre in the same country. "–" when no other
  centre in the country has that plan.
- **Total sessions/week** = no. members × sessions a week each, added up (group overview: all
  the centre's members). In the ideal table, members are fractions, so sessions are too.
- **Sessions per member** = total sessions a week ÷ members: how often the average member
  comes. ARPM = revenue per session × sessions per member.
- **Revenue per session** = membership price ÷ sessions a week (totals: total value ÷ all
  weekly sessions).
- **Prices under ${mc.FREEZE_BELOW}** (a freeze fee or discretionary amount) are left out of the
  rows but kept in the totals, ARPM and composition.
- **Composition table:** every member on each plan, including prices under ${mc.FREEZE_BELOW}.
  Average membership price = the plan's total value ÷ its members; revenue per session =
  total value ÷ all its weekly sessions.
- **Ideal composition:** the centre's members spread {mc.TARGET_MIX['1×']}% 1×,
  {mc.TARGET_MIX['2×']}% 2×, {mc.TARGET_MIX['3+']}% 3 or more and {mc.TARGET_MIX['1:1']}% 1:1,
  each at RRP, or at the centre's own average price per session where that is higher.
  Composition variance = current % − ideal %.
- **Trends:** one point per uploaded week (gaps = no upload); school holidays shaded for the
  centre's state. Group trends add up the centres in view (one country at a time) and work
  out prices from the totals. The dashed typical-centre line is the median of the other
  centres in the country, shown in weeks with at least 5 of them, and not for
  centres trading 12 months or less.
- **Margin per member a week** = (total value − hubs needed × tutor cost an hour) ÷ members,
  with hubs needed = group sessions ÷ the target fill for the centre's size, plus 1:1
  sessions. Tutor cost: AU $37.90 an hour ($33.85 + 12% super); other countries to come.
- **4×+ group plans** (4 or more sessions a week) are not recommended: they are flagged, kept
  at today's price (no RRP uplift), left out of the ideal mix and out of % of RRP.
- **% of RRP** (group overview) = what members pay ÷ what the same members would pay at RRP.
"""


def _m(v, cur=""):
    return "–" if v is None or pd.isna(v) else f"${v:,.2f}"


def _n(v, dp=0):
    """A count of sessions: whole numbers, or dp decimals (the ideal table's members are
    fractions)."""
    if v is None or pd.isna(v):
        return "–"
    return f"{v:,.{dp}f}"


def _var(v, p):
    if v is None or pd.isna(v):
        return "–"
    cls = "neg" if v < -0.005 else "pos" if v > 0.005 else ""
    return f'<span class="{cls}">{"+" if v > 0.005 else "−" if v < -0.005 else ""}' \
           f'${abs(v):,.2f} ({"+" if p > 0.5 else "−" if p < -0.5 else ""}{abs(p):.0f}%)</span>'


def _pts(v):
    sign = "+" if v >= 0.5 else "−" if v <= -0.5 else ""
    return f"{sign}{abs(v):.0f} pts"


def _table(headers, rows, classes=None):
    head = "".join(f"<th>{html.escape(h)}</th>" for h in headers)
    body = ""
    for i, r in enumerate(rows):
        cls = (classes or {}).get(i, "")
        body += f'<tr class="{cls}">' + "".join(
            f'<td class="{"t" if j == 0 and not cls else ""}">{c}</td>'
            for j, c in enumerate(r)) + "</tr>"
    return f'<table class="mc"><tr>{head}</tr>{body}</table>'


def _card(title, sub, inner):
    st.markdown(f'<div class="mc-card"><div class="mc-head"><b>{title}</b>'
                f'<span>{sub}</span></div>{inner}</div>', unsafe_allow_html=True)


def _per_member(t):
    return t["sessions"] / t["members"] if t["members"] else None


def current_html(t):
    rows = [[mc.plan_label(r["Type"], r["Sessions/week"]) +
             (f' ({r["Sessions/week"]}×)' if r["Type"] == "1:1" else ""),
             f'{r["Members"]:,}', _n(r["Total sessions"]),
             _m(r["Membership price"]), _m(r["Revenue per session"]), _m(r["Total value"]),
             '<span class="mc-nr">Not recommended</span>' if r.get("Not recommended")
             else _var(r["RRP variance"], r["RRP variance %"]),
             _var(r["Median variance"], r["Median variance %"])]
            for _, r in t["rows"].iterrows()]
    k = len(rows)
    rows.append(["Total", f'{t["members"]:,}', _n(t["sessions"]), "", _m(t["per_session"]),
                 _m(t["total"]), "", ""])
    rows.append(["Per member", "", f'{_n(_per_member(t), 2)} sessions',
                 f'{_m(t["arpm"])} ARPM', "", "", "", ""])
    return _table(["Plan", "No. members", "Total sessions/week", "Membership price",
                   "Revenue per session", "Total value", "RRP variance",
                   "Variance from median"], rows, {k: "total", k + 1: "arpm"})


def _is_nr_plan(label):
    m = re.match(r"Group (\d+)×", label)
    return bool(m) and int(m.group(1)) >= mc.NOT_RECOMMENDED_FROM


def composition_html(comp):
    rows = [[html.escape(r["Plan"]) + (' <span class="mc-nr">⚠ not recommended</span>'
                                       if _is_nr_plan(r["Plan"]) else ""),
             f'{int(r["Members"]):,}', f'{r["% of members"]:.0f}%',
             _n(r["Total sessions"]), _n(r["Sessions per member"], 2), _m(r["Average price"]),
             _m(r["Revenue per session"]), _m(r["Total value"])] for _, r in comp.iterrows()]
    n, sess, value = comp["Members"].sum(), comp["Total sessions"].sum(), comp["Total value"].sum()
    rows.append(["Total", f"{int(n):,}", "100%" if n else "–", _n(sess),
                 _n(sess / n if n else None, 2), _m(value / n if n else None),
                 _m(value / sess if sess else None), _m(value)])
    return _table(["Plan", "No. members", "% of members", "Total sessions/week",
                   "Sessions per member", "Average membership price", "Revenue per session",
                   "Total value"], rows, {len(rows) - 1: "total"})


def current_vs_ideal_html(plans, ideal):
    """One table, now → ideal for each plan: members (with % of members), weekly price and
    value a week. Rows empty now and in the ideal are hidden; totals use unrounded figures."""
    n = int(plans["Members"].sum())
    ideal_rows = ideal["rows"]
    pct = lambda m: f" ({m / n * 100:.0f}%)" if n else ""
    whole = lambda x: f"{math.floor(x + 0.5 + 1e-9):,}"           # 6.5 members -> 7
    arrow = lambda a, b: f'{a} → <b>{b}</b>'

    def now(mask):
        g = plans[mask]
        m = float(g["Members"].sum())
        v = float((g["Members"] * g["Weekly price"]).sum())
        return m, v, (v / m if m else None)
    rows = []
    for kind, plan, lo, hi, _ in mc.IDEAL_GROUPS:
        cm, cv, cp = now((plans["Type"] == kind) & plans["Sessions/week"].between(lo, hi))
        r = ideal_rows[ideal_rows["Plan"] == plan].iloc[0]
        im, iv, ip = float(r["Members"]), float(r["Total value"] or 0), r["Membership price"]
        if math.floor(cm + 0.5) == 0 and math.floor(im + 0.5) == 0:
            continue
        label = ("Group " + plan if kind == "Group" else plan).replace("3+", "3×")
        rows.append([label, arrow(f"{whole(cm)}{pct(cm)}", f"{whole(im)}{pct(im)}"),
                     arrow(_m(cp) if cm else "–",
                           _m(ip) + (" ¹" if r["Own price used"] else "")),
                     arrow(f"${cv:,.0f}", f"${iv:,.0f}")])
    nr = ideal_rows[ideal_rows["Not recommended"]]
    if len(nr):
        r = nr.iloc[0]
        rows.append([f'Group 4×+ <span class="mc-nr">⚠ not recommended</span>',
                     f'{whole(r["Members"])}{pct(r["Members"])} (unchanged)',
                     f'{_m(r["Membership price"])} (kept)',
                     arrow(f'${r["Total value"]:,.0f}', f'${r["Total value"]:,.0f}')])
    v0 = float((plans["Members"] * plans["Weekly price"]).sum())
    diff = ideal["total"] - v0
    rows.append(["Total", arrow(f"{n:,}", f"{n:,}"),
                 arrow(f"ARPM {_m(v0 / n if n else None)}", f"ARPM {_m(ideal['arpm'])}"),
                 arrow(f"${v0:,.0f}", f"${ideal['total']:,.0f}")
                 + f' ({"+" if diff >= 0 else "−"}${abs(diff):,.0f})'])
    return _table(["Plan", "Members: now → ideal", "Weekly price: now → ideal",
                   "Value a week: now → ideal"], rows, {len(rows) - 1: "total"})


# ── Views ─────────────────────────────────────────────────────────────────────
MIN_FOR_MEDIAN = 5   # non-admins: country median only shown with at least this many centres


def _visible(week, vl, table, allowed):
    """Group table for the centres this viewer may see. Medians use every centre, but for a
    non-admin they are only shown where the country has enough centres to stay anonymous."""
    t = mc.group_table(week, vl, table)
    if allowed is None or t.empty:
        return t
    count = t.groupby("Country")["Centre"].transform("count")
    t = t.assign(**{"ARPM vs country median":
                    t["ARPM vs country median"].where(count >= MIN_FOR_MEDIAN)})
    return t[t["Location"].isin(allowed)].reset_index(drop=True)


SMALL = 15            # centres under this many members are faded on the chart
QUADRANTS = [("Efficient", "fewer sessions, higher price", "left", "top"),
             ("Premium", "more sessions, higher price", "right", "top"),
             ("Underpriced", "fewer sessions, lower price", "left", "bottom"),
             ("Discount trap", "more sessions, lower price", "right", "bottom")]


def country_medians(full, country, allowed):
    """(sessions per member, revenue per session) medians for a country, from every centre.
    None for a non-admin where the country has too few centres to stay anonymous."""
    c = full[(full["Country"] == country) & (full["Members"] > 0)]
    if len(c) < 2 or (allowed is not None and len(c) < MIN_FOR_MEDIAN):
        return None
    return float(c["Sessions per member"].median()), float(c["Revenue per session"].median())


def price_frequency_chart(t, med, country, title, highlight=None, ideal=None, key="pf"):
    """Sessions per member (across) vs revenue per session (up), one dot per centre in t,
    dashed country medians and the four quarters. highlight: a centre to draw strongly;
    ideal: (sessions per member, revenue per session) to point an arrow to."""
    import plotly.graph_objects as go
    fig = go.Figure()
    cur = mc.CURRENCY.get(country, "")
    for small in (False, True):
        d = t[(t["Members"] < SMALL) == small]
        if d.empty:
            continue
        strong = d["Location"] == highlight if highlight else pd.Series(True, index=d.index)
        fig.add_scatter(
            x=d["Sessions per member"], y=d["Revenue per session"], mode="markers+text",
            text=d["Centre"] if highlight is None else
            [c if s_ else "" for c, s_ in zip(d["Centre"], strong)],
            textposition="bottom center" if ideal is not None else "top center",
            textfont=dict(size=11, color="#1c2a30"),
            marker=dict(size=[max(10, min(40, m ** 0.5 * 4)) for m in d["Members"]],
                        color=["#1f9a9a" if s_ else "#c9d3d6" for s_ in strong],
                        opacity=0.35 if small else 0.85, line=dict(width=1, color="#fff")),
            customdata=d[["Members", "ARPM", "% of RRP"]].to_numpy(),
            hovertemplate=("<b>%{text}</b><br>" if highlight is None else "") +
            "Sessions per member: %{x:.2f}<br>Revenue per session: $%{y:.2f}<br>"
            "Members: %{customdata[0]:.0f}<br>ARPM: $%{customdata[1]:.2f}<extra></extra>",
            name=f"Under {SMALL} members" if small else "Centres", showlegend=False)
    if ideal is not None and highlight is not None and (t["Location"] == highlight).any():
        r = t[t["Location"] == highlight].iloc[0]
        fig.add_annotation(x=ideal[0], y=ideal[1], ax=r["Sessions per member"],
                           ay=r["Revenue per session"], xref="x", yref="y", axref="x",
                           ayref="y", showarrow=True, arrowhead=3, arrowwidth=2,
                           arrowcolor="#2e8540")
        fig.add_scatter(x=[ideal[0]], y=[ideal[1]], mode="markers+text", text=["Achievable"],
                        textposition="middle right", showlegend=False, hoverinfo="skip",
                        textfont=dict(size=11, color="#2e8540"),
                        marker=dict(size=14, color="#2e8540", symbol="star"))
    if med:
        fig.add_vline(x=med[0], line=dict(dash="dash", color="#8a9a9f", width=1))
        fig.add_hline(y=med[1], line=dict(dash="dash", color="#8a9a9f", width=1))
        for name, sub, xa, ya in QUADRANTS:
            fig.add_annotation(
                text=f"<b>{name}</b><br><span style='font-size:10px'>{sub}</span>",
                xref="x domain", yref="y domain", x=0.01 if xa == "left" else 0.99,
                y=0.99 if ya == "top" else 0.01, xanchor=xa, yanchor=ya, showarrow=False,
                font=dict(size=11, color={"Efficient": GREEN, "Discount trap": RED}.get(
                    name, "#5d6d73")), align=xa)
    fig.update_layout(title=title, height=460, margin=dict(l=10, r=10, t=50, b=10),
                      plot_bgcolor="#fff", paper_bgcolor="#fff",
                      xaxis=dict(title="Sessions per member a week", gridcolor="#eef1f3",
                                 zeroline=False),
                      yaxis=dict(title=f"Revenue per session ({cur})", gridcolor="#eef1f3",
                                 tickprefix="$", zeroline=False))
    st.plotly_chart(fig, width="stretch", key=key)


# ── Trends ────────────────────────────────────────────────────────────────────
SPANS = {"Last 12 weeks": 12, "Last 26 weeks": 26, "Last 52 weeks": 52, "All uploads": None}
TRENDS = [("Members", "{:,.0f}", ""), ("Total sessions/week", "{:,.0f}", ""),
          ("Sessions per member", "{:.2f}", ""), ("ARPM", "${:,.2f}", "$"),
          ("Revenue per session", "${:,.2f}", "$"), ("Margin per member a week", "${:,.2f}", "$")]
MIX_COLOURS = {"1×": "#8fcfcf", "2×": "#1f9a9a", "3+": "#1f3b57", "1:1": "#e0a100"}


def holiday_periods(holidays, state):
    """[(first day, last day)] of school holidays for a state (School Holidays tab: State,
    Year, Start_Week, End_Week as ISO weeks; an end week below the start runs into the
    next year)."""
    if holidays is None or holidays.empty or not state or "State" not in holidays:
        return []
    out = []
    for _, h in holidays[holidays["State"] == state].iterrows():
        try:
            y, a, b = int(h["Year"]), int(h["Start_Week"]), int(h["End_Week"])
            d0 = pd.Timestamp.fromisocalendar(y, a, 1)
            d1 = pd.Timestamp.fromisocalendar(y + (b < a), b, 7)
        except (ValueError, TypeError):
            continue
        out.append((d0, d1))
    return out


def _weeks(comp, span_key):
    """Every week from the first upload to the last (weeks with no upload included, so they
    show as gaps), cut to the chosen span."""
    up = comp["Week ending"].dropna()
    if up.empty:
        return []
    weeks = list(pd.date_range(up.max() - pd.Timedelta(weeks=int(
        (up.max() - up.min()).days // 7)), up.max(), freq="7D"))
    n = SPANS[st.session_state.get(span_key, "Last 26 weeks")]
    return weeks[-n:] if n else weeks


def trend_charts(series, weeks, typical=None, periods=(), cur="", note=None, key="t"):
    """Six small trend charts and the plan mix. series / typical: measures (trend_measures)
    indexed by week ending; weeks with no upload are left as gaps."""
    import plotly.graph_objects as go
    s = series.reindex(weeks)
    ty = typical.reindex(weeks) if typical is not None else None
    lo, hi = (weeks[0] - pd.Timedelta(days=4), weeks[-1] + pd.Timedelta(days=4)) \
        if weeks else (None, None)

    def frame(fig, title, prefix="", pct=False, height=250):
        for d0, d1 in periods:
            if lo is not None and d1 >= lo and d0 <= hi:
                fig.add_vrect(x0=max(d0, lo), x1=min(d1, hi), fillcolor="#e0a100",
                              opacity=0.10, layer="below", line_width=0)
        fig.update_layout(title=dict(text=title, font=dict(size=13)), height=height,
                          margin=dict(l=10, r=10, t=40, b=10), plot_bgcolor="#fff",
                          paper_bgcolor="#fff", showlegend=False, hovermode="x unified",
                          xaxis=dict(gridcolor="#eef1f3", tickformat="%d %b", range=[lo, hi]),
                          yaxis=dict(gridcolor="#eef1f3", tickprefix=prefix,
                                     ticksuffix="%" if pct else "", zeroline=False))
        return fig

    cols = st.columns(3)
    for i, (name, fmt, prefix) in enumerate(TRENDS):
        fig = go.Figure()
        if ty is not None and ty[name].notna().any():
            fig.add_scatter(x=weeks, y=ty[name], mode="lines", name="Typical centre",
                            line=dict(color="#8a9a9f", dash="dash", width=1.5),
                            connectgaps=False)
        fig.add_scatter(x=weeks, y=s[name], mode="lines+markers", name="This view",
                        line=dict(color="#1f9a9a", width=2.5), marker=dict(size=5),
                        connectgaps=False)
        vals = pd.concat([s[name], ty[name] if ty is not None else pd.Series(dtype=float)])
        vals = vals.dropna()
        if len(vals):                                   # at least ±10%, so small wobbles
            mid, half = (vals.max() + vals.min()) / 2, (vals.max() - vals.min()) / 2
            half = max(half * 1.15, abs(mid) * 0.10, 1e-6)  # don't look like big swings
            fig.update_yaxes(range=[max(mid - half, 0) if vals.min() >= 0 else mid - half,
                                    mid + half])
        last = s[name].dropna()
        title = f"{name}" + (f" · {fmt.format(last.iloc[-1])}" if len(last) else "")
        with cols[i % 3]:
            if name == "Margin per member a week" and not s[name].notna().any():
                st.caption(f"**{name}**: no tutor rate set for this country yet.")
            else:
                st.plotly_chart(frame(fig, title, prefix), width="stretch",
                                key=f"trend_{key}_{name}")
    # Plan mix: the first week shown, the latest week and the ideal mix, as pies.
    have = s[mc.MIX].dropna(how="all")
    have = have[have.sum(axis=1) > 0]
    pies = []
    if len(have):
        if len(have) > 1:
            pies.append((f"{have.index[0]:%d %b %Y}", have.iloc[0]))
        pies.append((f"Latest · {have.index[-1]:%d %b %Y}", have.iloc[-1]))
    pies.append(("Ideal mix", pd.Series(mc.TARGET_MIX)))
    st.markdown("**Plan mix · % of members**")
    for n, (col, (title, mix)) in enumerate(zip(st.columns(3), pies)):
        fig = go.Figure(go.Pie(
            labels=list(mc.MIX), values=[float(mix.get(k, 0) or 0) for k in mc.MIX],
            sort=False, hole=0.45, marker=dict(colors=[MIX_COLOURS[k] for k in mc.MIX],
                                               line=dict(color="#fff", width=2)),
            texttemplate="%{label}<br>%{percent:.0%}", textposition="inside",
            hovertemplate="%{label}: %{percent:.0%}<extra></extra>"))
        fig.update_layout(title=dict(text=title, font=dict(size=13)), height=280,
                          margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor="#fff",
                          legend=dict(orientation="h", yanchor="top", y=-0.02, x=0.5,
                                      xanchor="center"), uniformtext_minsize=10,
                          uniformtext_mode="hide")
        with col:
            st.plotly_chart(fig, width="stretch", key=f"trend_{key}_mix_{n}")
    st.caption((note + " " if note else "") +
               "Gaps = weeks with no upload. Shaded = school holidays. Plan mix compares the first and "
               "latest week shown with the ideal mix. Amounts weekly, in the "
               f"centre's own currency{f' ({cur})' if cur else ''}. Margin per member is after "
               "tutor cost at the target fill for the centre's size.")


def _trend_controls(key):
    st.selectbox("Weeks shown", list(SPANS), index=1, key=key)


def _state(vl, locs, state_of):
    """The (most common) state of these centres, for school holiday shading."""
    if state_of is None or vl is None or vl.empty or mc.LOC not in vl.columns or \
            "Region" not in vl.columns:
        return None
    states = [state_of(r) for r in vl.loc[vl[mc.LOC].isin(locs), "Region"]]
    states = [x for x in states if x]
    return max(set(states), key=states.count) if states else None


GROUP_VIEWS = {
    "Summary": ["Centre", "Country", "Members", "ARPM", "% of RRP", "Students per hub",
                "Extra hubs a week", "Margin per member a week"],
    "Pricing": ["Centre", "Country", "Members", "ARPM", "Revenue per session", "% of RRP",
                "Members on 4×+", "ARPM vs country median", "RRP used"],
    "Scheduling": ["Centre", "Country", "Members", "Total sessions/week",
                   "Sessions per member", "Target fill", "Hubs needed a week",
                   "Hubs run a week", "Students per hub", "Extra hubs a week", "Check"],
    "Margin": ["Centre", "Country", "Members", "Revenue per session", "Tutor cost %",
               "Margin per member a week"],
}


def _group(comp, week, vl, table, allowed, holidays=None, state_of=None, revenue=None):
    full = mc.group_table(week, vl, table)
    t = _visible(week, vl, table, allowed)
    if not t.empty:
        t = mc.add_measured(t, revenue, week["Week ending"].iloc[0])
    countries = sorted(t["Country"].unique())
    pick = st.selectbox("Country", ["All countries"] + countries, key="comp_country")
    if pick != "All countries":
        t = t[t["Country"] == pick]
    st.caption("Lowest % of RRP first. Amounts weekly, in each centre's own currency. "
               "ARPM vs country median: \"–\" where the country has one centre"
               + ("." if allowed is None else
                  f" or fewer than {MIN_FOR_MEDIAN} centres.") +
               " Target fill rises with members (3.0 up to 20, 5.0 at 120, 6.0 from 180); "
               "tutor cost and margin are AU only for now. Hubs run come from the CRM's "
               "weekly export; left blank where its students differ from the members by "
               "more than 20% (see Check).")
    view = st.radio("Show", list(GROUP_VIEWS), horizontal=True, key="comp_group_view")
    show = t[GROUP_VIEWS[view]]
    money = lambda: st.column_config.NumberColumn(format="$%.2f")
    st.dataframe(show.style.map(lambda v: "" if pd.isna(v) else
                                f"color:{GREEN if v >= 99.5 else RED};font-weight:600",
                                subset=[c for c in ["% of RRP"] if c in show.columns]),
                 hide_index=True, width="stretch", key="comp_group",
                 column_config={"ARPM": money(), "Revenue per session": money(),
                                "Margin per member a week": money(),
                                "Total sessions/week": st.column_config.NumberColumn(
                                    format="%d"),
                                "Sessions per member": st.column_config.NumberColumn(
                                    format="%.2f"),
                                "Target fill": st.column_config.NumberColumn(
                                    format="%.1f", help="Target students per study hub for "
                                    "a centre this size"),
                                "Hubs needed a week": st.column_config.NumberColumn(
                                    format="%.0f", help="Group sessions ÷ target fill, plus "
                                    "1:1 sessions"),
                                "Hubs run a week": st.column_config.NumberColumn(
                                    format="%.0f", help="Sessions run that week (CRM, "
                                    "weekly Revenue export)"),
                                "Students per hub": st.column_config.NumberColumn(
                                    format="%.1f", help="Group sessions booked ÷ group hubs "
                                    "run (1:1 left out). Booked, not attended: actual is a "
                                    "little lower."),
                                "Extra hubs a week": st.column_config.NumberColumn(
                                    format="%+.0f", help="Hubs run − hubs needed at the "
                                    "target fill. + = more hubs than needed"),
                                "Check": st.column_config.TextColumn(
                                    help="Why measured figures are missing"),
                                "Tutor cost %": st.column_config.NumberColumn(
                                    format="%.0f%%", help="Hubs needed × tutor cost an hour, "
                                    "as a % of total value"),
                                "% of RRP": st.column_config.NumberColumn(
                                    format="%.0f%%", help="Recommended plans only (4×+ "
                                    "group plans left out)"),
                                "Members on 4×+": st.column_config.NumberColumn(
                                    format="%d", help="Members on group plans with 4 or more "
                                    "sessions a week: not recommended"),
                                "ARPM vs country median": st.column_config.NumberColumn(
                                    format="$%+.2f")})
    chart_country = pick if pick != "All countries" else (
        "AU" if "AU" in countries else (countries[0] if countries else None))
    if chart_country:
        c = t[(t["Country"] == chart_country) & (t["Members"] > 0)]
        st.markdown(f"**Price vs frequency · {chart_country}**")
        st.caption("Each dot is a centre (bigger = more members; faded = under "
                   f"{SMALL} members). Dashed lines: the country's medians. Top-left is the "
                   "strongest position: members come less often and each session earns more.")
        price_frequency_chart(c, country_medians(full, chart_country, allowed),
                              chart_country, None)
        w = mc.weekly_totals(comp)
        w = w[(w["Country"] == chart_country) &
              (w["Location"].isin(allowed) if allowed is not None else True)]
        if not w.empty:
            st.markdown(f"**Group trends · {chart_country}**" +
                        ("" if allowed is None else " · your centres"))
            _trend_controls("comp_trend_g")
            cols = ["Members", "Sessions", "One", "Value"] + mc.MIX
            g = w.groupby("Week ending")[cols].sum()
            g["Margin"] = w.groupby("Week ending")["Margin"].sum(min_count=1)
            locs = sorted(w["Location"].unique())
            trend_charts(mc.trend_measures(g), _weeks(comp, "comp_trend_g"),
                         periods=holiday_periods(holidays, _state(vl, locs, state_of)),
                         cur=mc.CURRENCY.get(chart_country, ""),
                         note=f"All {len(locs)} centres added together; prices worked out "
                              "from the totals, so bigger centres count more.", key="g")


def _money_section(plans, rrp, ideal, row, week, vl, table, allowed, revenue):
    """Today vs achievable after tutor cost, the gain split into price / mix / scheduling,
    and this centre on the price vs frequency chart with an arrow to where it could be."""
    import pnl_share
    loc, country = row["Location"], row["Country"]
    run = mc.hubs_run(revenue, week["Week ending"].iloc[0]).get(loc)
    check = mc.measured(row, run)["Check"] if run else "No sessions-run data"
    r = mc.where_the_money_is(plans, rrp, ideal, country, None if check else run)
    t, a = r["today"], r["ideal"]
    cur = mc.CURRENCY.get(country, "")

    def m(v):
        return _m(v)

    def n(v, dp=1):
        return "–" if v is None or pd.isna(v) else f"{v:,.{dp}f}"

    def change(x, y, money=True, dp=1):
        if x is None or y is None or pd.isna(x) or pd.isna(y):
            return ""
        d = y - x
        txt = (f"{'+' if d >= 0 else '−'}${abs(d):,.2f}" if money else
               f"{'+' if d >= 0 else '−'}{abs(d):,.{dp}f}")
        return txt
    lines = [("Revenue per session", m(t["Revenue per session"]), m(a["Revenue per session"]),
              change(t["Revenue per session"], a["Revenue per session"])),
             ("Sessions per member", n(t["Sessions per member"], 2),
              n(a["Sessions per member"], 2),
              change(t["Sessions per member"], a["Sessions per member"], False, 2)),
             ("Study hubs a week", n(t["Hubs a week"], 0) + ("" if r["hubs_run"] else
                                                             " (needed)"),
              n(a["Hubs a week"], 0), change(t["Hubs a week"], a["Hubs a week"], False, 0)),
             ("Tutor cost % of fees", n(t["Tutor cost %"], 0) + "%" if t["Tutor cost %"]
              else "–", n(a["Tutor cost %"], 0) + "%" if a["Tutor cost %"] else "–", ""),
             ("Margin a week (after tutor cost)", m(t["Margin a week"]),
              m(a["Margin a week"]), change(t["Margin a week"], a["Margin a week"]))]
    k = len(lines) - 1
    table_html = _table(["", "Today", "Achievable", "Change"],
                        [list(x) for x in lines], {k: "total"})
    gains = [g for g in r["gains"]]
    total = sum(g[1] for g in gains)
    colours = {"Price": "#1f9a9a", "Mix": "#1f3b57", "Scheduling": "#e0a100"}
    tiles = "".join(
        f'<div class="mc-gain{"" if v >= 0.5 else " zero"}" style="--c:'
        f'{colours.get(name, "#1f9a9a") if v >= 0.5 else "#c9d3d6"}">'
        f'<div class="k">{html.escape(name)}</div>'
        f'<div class="v">{"+" + _m(v) if v >= 0.5 else "$0"} <small>a week</small></div>'
        f'<div class="w">{html.escape(why)}</div></div>' for name, v, why in gains)
    st.markdown('<div class="section-header">Where the money is</div>',
                unsafe_allow_html=True)
    head = (f'Achievable: <b>+{_m(total)} a week</b> ({cur}) after tutor cost, biggest '
            'first:' if total >= 0.5 else
            "Already at or above the achievable position.")
    st.markdown((f'<div class="mc-total">{head}</div><div class="mc-gains">{tiles}</div>')
                .replace("$", "&#36;"), unsafe_allow_html=True)
    sub = f"Margin a week: {_m(t['Margin a week'])} → {_m(a['Margin a week'])}" \
        if t["Margin a week"] is not None else "Fees only (no tutor rate for this country)"
    _card("Today vs achievable", sub, table_html)
    notes = [f"Fill used: {r['fill']:.1f} students a hub (target for {row['Members']:.0f} "
             f"members: {r['target']:.1f}"
             + (f"; measured {r['measured_fill']:.1f}" if r["measured_fill"] else "") + ")."]
    if check:
        notes.append(f"Hubs run not used: {check}.")
    if r["rate"] is None:
        notes.append("No tutor rate for this country yet: tutor cost and margins show –, and "
                     "the gains are fees only.")
    if not r["use_mix"]:
        notes.append("Achievable keeps today's mix, because it already earns more than the "
                     "ideal mix.")
    nr = r["not_recommended"]
    if nr["members"]:
        notes.append(f"{nr['members']} members ({nr['share']:.0f}%) on "
                     f"{mc.NOT_RECOMMENDED_FROM}×+ plans: not recommended, kept at today's "
                     f"price ({_m(nr['value'])} a week).")
    months = pnl_share.months_trading(vl, loc, pd.Timestamp(
        week["Week ending"].iloc[0]).strftime("%Y-%m"))
    if months is not None and months <= pnl_share.NEW_MONTHS:
        notes.append(f"Trading {months} months: treat the hubs figure as guidance while "
                     "groups are still forming.")
    st.caption(" ".join(notes))
    full = mc.group_table(week, vl, table)
    vis = _visible(week, vl, table, allowed)
    pool = full if allowed is None else vis
    c = pool[(pool["Country"] == country) & (pool["Members"] > 0)]
    price_frequency_chart(c, country_medians(full, country, allowed), country,
                          "Today → achievable", highlight=loc,
                          ideal=(a["Sessions per member"], a["Revenue per session"]),
                          key="pf_centre")


def _centre_trends(comp, vl, loc, pick, country, allowed, holidays, state_of):
    import pnl_share
    w = mc.weekly_totals(comp)
    w = w[w["Country"] == country]
    mine = w[w["Location"] == loc].set_index("Week ending")
    if mine.empty:
        return
    st.markdown(f"**Trends · {pick}**")
    _trend_controls("comp_trend_c")
    weeks = _weeks(comp, "comp_trend_c")
    latest = pd.Timestamp(max(comp["Week ending"])).strftime("%Y-%m")
    months = pnl_share.months_trading(vl, loc, latest)
    too_new = months is not None and months <= pnl_share.NEW_MONTHS
    typical, note = None, None
    if too_new:
        note = (f"No typical-centre line: {pick} has been trading {months} months "
                f"({pnl_share.NEW_MONTHS} or less).")
    else:
        others = w[(w["Location"] != loc) & (w["Members"] > 0)]
        if len(others):
            m = mc.trend_measures(others).assign(**{"Week ending": others["Week ending"]})
            count = m.groupby("Week ending")["Members"].count()
            typical = m.groupby("Week ending").median(numeric_only=True)
            typical = typical[count.reindex(typical.index) >= MIN_FOR_MEDIAN]
            note = (f"Dashed line: the typical (median) {country} centre, shown in weeks with "
                    f"at least {MIN_FOR_MEDIAN} other centres.")
    trend_charts(mc.trend_measures(mine), weeks, typical,
                 holiday_periods(holidays, _state(vl, [loc], state_of)),
                 mc.CURRENCY.get(country, ""), note, key="c")


def _centre(comp, week, vl, table, allowed, holidays=None, state_of=None, revenue=None):
    g = _visible(week, vl, table, allowed)
    names = dict(sorted(zip(g["Centre"], g["Location"]), key=lambda kv: kv[0].lower()))
    pick = st.selectbox("Centre", list(names), key="comp_centre")
    loc = names[pick]
    row = g[g["Location"] == loc].iloc[0]
    region = None
    if vl is not None and not vl.empty and mc.LOC in vl.columns and "Region" in vl.columns:
        match = vl[vl[mc.LOC] == loc]
        region = match["Region"].iloc[0] if not match.empty else None
    rrp = mc.rrp_for(table, row["Country"], region)
    plans = week[week["Location"] == loc]
    cur, ideal = mc.current_table(plans, rrp, mc.plan_medians(week)), mc.ideal_table(plans, rrp)
    title = f'{html.escape(loc)} <span>{html.escape(row["Country"])}</span>'
    rrp_txt = (f'RRP ({row["RRP used"]}): ${rrp["mkt"][0]:,.0f} / ${rrp["mkt"][1]:,.0f} / '
               f'${rrp["mkt"][2]:,.0f} a week for 1× / 2× / 3×, 1:1 ${rrp["one"]:,.0f} a '
               f'session ({row["Currency"]})') if rrp else "No RRP set for this country"
    st.caption(rrp_txt)

    note = (f'Members under ${mc.FREEZE_BELOW} (not shown in the rows, included in totals and '
            f'composition): {cur["freeze_members"]}'
            + (f' (${cur["freeze_value"]:,.2f} a week)' if cur["freeze_members"] else "") + ".")
    nr = cur["rows"][cur["rows"]["Not recommended"]] if len(cur["rows"]) else cur["rows"]
    nr_note = (f' {int(nr["Members"].sum())} members are on group plans with '
               f'{mc.NOT_RECOMMENDED_FROM}+ sessions a week (not recommended): they are kept at '
               "today's price and left out of the ideal mix and % of RRP." if len(nr) else "")
    _card(title, f'{cur["members"]} members · {_m(cur["total"])} · ARPM {_m(cur["arpm"])}',
          '<div class="mc-sub">Plan mix</div>' + composition_html(cur["composition"])
          + f'<div class="mc-note">{note}{nr_note}</div>')
    with st.expander("Plan details (every price)"):
        st.markdown(f'<div class="mc-card">{current_html(cur)}</div>'.replace("$", "&#36;"),
                    unsafe_allow_html=True)
    own = ideal["rows"]["Own price used"].any()
    diff = ideal["total"] - cur["total"]
    _card("Ideal plan mix at RRP",
          f'{"+" if diff >= 0 else "−"}${abs(diff):,.0f} a week vs today (fees)',
          current_vs_ideal_html(plans, ideal)
          + '<div class="mc-note">' + ("¹ The centre's own average price is used where it is "
                                       "higher than RRP. " if own else "")
          + "Fees only, before tutor cost. See 'Where the money is' for the best achievable "
            "after tutor cost.</div>")
    _money_section(plans, rrp, ideal, row, week, vl, table, allowed, revenue)
    _centre_trends(comp, vl, loc, pick, row["Country"], allowed, holidays, state_of)


def report_composition(df_comp, vl, df_rrp, allowed=None, holidays=None, state_of=None,
                       revenue=None):
    """allowed: None for admins (every centre), else the locations this viewer may see."""
    st.markdown('<div class="report-title">Membership Composition</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Each centre\'s plans and prices against RRP, '
                'and the ideal composition to achieve RRP.</div>', unsafe_allow_html=True)
    if st.session_state.get("access_level") != "admin" and allowed is None:
        allowed = []                     # never show every centre to a non-admin
    comp = mc.prepare(df_comp)
    if comp.empty:
        st.info("No Membership Composition data yet. An admin can add a week on the "
                "⚙ Composition Upload page.")
        return
    st.markdown(CSS, unsafe_allow_html=True)
    table = mc.rrp_table(df_rrp)
    weeks = sorted({d.strftime("%Y-%m-%d") for d in comp["Week ending"]}, reverse=True)
    with st.container(border=True):
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b = st.columns([2, 1])
        view = a.radio("View", VIEWS, horizontal=True, key="comp_view",
                       label_visibility="collapsed")
        wk = b.selectbox("Week ending", weeks, key="comp_week_pick",
                         format_func=lambda d: pd.Timestamp(d).strftime("%d/%m/%Y"))
    week = comp[comp["Week ending"] == pd.Timestamp(wk)]
    if allowed is not None and not week["Location"].isin(allowed).any():
        st.info("None of your centres is in this week's Membership Composition data.")
        return
    if view == VIEWS[0]:
        _group(comp, week, vl, table, allowed, holidays, state_of, revenue)
    else:
        _centre(comp, week, vl, table, allowed, holidays, state_of, revenue)
    with st.expander("How this is worked out"):
        st.markdown(HOW)
