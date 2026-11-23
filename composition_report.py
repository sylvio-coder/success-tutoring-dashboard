"""Membership Composition report (admins only): a group overview of every centre's
ARPM, revenue per session and % of RRP, and a centre view with its current composition
and the ideal composition to achieve RRP, in the portal's table layout."""
import html

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
</style>"""

HOW = f"""
- **Source:** the portal's Membership Composition export, uploaded weekly by an admin
  (⚙ Composition Upload). Prices are weekly, in the centre's own currency; US prices are
  billed monthly and divided by 4.
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
             _var(r["RRP variance"], r["RRP variance %"]),
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


def composition_html(comp):
    rows = [[html.escape(r["Plan"]), f'{int(r["Members"]):,}', f'{r["% of members"]:.0f}%',
             _n(r["Total sessions"]), _m(r["Average price"]), _m(r["Revenue per session"]),
             _m(r["Total value"])] for _, r in comp.iterrows()]
    n, sess, value = comp["Members"].sum(), comp["Total sessions"].sum(), comp["Total value"].sum()
    rows.append(["Total", f"{int(n):,}", "100%" if n else "–", _n(sess),
                 _m(value / n if n else None), _m(value / sess if sess else None), _m(value)])
    return _table(["Composition", "No. members", "% of members", "Total sessions/week",
                   "Average membership price", "Revenue per session", "Total value"],
                  rows, {len(rows) - 1: "total"})


def ideal_html(t):
    rows = [[f'Group {r["Plan"]}' if r["Type"] == "Group" else r["Plan"],
             f'{r["Members"]:.1f}', _n(r["Total sessions"], 1),
             _m(r["Membership price"]) + (" ¹" if r["Own price used"] else ""),
             _m(r["Revenue per session"]), _m(r["Total value"])]
            for _, r in t["rows"].iterrows()]
    k = len(rows)
    sign = "+" if t["vs_today"] >= 0 else "−"
    rows.append(["Total", f'{t["members"]:,}', _n(t["sessions"], 1), "",
                 _m(t["per_session"]), _m(t["total"])])
    rows.append(["Per member", "", f'{_n(_per_member(t), 2)} sessions',
                 f'{_m(t["arpm"])} ARPM', "",
                 f'{sign}${abs(t["vs_today"]):,.2f} a week vs today'])
    return _table(["Plan", "No. members", "Total sessions/week", "Membership price",
                   "Revenue per session", "Total value"], rows, {k: "total", k + 1: "arpm"})


def ideal_composition_html(comp):
    return _table(["Composition", "Current %", "Ideal %", "Variance"],
                  [[html.escape(r["Plan"]), f'{r["Current %"]:.0f}%', f'{r["Ideal %"]:.0f}%',
                    _pts(r["Variance"])] for _, r in comp.iterrows()])


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


def _group(week, vl, table, allowed):
    t = _visible(week, vl, table, allowed)
    countries = sorted(t["Country"].unique())
    pick = st.selectbox("Country", ["All countries"] + countries, key="comp_country")
    if pick != "All countries":
        t = t[t["Country"] == pick]
    st.caption("Lowest % of RRP first. Amounts weekly, in each centre's own currency. "
               "ARPM vs country median: \"–\" where the country has one centre"
               + ("." if allowed is None else
                  f" or fewer than {MIN_FOR_MEDIAN} centres."))
    t = t.assign(**{"Sessions per member": t["Total sessions/week"] /
                    t["Members"].where(t["Members"] > 0)})
    show = t[["Centre", "Country", "Members", "Total sessions/week", "Sessions per member",
              "ARPM", "Revenue per session", "% of RRP", "ARPM vs country median",
              "RRP used"]]
    st.dataframe(show.style.map(lambda v: "" if pd.isna(v) else
                                f"color:{GREEN if v >= 99.5 else RED};font-weight:600",
                                subset=["% of RRP"]),
                 hide_index=True, width="stretch", key="comp_group",
                 column_config={"ARPM": st.column_config.NumberColumn(format="$%.2f"),
                                "Total sessions/week": st.column_config.NumberColumn(
                                    format="%d"),
                                "Sessions per member": st.column_config.NumberColumn(
                                    format="%.2f"),
                                "Revenue per session": st.column_config.NumberColumn(
                                    format="$%.2f"),
                                "% of RRP": st.column_config.NumberColumn(format="%.0f%%"),
                                "ARPM vs country median": st.column_config.NumberColumn(
                                    format="$%+.2f")})


def _centre(comp, week, vl, table, allowed):
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
    _card(title, f'{cur["members"]} members · {_m(cur["total"])} · ARPM {_m(cur["arpm"])}',
          current_html(cur) + '<div style="height:10px"></div>'
          + composition_html(cur["composition"]) + f'<div class="mc-note">{note}</div>')
    own = ideal["rows"]["Own price used"].any()
    _card("Ideal composition to achieve RRP",
          f'{ideal["members"]} members · {_m(ideal["total"])} · ARPM {_m(ideal["arpm"])}',
          ideal_html(ideal) + '<div style="height:10px"></div>'
          + ideal_composition_html(ideal["composition"])
          + ('<div class="mc-note">¹ The centre\'s own average price is used where it is '
             'higher than RRP.</div>' if own else ""))


def report_composition(df_comp, vl, df_rrp, allowed=None):
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
        _group(week, vl, table, allowed)
    else:
        _centre(comp, week, vl, table, allowed)
    with st.expander("How this is worked out"):
        st.markdown(HOW)
