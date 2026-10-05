"""P&L reports: summary in the centres side-by-side layout, and the break-even model."""
import html
import io

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import pnl

LOC = "Success Tutoring - Business name"
PREFIX = "Success Tutoring - "
VIEWS = ["Centre", "State", "Country", "GPM", "Network"]
PERIODS = ["Single month", "Last 3 months", "Year to date", "All months"]

# Rows of the side-by-side table: (label, key, style)
TABLE_ROWS = [
    ("Members", "Members", "input"),
    ("Average Membership Value $ (per week)", "AMV", "input"),
    ("Months reported", "Months", "input"),
    (None, None, "gap"),
    ("Total Revenue", "Total Revenue", "total"),
    ("Tutor Wages + Super", "Tutor Wages + Super", "line"),
    ("Rent", "Rent", "line"),
    ("Wages - Manager/Admin + Super", "Wages - Manager/Admin + Super", "line"),
    ("Royalties", "Royalties", "line"),
    ("Marketing - Agency Fees", "Marketing - Agency Fees", "line"),
    ("Marketing", "Marketing", "line"),
    ("OPEX", "OPEX", "line"),
    ("EBITDA", "EBITDA", "result"),
    ("Interest", "Interest", "line"),
    ("Net Profit", "Net Profit", "result"),
    (None, None, "gap"),
    ("Break-even revenue", "BE revenue", "money"),
    ("Break-even members", "BE members", "count"),
    ("Members above / below break-even", "BE gap", "count"),
]
AMOUNT_KEYS = pnl.ROW_ORDER + ["EBITDA", "Interest", "Net Profit"]

CSS = """<style>
.pnl-wrap{overflow-x:auto;border:1px solid #dde5e6;border-radius:6px;background:#fff;margin-bottom:8px}
.pnl{border-collapse:collapse;font-size:13px;white-space:nowrap;width:100%}
.pnl th{background:#1f3b57;color:#fff;font-weight:600;padding:6px 10px;text-align:center}
.pnl th.sub{background:#2c4f72;font-weight:500;font-size:12px}
.pnl td{padding:3px 10px;text-align:right;border-bottom:1px solid #eef2f3}
.pnl td.lbl{text-align:left;position:sticky;left:0;background:#fff;min-width:250px;z-index:1}
.pnl th.lbl{text-align:left;position:sticky;left:0;z-index:2}
.pnl td.note{text-align:left;color:#5d6d73;font-size:12px}
.pnl tr.input td{background:#eef1f3}
.pnl tr.input td.lbl{background:#eef1f3}
.pnl tr.total td,.pnl tr.result td{font-weight:700}
.pnl tr.result td{border-top:1.5px solid #1c2a30;border-bottom:2px solid #1c2a30}
.pnl tr.gap td{border:none;height:10px}
.pnl td.neg{color:#c8322f}
.pnl td.pct{color:#5d6d73}
.pnl td.hl{background:#dbeaf6}
</style>"""


def _money(v):
    if v is None or pd.isna(v):
        return "–"
    s = f"${abs(v):,.0f}"
    return f"({s})" if v < -0.5 else s


def _pct(v, rev):
    if v is None or pd.isna(v) or not rev or rev <= 0:
        return ""
    return f"{v / rev * 100:.1f}%"


def _count(v, signed=False):
    if v is None or pd.isna(v):
        return "–"
    return f"{v:+,.0f}" if signed else f"{v:,.0f}"


# ── Data ──────────────────────────────────────────────────────────────────────
def members_by_month(df_wm):
    """Average active members per location per month."""
    if df_wm.empty or "Date" not in df_wm.columns:
        return pd.Series(dtype=float)
    d = df_wm.assign(Period=df_wm["Date"].dt.strftime("%Y-%m"))
    return d.groupby([LOC, "Period"])["# Active members"].mean()


def select_periods(periods, mode, month):
    periods = sorted(periods)
    if mode == "Single month":
        return [month]
    if mode == "Last 3 months":
        return [p for p in periods if p <= month][-3:]
    if mode == "Year to date":
        # Calendar year (January onwards), the same for every country.
        start = f"{month[:4]}-01"
        return [p for p in periods if start <= p <= month]
    return periods


def centre_table(df_pnl, members, vl, periods, age_group):
    """One row per location: monthly averages over the chosen months plus its details."""
    d = df_pnl[df_pnl["Period"].isin(periods)]
    if d.empty:
        return pd.DataFrame()
    lm = pnl.location_months(d)
    months = lm.groupby(level="Location").size()
    out = lm.groupby(level="Location").mean()
    out["Months"] = months
    mem = members[members.index.get_level_values("Period").isin(periods)] \
        if not members.empty else members
    out["Members"] = mem.groupby(level=0).mean().reindex(out.index) if not mem.empty else None
    last = d.sort_values("Period").groupby("Location").last()
    for c in ["Country", "Currency", "Status"]:
        out[c] = last[c].reindex(out.index)
    out["FX"] = pd.to_numeric(last["FX rate to AUD"], errors="coerce").reindex(out.index)
    out.loc[out["Currency"] == "AUD", "FX"] = 1.0
    info = vl.set_index(LOC) if LOC in vl.columns else pd.DataFrame()
    for c in ["Region", "Stage", "GPM", "Age (Months)", "Royalty Min"]:
        out[c] = info[c].reindex(out.index) if c in info.columns else None
    if "Country" in info.columns:
        blank = out["Country"].fillna("").astype(str).str.strip() == ""
        out.loc[blank, "Country"] = info["Country"].reindex(out.index)[blank]
    out["Age group"] = out["Age (Months)"].map(age_group) if "Age (Months)" in out else "Unknown"
    rmin = pd.to_numeric(out["Royalty Min"].astype(str).str.replace(r"[$,\s]", "", regex=True),
                         errors="coerce")
    out["Royalty Min"] = rmin.fillna(pnl.DEFAULT_ROYALTY_MIN)
    out["Band"] = out["Members"].map(pnl.membership_band)
    return out


def combine(rows, to_aud):
    """Add centres together (monthly averages). to_aud converts each centre first."""
    keys = AMOUNT_KEYS + ["Royalty Min"]
    r = rows.copy()
    if to_aud:
        r = r[r["FX"].notna()]
        for k in keys:
            r[k] = r[k] * r["FX"]
    s = r[keys].sum()
    s["Members"] = r["Members"].sum(min_count=1)
    s["Months"] = r["Months"].max()
    s["Centres"] = len(r)
    return s


def add_break_even(col, rates):
    """Break-even with tutor % from the members band. A group counts as its number of
    centres, each with the group's average members and costs."""
    rev, mem = col["Total Revenue"], col.get("Members")
    mem = 0 if mem is None or pd.isna(mem) else mem
    n = int(col.get("Centres", 1) or 1)
    fixed = sum(col[k] for k in pnl.FIXED_ROWS)
    per_member = rev / mem if mem else 0
    be_rev, be_mem = pnl.break_even_banded(per_member, fixed, col["Royalty Min"], rates, n)
    col["BE revenue"], col["BE members"] = be_rev, be_mem
    col["BE gap"] = (mem - be_mem) if be_mem is not None and mem else None
    col["AMV"] = rev / mem / pnl.WEEKS_PER_MONTH if mem else None
    return col


def build_columns(cent, view, state_of, rates):
    """[(heading, currency, column)] for the chosen view."""
    cols = []
    if view == "Centre":
        for loc, r in cent.sort_values("Total Revenue", ascending=False).iterrows():
            col = add_break_even(r.copy(), rates)
            col["Locations"] = (loc,)
            cols.append((loc.replace(PREFIX, ""), r["Currency"], col))
        return cols, []
    if view == "Network":
        groups = {"Network": cent}
    else:
        key = cent["Country"].fillna("Not set") if view == "Country" else \
            cent["GPM"].replace("", None).fillna("Not set") if view == "GPM" else \
            cent["Region"].map(state_of).fillna("Not set")
        groups = {k: g for k, g in cent.groupby(key)}
    left_out = []
    for name, g in sorted(groups.items(), key=lambda kv: -kv[1]["Total Revenue"].sum()):
        mixed = g["Currency"].nunique() > 1
        if mixed:
            left_out += g[g["FX"].isna()].index.tolist()
        col = add_break_even(combine(g, mixed), rates)
        countries = g["Country"].dropna().unique()
        col["Country"] = countries[0] if len(countries) == 1 else None
        col["Locations"] = tuple(g.index)
        cols.append((f"{name} ({int(col['Centres'])})", "AUD" if mixed else g["Currency"].iloc[0],
                     col))
    return cols, left_out


# ── Rendering ─────────────────────────────────────────────────────────────────
def side_by_side_html(cols, notes=None, rows=None):
    rows = rows or TABLE_ROWS
    head = '<tr><th class="lbl">Account</th>' + "".join(
        f'<th colspan="2">{html.escape(h)}<br><span style="font-weight:400;font-size:11px">'
        f'{html.escape(str(cur))}</span></th>' for h, cur, _ in cols) + \
        ("<th></th>" if notes else "") + "</tr>"
    body = []
    for label, key, style in rows:
        if style == "gap":
            body.append(f'<tr class="gap"><td colspan="{2 * len(cols) + 1}"></td></tr>')
            continue
        cells = []
        for _, _, c in cols:
            v = c.get(key)
            rev = c.get("Total Revenue")
            if key in AMOUNT_KEYS:
                neg = " neg" if v is not None and not pd.isna(v) and v < -0.5 else ""
                cells.append(f'<td class="{neg}">{_money(v)}</td><td class="pct">{_pct(v, rev)}</td>')
            elif key == "AMV" or style == "money":
                cells.append(f"<td>{_money(v)}</td><td></td>")
            elif key == "BE gap":
                neg = " neg" if v is not None and not pd.isna(v) and v < 0 else ""
                cells.append(f'<td class="{neg}">{_count(v, True)}</td><td></td>')
            else:
                cells.append(f"<td>{_count(v)}</td><td></td>")
        note = f'<td class="note">{html.escape(notes.get(key, ""))}</td>' if notes else ""
        body.append(f'<tr class="{style}"><td class="lbl">{html.escape(label)}</td>'
                    + "".join(cells) + note + "</tr>")
    return CSS + f'<div class="pnl-wrap"><table class="pnl">{head}{"".join(body)}</table></div>'


def to_excel(cols, sheet="P&L"):
    data = {}
    for h, cur, c in cols:
        rev = c.get("Total Revenue")
        data[(f"{h} ({cur})", "$")] = [c.get(k) if k else None for _, k, _ in TABLE_ROWS]
        data[(f"{h} ({cur})", "%")] = [
            (c.get(k) / rev if k in AMOUNT_KEYS and rev and rev > 0 else None) if k else None
            for _, k, _ in TABLE_ROWS]
    df = pd.DataFrame(data, index=[l or "" for l, _, _ in TABLE_ROWS])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name=sheet)
    return buf.getvalue()


# ── Report ────────────────────────────────────────────────────────────────────
def tutor_rate_settings(bench_cent, key):
    """Editable tutor % by members band, starting from the benchmark. Returns {band: %}."""
    bench = pnl.tutor_benchmarks(bench_cent).rename_axis("Members band").reset_index()
    with st.expander("Break-even settings: tutor wages + super % of revenue by members band"):
        st.caption("Tutor costs fall as a share of revenue as centres grow, so the break-even "
                   "uses the typical (median) % for centres of each size instead of carrying a "
                   "centre's current % forward. Worked out from all reporting centres; bands "
                   f"with fewer than {pnl.MIN_BENCH_CENTRES} centres borrow the nearest band. "
                   "Change 'Use %' to test higher or lower costs.")
        edited = st.data_editor(
            bench, key=key, hide_index=True, use_container_width=True,
            disabled=["Members band", "Centres", "Benchmark %", "Source"],
            column_config={
                "Benchmark %": st.column_config.NumberColumn(format="%.1f%%"),
                "Use %": st.column_config.NumberColumn(min_value=0.0, max_value=100.0,
                                                       step=0.5, format="%.1f%%")})
    return dict(zip(edited["Members band"], edited["Use %"].fillna(0).astype(float)))


def report_pnl(df_pnl, df_wm, vl, helpers, bench=None):
    """helpers: age_group, age_order, state_of, show_chart, std_layout, report_filters.
    bench: (P&L, Weekly Membership) for every centre, used only for band benchmarks, so
    someone who sees a few centres still gets network-wide typical values."""
    st.markdown('<div class="report-title">P&L Summary & Break-even</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">Source: P&L tab (monthly P&L uploads) with members '
                'from Weekly Membership. Amounts are average per month, in each centre\'s own '
                'currency.</div>', unsafe_allow_html=True)
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return

    members = members_by_month(df_wm)
    all_periods = sorted(df_pnl["Period"].unique())
    panel = st.container(border=True)
    with panel:
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b, c, d = st.columns(4)
        mode = a.selectbox("Period", PERIODS, key="pnl_mode")
        month = b.selectbox("Month" if mode == "Single month" else "Up to", all_periods[::-1],
                            format_func=pnl.period_name, key="pnl_month")
        view = c.selectbox("Show by", VIEWS, key="pnl_view")
        band_opts = pnl.membership_bands()
        bands = d.multiselect("Members", band_opts, key="pnl_band", placeholder="All sizes")
    periods = select_periods(all_periods, mode, month)

    cent = centre_table(df_pnl, members, vl, periods, helpers["age_group"])
    if cent.empty:
        st.warning("No P&L figures for the chosen months.")
        return
    cent = cent.reset_index().rename(columns={"Location": LOC})
    with panel:
        cent = helpers["report_filters"](cent, key_prefix="pnl", show_date=False,
                                         show_location=view == "Centre", show_status=False,
                                         panel=st.container(), show_label=False)
        e, f = st.columns(2)
        ages = e.multiselect("Age", helpers["age_order"], key="pnl_age", placeholder="All ages")
    if bands:
        cent = cent[cent["Band"].isin(bands)]
    if ages:
        cent = cent[cent["Age group"].isin(ages)]
    cent = cent.set_index(LOC)
    if cent.empty:
        st.warning("No locations match the filters.")
        return

    label = pnl.period_name(periods[0]) if len(periods) == 1 else \
        f"{pnl.period_name(periods[0])} – {pnl.period_name(periods[-1])} ({len(periods)} months)"
    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    bench_cent = centre_table(b_pnl, members_by_month(b_wm), vl, periods, helpers["age_group"])
    rates = tutor_rate_settings(bench_cent, key=f"pnl_rates_{periods[0]}_{periods[-1]}")
    cols, left_out = build_columns(cent, view, helpers["state_of"], rates)

    st.markdown(f'<div class="section-header">P&L by {view.lower()} — {label}</div>',
                unsafe_allow_html=True)
    st.markdown(side_by_side_html(cols), unsafe_allow_html=True)
    st.caption(
        "Average per month over the months each centre reported. Members = average active "
        "members (Weekly Membership); Average Membership Value = revenue ÷ members ÷ 4.2 weeks. "
        "Break-even: tutor wages at the typical % for the members band (see settings above); "
        "rent, manager wages, marketing and OPEX stay fixed; royalties are the higher of the "
        "centre's Royalty Min (Vlookup, blank = 2,200) or 8% of revenue. A state, country or "
        "network counts as its number of centres, each with the average members and costs. "
        "Groups mixing currencies are in AUD.")
    if left_out:
        st.warning("Left out of AUD totals (no exchange rate): " +
                   ", ".join(x.replace(PREFIX, "") for x in left_out))
    st.download_button("Download table (Excel)", to_excel(cols),
                       file_name=f"pnl-{view.lower()}-{periods[-1]}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="pnl_dl")

    acct_bench = pnl.account_table(b_pnl, periods)
    outliers_section(cols, cent, bench_cent, acct_bench, label, helpers)
    network_outliers(cent, bench_cent, label)
    break_even_model(cols, label, helpers, rates)


def break_even_model(cols, label, helpers, rates):
    st.markdown('<div class="section-header">Break-even model</div>', unsafe_allow_html=True)
    names = [h for h, _, _ in cols]
    a, b = st.columns([2, 3])
    pick = a.selectbox("Centre or group", names, key="pnl_model_pick")
    heading, cur, col = next(c for c in cols if c[0] == pick)
    rev, mem = col["Total Revenue"], col.get("Members")
    if not mem or pd.isna(mem) or rev <= 0:
        st.info(f"{pick} has no members data or no revenue for {label}, so the model can't be "
                "worked out.")
        return
    n = int(col.get("Centres", 1) or 1)
    unit = "members" if n == 1 else "members per centre"
    amv0 = rev / mem / pnl.WEEKS_PER_MONTH
    tutor0 = col["Tutor Wages + Super"] / rev * 100
    with b:
        c1, c2 = st.columns(2)
        amv = c1.number_input("Average Membership Value $ / week", min_value=0.0,
                              value=round(float(amv0), 2), step=1.0, key=f"pnl_amv_{pick}")
        rmin = c2.number_input("Royalty minimum $ / month" + (" (all centres)" if n > 1 else ""),
                               min_value=0.0, value=float(col["Royalty Min"]), step=100.0,
                               key=f"pnl_rmin_{pick}")
    r1, r2 = st.columns([3, 1])
    lo, hi = r1.select_slider(f"Range ({unit})", options=list(range(10, 310, 10)),
                              value=(20, 200), key="pnl_model_range")
    step = r2.selectbox("Step", [10, 20, 50], index=1, key="pnl_model_step")

    per_member = amv * pnl.WEEKS_PER_MONTH
    fixed_lines = {k: col[k] for k in pnl.FIXED_ROWS}
    fixed = sum(fixed_lines.values())
    be_rev, be_mem = pnl.break_even_banded(per_member, fixed, rmin, rates, n)

    def model(per_centre):
        """The whole centre or group with this many members per centre."""
        total = per_centre * n
        c = pnl.model_column(total, per_member, pnl.rate_for(per_centre, rates), fixed_lines,
                             rmin)
        c.update({"Members": total, "AMV": amv})
        return c

    actual = {k: col.get(k) for k in AMOUNT_KEYS}
    actual.update({"Members": mem, "AMV": amv0})
    model_cols = [("Actual", label, actual)]
    if be_mem:
        model_cols.append(("Break-even", f"{be_mem / n:,.1f} {unit}", model(be_mem / n)))
    model_cols += [(f"{m} {unit}", "Model", model(m)) for m in range(lo, hi + 1, step)]

    notes = {"Total Revenue": f"Members × 4.2 wks × ${amv:,.2f}",
             "Tutor Wages + Super": "% for the members band (break-even settings)",
             "Royalties": f"Higher of ${rmin:,.0f} or 8% of revenue"}
    notes.update({k: "Dollar value fixed (actual)" for k in pnl.FIXED_ROWS})
    rows = [r for r in TABLE_ROWS[:13] if r[1] != "Months"]
    st.markdown(side_by_side_html(model_cols, notes, rows), unsafe_allow_html=True)

    band_now = pnl.membership_band(mem / n)
    tutor_note = (f" Tutor costs are {tutor0:.0f}% of revenue now; the break-even uses "
                  f"{rates.get(band_now, 0):.0f}% at its current size ({band_now} {unit}), "
                  "changing with size.")
    if be_mem:
        gap = mem - be_mem
        st.markdown(_md(f"**{pick}** needs **{be_mem:,.0f} members** "
                        f"({cur} {_money(be_rev)} a month) to break even and has "
                        f"**{mem:,.0f}**: " + (f"**{gap:,.0f} above** break-even." if gap >= 0
                                              else f"**{-gap:,.0f} short**.") + tutor_note))
    else:
        st.warning("At these settings tutor wages and royalties use up all the revenue, so more "
                   "members never reach break-even." + tutor_note)

    top = max(hi, int((be_mem or 0) / n * 1.2), int(mem / n * 1.2))
    xs = [x / 2 for x in range(0, top * 2 + 1)]
    ys = [model(x)["EBITDA"] for x in xs]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", name="EBITDA (model)",
                             line=dict(color="#1f9a9a", width=3),
                             hovertemplate="%{x:,.0f} " + unit + ": $%{y:,.0f}<extra></extra>"))
    fig.add_hline(y=0, line=dict(color="#5d6d73", width=1))
    fig.add_trace(go.Scatter(x=[mem / n], y=[col["EBITDA"]], mode="markers+text", name="Actual",
                             marker=dict(color="#e0562a", size=11), text=["Actual"],
                             textposition="top center",
                             hovertemplate="Actual: %{x:,.0f} " + unit +
                             ", $%{y:,.0f}<extra></extra>"))
    if be_mem:
        fig.add_vline(x=be_mem / n, line=dict(color="#2f6fb5", dash="dash"),
                      annotation_text=f"Break-even {be_mem / n:,.0f}", annotation_position="top")
    fig.update_layout(**helpers["std_layout"](f"EBITDA by members — {pick}",
                                               f"EBITDA per month ({cur})", height=420))
    fig.update_xaxes(title_text=unit.capitalize())
    helpers["show_chart"](fig)


# ── Outliers ──────────────────────────────────────────────────────────────────
STATUS_ICON = {"Very high": "▲▲", "High": "▲", "Typical": "", "Low": "▼", "Very low": "▼▼"}


def _md(text):
    """Markdown with dollar signs shown as dollars (not maths)."""
    return text.replace("$", "\\$")


def _fmt_pct(v):
    return "" if v is None or pd.isna(v) else f"{v:.1f}%"


def comparison_rows(values, peers, lines):
    cmp = pnl.compare_lines(values, peers, lines)
    if cmp.empty:
        return cmp
    out = []
    for _, r in cmp.iterrows():
        amv = r["Line"] == "Average Membership Value"
        bad = pnl.is_bad(r["Line"], r["Status"])
        good = r["Status"] not in ("", "Typical") and not bad
        flag = ("🔴 " if bad else "🟢 " if good else "") + \
            (f"{r['Status']} {STATUS_ICON.get(r['Status'], '')}".strip() if r["Status"] else "")
        out.append({
            "Line": r["Line"],
            "Actual": _money(r["Actual $"]) if not amv else f"${r['Actual $']:,.2f}/wk",
            "% of revenue": _fmt_pct(r["Actual %"]),
            "Typical": f"${r['Typical $']:,.2f}/wk" if amv else _fmt_pct(r["Typical %"]),
            "Normal range": (f"${r['Range low $']:,.2f}–${r['Range high $']:,.2f}" if amv else
                             f"{_fmt_pct(r['Range low %'])} – {_fmt_pct(r['Range high %'])}"),
            "Compared with": f"{int(r['Centres'])} centres",
            "Cost vs typical $/month": r["Gap $"],
            "Status": flag})
    return pd.DataFrame(out)


def _peers_for(col, bench, exclude):
    members = col.get("Members")
    n = int(col.get("Centres", 1) or 1)
    per_centre = members / n if members and not pd.isna(members) else None
    country = col.get("Country") if isinstance(col.get("Country"), str) else None
    return pnl.peer_set(bench, country, per_centre if n == 1 else None, exclude)


def outliers_section(cols, cent, bench, acct_bench, label, helpers):
    st.markdown('<div class="section-header">Outliers vs similar centres</div>',
                unsafe_allow_html=True)
    names = [h for h, _, _ in cols]
    a, b = st.columns([2, 2])
    pick = a.selectbox("Centre or group", names, key="pnl_out_pick")
    single = len(cols) and next(c for c in cols if c[0] == pick)[2].get("Centres") is None
    detail = b.toggle("Every account line (not just the grouped lines)", key="pnl_out_detail",
                      disabled=not single,
                      help="Available for a single centre (Show by: Centre).")
    heading, cur, col = next(c for c in cols if c[0] == pick)
    rev = col["Total Revenue"]
    if not rev or rev < pnl.LOW_REVENUE:
        st.info(f"{pick} has little or no revenue for {label}, so comparing its lines as a % "
                "of revenue wouldn't mean much.")
        return
    locs = list(col.get("Locations") or [])
    loc = locs[0] if single and locs else None
    # A centre is compared with other centres; a group with typical centres (its own included).
    peers, desc = _peers_for(col, bench, [loc] if loc else None)
    if detail and single and loc in acct_bench.index:
        acct = acct_bench.copy()
        for c in ["Total Revenue", "Members", "Country"]:
            acct[c] = bench[c].reindex(acct.index)
        peers = acct.loc[acct.index.intersection(peers.index)]
        values = acct.loc[loc].copy()
        lines = [c for c in acct_bench.columns if acct_bench[c].abs().sum() > 0]
    else:
        values, lines = col, pnl.BENCH_LINES
    if len(peers) < 2:
        st.info("Not enough other centres have reported yet to compare with.")
        return
    table = comparison_rows(values, peers, lines)
    n = int(col.get("Centres", 1) or 1)
    who = desc if n == 1 else f"a typical centre ({desc})"
    st.caption(f"{pick} compared with {who} for {label}: each line as a % of revenue against "
               "the typical (middle) value and the normal range (middle half of centres). "
               "'Cost vs typical' is what the difference is worth each month: positive costs "
               "money, negative saves it. Other centres' names and figures are never shown.")
    show = table.sort_values("Cost vs typical $/month", ascending=False, key=lambda s: s.fillna(0))
    st.dataframe(show, hide_index=True, use_container_width=True, column_config={
        "Cost vs typical $/month": st.column_config.NumberColumn(format="%,.0f")})

    # Plain-words summary and chart of the biggest gaps.
    lines_only = show[~show["Line"].isin(["EBITDA", "↳ vs centres that pay a manager"])]
    costly = lines_only[lines_only["Cost vs typical $/month"].fillna(0) > 0]
    ebitda = col.get("EBITDA") if not detail else None
    parts = []
    if ebitda is not None and not pd.isna(ebitda):
        parts.append(f"**{pick}** {'made' if ebitda >= 0 else 'lost'} **{_money(abs(ebitda))}** "
                     f"a month ({label}).")
    top = costly.head(3)
    if not top.empty:
        parts.append("Biggest costs above typical: " + "; ".join(
            f"**{r['Line']}** ({_money(r['Cost vs typical $/month'])} more)"
            for _, r in top.iterrows()) + ".")
        if ebitda is not None and not pd.isna(ebitda):
            parts.append(f"If every line above typical were typical, EBITDA would be about "
                         f"**{_money(ebitda + costly['Cost vs typical $/month'].sum())}**.")
    if parts:
        st.markdown(_md(" ".join(parts)))
    chart = lines_only.dropna(subset=["Cost vs typical $/month"])
    chart = chart[chart["Cost vs typical $/month"].abs() >= 1]
    if not chart.empty:
        chart = chart.sort_values("Cost vs typical $/month")
        fig = go.Figure(go.Bar(
            x=chart["Cost vs typical $/month"], y=chart["Line"], orientation="h",
            marker_color=["#c8322f" if v > 0 else "#2e8540"
                          for v in chart["Cost vs typical $/month"]],
            hovertemplate="%{y}: $%{x:,.0f} a month vs typical<extra></extra>"))
        fig.add_vline(x=0, line=dict(color="#5d6d73", width=1))
        fig.update_layout(**helpers["std_layout"](
            f"Cost vs typical — {pick} ({cur} per month; red costs money, green saves it)", "",
            height=max(280, 34 * len(chart) + 120)))
        fig.update_layout(showlegend=False)
        helpers["show_chart"](fig)


def network_outliers(cent, bench, label):
    st.markdown('<div class="section-header">Outlier list</div>', unsafe_allow_html=True)
    rows = []
    for loc, r in cent.iterrows():
        if not r["Total Revenue"] or r["Total Revenue"] < pnl.LOW_REVENUE:
            continue
        peers, desc = pnl.peer_set(bench, r.get("Country"), r.get("Members"), loc)
        if len(peers) < 2:
            continue
        cmp = pnl.compare_lines(r, peers, pnl.BENCH_LINES)
        for _, c in cmp.iterrows():
            if c["Status"] in ("", "Typical") or c["Line"].startswith("↳"):
                continue
            rows.append({"Location": loc.replace(PREFIX, ""), "Line": c["Line"],
                         "Bad": pnl.is_bad(c["Line"], c["Status"]), "Status": c["Status"],
                         "Actual %": c["Actual %"], "Typical %": c["Typical %"],
                         "Cost vs typical $/month": c["Gap $"], "Currency": r["Currency"],
                         "Compared with": desc})
    if not rows:
        st.info("No lines outside the normal range for the chosen centres.")
        return
    df = pd.DataFrame(rows)
    a, b, c = st.columns([2, 2, 1])
    which = a.selectbox("Show", ["Costing more than typical", "Doing better than typical",
                                 "Both"], key="pnl_out_which")
    by_line = b.multiselect("Lines", sorted(df["Line"].unique()), key="pnl_out_lines",
                            placeholder="All lines")
    strong = c.toggle("Very high/low only", key="pnl_out_strong")
    if which != "Both":
        df = df[df["Bad"] == (which == "Costing more than typical")]
    if by_line:
        df = df[df["Line"].isin(by_line)]
    if strong:
        df = df[df["Status"].str.startswith("Very")]
    df = df.sort_values("Cost vs typical $/month", ascending=(which == "Doing better than typical"),
                        key=lambda s: s.fillna(0))
    st.caption(f"Every line outside the normal range for each centre's comparison group "
               f"({label}). 'Very' = far outside the range. Centres with under "
               f"{pnl.LOW_REVENUE:,} revenue a month are left out.")
    st.dataframe(df.drop(columns="Bad"), hide_index=True, use_container_width=True,
                 column_config={"Actual %": st.column_config.NumberColumn(format="%.1f%%"),
                                "Typical %": st.column_config.NumberColumn(format="%.1f%%"),
                                "Cost vs typical $/month":
                                    st.column_config.NumberColumn(format="%,.0f")})


# ── Deep dive ─────────────────────────────────────────────────────────────────
UNITS = ["$", "% of revenue", "$ per member"]
TREND_LINES = pnl.ROW_ORDER + ["EBITDA", "Net Profit"]


def location_info(df_pnl, vl):
    """Country, currency, latest rate to AUD, state region and GPM per location."""
    last = df_pnl.sort_values("Period").groupby("Location").last()
    info = pd.DataFrame(index=last.index)
    info["Country"] = last["Country"].replace("", None)
    info["Currency"] = last["Currency"]
    info["FX"] = pd.to_numeric(last["FX rate to AUD"], errors="coerce")
    info.loc[info["Currency"] == "AUD", "FX"] = 1.0
    v = vl.set_index(LOC) if LOC in vl.columns else pd.DataFrame()
    for c in ["Region", "GPM", "Country"]:
        col = v[c].reindex(info.index) if c in v.columns else pd.Series(None, index=info.index)
        info[c] = info[c].fillna(col) if c in info else col
    return info


def group_options(info, view, state_of):
    """{option label: [locations]} for the chosen view."""
    if view == "Centre":
        return {loc.replace(PREFIX, ""): [loc] for loc in sorted(info.index)}
    if view == "Network":
        return {"Network": list(info.index)}
    key = info["Country"].fillna("Not set") if view == "Country" else \
        info["GPM"].replace("", None).fillna("Not set") if view == "GPM" else \
        info["Region"].map(state_of).fillna("Not set")
    return {f"{k} ({len(g)})": list(g.index) for k, g in info.groupby(key)}


def entity_months(lm, members, info, locs):
    """Month-by-month totals for these locations (AUD if they mix currencies).
    Returns (DataFrame indexed by Period, currency, locations left out)."""
    rows = lm[lm.index.get_level_values("Location").isin(locs)].copy()
    mixed = info.loc[info.index.intersection(locs), "Currency"].nunique() > 1
    left_out = []
    if mixed:
        fx = info["FX"].reindex(rows.index.get_level_values("Location")).values
        left_out = sorted(info.index[info.index.isin(locs) & info["FX"].isna()])
        rows = rows.mul(fx, axis=0).dropna(how="all")
    out = rows.groupby(level="Period").sum()
    out["Centres"] = rows.groupby(level="Period").size()
    if not members.empty:
        m = members[members.index.get_level_values(0).isin(locs)]
        # Members of centres that reported a P&L that month.
        reported = set(rows.index)
        m = m[[k in reported for k in m.index]]
        out["Members"] = m.groupby(level="Period").sum().reindex(out.index)
    else:
        out["Members"] = None
    cur = "AUD" if mixed else (info.loc[info.index.intersection(locs), "Currency"].iloc[0]
                               if len(info.index.intersection(locs)) else "")
    return out.sort_index(), cur, left_out


def _in_units(df, unit, lines):
    if unit == "% of revenue":
        rev = df["Total Revenue"].where(df["Total Revenue"] > 0)
        return df[lines].div(rev, axis=0) * 100
    if unit == "$ per member":
        mem = df["Members"].where(df["Members"] > 0)
        return df[lines].div(mem, axis=0)
    return df[lines]


def _fmt_unit(v, unit):
    if v is None or pd.isna(v):
        return "–"
    if unit == "% of revenue":
        return f"{v:.1f}%"
    if unit == "$ per member":
        s = f"${abs(v):,.2f}"
        return f"({s})" if v < 0 else s
    return _money(v)


def _card(label, value, delta=None, delta_text="", good=None):
    d = ""
    if delta is not None and not pd.isna(delta):
        arrow = "▲" if delta > 0 else "▼" if delta < 0 else "●"
        colour = "#5d6d73" if good is None or delta == 0 else ("#2e8540" if good else "#c8322f")
        d = (f'<div class="metric-delta" style="color:{colour}">{arrow} {delta_text} '
             '<span>vs previous month</span></div>')
    return (f'<div class="metric-card"><div class="metric-label">{html.escape(label)}</div>'
            f'<div class="metric-value">{html.escape(value)}</div>{d}</div>')


def peer_typical_by_month(bench_lm, bench_members, binfo, line, country, per_centre, exclude):
    """Typical (median) % of revenue for this line among similar centres, each month."""
    out = {}
    for period, g in bench_lm.groupby(level="Period"):
        t = g.droplevel("Period").copy()
        t["Country"] = binfo["Country"].reindex(t.index)
        t["Members"] = bench_members.xs(period, level="Period").reindex(t.index) \
            if not bench_members.empty and period in bench_members.index.get_level_values(
                "Period") else None
        peers, _ = pnl.peer_set(t, country, per_centre, exclude)
        peers = peers[peers["Total Revenue"] > 0]
        if len(peers) >= 2:
            out[period] = (peers[line] / peers["Total Revenue"] * 100).median()
    return pd.Series(out, dtype=float)


def report_deep_dive(df_pnl, df_wm, vl, helpers, bench=None):
    """One centre or group month by month: headline numbers, where the revenue goes,
    every line over time (with the typical value for similar centres), every account."""
    st.markdown('<div class="report-title">P&L Deep Dive</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">One centre or group month by month: headline '
                'numbers, where the money goes, and each revenue and expense line over time '
                'against similar centres.</div>', unsafe_allow_html=True)
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return
    info = location_info(df_pnl, vl)
    lm = pnl.location_months(df_pnl)
    members = members_by_month(df_wm)
    periods = sorted(df_pnl["Period"].unique())

    panel = st.container(border=True)
    with panel:
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b, c, d = st.columns([1, 2, 1, 1])
        view = a.selectbox("Show by", VIEWS, key="dd_view")
        opts = group_options(info, view, helpers["state_of"])
        pick = b.selectbox("Centre or group", list(opts), key=f"dd_pick_{view}")
        month = c.selectbox("Month", periods[::-1], format_func=pnl.period_name, key="dd_month")
        unit = d.selectbox("Show as", UNITS, key="dd_unit")
    locs = opts[pick]
    data, cur, left_out = entity_months(lm, members, info, locs)
    data = data[data.index <= month]
    if month not in data.index:
        st.warning(f"{pick} has no P&L for {pnl.period_name(month)}.")
        return
    if left_out:
        st.warning("Left out of AUD totals (no exchange rate): " +
                   ", ".join(x.replace(PREFIX, "") for x in left_out))
    now = data.loc[month]
    prev = data.iloc[-2] if len(data) >= 2 else None
    n_now = int(now["Centres"])

    # ── Headline numbers ──
    st.markdown(f'<div class="section-header">{html.escape(pick)} — {pnl.period_name(month)} '
                f'({html.escape(cur)})</div>', unsafe_allow_html=True)
    rev = now["Total Revenue"]
    gm = rev - now["Tutor Wages + Super"]
    mem = now.get("Members")
    rpm = rev / mem / pnl.WEEKS_PER_MONTH if mem and not pd.isna(mem) and mem > 0 else None

    def delta(key, fn=None):
        if prev is None:
            return None
        f = fn or (lambda r: r[key])
        try:
            return f(now) - f(prev)
        except (TypeError, ZeroDivisionError):
            return None

    gm_pct = lambda r: (r["Total Revenue"] - r["Tutor Wages + Super"]) / r["Total Revenue"] * 100 \
        if r["Total Revenue"] > 0 else None
    eb_pct = lambda r: r["EBITDA"] / r["Total Revenue"] * 100 if r["Total Revenue"] > 0 else None
    cards = [
        ("Revenue", _money(rev), delta("Total Revenue"), True, _money),
        ("Gross margin (after tutor wages)", f"{gm_pct(now):.1f}%" if rev > 0 else "–",
         delta(None, gm_pct), True, lambda v: f"{abs(v):.1f} pts"),
        ("EBITDA", _money(now["EBITDA"]), delta("EBITDA"), True, _money),
        ("EBITDA % of revenue", f"{eb_pct(now):.1f}%" if rev > 0 else "–",
         delta(None, eb_pct), True, lambda v: f"{abs(v):.1f} pts"),
        ("Members", _count(mem), delta("Members"), True, lambda v: f"{abs(v):,.0f}"),
        ("Revenue per member / week", f"${rpm:,.2f}" if rpm else "–",
         delta(None, lambda r: r["Total Revenue"] / r["Members"] / pnl.WEEKS_PER_MONTH
               if r.get("Members") and r["Members"] > 0 else None), True,
         lambda v: f"${abs(v):,.2f}"),
    ]
    html_cards = []
    for label, value, dlt, up_good, fmt in cards:
        ok = dlt is not None and not pd.isna(dlt)
        html_cards.append(_card(label, value, dlt if ok else None,
                                fmt(abs(dlt)) if ok else "", (dlt > 0) == up_good if ok else None))
    cols = st.columns(len(html_cards))
    for col, h in zip(cols, html_cards):
        col.markdown(h, unsafe_allow_html=True)
    if n_now > 1:
        st.caption(f"{n_now} centres reported for {pnl.period_name(month)}.")

    # ── Where the revenue goes ──
    st.markdown('<div class="section-header">Where the revenue goes</div>',
                unsafe_allow_html=True)
    costs = [k for k in pnl.ROW_ORDER if k != "Total Revenue"]
    fig = go.Figure(go.Waterfall(
        x=["Revenue"] + costs + ["EBITDA"],
        y=[rev] + [-now[k] for k in costs] + [0],
        measure=["absolute"] + ["relative"] * len(costs) + ["total"],
        text=[_money(rev)] + [_money(-now[k]) for k in costs] + [_money(now["EBITDA"])],
        textposition="outside",
        increasing=dict(marker=dict(color="#2e8540")),
        decreasing=dict(marker=dict(color="#c8322f")),
        totals=dict(marker=dict(color="#1f9a9a" if now["EBITDA"] >= 0 else "#c8322f")),
        connector=dict(line=dict(color="#c9d3d6")),
        hovertemplate="%{x}: %{text}<extra></extra>"))
    fig.update_layout(**helpers["std_layout"](f"From revenue to EBITDA — {pick}, "
                                               f"{pnl.period_name(month)}",
                                               f"{cur} per month", height=440))
    fig.update_layout(showlegend=False)
    helpers["show_chart"](fig)

    # ── Every line by month ──
    recent = data.tail(12)
    st.markdown(f'<div class="section-header">Every line by month ({unit})</div>',
                unsafe_allow_html=True)
    table = _in_units(recent, unit, TREND_LINES).T
    table.columns = [pnl.period_name(p) for p in table.columns]
    shown = table.apply(lambda col: col.map(lambda v: _fmt_unit(v, unit)))
    if prev is not None:
        chg = _in_units(data.tail(2), unit, TREND_LINES).T
        diff = chg.iloc[:, -1] - chg.iloc[:, -2]
        shown["Change vs previous month"] = [
            "–" if pd.isna(v) else (f"{v:+.1f} pts" if unit == "% of revenue" else
                                   ("+" if v >= 0 else "−") + _fmt_unit(abs(v), unit))
            for v in diff]
    ytd = data[(data.index >= f"{month[:4]}-01")]
    if len(ytd) > 1:
        shown[f"Year to date ({len(ytd)} months, avg/month)"] = [
            _fmt_unit(v, unit) for v in _in_units(
                ytd.mean().to_frame().T.assign(Members=ytd["Members"].mean()), unit,
                TREND_LINES).iloc[0]]
    members_row = pd.DataFrame([[_count(v) for v in recent["Members"]] +
                                [""] * (shown.shape[1] - len(recent))],
                               index=["Members"], columns=shown.columns)
    st.dataframe(pd.concat([members_row, shown]), use_container_width=True)
    if len(data) < 2:
        st.caption("Only one month uploaded so far: trends and month-on-month changes appear "
                   "as more months are added.")

    # ── A line over time against similar centres ──
    st.markdown('<div class="section-header">A line over time, against similar centres</div>',
                unsafe_allow_html=True)
    line = st.selectbox("Line", TREND_LINES[1:], key="dd_line")
    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    b_lm = pnl.location_months(b_pnl)
    b_info = location_info(b_pnl, vl)
    countries = info.loc[info.index.intersection(locs), "Country"].dropna().unique()
    country = countries[0] if len(countries) == 1 else None
    per_centre = mem / n_now if mem and not pd.isna(mem) and n_now else None
    typical = peer_typical_by_month(b_lm, members_by_month(b_wm), b_info, line, country,
                                    per_centre if len(locs) == 1 else None,
                                    locs if len(locs) == 1 else None)
    own = _in_units(data, "% of revenue", [line])[line]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=[pnl.period_name(p) for p in own.index], y=own.values,
                             mode="lines+markers", name=pick, line=dict(color="#1f9a9a", width=3),
                             hovertemplate="%{x}: %{y:.1f}%<extra>" + html.escape(pick) +
                             "</extra>"))
    t = typical.reindex(own.index)
    if t.notna().any():
        fig.add_trace(go.Scatter(x=[pnl.period_name(p) for p in t.index], y=t.values,
                                 mode="lines+markers", name="Typical for similar centres",
                                 line=dict(color="#5d6d73", width=2, dash="dash"),
                                 hovertemplate="%{x}: %{y:.1f}%<extra>Typical</extra>"))
    fig.update_layout(**helpers["std_layout"](f"{line} as % of revenue — {pick}",
                                               "% of revenue", height=380))
    helpers["show_chart"](fig)
    st.caption("Typical = the middle value for centres in the same country and members band "
               "each month (other centres are never named). Shown as % of revenue so centres "
               "of different sizes and currencies compare fairly.")

    # ── Every account ──
    with st.expander("Every account line by month"):
        d = df_pnl[df_pnl["Location"].isin(locs) & (df_pnl["Period"] <= month)].copy()
        d["Amount"] = d["Amount"].astype(float)
        if len(locs) > 1 and info.loc[info.index.intersection(locs), "Currency"].nunique() > 1:
            d["Amount"] = d["Amount"] * d["Location"].map(info["FX"])
        acc = d.pivot_table(index=["P&L section", "Account"], columns="Period", values="Amount",
                            aggfunc="sum", fill_value=0.0)
        acc = acc[sorted(acc.columns)[-12:]]
        rev_m = data["Total Revenue"].reindex(acc.columns)
        if unit == "% of revenue":
            acc = acc.div(rev_m.where(rev_m > 0), axis=1) * 100
        elif unit == "$ per member":
            acc = acc.div(data["Members"].reindex(acc.columns).where(lambda m: m > 0), axis=1)
        acc.columns = [pnl.period_name(p) for p in acc.columns]
        st.dataframe(acc.apply(lambda col: col.map(lambda v: _fmt_unit(v, unit))),
                     use_container_width=True)
        st.caption("Unrelated Expenses (479) are a memo line and aren't in EBITDA or Net Profit.")
