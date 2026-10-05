"""P&L reports: summary in the centres side-by-side layout, and the break-even model."""
import html
import io

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import pnl

LOC = "Success Tutoring - Business name"
PREFIX = "Success Tutoring - "
VIEWS = ["Centre", "State", "Country", "Network"]
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
            cols.append((loc.replace(PREFIX, ""), r["Currency"], add_break_even(r.copy(), rates)))
        return cols, []
    if view == "Network":
        groups = {"Network": cent}
    else:
        key = cent["Country"].fillna("Not set") if view == "Country" else \
            cent["Region"].map(state_of).fillna("Not set")
        groups = {k: g for k, g in cent.groupby(key)}
    left_out = []
    for name, g in sorted(groups.items(), key=lambda kv: -kv[1]["Total Revenue"].sum()):
        mixed = g["Currency"].nunique() > 1
        if mixed:
            left_out += g[g["FX"].isna()].index.tolist()
        col = add_break_even(combine(g, mixed), rates)
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
        st.markdown(f"**{pick}** needs **{be_mem:,.0f} members** "
                    f"({cur} {_money(be_rev)} a month) to break even and has **{mem:,.0f}**: "
                    + (f"**{gap:,.0f} above** break-even." if gap >= 0 else
                       f"**{-gap:,.0f} short**.") + tutor_note)
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
