"""P&L reports: summary in the centres side-by-side layout, and the break-even model."""
import base64
import html
import io
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import pnl
import pnl_forecast as pf
import pnl_model as pm
import pnl_share

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
    ("Gross Margin", pnl.GROSS_MARGIN, "subtotal"),
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
AMOUNT_KEYS = pnl.ROW_ORDER + [pnl.GROSS_MARGIN, "EBITDA", "Interest", "Net Profit"]
# Members, AMV and the condensed P&L down to EBITDA (no months reported, interest or net profit)
MODEL_ROWS = [r for r in TABLE_ROWS[:TABLE_ROWS.index(("EBITDA", "EBITDA", "result")) + 1]
              if r[1] != "Months"]
GM_NOTE = "Total Revenue − Tutor Wages + Super"

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
.pnl tr.total td,.pnl tr.result td,.pnl tr.subtotal td{font-weight:700}
.pnl tr.subtotal td,.pnl tr.subtotal td.lbl{background:#f2f8f8;border-top:1px solid #b9d3d6}
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


def _money_short(v):
    """-$11,408 rather than ($11,408), to fit a tile."""
    if v is None or pd.isna(v):
        return "–"
    return f"-${abs(v):,.0f}" if v < -0.5 else f"${v:,.0f}"


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
    s["Months short"] = int((r["Months"] < r["Months"].max()).sum())
    return s


def add_break_even(col, rates, samples=None):
    """Break-even with tutor % from the members band. A group counts as its number of
    centres, each with the group's average members and costs. samples: {band: centres
    behind that band's %}, to flag break-evens resting on a small sample."""
    rev, mem = col["Total Revenue"], col.get("Members")
    mem = 0 if mem is None or pd.isna(mem) else mem
    n = int(col.get("Centres", 1) or 1)
    fixed = sum(col[k] for k in pnl.FIXED_ROWS)
    per_member = rev / mem if mem else 0
    be_rev, be_mem = pnl.break_even_banded(per_member, fixed, col["Royalty Min"], rates, n)
    col["BE revenue"], col["BE members"] = be_rev, be_mem
    col["BE sample"] = (samples or {}).get(pnl.membership_band(be_mem / n)) \
        if be_mem else None
    col["BE gap"] = (mem - be_mem) if be_mem is not None and mem else None
    col["AMV"] = rev / mem / pnl.WEEKS_PER_MONTH if mem else None
    return col


def build_columns(cent, view, state_of, rates, samples=None, n_periods=1):
    """[(heading, currency, column)] for the chosen view."""
    cols = []
    if view == "Centre":
        for loc, r in cent.sort_index(key=lambda i: i.str.replace(PREFIX, "").str.lower()) \
                .iterrows():
            col = add_break_even(r.copy(), rates, samples)
            col["Locations"] = (loc,)
            col["Months of"] = n_periods
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
        col = add_break_even(combine(g, mixed), rates, samples)
        col["Months of"] = n_periods
        col["Months short"] = int((g["Months"] < n_periods).sum())
        countries = g["Country"].dropna().unique()
        col["Country"] = countries[0] if len(countries) == 1 else None
        col["Locations"] = tuple(g.index)
        cols.append((f"{name} ({int(col['Centres'])})", "AUD" if mixed else g["Currency"].iloc[0],
                     col))
    return cols, left_out


DISCLAIMER = ("Break-even figures are estimates from network benchmarks and this centre's "
              "recent figures. A true break-even should be worked out case by case with the "
              "centre's own costs, lease and staffing plans.")


# ── Rendering ─────────────────────────────────────────────────────────────────
def _gross_margin(c):
    """A column's Gross Margin (worked out if it isn't there; None if it can't be)."""
    v = c.get(pnl.GROSS_MARGIN)
    if v is not None and not pd.isna(v):
        return v
    rev, tutor = c.get("Total Revenue"), c.get(pnl.TUTOR)
    if rev is None or tutor is None or pd.isna(rev) or pd.isna(tutor):
        return None
    return rev - tutor


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
            v = _gross_margin(c) if key == pnl.GROSS_MARGIN else c.get(key)
            rev = c.get("Total Revenue")
            if key in AMOUNT_KEYS:
                neg = " neg" if v is not None and not pd.isna(v) and v < -0.5 else ""
                cells.append(f'<td class="{neg}">{_money(v)}</td><td class="pct">{_pct(v, rev)}</td>')
            elif key == "AMV" or style == "money":
                cells.append(f"<td>{_money(v)}</td><td></td>")
            elif key == "Months":
                of = int(c.get("Months of") or 0)
                short = int(c.get("Months short") or 0)
                if of and v is not None and not pd.isna(v):
                    if c.get("Centres"):
                        txt = f"{int(v)} of {of}" + (f" · {short} missed a month ⚠" if short else "")
                    else:
                        txt = f"{int(v)} of {of}" + (" ⚠" if v < of else "")
                else:
                    txt = _count(v)
                cells.append(f"<td>{html.escape(txt)}</td><td></td>")
            elif key == "BE members":
                flag = " ⚠" if pnl.small_sample(c.get("BE sample")) else ""
                cells.append(f"<td>{_count(v)}{flag}</td><td></td>")
            elif key == "BE gap":
                neg = " neg" if v is not None and not pd.isna(v) and v < 0 else ""
                cells.append(f'<td class="{neg}">{_count(v, True)}</td><td></td>')
            else:
                cells.append(f"<td>{_count(v)}</td><td></td>")
        note = f'<td class="note">{html.escape(notes.get(key, ""))}</td>' if notes else ""
        body.append(f'<tr class="{style}"><td class="lbl">{html.escape(label)}</td>'
                    + "".join(cells) + note + "</tr>")
    return CSS + f'<div class="pnl-wrap"><table class="pnl">{head}{"".join(body)}</table></div>'


def tutor_rate_settings(bench_cent, key):
    """Editable tutor % by members band from established centres (12 months or older).
    Returns ({band: %}, {band: centres})."""
    established, _ = pnl.split_by_age(bench_cent)
    bench = pnl.tutor_benchmarks(established).rename_axis("Members band").reset_index()
    st.markdown('<div class="section-header">Benchmark settings: Tutor Wages + Super by '
                'members band</div>', unsafe_allow_html=True)
    with st.container():
        st.caption("Typical tutor wages + super as a % of revenue for established centres "
                   "(12 months or older) of each size, used for break-even. Change 'Use %' to "
                   f"test higher or lower costs. ⚠ = fewer than {pnl.SMALL_SAMPLE} centres "
                   "behind the %.")
        edited = st.data_editor(
            bench.drop(columns="Based on"), key=key, hide_index=True,
            use_container_width=True,
            disabled=["Members band", "Centres", "Benchmark %", "Source", "Sample"],
            column_config={
                "Benchmark %": st.column_config.NumberColumn(format="%.1f%%"),
                "Use %": st.column_config.NumberColumn(min_value=0.0, max_value=100.0,
                                                       step=0.5, format="%.1f%%")})
    rates = dict(zip(edited["Members band"], edited["Use %"].fillna(0).astype(float)))
    return rates, dict(zip(bench["Members band"], bench["Based on"]))


# ── Outliers ──────────────────────────────────────────────────────────────────
STATUS_TEXT = {"Very high": "Well above typical", "High": "Above typical",
               "Typical": "Typical", "Low": "Below typical", "Very low": "Well below typical"}


def _md(text):
    """Markdown with dollar signs shown as dollars (not maths)."""
    return text.replace("$", "\\$")


def _fmt_pct(v):
    return "" if v is None or pd.isna(v) else f"{v:.1f}%"


def bold_gross_margin(df, col):
    """The table styled with its Gross Margin row(s) in bold (for st.dataframe)."""
    return df.style.apply(lambda r: ["font-weight:700; background-color:#f2f8f8"] * len(r)
                          if r[col] == pnl.GROSS_MARGIN else [""] * len(r), axis=1)


def gross_margin_first(df, col="Line"):
    """Gross Margin pinned to the top, the other rows in their order."""
    gm = df[col] == pnl.GROSS_MARGIN
    return pd.concat([df[gm], df[~gm]])


def comparison_rows(values, peers, lines):
    cmp = pnl.compare_lines(values, peers, lines)
    if cmp.empty:
        return cmp
    out = []
    for _, r in cmp.iterrows():
        amv = r["Line"] == "Average Membership Value"
        bad = pnl.is_bad(r["Line"], r["Status"])
        good = r["Status"] not in ("", "Typical") and not bad
        flag = ("🔴 " if bad else "🟢 " if good else "") + STATUS_TEXT.get(r["Status"], "")
        out.append({
            "Line": r["Line"],
            "Actual": _money(r["Actual $"]) if not amv else f"${r['Actual $']:,.2f}/wk",
            "% of revenue": _fmt_pct(r["Actual %"]),
            "Typical": f"${r['Typical $']:,.2f}/wk" if amv else _fmt_pct(r["Typical %"]),
            "Normal range": (f"${r['Range low $']:,.2f}–${r['Range high $']:,.2f}" if amv else
                             f"{_fmt_pct(r['Range low %'])} – {_fmt_pct(r['Range high %'])}"),
            "Compared with": f"{int(r['Centres'])} centres" +
                             (" ⚠ small sample" if pnl.small_sample(r["Centres"]) else ""),
            "Cost vs typical $/month": r["Gap $"],
            "Status": flag})
    return pd.DataFrame(out)


def _describe(stage, desc):
    """'established' + 'all Australian centres' -> 'all established Australian centres'."""
    if desc.startswith("all "):
        return f"all {stage} {desc[4:]}"
    if desc.startswith("Centres"):
        return f"{stage} c{desc[1:]}"
    return f"{stage} {desc}"


def peers_by_age(bench, age, country, members, exclude):
    """Similar centres of the same stage: a centre under 12 months old is compared with
    other developing centres, everyone else with established centres."""
    established, developing = pnl.split_by_age(bench)
    if pnl.is_developing(age):
        peers, desc = pnl.peer_set(developing, country, members, exclude)
        return peers, _describe("developing (under 12 months)", desc)
    peers, desc = pnl.peer_set(established, country, members, exclude)
    return peers, _describe("established", desc)


def _peers_for(col, bench, exclude):
    members = col.get("Members")
    n = int(col.get("Centres", 1) or 1)
    per_centre = members / n if members and not pd.isna(members) else None
    country = col.get("Country") if isinstance(col.get("Country"), str) else None
    age = col.get("Age (Months)") if n == 1 else None
    return peers_by_age(bench, age, country, per_centre if n == 1 else None, exclude)


# ── Deep dive ─────────────────────────────────────────────────────────────────
UNITS = ["$", "% of revenue", "$ per member"]
DD_SPANS = ["Latest month", "Last 3 months", "Last 12 months", "Year to date", "All months"]
DEFAULT_SPAN = "Latest month"
GROUP_VIEWS = ["State", "Country", "GPM", "Network"]


def months_shown(index, span, month):
    """Months up to `month` for the Deep Dive tables."""
    ps = [p for p in index if p <= month]
    if span == "Latest month":
        return ps[-1:]
    if span == "Last 3 months":
        return ps[-3:]
    if span == "Last 12 months":
        return ps[-12:]
    if span == "Year to date":
        return [p for p in ps if p >= f"{month[:4]}-01"]
    return ps
TREND_LINES = [pnl.ROW_ORDER[0], pnl.TUTOR, pnl.GROSS_MARGIN] + \
    [r for r in pnl.ROW_ORDER[1:] if r != pnl.TUTOR] + ["EBITDA", "Net Profit"]


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
    df = pnl.with_gross_margin(df)
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


def average_months(data, months):
    """Average per month over these months (amounts and members)."""
    return data.loc[[m for m in months if m in data.index]].mean(numeric_only=True)


def _card(label, value, delta=None, delta_text="", good=None, vs="vs previous month"):
    d = ""
    if delta is not None and not pd.isna(delta):
        arrow = "▲" if delta > 0 else "▼" if delta < 0 else "●"
        colour = "#5d6d73" if good is None or delta == 0 else ("#2e8540" if good else "#c8322f")
        d = (f'<div class="metric-delta" style="color:{colour}">{arrow} {delta_text} '
             f'<span>{html.escape(vs)}</span></div>')
    return (f'<div class="metric-card"><div class="metric-label">{html.escape(label)}</div>'
            f'<div class="metric-value">{html.escape(value)}</div>{d}</div>')




def _section(title, caption=None):
    st.markdown(f'<div class="section-header">{html.escape(title)}</div>',
                unsafe_allow_html=True)
    if caption:
        st.caption(caption)


def _period_label(periods):
    return pnl.period_name(periods[0]) if len(periods) == 1 else \
        f"{pnl.period_name(periods[0])} – {pnl.period_name(periods[-1])} ({len(periods)} months)"


HOW_GROUP = f"""
- **Amounts** are the average per month over the months each centre reported, in each
  centre's own currency. **Months reported** shows ⚠ when a centre missed a month.
- **Members** = average active members (Weekly Membership). **Average Membership Value** =
  revenue ÷ members ÷ 4.2 weeks.
- **Break-even**: tutor wages at the benchmark % for established centres of that size (the
  settings above); rent, manager wages, marketing and OPEX stay at the centre's own amounts;
  royalties are the higher of the centre's Royalty Min (Vlookup, blank = 2,200) or 8% of
  revenue. ⚠ = the tutor % used comes from fewer than {pnl.SMALL_SAMPLE} centres.
- **Gross Margin** = Total Revenue − Tutor Wages + Super (accounts 472 and 474), with its % of revenue. Higher is better. It is a subtotal, so it is never added into costs or EBITDA.
- **Outliers**: each centre is compared with similar centres across the whole network (same
  country and members band; centres under 12 months old with other developing centres).
  Other centres are counted but never named. Centres with under {pnl.LOW_REVENUE:,} revenue
  a month are left out. ⚠ = fewer than {pnl.SMALL_SAMPLE} centres to compare with.
"""


def report_group(df_pnl, df_wm, vl, helpers, bench=None):
    """P&L – Group Level: every centre in a state, country, GPM's group or the network,
    side by side, with break-even and the group's outliers."""
    st.markdown('<div class="report-title">P&L – Group Level</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">How a group of centres is doing, centre by centre: '
                'P&L, break-even and the lines out of step with similar centres.</div>',
                unsafe_allow_html=True)
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return
    members = members_by_month(df_wm)
    all_periods = sorted(df_pnl["Period"].unique())
    panel = st.container(border=True)
    with panel:
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b, c, d, e = st.columns(5)
        mode = a.selectbox("Period", PERIODS, key="pnl_mode")
        month = b.selectbox("Month" if mode == "Single month" else "Up to", all_periods[::-1],
                            format_func=pnl.period_name, key="pnl_month")
        view = c.selectbox("Group by", GROUP_VIEWS, key="pnl_group_view")
        bands = e.multiselect("Members", pnl.membership_bands(), key="pnl_band",
                              placeholder="All sizes")
    periods = select_periods(all_periods, mode, month)
    cent = centre_table(df_pnl, members, vl, periods, helpers["age_group"])
    if cent.empty:
        st.warning("No P&L figures for the chosen months.")
        return
    cent = cent.reset_index().rename(columns={"Location": LOC})
    with panel:
        cent = helpers["report_filters"](cent, key_prefix="pnl", show_date=False,
                                         show_location=False, show_status=False,
                                         panel=st.container(), show_label=False)
        f, _ = st.columns(2)
        ages = f.multiselect("Age", helpers["age_order"], key="pnl_age", placeholder="All ages")
    if bands:
        cent = cent[cent["Band"].isin(bands)]
    if ages:
        cent = cent[cent["Age group"].isin(ages)]
    cent = cent.set_index(LOC)
    if cent.empty:
        st.warning("No locations match the filters.")
        return
    if view == "Network":
        groups = {f"Network ({len(cent)})": cent}
    else:
        key = cent["Country"].fillna("Not set") if view == "Country" else \
            cent["GPM"].replace("", None).fillna("Not set") if view == "GPM" else \
            cent["Region"].map(helpers["state_of"]).fillna("Not set")
        groups = {f"{k} ({len(g)})": g for k, g in cent.groupby(key)}
    pick = d.selectbox("Group", sorted(groups), key=f"pnl_group_{view}")
    group = groups[pick]
    label = _period_label(periods)

    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    bench_cent = centre_table(b_pnl, members_by_month(b_wm), vl, periods, helpers["age_group"])
    # The settings sit at the end of the page but feed the table, so lay out the table and
    # outliers first (in containers) and fill them once the settings are read.
    table_slot, outliers_slot = st.container(), st.container()
    rates, samples = tutor_rate_settings(bench_cent, key=f"pnl_rates_{periods[0]}_{periods[-1]}")
    # Centre Level uses these rates too, while they're for the same month.
    st.session_state["pnl_bench_rates"] = {"periods": tuple(periods), "rates": rates,
                                           "samples": samples}
    cols, _ = build_columns(group, "Centre", helpers["state_of"], rates, samples, len(periods))
    # Group total at the end (in AUD when the group mixes currencies).
    if len(group) > 1:
        total, left_out = build_columns(group, "Network", helpers["state_of"], rates, samples,
                                        len(periods))
        name = pick.rsplit(" (", 1)[0]
        cols.append((f"Total — {name}", total[0][1], total[0][2]))
    else:
        left_out = []

    with table_slot:
        _section(f"P&L by centre — {pick}, {label}",
                 "Each centre in its own currency, A–Z, then the group total (AUD if the group "
                 "mixes currencies). ⚠ = missed a month, or a break-even based on a small "
                 "benchmark sample. Break-even uses the benchmark settings at the end of the "
                 "page.")
        st.markdown(side_by_side_html(cols), unsafe_allow_html=True)
        if left_out:
            st.warning("Left out of the AUD total (no exchange rate): " +
                       ", ".join(x.replace(PREFIX, "") for x in left_out))
    with outliers_slot:
        group_outliers(group, bench_cent, label)
    sessions_section(group, bench_cent, helpers.get("revenue"), periods, label, helpers,
                     young_centres(vl))
    with st.expander("How this is worked out"):
        st.markdown(HOW_GROUP + "\n- " + DISCLAIMER)


def group_outliers(cent, bench, label):
    _section("Outliers in this group",
             f"Lines outside the normal range for similar centres ({label}).")
    rows = []
    for loc, r in cent.iterrows():
        if not r["Total Revenue"] or r["Total Revenue"] < pnl.LOW_REVENUE:
            continue
        peers, desc = peers_by_age(bench, r.get("Age (Months)"), r.get("Country"),
                                   r.get("Members"), loc)
        if len(peers) < 2:
            continue
        cmp = pnl.compare_lines(r, peers, pnl.BENCH_LINES)
        for _, c in cmp.iterrows():
            if c["Status"] in ("", "Typical") or c["Line"].startswith("↳"):
                continue
            rows.append({"Location": loc.replace(PREFIX, ""), "Line": c["Line"],
                         "Bad": pnl.is_bad(c["Line"], c["Status"]),
                         "Status": STATUS_TEXT.get(c["Status"], c["Status"]),
                         "Actual %": c["Actual %"], "Typical %": c["Typical %"],
                         "Cost vs typical $/month": c["Gap $"], "Currency": r["Currency"],
                         "Compared with": f"{desc} ({int(c['Centres'])})" +
                         (" ⚠ small sample" if pnl.small_sample(c["Centres"]) else "")})
    if not rows:
        st.info("No lines outside the normal range for these centres.")
        return
    df = pd.DataFrame(rows)
    a, b, c = st.columns([2, 2, 1])
    which = a.selectbox("Show", ["Costing more than typical", "Doing better than typical",
                                 "Both"], key="pnl_out_which")
    by_line = b.multiselect("Lines", sorted(df["Line"].unique()), key="pnl_out_lines",
                            placeholder="All lines")
    strong = c.toggle("Well above/below only", key="pnl_out_strong")
    if which != "Both":
        df = df[df["Bad"] == (which == "Costing more than typical")]
    if by_line:
        df = df[df["Line"].isin(by_line)]
    if strong:
        df = df[df["Status"].str.startswith("Well")]
    df = df.assign(_gap=df["Cost vs typical $/month"].fillna(0).abs()) \
        .sort_values(["Location", "_gap"], ascending=[True, False]).drop(columns="_gap")
    st.dataframe(df.drop(columns="Bad"), hide_index=True, use_container_width=True,
                 column_config={"Actual %": st.column_config.NumberColumn(format="%.1f%%"),
                                "Typical %": st.column_config.NumberColumn(format="%.1f%%"),
                                "Cost vs typical $/month":
                                    st.column_config.NumberColumn(format="%,.0f")})


# ── Centre Level ──────────────────────────────────────────────────────────────
TILE_CSS = """<style>
.otile{background:#fff;border:1px solid #dde5e6;border-left:4px solid var(--c);border-radius:8px;
 padding:10px 12px;height:100%;margin-bottom:10px}
.otile .t{font-size:12px;color:#5d6d73;font-weight:600}
.otile .v{font-size:22px;font-weight:700;color:var(--c);margin:2px 0}
.otile .s{font-size:12px;color:#5d6d73}
</style>"""
TILES_PER_ROW = 4


def _outlier_tiles(rows, bad, cur):
    """A tile for every account, 4 to a row, adding rows as needed."""
    if rows.empty:
        st.caption("None for this period.")
        return
    colour = "#c8322f" if bad else "#2e8540"
    st.markdown(TILE_CSS, unsafe_allow_html=True)
    items = list(rows.iterrows())
    for start in range(0, len(items), TILES_PER_ROW):
        tiles = st.columns(TILES_PER_ROW)
        for col, (_, r) in zip(tiles, items[start:start + TILES_PER_ROW]):
            gap = abs(r["Gap $"])
            flag = " ⚠" if pnl.small_sample(r["Centres"]) else ""
            col.markdown(
                f'<div class="otile" style="--c:{colour}"><div class="t">'
                f'{html.escape(r["Line"])}</div><div class="v">{"+" if bad else "−"}'
                f'{_money(gap)}<span style="font-size:12px;font-weight:500"> '
                f'{html.escape(cur)}/month</span></div><div class="s">{r["Actual %"]:.1f}% of '
                f'revenue vs typical {r["Typical %"]:.1f}%{flag}</div></div>',
                unsafe_allow_html=True)


def _waterfall(now, pick, month, cur, helpers):
    """Revenue (green) down through each cost (black) to EBITDA (black, red if a loss)."""
    rev = now["Total Revenue"]
    costs = [k for k in pnl.ROW_ORDER if k != "Total Revenue"]
    x, base, height, colours, text = ["Revenue"], [0], [rev], ["#2e8540"], [_money(rev)]
    running = rev
    for k in costs:
        v = now[k]
        x.append(k)
        base.append(running - v if v >= 0 else running)
        height.append(abs(v))
        colours.append("#1c2a30")
        text.append(_money(-v))
        running -= v
    e = now["EBITDA"]
    x.append("EBITDA")
    base.append(min(0, e))
    height.append(abs(e))
    colours.append("#c8322f" if e < 0 else "#1c2a30")
    text.append(_money(e))
    x = [l.replace(" + ", "<br>+ ").replace(" - ", "<br>") for l in x]
    fig = go.Figure(go.Bar(x=x, y=height, base=base, marker_color=colours, text=text,
                           textposition="outside", cliponaxis=False,
                           hovertemplate="%{x}: %{text}<extra></extra>"))
    fig.add_hline(y=0, line=dict(color="#5d6d73", width=1))
    fig.update_layout(**helpers["std_layout"](f"From revenue to EBITDA — {pick}, {month}",
                                               f"{cur} per month", height=440))
    fig.update_layout(showlegend=False)
    helpers["show_chart"](fig)


def _centre_outliers(pick, values, peers, desc, acct_values, acct_peers, label):
    detail = st.toggle("Every account line (not just the grouped lines)", key="pnl_out_detail")
    if detail and acct_values is not None:
        lines = [c for c in acct_peers.columns
                 if c not in ("Total Revenue", "Members", "Country", "Age (Months)")
                 and (acct_peers[c].abs().sum() > 0 or abs(acct_values.get(c, 0)) > 0)]
        table = comparison_rows(acct_values, acct_peers, lines)
    else:
        table = comparison_rows(values, peers, pnl.BENCH_LINES)
    st.caption(f"{pick} compared with {desc} ({label}). 'Cost vs typical' is what the "
               "difference is worth each month: positive costs money, negative saves it.")
    show = gross_margin_first(table.sort_values("Cost vs typical $/month", ascending=False,
                                                key=lambda s: s.fillna(0)))
    st.dataframe(bold_gross_margin(show, "Line"), hide_index=True, use_container_width=True,
                 column_config={"Cost vs typical $/month":
                                st.column_config.NumberColumn(format="%,.0f")})
    lines_only = show[~show["Line"].isin(pnl.NOT_COSTS | {"↳ vs centres that pay a manager"})]
    costly = lines_only[lines_only["Cost vs typical $/month"].fillna(0) > 0]
    ebitda = values.get("EBITDA")
    parts = [f"**{pick}** {'made' if ebitda >= 0 else 'lost'} **{_money(abs(ebitda))}** "
             f"({label})."]
    if not costly.empty:
        parts.append("Biggest costs above typical: " + "; ".join(
            f"**{r['Line']}** ({_money(r['Cost vs typical $/month'])} more)"
            for _, r in costly.head(3).iterrows()) + ".")
        if not detail:
            parts.append(f"If every line above typical were typical, EBITDA would be about "
                         f"**{_money(ebitda + costly['Cost vs typical $/month'].sum())}**.")
    st.markdown(_md(" ".join(parts)))


GM_SECTION = "Revenue − tutor wages + super"


def with_gross_margin_row(acc):
    """Every-account table ($, index P&L section / Account code / Account) with a Gross
    Margin row after the revenue and cost of sales accounts."""
    codes = pd.to_numeric(pd.Series(acc.index.get_level_values("Account code")),
                          errors="coerce").values
    rev = acc[pd.Series(codes).isin(pnl.REVENUE_CODES).values].sum()
    tutor = acc[pd.Series(codes).isin(pnl.TUTOR_CODES).values].sum()
    row = pd.DataFrame([rev - tutor], columns=acc.columns, index=pd.MultiIndex.from_tuples(
        [(GM_SECTION, -1, pnl.GROSS_MARGIN)], names=acc.index.names))
    top = acc.index.get_level_values("P&L section").isin(["Revenue", "Cost of Sales"])
    return pd.concat([acc[top], row, acc[~top]])


def _break_even(pick, cur, col, label, rates, samples):
    rev, mem = col["Total Revenue"], col.get("Members")
    if not mem or pd.isna(mem) or rev <= 0:
        st.info(f"{pick} has no members data or no revenue for {label}, so its break-even "
                "can't be worked out.")
        return
    amv0 = rev / mem / pnl.WEEKS_PER_MONTH
    tutor0 = col["Tutor Wages + Super"] / rev * 100
    c1, c2, _ = st.columns(3)
    amv = c1.number_input("Average Membership Value $ / week", min_value=0.0,
                          value=round(float(amv0), 2), step=1.0, key=f"pnl_amv_{pick}")
    rmin = c2.number_input("Royalty minimum $ / month", min_value=0.0,
                           value=float(col["Royalty Min"]), step=100.0, key=f"pnl_rmin_{pick}")
    per_member = amv * pnl.WEEKS_PER_MONTH
    fixed_lines = {k: col[k] for k in pnl.FIXED_ROWS}
    be_rev, be_mem = pnl.break_even_banded(per_member, sum(fixed_lines.values()), rmin, rates)
    actual = {k: col.get(k) for k in AMOUNT_KEYS}
    actual.update({"Members": mem, "AMV": amv0})
    model_cols = [("Actual", label, actual)]
    sample = None
    if be_mem:
        band = pnl.membership_band(be_mem)
        sample = samples.get(band)
        c = pnl.model_column(be_mem, per_member, pnl.rate_for(be_mem, rates), fixed_lines, rmin)
        c.update({"Members": be_mem, "AMV": amv})
        model_cols.append(("Break-even" + (" ⚠" if pnl.small_sample(sample) else ""),
                           f"{be_mem:,.1f} members", c))
    notes = {"Total Revenue": f"Members × 4.2 wks × ${amv:,.2f}",
             "Tutor Wages + Super": "Benchmark % for established centres of this size",
             "Royalties": f"Higher of ${rmin:,.0f} or 8% of revenue"}
    notes.update({k: "Centre's own amount (fixed)" for k in pnl.FIXED_ROWS})
    notes[pnl.GROSS_MARGIN] = GM_NOTE
    st.markdown(side_by_side_html(model_cols, notes, MODEL_ROWS), unsafe_allow_html=True)
    if be_mem:
        gap = mem - be_mem
        st.markdown(_md(
            f"**{pick}** needs about **{be_mem:,.0f} members** ({cur} {_money(be_rev)} a "
            f"month) to break even and has **{mem:,.0f}**: " +
            (f"**{gap:,.0f} above** break-even." if gap >= 0 else f"**{-gap:,.0f} short**.") +
            f" Tutor costs are {tutor0:.0f}% of revenue now; the benchmark at break-even size "
            f"is {pnl.rate_for(be_mem, rates) * 100:.0f}%."))
        if pnl.small_sample(sample):
            st.warning(f"⚠ The tutor % at break-even size ({pnl.membership_band(be_mem)} "
                       f"members) is based on {int(sample)} centres, fewer than "
                       f"{pnl.SMALL_SAMPLE}. Treat this break-even as indicative.")
    else:
        st.warning("At these settings tutor wages and royalties use up all the revenue, so more "
                   "members never reach break-even.")
    st.info(DISCLAIMER)


HOW_CENTRE = f"""
- **Headline, outliers, where the revenue goes and break-even** are for the chosen **Month**.
  **Months shown** sets the months in the two tables at the bottom.
- **Accounts costing more / less than typical**: each account as a % of revenue against similar
  centres (same country and members band; centres under 12 months old with other developing
  centres). The amount is what the difference is worth each month. Other centres are never
  named. ⚠ = fewer than {pnl.SMALL_SAMPLE} centres to compare with.
- **Break-even**: tutor wages at the benchmark % for established centres of that size (set on
  P&L – Group Level); rent, manager wages, marketing and OPEX at the centre's own amounts;
  royalties the higher of the Royalty Min (blank = 2,200) or 8% of revenue.
- **Gross Margin** = Total Revenue − Tutor Wages + Super (accounts 472 and 474), with its % of revenue. Higher is better. It is a subtotal, so it is never added into costs or EBITDA. In the outliers table it is pinned to the top and left
  out of "biggest costs above typical".
- Unrelated Expenses (479) are a memo line and aren't in EBITDA or Net Profit.
"""


def _headline_cards(now, prev, vs):
    """Revenue, gross margin, EBITDA, EBITDA %, members and revenue per member, each with
    the change against the previous period."""
    rev = now["Total Revenue"]
    mem = now.get("Members")
    rpm = rev / mem / pnl.WEEKS_PER_MONTH if mem and not pd.isna(mem) and mem > 0 else None

    def delta(fn):
        if prev is None:
            return None
        try:
            a_, b_ = fn(now), fn(prev)
            return None if a_ is None or b_ is None else a_ - b_
        except (TypeError, ZeroDivisionError, KeyError):
            return None

    gm_pct = lambda r: (r["Total Revenue"] - r["Tutor Wages + Super"]) / r["Total Revenue"] * 100 \
        if r["Total Revenue"] > 0 else None
    eb_pct = lambda r: r["EBITDA"] / r["Total Revenue"] * 100 if r["Total Revenue"] > 0 else None
    rpm_f = lambda r: r["Total Revenue"] / r["Members"] / pnl.WEEKS_PER_MONTH \
        if r.get("Members") and r["Members"] > 0 else None
    cards = [
        ("Revenue", _money(rev), delta(lambda r: r["Total Revenue"]), _money),
        ("Gross margin (after tutor wages)", f"{gm_pct(now):.1f}%" if rev > 0 else "–",
         delta(gm_pct), lambda v: f"{abs(v):.1f} pts"),
        ("EBITDA", _money_short(now["EBITDA"]), delta(lambda r: r["EBITDA"]), _money),
        ("EBITDA % of revenue", f"{eb_pct(now):.1f}%" if rev > 0 else "–", delta(eb_pct),
         lambda v: f"{abs(v):.1f} pts"),
        ("Members", _count(mem), delta(lambda r: r.get("Members")), lambda v: f"{abs(v):,.0f}"),
        ("Revenue per member / week", f"${rpm:,.2f}" if rpm else "–", delta(rpm_f),
         lambda v: f"${abs(v):,.2f}"),
    ]
    row = st.columns(len(cards))
    for col, (lbl, value, dlt, fmt) in zip(row, cards):
        ok = dlt is not None and not pd.isna(dlt)
        col.markdown(_card(lbl, value, dlt if ok else None, fmt(abs(dlt)) if ok else "",
                           dlt > 0 if ok else None, vs), unsafe_allow_html=True)



def report_centre(df_pnl, df_wm, vl, helpers, bench=None):
    """P&L – Centre Level: one centre's month, what's out of step with similar centres,
    its break-even, and every account by month."""
    st.markdown('<div class="report-title">P&L – Centre Level</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">One centre: how the month went, what costs more '
                'or less than similar centres, its break-even, and every account by month.'
                '</div>', unsafe_allow_html=True)
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return
    info = location_info(df_pnl, vl)
    lm = pnl.location_months(df_pnl)
    members = members_by_month(df_wm)
    periods = sorted(df_pnl["Period"].unique())
    centres = {loc.replace(PREFIX, ""): loc for loc in
               sorted(info.index, key=lambda l: l.replace(PREFIX, "").lower())}

    panel = st.container(border=True)
    with panel:
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b, c, d = st.columns([2, 1, 1, 1])
        pick = a.selectbox("Centre", list(centres), key="pc_centre")
        month = b.selectbox("Month", periods[::-1], format_func=pnl.period_name, key="pc_month")
        span = c.selectbox("Months shown", DD_SPANS, index=DD_SPANS.index(DEFAULT_SPAN),
                           key="pc_span")
        unit = d.selectbox("Show as", UNITS, key="pc_unit")
    loc = centres[pick]
    data, cur, _ = entity_months(lm, members, info, [loc])
    data = data[data.index <= month]
    if month not in data.index:
        st.warning(f"{pick} has no P&L for {pnl.period_name(month)}.")
        return
    # The months chosen: one month, or several (shown as the average per month), compared
    # with the same number of months just before.
    sel = months_shown(list(data.index), span, month)
    n_sel = len(sel)
    earlier = [p for p in data.index if p < sel[0]]
    prev_sel = earlier[-n_sel:] if len(earlier) >= n_sel else []
    now = average_months(data, sel)
    prev = average_months(data, prev_sel) if prev_sel else None
    label = pnl.period_name(month) if n_sel == 1 else \
        f"Average per month, {pnl.period_name(sel[0])} – {pnl.period_name(sel[-1])}"
    vs = "vs previous month" if n_sel == 1 else f"vs previous {n_sel} months"
    missed = [p for p in select_periods(periods, "All months", month)
              if sel[0] <= p <= month and p not in sel]

    # ── Row 1: headline ──
    _section(f"Headline — {label} ({cur})")
    if missed:
        st.caption(f"⚠ No P&L for {', '.join(pnl.period_name(p) for p in missed)}: the average "
                   "is over the months reported.")
    _headline_cards(now, prev, vs)
    rev = now["Total Revenue"]

    # Comparison with similar centres, for this month.
    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    bench_cent = centre_table(b_pnl, members_by_month(b_wm), vl, sel, helpers["age_group"])
    mine = centre_table(df_pnl[df_pnl["Location"] == loc], members, vl, sel,
                        helpers["age_group"])
    values = mine.loc[loc] if loc in mine.index else None
    peers = acct_values = acct_peers = None
    desc = ""
    if values is not None and rev >= pnl.LOW_REVENUE:
        peers, desc = peers_by_age(bench_cent, values.get("Age (Months)"), values.get("Country"),
                                   values.get("Members"), loc)
        acct = pnl.account_table(b_pnl, sel)
        if loc in acct.index and len(peers) >= 2:
            for c in ["Total Revenue", "Members", "Country"]:
                acct[c] = bench_cent[c].reindex(acct.index)
            acct_values = acct.loc[loc].copy()
            acct_peers = acct.loc[acct.index.intersection(peers.index)]
    tiles = None
    if acct_values is not None and len(acct_peers) >= 2:
        lines = [c for c in acct_peers.columns if c not in ("Total Revenue", "Members", "Country")]
        tiles = pnl.compare_lines(acct_values, acct_peers, lines)
        tiles = tiles[~tiles["Line"].isin(["Average Membership Value"])]
        tiles = tiles[tiles["Status"].isin(["High", "Very high", "Low", "Very low"])]

    # ── Rows 2 and 3: accounts out of step ──
    for bad, title in ((True, "Accounts costing more than typical"),
                       (False, "Accounts costing less than typical")):
        _section(title)
        if tiles is None:
            st.caption("Not enough similar centres reported this month to compare with."
                       if rev >= pnl.LOW_REVENUE else
                       "Revenue is too low this month for a meaningful comparison.")
            continue
        side = tiles[[pnl.is_bad(l, s_) == bad for l, s_ in zip(tiles["Line"], tiles["Status"])]]
        side = side.assign(_g=side["Gap $"].abs()).sort_values("_g", ascending=False)
        _outlier_tiles(side, bad, cur)

    # ── Row 4: scheduling efficiency ──
    scheduling_efficiency(pick, loc, values, peers, bench_cent, helpers.get("revenue"), vl,
                          sel, cur)

    # ── Where the revenue goes ──
    _section("Where the revenue goes")
    _waterfall(now, pick, label, cur, helpers)

    # ── Outliers vs similar centres ──
    _section("Outliers vs similar centres")
    if peers is None or len(peers) < 2:
        st.info("Not enough similar centres reported this month to compare with, or revenue "
                "is too low for a meaningful comparison.")
    else:
        _centre_outliers(pick, values, peers, desc, acct_values, acct_peers, label)

    # ── Break-even ──
    _section("Break-even")
    stored = st.session_state.get("pnl_bench_rates")
    if stored and stored["periods"] == tuple(sel):
        rates, samples = stored["rates"], stored["samples"]
    else:
        est, _ = pnl.split_by_age(bench_cent)
        t = pnl.tutor_benchmarks(est)
        rates, samples = dict(zip(t.index, t["Use %"])), dict(zip(t.index, t["Based on"]))
    if values is not None:
        _break_even(pick, cur, values, label, rates, samples)

    # ── By month ──
    table_months = sel if n_sel > 1 else months_shown(list(data.index), "Last 12 months", month)
    recent = data.loc[table_months]
    span_label = span.lower() if n_sel > 1 else "last 12 months"
    with st.expander(f"Every account line by month ({unit}, {span_label})", expanded=True):
        d = df_pnl[(df_pnl["Location"] == loc) & (df_pnl["Period"] <= month)].copy()
        d["Amount"] = d["Amount"].astype(float)
        acc = d.pivot_table(index=["P&L section", "Account code", "Account"], columns="Period",
                            values="Amount", aggfunc="sum", fill_value=0.0)
        acc = acc[[p for p in sorted(acc.columns) if p in set(recent.index)]]
        order = {"Revenue": 0, "Cost of Sales": 1, "Operating Costs": 2,
                 "Interest, Tax & Depreciation": 3}
        acc = acc.sort_index(key=lambda i: i.map(lambda v: order.get(v, 9))
                             if i.name == "P&L section" else i)
        acc = with_gross_margin_row(acc)
        acc = acc.droplevel("Account code")
        if unit == "% of revenue":
            rev_m = data["Total Revenue"].reindex(acc.columns)
            acc = acc.div(rev_m.where(rev_m > 0), axis=1) * 100
        elif unit == "$ per member":
            acc = acc.div(data["Members"].reindex(acc.columns).where(lambda m: m > 0), axis=1)
        acc.columns = [pnl.period_name(p) for p in acc.columns]
        acc = acc.apply(lambda col: col.map(lambda v: _fmt_unit(v, unit))).reset_index()
        acc = acc[["Account", "P&L section"] + [c for c in acc.columns
                                                if c not in ("Account", "P&L section")]]
        st.dataframe(bold_gross_margin(acc, "Account"), hide_index=True,
                     use_container_width=True,
                     column_config={"Account": st.column_config.TextColumn(
                         "Account", width="large", pinned=True)})

    with st.expander(f"Every line by month ({unit}, {span_label})"):
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
        full = pd.concat([members_row, shown])
        full.insert(0, "Line", full.index)
        st.dataframe(bold_gross_margin(full, "Line"), hide_index=True, use_container_width=True,
                     column_config={"Line": st.column_config.TextColumn(
                         "Line", width="large", pinned=True)})

    with st.expander("How this is worked out"):
        st.markdown(HOW_CENTRE + "\n- " + DISCLAIMER)


# ── Sessions vs students (Group Level) ────────────────────────────────────────


def BEST_WORST_AXIS(limit, title, unit="%"):
    """Diverging colour scale: red (worst) – grey (typical) – green (best)."""
    return dict(cmin=-limit, cmax=limit, cmid=0,
                colorscale=[[0, "#c8322f"], [0.5, "#c9d3d6"], [1, "#2e8540"]],
                colorbar=dict(title=dict(text=title, side="right"), thickness=12, len=0.8,
                              tickvals=[-limit, 0, limit],
                              ticktext=[f"Worst (−{limit}{unit})", "Typical",
                                        f"Best (+{limit}{unit})"]))
MAX_LABELS = 12          # name every centre in a small group; in a big one, only the outliers
MIN_STUDENTS = 5         # centres with fewer students a week are left out of these charts
YOUNG_WEEKS = 18         # centres this many weeks old or younger are left out too


def young_centres(vl):
    """Centres 18 weeks old or younger (Vlookup 'Weeks old'), left out of the sessions
    charts: their groups are still forming."""
    if LOC not in vl.columns or "Weeks old" not in vl.columns:
        return set()
    w = pd.to_numeric(vl["Weeks old"], errors="coerce")
    return set(vl.loc[w.notna() & (w <= YOUNG_WEEKS), LOC])


def sessions_section(group, bench_cent, revenue, periods, label, helpers, young=()):
    """Active students vs sessions for every centre around the network average, the
    group's centres named; then students per session vs tutor wages %."""
    loc_col = LOC
    table = pnl.sessions_table(revenue, periods, loc_col)
    left_young = sorted(set(group.index) & set(young) & set(table.index)) \
        if not table.empty else []
    if not table.empty:
        table = table[table["Students"] >= MIN_STUDENTS]     # tiny centres distort the ratio
        table = table[~table.index.isin(young)]              # still building their groups
    _section("Sessions vs active students",
             f"Weekly averages for {label}. The line is the network average "
             "(students per session); the shaded band is ±10%. Above the band = more sessions "
             "than typical for its students (smaller groups, higher tutor cost). Colour runs "
             "from green (fullest groups, best) to red (smallest groups, worst). This group's "
             "centres are outlined and labelled; other centres are paler. Hover for names of "
             f"centres you have access to. Centres under {MIN_STUDENTS} students a week or "
             f"{YOUNG_WEEKS} weeks old or younger are left out.")
    if left_young:
        st.caption(f"Left out ({YOUNG_WEEKS} weeks old or younger): " +
                   ", ".join(l.replace(PREFIX, "") for l in left_young))
    if table.empty:
        st.info("No sessions data in the Revenue tab for these months.")
        return
    typical = pnl.sessions_average(table)
    mine = table[table.index.isin(group.index)].copy()
    others = table[~table.index.isin(group.index)]
    mine["Status"] = mine["Students per session"].map(
        lambda v: pnl.sessions_status(v, typical))
    name = lambda l: l.replace(PREFIX, "")
    # Names on hover for every centre the viewer may see (admins: all; GPMs: their own).
    can_name = helpers.get("can_name") or (lambda l: l in group.index)
    who = lambda l: name(l) if can_name(l) else "Another centre"
    size = lambda s: (8 + 22 * (s / max(table["Students"].max(), 1)) ** 0.5).tolist()

    top = float(table["Students"].max()) * 1.08
    xs = [0, top]
    fig = go.Figure()
    # ±10% band and the average line (sessions = students ÷ typical students per session)
    fig.add_trace(go.Scatter(x=xs + xs[::-1],
                             y=[x / typical * (1 + pnl.SESSION_BAND) for x in xs] +
                               [x / typical * (1 - pnl.SESSION_BAND) for x in xs[::-1]],
                             fill="toself", fillcolor="rgba(93,109,115,0.10)", mode="lines",
                             line=dict(width=0), hoverinfo="skip", name="Typical (±10%)"))
    fig.add_trace(go.Scatter(x=xs, y=[x / typical for x in xs], mode="lines",
                             line=dict(color="#5d6d73", width=2, dash="dash"),
                             name=f"Network average ({typical:.2f} students per session)",
                             hoverinfo="skip"))
    # Colour best to worst: students per session against the average (fuller groups = best).
    table["Score"] = (table["Students per session"] / typical - 1) * 100
    label_all = len(mine) <= MAX_LABELS
    for part, is_group in ((others, False), (mine, True)):
        if part.empty:
            continue
        g = table.loc[part.index]
        labels = [name(l) if is_group and (label_all or abs(sc) > pnl.SESSION_BAND * 100)
                  else "" for l, sc in zip(g.index, g["Score"])]
        fig.add_trace(go.Scatter(
            x=g["Students"], y=g["Sessions"], mode="markers+text", text=labels,
            textposition="top center", textfont=dict(size=11, color="#1c2a30"),
            hovertext=[who(l) for l in g.index],
            customdata=g[["Students per session", "Score"]].values,
            marker=dict(color=g["Score"], coloraxis="coloraxis", size=size(g["Students"]),
                        opacity=1 if is_group else 0.5,
                        line=dict(color="#1c2a30" if is_group else "#ffffff",
                                  width=1.5 if is_group else 1)),
            name="This group" if is_group else "Other centres",
            hovertemplate="<b>%{hovertext}</b><br>%{x:.0f} students, %{y:.0f} sessions a "
                          "week<br>%{customdata[0]:.2f} students per session "
                          "(%{customdata[1]:+.0f}% vs average)<extra></extra>"))
    fig.update_layout(coloraxis=BEST_WORST_AXIS(30, "Students per session vs average"))
    fig.update_layout(**helpers["std_layout"]("Active students vs sessions (weekly average)",
                                               "Sessions per week", height=520))
    fig.update_xaxes(title_text="Active students per week", rangemode="tozero")
    fig.update_yaxes(rangemode="tozero")
    helpers["show_chart"](fig)

    if len(mine):
        far = mine.assign(
            Typical=typical,
            Diff=(mine["Students per session"] / typical - 1) * 100)
        far = far.reindex(far["Diff"].abs().sort_values(ascending=False).index)
        st.markdown("**Furthest from average in this group**")
        st.dataframe(pd.DataFrame({
            "Location": [name(l) for l in far.index],
            "Active students / week": far["Students"].round(0).values,
            "Sessions / week": far["Sessions"].round(0).values,
            "Students per session": far["Students per session"].values,
            "Typical": far["Typical"].values,
            "vs typical": far["Diff"].values,
            "Status": far["Status"].values}),
            hide_index=True, use_container_width=True, column_config={
                "Active students / week": st.column_config.NumberColumn(format="%d"),
                "Sessions / week": st.column_config.NumberColumn(format="%d"),
                "Students per session": st.column_config.NumberColumn(format="%.2f"),
                "Typical": st.column_config.NumberColumn(format="%.2f"),
                "vs typical": st.column_config.NumberColumn(
                    format="%+.0f%%", help="Students per session vs the network average. "
                    "Negative = smaller groups than typical.")})

    # ── Students per session vs tutor wages % ──
    cost = bench_cent[bench_cent["Total Revenue"] >= pnl.LOW_REVENUE]
    cost = cost.assign(**{"Tutor %": cost["Tutor Wages + Super"] / cost["Total Revenue"] * 100})
    both = table.join(cost[["Tutor %"]], how="inner")
    _section("Group size vs tutor cost",
             "Students per session (across) against tutor wages + super as a % of revenue (up). "
             "Colour runs from green (lowest tutor cost, best) to red (highest, worst). Top-left "
             "= small groups and high tutor cost: the centres to look at first. Dashed lines are "
             "the network middle values.")
    if len(both) < 2:
        st.info("Not enough centres with both sessions and P&L figures for these months.")
        return
    mx, my = both["Students per session"].median(), both["Tutor %"].median()
    g_both = both[both.index.isin(group.index)]
    o_both = both[~both.index.isin(group.index)]
    fig = go.Figure()
    # Colour best to worst: tutor wages % against the middle value (lower = best).
    both["Score"] = my - both["Tutor %"]
    for part, is_group in ((o_both, False), (g_both, True)):
        if part.empty:
            continue
        g = both.loc[part.index]
        flag = (g["Students per session"] < mx) & (g["Tutor %"] > my)
        labels = [name(l) if is_group and (len(g_both) <= MAX_LABELS or f) else ""
                  for l, f in zip(g.index, flag)]
        fig.add_trace(go.Scatter(
            x=g["Students per session"], y=g["Tutor %"], mode="markers+text", text=labels,
            textposition="top center", textfont=dict(size=11, color="#1c2a30"),
            hovertext=[who(l) for l in g.index],
            marker=dict(color=g["Score"], coloraxis="coloraxis", size=size(g["Students"]),
                        opacity=1 if is_group else 0.5,
                        line=dict(color="#1c2a30" if is_group else "#ffffff",
                                  width=1.5 if is_group else 1)),
            name="This group" if is_group else "Other centres",
            hovertemplate="<b>%{hovertext}</b><br>%{x:.2f} students per session<br>Tutor "
                          "cost %{y:.1f}% of revenue<extra></extra>"))
    fig.update_layout(coloraxis=BEST_WORST_AXIS(15, "Tutor wages % vs middle", unit=" pts"))
    fig.add_vline(x=mx, line=dict(color="#5d6d73", dash="dash", width=1),
                  annotation_text=f"Middle {mx:.2f}", annotation_position="top")
    fig.add_hline(y=my, line=dict(color="#5d6d73", dash="dash", width=1),
                  annotation_text=f"Middle {my:.0f}%", annotation_position="bottom right")
    fig.update_layout(**helpers["std_layout"]("Students per session vs tutor wages + super %",
                                               "Tutor wages + super, % of revenue", height=480))
    fig.update_xaxes(title_text="Students per session")
    helpers["show_chart"](fig)


# ── Scheduling efficiency (Centre Level) ──────────────────────────────────────
def _sparkline(centre, network, w=200, h=44):
    """Tiny inline line: this centre (teal) against the network average (dashed grey)."""
    vals = [v for v in list(centre) + list(network) if v is not None and not pd.isna(v)]
    if not vals:
        return ""
    lo, hi = min(vals), max(vals)
    pad = (hi - lo) * 0.15 or 0.1
    lo, hi = lo - pad, hi + pad
    n = max(len(centre), 2)

    def path(series):
        pts = [(i * (w - 4) / (n - 1) + 2, h - 2 - (v - lo) / (hi - lo) * (h - 4))
               for i, v in enumerate(series) if v is not None and not pd.isna(v)]
        return " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)

    return (f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" role="img" '
            f'aria-label="Students per session, last {len(centre)} weeks">'
            f'<polyline points="{path(network)}" fill="none" stroke="#9fb0b5" '
            f'stroke-width="1.5" stroke-dasharray="4 3"/>'
            f'<polyline points="{path(centre)}" fill="none" stroke="#1f9a9a" stroke-width="2"/>'
            '</svg>')


def _tile(title, value, sub="", colour="#1c2a30", extra=""):
    return (f'<div class="otile" style="--c:{colour}"><div class="t">{html.escape(title)}</div>'
            f'<div class="v">{value}</div>{extra}<div class="s">{sub}</div></div>')


def scheduling_efficiency(pick, loc, values, peers, bench_cent, revenue, vl, sel, cur):
    _section("Scheduling Efficiency")
    if loc in young_centres(vl):
        st.caption(f"Too new to compare: {YOUNG_WEEKS} weeks old or younger, so its groups are "
                   "still forming.")
        return
    table = pnl.sessions_table(revenue, sel, LOC)
    if table.empty or loc not in table.index:
        st.caption("No sessions data in the Revenue tab for this period.")
        return
    mine = table.loc[loc]
    network = table[(table["Students"] >= MIN_STUDENTS) & ~table.index.isin(young_centres(vl))]
    typical = pnl.sessions_average(network)
    if not typical:
        st.caption("Not enough centres with sessions data to compare with.")
        return
    students, sessions, sps = mine["Students"], mine["Sessions"], mine["Students per session"]
    status = pnl.sessions_status(sps, typical)
    diff = (sps / typical - 1) * 100
    extra = sessions - students / typical                       # sessions a week vs typical
    tutor = values.get("Tutor Wages + Super") if values is not None else None
    per_session = tutor / (sessions * pnl.WEEKS_PER_MONTH) if tutor and sessions else None
    cost = extra * pnl.WEEKS_PER_MONTH * per_session if per_session else None

    colour = {"More sessions than typical": "#c8322f", "Fewer sessions than typical": "#2e8540"
              }.get(status, "#5d6d73")
    end = pd.Timestamp(sel[-1] + "-01") + pd.offsets.MonthEnd(0)
    weekly = pnl.weekly_sessions(revenue, LOC, end, 12, young_centres(vl))
    spark, trend = "", None
    if loc in weekly.columns:
        trend = pnl.trend_word(weekly[loc])
        spark = _sparkline(weekly[loc].tolist(), weekly["Network"].tolist())
    t_colour = {"Improving": "#2e8540", "Getting worse": "#c8322f"}.get(trend, "#5d6d73")

    if cost is None:
        cost_txt, cost_sub = "–", "No tutor wages in the P&L for this period"
    elif extra >= 0:
        cost_txt = f"+{extra:.1f}/wk"
        cost_sub = f"≈ {_money(cost)} {cur} a month in tutor wages"
    else:
        cost_txt = f"{extra:.1f}/wk"
        cost_sub = f"≈ {_money(-cost)} {cur} a month saved"
    tiles = [
        _tile("Active students / week", f"{students:,.0f}"),
        _tile("Sessions / week", f"{sessions:,.0f}"),
        _tile("Students per session", f"{sps:.2f}",
              f"vs typical {typical:.2f} ({diff:+.0f}%) · {status.lower()}", colour),
        _tile("Extra sessions vs typical", cost_txt, cost_sub,
              "#c8322f" if extra > 0.5 else "#2e8540" if extra < -0.5 else "#5d6d73"),
        _tile("Last 12 weeks", trend or "–",
              "Students per session (teal) vs network average (dashed)", t_colour, spark),
    ]
    st.markdown(TILE_CSS, unsafe_allow_html=True)
    for col, t in zip(st.columns(len(tiles)), tiles):
        col.markdown(t, unsafe_allow_html=True)

    # One sentence pulling it together.
    rev = values.get("Total Revenue") if values is not None else None
    tutor_pct = tutor / rev * 100 if tutor and rev and rev > 0 else None
    typical_pct = None
    pool = peers if peers is not None and len(peers) else bench_cent
    if pool is not None and len(pool):
        p = pool[pool["Total Revenue"] > 0]
        typical_pct = (p["Tutor Wages + Super"] / p["Total Revenue"] * 100).median()
    wages = (f" ({tutor_pct:.0f}% of revenue vs {typical_pct:.0f}% typical)"
             if tutor_pct is not None and typical_pct is not None else "")
    if abs(extra) < 0.5 or cost is None:
        text = (f"**{pick}** runs about as many sessions as a typical centre with "
                f"{students:,.0f} students ({sps:.2f} students per session vs {typical:.2f}).")
    elif extra > 0:
        text = (f"**{pick}** runs about **{extra:.0f} more sessions a week** than a typical "
                f"centre with {students:,.0f} students ({sps:.2f} students per session vs "
                f"{typical:.2f}). At its tutor cost per session, that's about "
                f"**{_money(cost)} a month**, which helps explain its tutor wages{wages}.")
    else:
        text = (f"**{pick}** runs about **{-extra:.0f} fewer sessions a week** than a typical "
                f"centre with {students:,.0f} students ({sps:.2f} students per session vs "
                f"{typical:.2f}), saving about **{_money(-cost)} a month** in tutor wages"
                f"{wages}.")
    if trend:
        text += f" Over the last 12 weeks its group sizes are **{trend.lower()}**."
    st.markdown(_md(text))
    return status


# ── P&L Share (for franchise partners) ────────────────────────────────────────
SHARE_UNITS = ["$", "$ per member"]
CHART_WIDTH = 700               # px: fits the printable width of an A4 page
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

SHARE_CSS = """<style>
.share-head{display:flex;align-items:center;gap:16px;border-bottom:3px solid #1f9a9a;
 padding:4px 0 10px;margin:4px 0 12px}
.share-head img{height:46px}
.share-head .c{font-size:1.55em;font-weight:700;color:#1c2a30;line-height:1.15}
.share-head .m{color:#5d6d73;font-size:.92em}
.share-head .conf{margin-left:auto;text-align:right;font-size:.82em;color:#5d6d73}
.share-head .conf b{color:#c8322f;letter-spacing:.08em;text-transform:uppercase;font-size:1.05em}
.share-tbl{border-collapse:collapse;width:100%;font-size:12.5px;background:#fff;
 border:1px solid #dde5e6;margin-bottom:8px}
.share-tbl th{background:#1f3b57;color:#fff;font-weight:600;padding:6px 8px;text-align:right}
.share-tbl th.l,.share-tbl td.l{text-align:left}
.share-tbl td{padding:4px 8px;border-bottom:1px solid #eef2f3;text-align:right;color:#1c2a30}
.share-tbl td.muted{color:#8a9a9f}
.share-tbl tr.sens td{background:#f6f8f9}
.share-tbl tr.sub td{font-weight:700;background:#f2f8f8;border-bottom:1px solid #b9d3d6}
.pill{display:inline-block;border-radius:10px;padding:1px 8px;font-size:11.5px;font-weight:600;
 white-space:nowrap}
.pill.bad{background:#fbe9e8;color:#c8322f}.pill.good{background:#e7f3ea;color:#2e8540}
.pill.ok{background:#eef1f3;color:#5d6d73}.pill.none{background:transparent;color:#8a9a9f;
 font-weight:500;padding:0}
.share-box{background:#fff;border:1px solid #dde5e6;border-left:4px solid var(--c);
 border-radius:8px;padding:10px 14px;margin-bottom:10px;height:100%}
.share-box .h{font-weight:700;color:var(--c);margin-bottom:6px;font-size:1.02em}
.share-box ul{margin:0;padding-left:18px}.share-box li{margin:4px 0;color:#1c2a30;font-size:.92em}
.bands{display:flex;flex-wrap:wrap;gap:8px;margin:2px 0 8px}
.band{background:#fff;border:1px solid #dde5e6;border-radius:16px;padding:3px 12px;
 font-size:.85em;color:#1c2a30}.band b{color:#1f9a9a}
.check-item{margin:2px 0;color:#1c2a30}
.otile svg{max-width:100%;height:auto}
@media print {
  @page { size: A4 portrait; margin: 10mm; }
  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  html, body, .stApp, div[data-testid="stAppViewContainer"], section[data-testid="stMain"],
  .main, div[data-testid="stMainBlockContainer"] {
    height: auto !important; min-height: 0 !important; overflow: visible !important;
    position: static !important; background: #fff !important; }
  section[data-testid="stSidebar"], div[data-testid="stSidebarCollapsedControl"],
  header[data-testid="stHeader"], div[data-testid="stToolbar"],
  div[data-testid="stDecoration"], div[data-testid="stStatusWidget"],
  .topbar, .st-key-ps_filters, .st-key-ps_check, .st-key-ps_screen, .st-key-ps_how,
  details[data-testid="stExpander"] { display: none !important; }
  div[data-testid="stMainBlockContainer"], .main .block-container {
    padding: 0 !important; max-width: 100% !important; zoom: 0.8; }
  div[data-testid="stBottom"], .st-key-ps_how { display: none !important; }
  .otile, .metric-card, .share-box, .share-head, .bands, .share-tbl tr,
  div[data-testid="stPlotlyChart"], div[data-testid="stHorizontalBlock"] {
    break-inside: avoid; page-break-inside: avoid; }
  .section-header { break-after: avoid; page-break-after: avoid; }
  .st-key-ps_sched, .st-key-ps_members { break-inside: avoid; page-break-inside: avoid; }
  .page-break { break-before: page; page-break-before: always; }
  div[data-testid="stPlotlyChart"] { border: none !important; }
  .metric-card { padding: 8px 10px !important; }
  .metric-value { font-size: 1.35em !important; }
}
</style>"""

HOW_SHARE = f"""
- **Similar centres** are centres at the same stage (developing: under 12 months; or
  established), in the same country and of a similar size. If fewer than
  {pnl_share.MIN_PEERS} reported, the group widens (nearby sizes, then the whole country, then
  any stage). With fewer than {pnl_share.MIN_PEERS}, nothing is compared. Centres
  {pnl_share.YOUNG_WEEKS} weeks old or younger and centres with under {pnl.LOW_REVENUE:,} revenue
  a month are left out. Other centres are never named.
- **Typical** is the middle value of similar centres. **Normal range** is the middle half of
  them (the 25th to 75th percentile), so no single centre's figure is shown.
- A line is only compared when **most similar centres use that account**; otherwise it shows
  "No benchmark".
- **Rent, royalties, manager wages, interest and CRM software** are shown as this centre's own figures
  only. They are never compared and are left out of every total and opportunity.
- **Cost vs typical** is what the difference from typical is worth each month: positive costs
  money, negative saves it. **Bands** say which quarter of similar centres this centre is in,
  never an exact position.
- **Gross Margin** = Total Revenue − Tutor Wages + Super (accounts 472 and 474), with its % of revenue. Higher is better. It is a subtotal, so it is never added into costs or EBITDA. It is pinned to the top of the table and compared with similar
  centres like the other lines, but never counted in the opportunities.
- **Members trend**: average active members over the last 4 weeks against the 4 weeks before,
  compared with the network's established centres.
- Unrelated Expenses (479) are a memo line and aren't in EBITDA.
"""


def _logo_html():
    if not os.path.exists(LOGO):
        return ""
    with open(LOGO, "rb") as f:
        return f'<img src="data:image/png;base64,{base64.b64encode(f.read()).decode()}" alt="">'


def _share_header(pick, label, viewer):
    today = datetime.now(ZoneInfo("Australia/Sydney")).strftime("%-d %B %Y")
    by = f"Prepared by {html.escape(viewer)} · " if viewer else ""
    return (f'<div class="share-head">{_logo_html()}<div><div class="c">{html.escape(pick)}</div>'
            f'<div class="m">P&amp;L review — {html.escape(label)}</div></div>'
            f'<div class="conf"><b>Confidential</b><br>{by}{today}</div></div>')


def _signed_money(v):
    if v is None or pd.isna(v):
        return ""
    return ("+" if v > 0.5 else "−" if v < -0.5 else "") + f"${abs(v):,.0f}"


def _pill(line, status, note):
    if note:
        return f'<span class="pill none">{html.escape(note)}</span>'
    if not status:
        return ""
    word = pnl_share.soft_status(line, status)
    cls = "ok" if status == "Typical" else "bad" if pnl.is_bad(line, status) else "good"
    return f'<span class="pill {cls}">{html.escape(word)}</span>'


def share_table_html(table, top=None):
    """Every account line, all rows, as a plain table (prints in full, no scroll box).
    top: rows pinned first in bold (the Gross Margin row)."""
    t = table.copy()
    t["_b"] = t["Benchmarked"].astype(bool)
    t["_g"] = t["Gap $"].astype(float).fillna(0)
    t["_a"] = t["Actual $"].astype(float).fillna(0)
    t = pd.concat([t[t["_b"]].sort_values("_g", ascending=False),
                   t[~t["_b"]].sort_values("_a", ascending=False)])
    if top is not None and len(top):
        t = pd.concat([top.assign(_sub=True), t.assign(_sub=False)])
    head = ('<tr><th class="l">Line</th><th>Actual</th><th>% of revenue</th><th>Typical</th>'
            '<th>Normal range</th><th>Similar centres</th><th>Cost vs typical / month</th>'
            '<th class="l">Status</th></tr>')
    body = []
    for _, r in t.iterrows():
        amv = r["Line"] == "Average Membership Value"
        ok = bool(r["Benchmarked"])
        actual = f"${r['Actual $']:,.2f}/wk" if amv else _money(r["Actual $"])
        if ok and amv:
            typ = f"${r['Typical $']:,.2f}/wk"
            rng = f"${r['Range low $']:,.2f} – ${r['Range high $']:,.2f}"
        elif ok:
            typ = _fmt_pct(r["Typical %"])
            rng = f"{_fmt_pct(r['Range low %'])} – {_fmt_pct(r['Range high %'])}"
        else:
            typ = rng = ""
        comp = (f"{int(r['Centres'])}" + (" ⚠" if pnl.small_sample(r["Centres"]) else "")) \
            if ok else ""
        sens = r["Note"] == pnl_share.NOT_BENCHMARKED
        cls = "sub" if r.get("_sub") is True else "sens" if sens else ""
        body.append(
            f'<tr class="{cls}"><td class="l">{html.escape(str(r["Line"]))}</td>'
            f'<td>{actual}</td><td>{_fmt_pct(r["Actual %"])}</td><td>{typ}</td><td>{rng}</td>'
            f'<td class="muted">{comp}</td><td>{_signed_money(r["Gap $"]) if ok else ""}</td>'
            f'<td class="l">{_pill(r["Line"], r["Status"], r["Note"])}</td></tr>')
    return f'<table class="share-tbl">{head}{"".join(body)}</table>'.replace("$", "&#36;")


def _box(title, items, colour, empty):
    li = "".join(f"<li>{_md_html(i)}</li>" for i in items) or f"<li>{html.escape(empty)}</li>"
    return (f'<div class="share-box" style="--c:{colour}"><div class="h">{html.escape(title)}'
            f'</div><ul>{li}</ul></div>').replace("$", "&#36;")


def _md_html(text):
    """**bold** to <b>, everything else escaped."""
    parts = html.escape(text).split("**")
    return "".join(f"<b>{p}</b>" if i % 2 else p for i, p in enumerate(parts))


def _trend_charts(data, months, typical, unit, cur, helpers):
    x = [pd.Timestamp(p + "-01").strftime("%b %y") for p in months]
    d = data.reindex(months)
    per_member = unit == "$ per member"
    if per_member:
        y = d["Total Revenue"] / d["Members"].where(d["Members"] > 0) / pnl.WEEKS_PER_MONTH
        ty = typical["Revenue per member"] / pnl.WEEKS_PER_MONTH if len(typical) else None
        title, fmt = "Revenue per member / week", "$%{y:,.2f}"
    else:
        y = d["Total Revenue"]
        ty = typical["Revenue"] if len(typical) else None
        title, fmt = "Revenue", "$%{y:,.0f}"
    eb = d["EBITDA"] / d["Total Revenue"].where(d["Total Revenue"] > 0) * 100
    tey = typical["EBITDA %"] if len(typical) else None
    for name, ys, tys, yfmt, ytitle in [
            (title, y, ty, fmt, f"{cur} per member a week" if per_member else f"{cur} a month"),
            ("EBITDA % of revenue", eb, tey, "%{y:.1f}%", "% of revenue")]:
        fig = go.Figure()
        if tys is not None and pd.Series(tys).notna().any():
            fig.add_trace(go.Scatter(
                x=x, y=list(pd.Series(tys).reindex(months)), mode="lines",
                name="Typical similar centre", line=dict(color="#5d6d73", width=2, dash="dash"),
                hovertemplate=f"%{{x}}<br>Typical similar centre: {yfmt}<extra></extra>"))
        fig.add_trace(go.Scatter(
            x=x, y=list(ys), mode="lines+markers", name="This centre",
            line=dict(color="#1f9a9a", width=3), marker=dict(size=7),
            hovertemplate=f"%{{x}}<br>This centre: {yfmt}<extra></extra>"))
        if name.startswith("EBITDA"):
            fig.add_hline(y=0, line=dict(color="#c9d3d6", width=1))
        fig.update_layout(**helpers["std_layout"](name, ytitle, height=250))
        # A fixed width that fits an A4 page: charts don't resize when printed.
        fig.update_layout(width=CHART_WIDTH, legend=dict(orientation="h", x=1, xanchor="right",
                                                         y=1.12, yanchor="bottom"))
        fig.update_xaxes(type="category")
        helpers["show_chart"](fig, use_container_width=False)


def _members_section(trend, cur_members):
    _section("Members trend",
             "Average active members over the last 4 weeks against the 4 weeks before, "
             "compared with the network average.")
    if trend is None:
        st.caption("Not enough weekly membership history to show a trend yet.")
        return
    if trend["too_new"]:
        st.caption(f"Too new to compare: trading for {trend['weeks']} weeks "
                   f"(trends start at {pnl_share.YOUNG_WEEKS} weeks).")
        return
    colour = {"Green": "#2e8540", "Red": "#c8322f"}.get(trend["status"], "#5d6d73")
    word = {"Significantly above average trend": "Well ahead of the network",
            "Above average trend": "Ahead of the network",
            "About average": "In line with the network",
            "Below average trend": "A little behind the network",
            "Significantly below average trend": "Behind the network"}.get(trend["category"],
                                                                           trend["category"])
    tiles = [
        _tile("Members now", f"{trend['members']:,.0f}",
              f"{trend['change']:+,.0f} over 4 weeks"),
        _tile("This centre's trend", f"{trend['trend']:+.1f}%", "last 4 weeks vs the 4 before",
              colour),
        _tile("Network average", f"{trend['average']:+.1f}%", "established centres"),
        _tile("Compared with the network", word, f"{trend['vs']:+.1f} pts", colour),
        _tile("Last 12 weeks", "", "Active members each week", "#1f9a9a",
              _sparkline(trend["spark"], [])),
    ]
    st.markdown(TILE_CSS, unsafe_allow_html=True)
    for col, t in zip(st.columns(len(tiles)), tiles):
        col.markdown(t, unsafe_allow_html=True)
    if trend["small"]:
        st.caption("⚠ Under 20 members: percentages swing a lot from week to week.")


def report_share(df_pnl, df_wm, vl, helpers, bench=None):
    """P&L Share: one centre's month in a partner-friendly layout that GPMs can screenshot
    or print to PDF. Benchmarks use every centre, but no other centre is ever named and no
    single centre's figures can be worked out."""
    st.markdown(SHARE_CSS, unsafe_allow_html=True)
    with st.container(key="ps_screen"):
        st.markdown('<div class="report-title">P&L Share</div>', unsafe_allow_html=True)
        st.markdown('<div class="report-subtitle">A one-centre P&L review to share with the '
                    'franchise partner: how the month went, what is going well and where the '
                    'opportunities are.</div>', unsafe_allow_html=True)
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return
    info = location_info(df_pnl, vl)
    lm = pnl.location_months(df_pnl)
    members = members_by_month(df_wm)
    periods = sorted(df_pnl["Period"].unique())
    centres = {loc.replace(PREFIX, ""): loc for loc in
               sorted(info.index, key=lambda l: l.replace(PREFIX, "").lower())}

    with st.container(border=True, key="ps_filters"):
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b, c, d = st.columns([2, 1, 1, 1])
        pick = a.selectbox("Centre", list(centres), key="ps_centre")
        month = b.selectbox("Month", periods[::-1], format_func=pnl.period_name, key="ps_month")
        span = c.selectbox("Months shown", DD_SPANS, index=DD_SPANS.index(DEFAULT_SPAN),
                           key="ps_span")
        unit = d.selectbox("Show as", SHARE_UNITS, key="ps_unit",
                           help="Units for the 12-month revenue chart.")
        st.caption("To share: press Ctrl + P (Cmd + P on a Mac) and choose Save as PDF. The "
                   "menu, filters and the 'Check before sharing' box are left out of the "
                   "printout.")
    loc = centres[pick]
    data, cur, _ = entity_months(lm, members, info, [loc])
    data = data[data.index <= month]
    if month not in data.index:
        st.warning(f"{pick} has no P&L for {pnl.period_name(month)}.")
        return
    sel = months_shown(list(data.index), span, month)
    n_sel = len(sel)
    earlier = [p for p in data.index if p < sel[0]]
    prev_sel = earlier[-n_sel:] if len(earlier) >= n_sel else []
    now = average_months(data, sel)
    prev = average_months(data, prev_sel) if prev_sel else None
    label = pnl.period_name(month) if n_sel == 1 else \
        f"Average per month, {pnl.period_name(sel[0])} – {pnl.period_name(sel[-1])}"
    vs = "vs previous month" if n_sel == 1 else f"vs previous {n_sel} months"
    missed = [p for p in periods if sel[0] <= p <= month and p not in sel]
    rev = now["Total Revenue"]

    # ── Numbers: this centre, similar centres, every account line ──
    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    young = young_centres(vl)
    b_members = members_by_month(b_wm)
    bench_cent = centre_table(b_pnl, b_members, vl, sel, helpers["age_group"])
    mine = centre_table(df_pnl[df_pnl["Location"] == loc], members, vl, sel,
                        helpers["age_group"])
    values = mine.loc[loc] if loc in mine.index else None
    if values is None:
        st.warning(f"{pick} has no P&L for these months.")
        return
    peers, desc = bench_cent.iloc[0:0], None
    if rev >= pnl.LOW_REVENUE:
        peers, desc = pnl_share.share_peers(bench_cent, values.get("Age (Months)"),
                                            values.get("Country"), values.get("Members"), loc,
                                            young)
    sens = pnl_share.sensitive_accounts(b_pnl)
    grouped = pnl_share.benchmark_lines(values, peers, pnl.BENCH_LINES, sens)
    acct = pnl.account_table(b_pnl, sel)
    accounts = None
    if loc in acct.index:
        for col_ in ["Total Revenue", "Members", "Country"]:
            acct[col_] = bench_cent[col_].reindex(acct.index)
        acct_values = acct.loc[loc].copy()
        acct_peers = acct.loc[acct.index.intersection(peers.index)]
        lines = [c_ for c_ in acct.columns
                 if c_ not in ("Total Revenue", "Members", "Country")
                 and (acct_peers[c_].abs().sum() > 0 or abs(acct_values.get(c_, 0)) > 0)]
        accounts = pnl_share.benchmark_lines(acct_values, acct_peers, lines, sens)
    more = less = pd.DataFrame()
    if accounts is not None:
        more, less = pnl_share.out_of_step(accounts)
    codes = b_pnl.groupby("Account")["Account code"].first()
    account_rows = {a_: pnl.row_of(int(c_)) for a_, c_ in codes.items()}
    opps = pnl_share.top_opportunities(pnl_share.opportunity_lines(
        grouped, accounts if accounts is not None else grouped.iloc[0:0], account_rows))
    bands = pnl_share.centre_bands(values, peers)
    mine_rows = df_pnl[(df_pnl["Location"] == loc) & df_pnl["Period"].isin(sel)]
    interest_income = float(mine_rows.loc[mine_rows["Account code"] == 270, "Amount"]
                            .astype(float).sum()) / max(n_sel, 1)

    # ── Check before sharing (screen only) ──
    found = pnl_share.checks(values, peers, accounts, [pnl.period_name(p) for p in missed],
                             interest_income, loc in young)
    with st.container(border=True, key="ps_check"):
        st.markdown('<div class="section-header" style="margin-top:0">Check before sharing'
                    '</div>', unsafe_allow_html=True)
        st.caption("Only you see this box: it is left out when you print.")
        if found:
            for f_ in found:
                st.markdown(f'<div class="check-item">⚠ {_md_html(f_)}</div>',
                            unsafe_allow_html=True)
        else:
            st.markdown('<div class="check-item">✓ Nothing unusual found.</div>',
                        unsafe_allow_html=True)

    # ── Header ──
    st.markdown(_share_header(pick, label, helpers.get("viewer_name", "")),
                unsafe_allow_html=True)

    # ── Headline ──
    _section(f"Headline — {label} ({cur})")
    if missed:
        st.caption(f"⚠ No P&L for {', '.join(pnl.period_name(p) for p in missed)}: the average "
                   "is over the months reported.")
    _headline_cards(now, prev, vs)
    if desc:
        st.markdown('<div class="bands">' + "".join(
            f'<span class="band"><b>{html.escape(b_)}</b> for {html.escape(m_)}</span>'
            for m_, b_ in bands.items()) + '</div>', unsafe_allow_html=True)
        st.caption(f"Compared with {desc} ({len(peers)} centres)" +
                   (" ⚠ small sample" if pnl.small_sample(len(peers)) else "") + ".")
    else:
        st.caption(pnl_share.NOT_ENOUGH + "." if rev >= pnl.LOW_REVENUE else
                   "Revenue is too low this month for a meaningful comparison.")

    # ── Accounts out of step ──
    for bad, title, rows in ((True, "Accounts costing more than typical", more),
                             (False, "Accounts costing less than typical", less)):
        _section(title)
        if not desc:
            st.caption(pnl_share.NOT_ENOUGH + ".")
            continue
        _outlier_tiles(rows, bad, cur)

    # ── What's going well / top opportunities (filled once scheduling is known) ──
    well_slot = st.container()

    # ── Scheduling efficiency ──
    with st.container(key="ps_sched"):
        sched = scheduling_efficiency(pick, loc, values, peers if len(peers) else None,
                                      bench_cent, helpers.get("revenue"), vl, sel, cur)

    # ── 12-month trend ──
    _section("Last 12 months", "This centre (teal) against the typical similar centre "
             f"(dashed). The typical line needs at least {pnl_share.MIN_PEERS} similar centres "
             "reporting that month.")
    t_months = months_shown(list(data.index), "Last 12 months", month)
    country_of = location_info(b_pnl, vl)["Country"].to_dict()
    typical = pnl_share.typical_by_month(pnl.location_months(b_pnl), b_members, list(peers.index),
                                         t_months, values.get("Country"), country_of) \
        if len(peers) else pd.DataFrame()
    data_m = data.copy()
    _trend_charts(data_m, t_months, typical, unit, cur, helpers)

    # ── Members trend ──
    end = pd.Timestamp(sel[-1] + "-01") + pd.offsets.MonthEnd(0)
    trend = None
    if b_wm is not None and not b_wm.empty and "Date" in b_wm:
        weeks = b_wm.loc[b_wm["Date"] <= end, "Date"]
        if len(weeks):
            trend = pnl_share.members_trend(b_wm, loc, weeks.max())
    with st.container(key="ps_members"):
        _members_section(trend, now.get("Members"))

    # ── What's going well ──
    good = []
    for _, r in less.head(3).iterrows():
        good.append(f"**{r['Line']}**: {r['Actual %']:.1f}% of revenue vs {r['Typical %']:.1f}% "
                    f"typical, about ${abs(r['Gap $']):,.0f} a month less.")
    for m_, b_ in bands.items():
        if b_ in pnl_share.BANDS[:2]:
            good.append(f"**{b_}** of similar centres for {m_}.")
    if now["EBITDA"] > 0:
        good.append(f"Made a profit (EBITDA) of **${now['EBITDA']:,.0f}**" +
                    (" a month." if n_sel > 1 else "."))
    if prev is not None and rev > prev["Total Revenue"] + 0.5:
        good.append(f"Revenue up **${rev - prev['Total Revenue']:,.0f}** {vs}.")
    if trend and not trend["too_new"] and trend["vs"] > 2:
        good.append(f"Members growing faster than the network ({trend['trend']:+.1f}% vs "
                    f"{trend['average']:+.1f}% over 4 weeks).")
    if sched == "Fewer sessions than typical":
        good.append("Fuller groups than typical: fewer sessions for the number of students.")
    tips = [o["Text"] for o in opps]
    with well_slot:
        g1, g2 = st.columns(2)
        g1.markdown(_box("What's going well", good, "#2e8540",
                         "A steady period: nothing stands out yet."), unsafe_allow_html=True)
        g2.markdown(_box("Top opportunities", tips, "#1f9a9a",
                         "No costs well above similar centres this period." if desc else
                         pnl_share.NOT_ENOUGH + "."), unsafe_allow_html=True)

    # ── Outliers vs similar centres: every account line, all rows ──
    st.markdown('<div class="page-break"></div>', unsafe_allow_html=True)
    _section("Outliers vs similar centres",
             (f"{pick} compared with {desc}. " if desc else "") +
             "Rent, royalties, manager wages, interest and CRM software are this centre's own "
             "figures only. "
             "'Cost vs typical' is what the difference is worth each month: + costs money, "
             "− saves it.")
    if accounts is None or accounts.empty:
        st.caption("No account lines for this period.")
    else:
        st.markdown(share_table_html(accounts, grouped[grouped["Line"] == pnl.GROSS_MARGIN]),
                    unsafe_allow_html=True)
    ebitda = now["EBITDA"]
    parts = [f"**{pick}** {'made' if ebitda >= 0 else 'lost'} **${abs(ebitda):,.0f}** "
             f"({label})."]
    if not more.empty:
        parts.append("Biggest costs above typical: " + "; ".join(
            f"**{r['Line']}** (${r['Gap $']:,.0f} more)" for _, r in more.head(3).iterrows()) +
            ".")
        parts.append(f"If these accounts were at typical, EBITDA would be about "
                     f"**{_money_short(ebitda + more['Gap $'].sum())}**.")
    st.markdown(_md(" ".join(parts)))

    with st.container(key="ps_how"):
        with st.expander("How this is worked out"):
            st.markdown(HOW_SHARE)


# ── P&L Centre Forecast (admins only) ─────────────────────────────────────────
FORECAST_VIEWS = ["Financial year", "Next 12 months"]
GRID_LABELS = {"Tutor Wages + Super": "Tutor wages %", "Rent": "Rent %",
               "Wages - Manager/Admin + Super": "Manager wages %", "Royalties": "Royalties %",
               "Marketing - Agency Fees": "Agency fees %", "Marketing": "Marketing %",
               "OPEX": "OPEX %"}
BASIS_COLOURS = {pf.ACTUAL: "#1f9a9a", pf.FORECAST: "#8fcfcf", pf.ESTIMATE: "#d5e3e4"}
FORECAST_ROWS = MODEL_ROWS + [
    (None, None, "gap"),
    ("Break-even members", "BE members", "count"),
    ("Members above / below break-even", "BE gap", "count"),
]
HOW_FORECAST = f"""
- **Actual** months are the centre's P&L. Every later month is worked out from the centre's
  last {pf.RECENT} P&L months (fewer if it has fewer).
- **Members**: months with Weekly Membership data use it. Later months carry on from the last
  known month: the usual change for that time of year plus the centre's recent growth, which
  counts a little less each month (×{pf.DAMP}) and is capped at ±{pf.MAX_TREND:.0%} a month. The time-of-year pattern
  is the centre's own once it has {pf.OWN_PATTERN_MONTHS} months of membership history, otherwise the network's
  (centres with {pf.NETWORK_MIN_MEMBERS}+ members).
- **Revenue** = members × Average Membership Value (average of the last {pf.RECENT} months) × 4.2 weeks.
- **Tutor wages and marketing**: the average % of revenue over the last {pf.RECENT} months.
- **Rent, manager wages, agency fees and OPEX** stay at their average $ of the last {pf.RECENT} months,
  shown as a % of that month's revenue. **Royalties**: the higher of the Royalty minimum or 8%
  of revenue, or $0 while royalty relief lasts.
- **Changing a %** sets that line to that % of revenue for that month. Lines left alone keep
  the behaviour above, so rent stays the same $ if you change members.
- **Estimate ⚠** = the centre has no P&L for that month of the year yet, so the forecast
  relies more on the network pattern. Treat these months with more caution.
- **Break-even members** = members needed that month for EBITDA of 0 at that month's AMV and
  costs.
- **Benchmarks**: the forecast months' average % of revenue against similar centres (same
  country and members band; centres under 12 months old against other developing centres).
  The normal range is the middle half (25th–75th percentile). ⚠ = fewer than
  {pnl.SMALL_SAMPLE} centres. Not shown with fewer than {pnl_share.MIN_PEERS} centres.
- **Gross Margin** = Total Revenue − Tutor Wages + Super (accounts 472 and 474), with its % of revenue. Higher is better. It is a subtotal, so it is never added into costs or EBITDA.
- Interest Income (account 270) is counted in revenue, as in the other P&L reports.
"""
FORECAST_NOTE = ("A forecast is an estimate from the centre's recent figures and network "
                 "patterns, not a promise. Check it against what you know about the centre: "
                 "price changes, staffing, lease and marketing plans.")


def _royalty_min(vl, loc):
    v = vl.set_index(LOC)["Royalty Min"].get(loc) if LOC in vl.columns and \
        "Royalty Min" in vl.columns else None
    v = pd.to_numeric(str(v).replace("$", "").replace(",", "").strip(), errors="coerce")
    return float(v) if not pd.isna(v) else pnl.DEFAULT_ROYALTY_MIN


def _forecast_chart(table, pick, cur, helpers):
    t = table[table["Basis"] != pf.MISSING]
    x = [pd.Period(p, "M").strftime("%b %y") for p in t.index]
    fig = go.Figure()
    for basis, colour in BASIS_COLOURS.items():
        mask = t["Basis"] == basis
        if mask.any():
            fig.add_bar(x=[x_ for x_, m in zip(x, mask) if m], y=t.loc[mask, "Total Revenue"],
                        name=f"Revenue ({basis.replace(' ⚠', '')})", marker_color=colour,
                        hovertemplate="%{x}: %{y:$,.0f}<extra>Revenue</extra>")
    fig.add_scatter(x=x, y=t["EBITDA"], name="EBITDA", mode="lines+markers",
                    line=dict(color="#1c2a30", width=2),
                    marker=dict(color=["#c8322f" if v < 0 else "#2e8540" for v in t["EBITDA"]],
                                size=8),
                    hovertemplate="%{x}: %{y:$,.0f}<extra>EBITDA</extra>")
    fig.add_hline(y=0, line=dict(color="#5d6d73", width=1))
    fig.update_layout(**helpers["std_layout"](f"Revenue and EBITDA by month — {pick}",
                                               f"{cur} per month", height=420))
    fig.update_layout(barmode="overlay", legend=dict(orientation="h", y=-0.25),
                      xaxis=dict(categoryorder="array", categoryarray=x))
    helpers["show_chart"](fig)


def _forecast_benchmarks(values, peers, desc, min_n=pnl_share.MIN_PEERS, label="Forecast"):
    """Benchmark % for each line, next to the forecast's average %."""
    if values is None:
        st.caption("No forecast months shown, so there's nothing to compare.")
        return
    if peers is None or len(peers) < min_n:
        st.caption(pnl_share.NOT_ENOUGH + ".")
        return
    rows = pnl.compare_lines(values, peers, pnl.BENCH_LINES)
    rows = rows[rows["Centres"] >= min_n] if not rows.empty else rows
    if rows.empty:
        st.caption(pnl_share.NOT_ENOUGH + ".")
        return
    small = pnl.small_sample(len(peers))
    out = []
    for _, r in rows.iterrows():
        amv = r["Line"] == "Average Membership Value"
        fmt = (lambda v: "–" if v is None or pd.isna(v) else f"${v:,.2f}") if amv else \
            (lambda v: "–" if v is None or pd.isna(v) else f"{v:.1f}%")
        lo, hi = (r.get("Range low $"), r.get("Range high $")) if amv else \
            (r["Range low %"], r["Range high %"])
        out.append({"Line": r["Line"],
                    f"{label} (avg)": fmt(r["Actual $"] if amv else r["Actual %"]),
                    "Typical": fmt(r.get("Typical $") if amv else r["Typical %"]),
                    "Normal range": f"{fmt(lo)} – {fmt(hi)}",
                    "Compared with": f"{int(r['Centres'])} centres" + (" ⚠" if small else ""),
                    f"{label} vs range": pnl_share.soft_status(r["Line"], r["Status"])})
    st.caption(f"Compared with {desc}. Use these when changing a line's %.")
    st.dataframe(pd.DataFrame(out), hide_index=True, use_container_width=True)
    if small:
        st.warning(f"⚠ Fewer than {pnl.SMALL_SAMPLE} similar centres: treat these ranges as "
                   "indicative.")


def _forecast_excel(table, grid, settings):
    """The months, the grid as changed and the settings, as an Excel file."""
    buf = io.BytesIO()
    lines = [pnl.ROW_ORDER[0], pnl.TUTOR, pnl.GROSS_MARGIN, "Gross Margin %"] + \
        [r for r in pnl.ROW_ORDER[1:] if r != pnl.TUTOR]
    cols = ["Basis", "Members", "AMV"] + lines + ["EBITDA", "BE members"]
    t = pnl.with_gross_margin(table)
    t["Gross Margin %"] = [pnl.gross_margin_pct(r, w) for r, w in
                           zip(t.get("Total Revenue", []), t.get(pnl.TUTOR, []))]
    out = t.reindex(columns=cols).copy()
    out.index = [pnl.period_name(p) for p in out.index]
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        out.rename(columns={"AMV": "AMV $ / week"}).to_excel(xl, sheet_name="Forecast",
                                                             index_label="Month")
        if grid is not None and not grid.empty:
            g = grid.copy()
            g.index = [pnl.period_name(p) for p in g.index]
            g.to_excel(xl, sheet_name="Adjustments", index_label="Month")
        pd.DataFrame(list(settings.items()), columns=["Setting", "Value"]).to_excel(
            xl, sheet_name="Settings", index=False)
    return buf.getvalue()


def report_forecast(df_pnl, df_wm, vl, helpers, bench=None):
    """P&L Centre Forecast (admins only): one centre's condensed P&L for a financial year or
    the next 12 months, worked out from its recent months, with every line adjustable."""
    st.markdown('<div class="report-title">P&L Centre Forecast</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">One centre\'s P&L forecast month by month, from '
                'its recent results. Change members, AMV or any line\'s % to test a plan.</div>',
                unsafe_allow_html=True)
    if st.session_state.get("access_level") != "admin":
        st.error("This report is for admins only.")
        return
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return
    info = location_info(df_pnl, vl)
    lm = pnl.location_months(df_pnl)
    members = members_by_month(df_wm)
    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    centres = {loc.replace(PREFIX, ""): loc for loc in
               sorted(info.index, key=lambda l: l.replace(PREFIX, "").lower())}

    with st.container(border=True):
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b, c = st.columns([2, 1, 1])
        pick = a.selectbox("Centre", list(centres), key="pf_centre")
        loc = centres[pick]
        data, cur, _ = entity_months(lm, members, info, [loc])
        last = data.index.max()
        view = b.selectbox("Show", FORECAST_VIEWS, key="pf_view")
        first_fy = pf.fy_start(pf.add_months(last, 1))
        fy = c.selectbox("Financial year", [first_fy, first_fy + 1], format_func=pf.fy_label,
                         key="pf_fy", disabled=view != "Financial year")
    months = pf.window(last, view, fy)
    fc_months = [p for p in months if p > last]
    start = pf.starting_point(data)
    own_series = members.xs(loc, level=0) if not members.empty and \
        loc in members.index.get_level_values(0) else pd.Series(dtype=float)
    if not start["amv"] or own_series.empty:
        st.warning(f"{pick} has no members data for its recent P&L months, so revenue can't "
                   "be forecast. Check its Weekly Membership upload.")
        return
    own = pf.own_pattern(own_series)
    pattern = own or pf.network_pattern(members_by_month(b_wm))
    member_fc = pf.forecast_members(own_series, fc_months, pattern)
    seen = set(data.index.str[5:7])
    if len(start["months"]) < pf.RECENT:
        st.warning(f"⚠ {pick} has only {len(start['months'])} month(s) of P&L, so the starting "
                   "point is based on fewer months than usual.")

    # ── Settings ──
    _section("Settings", "These set the starting forecast. Changing one resets the "
                         "month-by-month changes below.")
    s1, s2, s3, s4 = st.columns(4)
    rmin = s1.number_input("Royalty minimum $ / month", min_value=0.0, step=100.0,
                           value=_royalty_min(vl, loc), key=f"pf_rmin_{loc}")
    relief_opts = ["None"] + [f"Until {pnl.period_name(p)}" for p in fc_months[:-1]] + \
        ["Whole forecast"]
    relief = s2.selectbox("Royalty relief", relief_opts,
                          index=len(relief_opts) - 1 if start["royalties"] == 0 else 0,
                          key=f"pf_relief_{loc}",
                          help="Royalties at $0 now are normal (royalty relief). Choose when "
                               "the relief ends.")
    relief_until = None if relief == "None" or not fc_months else \
        fc_months[-1] if relief == "Whole forecast" else \
        fc_months[relief_opts.index(relief) - 1]
    rise_pct = s3.number_input("Price rise %", min_value=-50.0, max_value=100.0, value=0.0,
                               step=1.0, key=f"pf_rise_{loc}")
    rise_from = s4.selectbox("Price rise from", fc_months or [None],
                             format_func=lambda p: pnl.period_name(p) if p else "–",
                             key=f"pf_rise_from_{loc}", disabled=not rise_pct)
    base = pf.defaults(fc_months, member_fc, start, rmin, relief_until, rise_from, rise_pct)

    top = st.container()                    # headline and chart, filled in once worked out

    # ── Adjust the forecast ──
    grid = base
    if fc_months:
        resets = st.session_state.setdefault("pf_resets", 0)
        key = f"pf_grid_{loc}_{view}_{fy}_{rmin}_{relief}_{rise_pct}_{rise_from}_{resets}"
        shown = base.copy()
        shown.insert(0, "Basis", [pf.FORECAST if p[5:7] in seen else pf.ESTIMATE
                                  for p in fc_months])
        shown.index = [pnl.period_name(p) for p in fc_months]
        pct_cols = {pf.PCT[l]: st.column_config.NumberColumn(
            GRID_LABELS[l], format="%.1f%%", min_value=0.0, max_value=200.0, step=0.1)
            for l in pf.COST_LINES}
        _section("Adjust the forecast",
                 "Change members, AMV or any line's % of revenue for any month. Lines you "
                 "leave alone keep their usual behaviour (see How this is worked out).")
        edited = st.data_editor(
            shown, key=key, use_container_width=True,
            disabled=["Basis"],
            column_config={"Basis": st.column_config.TextColumn("Basis"),
                           "Members": st.column_config.NumberColumn(format="%.1f", min_value=0.0),
                           "AMV $ / week": st.column_config.NumberColumn(format="$%.2f",
                                                                         min_value=0.0),
                           **pct_cols})
        if st.button("Reset to forecast", icon=":material/restart_alt:", key="pf_reset"):
            st.session_state["pf_resets"] = resets + 1
            st.rerun()
        grid = edited.drop(columns="Basis")
        grid.index = fc_months
        grid = grid.fillna(base)
    fc = pf.compute(grid, base, start, rmin, relief_until)
    table = pf.months_table(data, fc, months, seen, rmin)
    table["BE gap"] = table["Members"] - table["BE members"] if "BE members" in table else None
    tot = pf.totals(table)

    # ── Headline ──
    title = pf.fy_label(fy) if view == "Financial year" else \
        f"Next 12 months ({pnl.period_name(months[0])} – {pnl.period_name(months[-1])})"
    with top:
        _section(f"Headline — {title} ({cur})",
                 f"{tot['Months'] - tot['Forecast months']} actual and "
                 f"{tot['Forecast months']} forecast months" +
                 (f", {tot['Estimate months']} of them estimates ⚠"
                  if tot["Estimate months"] else "") + ".")
        cards = [("Revenue", _money(tot["Revenue"])), ("EBITDA", _money_short(tot["EBITDA"])),
                 ("EBITDA % of revenue",
                  f"{tot['EBITDA %']:.1f}%" if tot["EBITDA %"] is not None else "–"),
                 (f"Members in {pnl.period_name(months[-1])}", _count(tot["End members"])),
                 ("Months below break-even", f"{tot['Months below 0']} of {tot['Months']}")]
        for col, (lbl, value) in zip(st.columns(len(cards)), cards):
            col.markdown(_card(lbl, value), unsafe_allow_html=True)
        missing = table.index[table["Basis"] == pf.MISSING]
        if len(missing):
            st.caption(f"⚠ No P&L for {', '.join(pnl.period_name(p) for p in missing)}: left "
                       "out of the totals.")
        _forecast_chart(table, pick, cur, helpers)

    # ── Month by month ──
    _section("Forecast by month",
             "Actual = the centre's P&L. Forecast / Estimate ⚠ = worked out (see How this is "
             "worked out).")
    cols = []
    for p in months:
        r = table.loc[p].to_dict()
        cols.append((pnl.period_name(p), r["Basis"], {k: (None if v is None or
                     (isinstance(v, float) and pd.isna(v)) else v) for k, v in r.items()}))
    t = table[table["Basis"] != pf.MISSING]
    total = pnl.with_gross_margin(t)[pnl.ROW_ORDER + [pnl.GROSS_MARGIN, "EBITDA"]].sum().to_dict()
    total["Members"] = t["Members"].mean()
    total["AMV"] = total["Total Revenue"] / total["Members"] / pnl.WEEKS_PER_MONTH / len(t) \
        if total["Members"] and len(t) else None
    cols.append(("Total", f"{len(t)} months", total))
    st.markdown(side_by_side_html(cols, rows=FORECAST_ROWS), unsafe_allow_html=True)
    st.caption("Total: amounts added up over the months; members and AMV are the average "
               "per month.")

    # ── Benchmarks ──
    _section("Benchmark % for each line")
    sel = start["months"]
    bench_cent = centre_table(b_pnl, members_by_month(b_wm), vl, sel, helpers["age_group"])
    values = pf.benchmark_values(table, info.at[loc, "Country"]
                                 if isinstance(info.at[loc, "Country"], str) else None)
    age = vl.set_index(LOC)["Age (Months)"].get(loc) if LOC in vl.columns and \
        "Age (Months)" in vl.columns else None
    peers, desc = None, ""
    if not bench_cent.empty:
        country = values.get("Country") if values is not None else None
        peers, desc = peers_by_age(bench_cent, age, country,
                                   data["Members"].tail(pf.RECENT).mean(), loc)
    _forecast_benchmarks(values, peers, desc)

    settings = {"Centre": pick, "Shown": title, "Royalty minimum $ / month": rmin,
                "Royalty relief": relief, "Price rise %": rise_pct,
                "Price rise from": pnl.period_name(rise_from) if rise_pct and rise_from else "–",
                "Starting point": ", ".join(pnl.period_name(p) for p in sel),
                "Time-of-year pattern": "centre's own" if own else "network",
                "Made": datetime.now(ZoneInfo("Australia/Sydney")).strftime("%d/%m/%Y %H:%M")}
    st.download_button("Download as Excel", _forecast_excel(table, grid if fc_months else None,
                                                            settings),
                       file_name=f"{datetime.now(ZoneInfo('Australia/Sydney')):%d-%m-%Y}-"
                                 f"pnl-forecast-{pick.lower().replace(' ', '-')}.xlsx",
                       icon=":material/download:", key="pf_download")
    st.info(FORECAST_NOTE)
    with st.expander("How this is worked out"):
        st.markdown(_md(HOW_FORECAST))


# ── Franchise Model (admins only) ─────────────────────────────────────────────
MODEL_TABLE_ROWS = MODEL_ROWS + [
    (None, None, "gap"),
    ("Owner's wage", pm.OWNER, "money"),
    ("Profit for payback", "Profit", "money"),
    ("Cumulative profit", "Cumulative", "money"),
]
MODEL_COLOURS = {pm.WITHOUT: "#1f9a9a", pm.WITH: "#1f3b57"}
HOW_MODEL = """
- **A theoretical new centre**, opening in the chosen month. Typical figures come from the
  country's centres over their last 3 P&L months (established centres, 12 months or older,
  where there are any): the middle value (median), not the average, so one unusual centre
  doesn't move it. ⚠ = fewer than 8 centres behind a figure.
- **Members**: a straight line from 0 to the members at week 12, then the new members a month
  until the mature members are reached. "Flat members" uses the mature members from month 1.
- **Revenue** = members × price per member × 4.2 weeks. The price rises by the yearly % each
  July after opening.
- **Tutor wages + super**: the typical % of revenue for that many members (members bands, as
  in the break-even). **Gross Margin** = Total Revenue − Tutor Wages + Super. **Marketing**:
  the typical % of revenue. **Rent, agency fees and OPEX**: the typical $ a month.
- **Manager** (with-manager scenario only): hired the first month members reach the trigger
  and kept after that, at the manager cost a month. The wage rise applies each July.
- **Royalties**: the higher of the minimum or 8% of revenue, and nothing in the relief weeks
  (a month partly in relief pays part). The minimum rises by its % each July.
- **Profit for payback** = EBITDA, less the owner's wage when it is switched on. **Payback** is
  the first month the profit added up since opening (early losses included) reaches the
  investment target, converted from AUD at the country's exchange rate.
- **Needed to pay back on time**: the mature members (or the price per member) at which
  payback falls within the target months, with everything else as set. These use the
  settings only, not the changes in the grid.
- **Changing the grid** sets that month's members or price, or a line to that % of revenue.
  Lines left alone keep the behaviour above. Changes apply to both scenarios (manager % to
  the with-manager scenario only).
- **Actual vs model**: each P&L month of the country's centres, placed by the centre's age
  then (from Weeks old in Vlookup); centres 18 weeks old or younger are left out. Actual =
  the typical (median) centre-month; model = the average of those model months.
"""


def _model_chart(tables, target, months, helpers, cur):
    fig = go.Figure()
    for name, t in tables.items():
        fig.add_trace(go.Scatter(
            x=t["Model month"], y=t["Cumulative"], name=name, mode="lines",
            line=dict(color=MODEL_COLOURS[name], width=3),
            customdata=[pnl.period_name(p) for p in t.index],
            hovertemplate=f"{name}<br>Month %{{x}} (%{{customdata}})<br>"
                          f"Cumulative %{{y:$,.0f}}<extra></extra>"))
    fig.add_hline(y=target, line=dict(color="#c8322f", dash="dash"),
                  annotation_text=f"Investment {_money(target)}")
    fig.add_vline(x=months, line=dict(color="#5d6d73", dash="dot"),
                  annotation_text=f"Month {months}")
    fig.update_layout(**helpers["std_layout"]("Profit added up since opening",
                                               f"{cur}", height=420))
    fig.update_xaxes(title="Months since opening")
    helpers["show_chart"](fig)


def _model_actual_chart(actual, model, helpers):
    fig = go.Figure()
    if len(actual):
        fig.add_trace(go.Scatter(
            x=actual["Model month"], y=actual["Members"], mode="markers", name="Actual centres",
            marker=dict(color="#8a9a9f", size=6, opacity=0.6),
            text=actual["Location"].str.replace(PREFIX, "", regex=False),
            hovertemplate="%{text}<br>Month %{x}: %{y:.0f} members<extra></extra>"))
    fig.add_trace(go.Scatter(x=model["Model month"], y=model["Members"], mode="lines",
                             name="Model", line=dict(color="#1f9a9a", width=3)))
    fig.update_layout(**helpers["std_layout"]("Members by centre age: actual vs model",
                                               "Members", height=380))
    fig.update_xaxes(title="Months since opening")
    helpers["show_chart"](fig)


def report_franchise_model(df_pnl, df_wm, vl, helpers, bench=None):
    """Franchise Model (admins only): a theoretical new centre's P&L from opening, with and
    without a manager, its payback month, and how actual centres compare by age."""
    st.markdown('<div class="report-title">Franchise Model</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">A new centre\'s monthly P&L from opening, with and '
                'without a campus manager, and when it pays back the investment. Built from '
                'the country\'s typical figures; change any input to test it.</div>',
                unsafe_allow_html=True)
    if st.session_state.get("access_level") != "admin":
        st.error("This report is for admins only.")
        return
    if df_pnl.empty:
        st.info("No P&L figures yet. An admin can add a month on the P&L Upload page.")
        return
    b_pnl, b_wm = bench if bench is not None else (df_pnl, df_wm)
    members = members_by_month(b_wm)
    periods = sorted(b_pnl["Period"].unique())[-pf.RECENT:]
    cent = centre_table(b_pnl, members, vl, periods, helpers["age_group"])
    countries = sorted(cent["Country"].dropna().astype(str).unique())
    if not countries:
        st.warning("No centres with a country set.")
        return
    this_month = datetime.now(ZoneInfo("Australia/Sydney")).strftime("%Y-%m")
    openings = [pf.add_months(this_month, i) for i in range(-24, 25)]

    with st.container(border=True):
        st.markdown('<div class="filter-label">Filters</div>', unsafe_allow_html=True)
        a, b = st.columns(2)
        country = a.selectbox("Country", countries, key="fm_country",
                              index=countries.index("Australia") if "Australia" in countries
                              else 0)
        start = b.selectbox("Opening month", openings, index=25, format_func=pnl.period_name,
                            key="fm_start")
    typ = pm.country_typical(cent, country)
    if typ is None:
        st.warning(f"No {country} centres with revenue in {pnl.period_name(periods[-1])}.")
        return
    cur = str(typ["peers"]["Currency"].mode().iloc[0]) if "Currency" in typ["peers"] else ""
    small = " ⚠" if pnl.small_sample(typ["centres"]) else ""
    st.caption(f"Typical figures from {typ['centres']} {'established ' if typ['established'] else ''}"
               f"{country} centre(s){small}, {pnl.period_name(periods[0])} – "
               f"{pnl.period_name(periods[-1])}.")

    # ── Settings ──
    _section("Settings", "Changing a setting resets the month-by-month changes below.")
    d, placeholder = pm.settings_for(country)
    if placeholder:
        st.warning(f"⚠ No royalty or investment figures set for {country} yet: using "
                   "Australia's as a placeholder. Change them below.")
    k = f"fm_{country}_"
    c = st.columns(4)
    s = dict(d)
    s["target_aud"] = c[0].number_input("Investment to pay back (AUD)", 0.0, value=d["target_aud"],
                                        step=5000.0, key=k + "target")
    s["payback_months"] = int(c[1].number_input("Payback target (months)", 1,
                                                value=d["payback_months"], key=k + "months"))
    amv = c[2].number_input(f"Price per member $ / week{small}", 0.0, value=round(typ["amv"], 2),
                            step=1.0, key=k + "amv")
    mature = c[3].number_input(f"Mature members{small}", 0.0, value=float(round(typ["members"])),
                               step=5.0, key=k + "mature")
    c = st.columns(4)
    s["ramp_members"] = c[0].number_input("Members at week 12", 0.0, value=d["ramp_members"],
                                          step=5.0, key=k + "ramp")
    s["growth"] = c[1].number_input("New members a month after that", 0.0, value=d["growth"],
                                    step=1.0, key=k + "growth")
    s["manager_trigger"] = c[2].number_input("Hire a manager at (members)", 0.0,
                                             value=d["manager_trigger"], step=5.0,
                                             key=k + "trigger")
    mgr_small = " ⚠" if pnl.small_sample(typ["manager_centres"]) else ""
    mgr_cost = c[3].number_input(f"Manager cost $ / month{mgr_small}", 0.0,
                                 value=float(round(typ["manager"])), step=100.0, key=k + "mgr",
                                 help=f"Typical at the {typ['manager_centres']} centre(s) that "
                                      "pay a manager.")
    c = st.columns(4)
    s["owner_on"] = c[0].toggle("Include an owner's wage", d["owner_on"], key=k + "owner_on")
    s["owner_wage"] = c[1].number_input("Owner's wage $ / year", 0.0, value=d["owner_wage"],
                                        step=5000.0, key=k + "owner", disabled=not s["owner_on"])
    s["owner_stops"] = c[2].toggle("Stop it once a manager is hired", d["owner_stops"],
                                   key=k + "owner_stops", disabled=not s["owner_on"])
    s["flat"] = c[3].toggle("Flat members (no ramp-up)", d["flat"], key=k + "flat")
    c = st.columns(5)
    s["royalty_min"] = c[0].number_input(f"Royalty minimum $ / month ({pf.fy_label(d['royalty_fy'])[:9]})",
                                         0.0, value=d["royalty_min"], step=10.0, key=k + "rmin")
    s["royalty_rise"] = c[1].number_input("Royalty minimum rise % each July", 0.0,
                                          value=d["royalty_rise"], step=0.5, key=k + "rrise")
    s["relief_weeks"] = c[2].number_input("Royalty relief (weeks)", 0.0, value=d["relief_weeks"],
                                          step=1.0, key=k + "relief")
    s["price_rise"] = c[3].number_input("Price rise % each July", 0.0, value=d["price_rise"],
                                        step=0.5, key=k + "price_rise")
    s["wage_rise"] = c[4].number_input("Wage rise % each July", 0.0, value=d["wage_rise"],
                                       step=0.5, key=k + "wage_rise",
                                       help="Applies to the manager and owner's wages.")
    typ = {**typ, "amv": amv, "members": mature, "manager": mgr_cost}
    target = s["target_aud"] / (typ["fx"] or 1.0)
    n = pm.horizon(start, s["payback_months"])
    months = pm.months_from(start, n)
    fys = sorted({pf.fy_start(p) for p in months})

    # ── Adjust the model (one tab per financial year) ──
    base = {name: pm.run(s, typ, start, n, name == pm.WITH) for name in (pm.WITH, pm.WITHOUT)}
    base_grid = pm.grid_from(base[pm.WITH])
    resets = st.session_state.setdefault("fm_resets", 0)
    gkey = f"fm_grid_{hash((country, start, tuple(sorted(s.items())), amv, mature, mgr_cost))}_{resets}"
    top = st.container()
    _section("Adjust the model",
             "Change members, price or any line's % of revenue for any month. Lines you leave "
             "alone keep their usual behaviour (see How this is worked out).")
    grid = base_grid.copy()
    pct_cols = {pm.PCT[l]: st.column_config.NumberColumn(
        GRID_LABELS[l], format="%.1f%%", min_value=0.0, max_value=200.0, step=0.1)
        for l in pm.LINES}
    for tab, fy in zip(st.tabs([pf.fy_label(f)[:9] for f in fys]), fys):
        ms = [p for p in months if pf.fy_start(p) == fy]
        shown = base_grid.loc[ms].copy()
        shown.index = [pnl.period_name(p) for p in ms]
        with tab:
            edited = st.data_editor(
                shown, key=f"{gkey}_{fy}", use_container_width=True,
                column_config={"Members": st.column_config.NumberColumn(format="%.1f",
                                                                        min_value=0.0),
                               "AMV $ / week": st.column_config.NumberColumn(
                                   "Price $ / week", format="$%.2f", min_value=0.0),
                               **pct_cols})
        edited.index = ms
        grid.loc[ms] = edited.fillna(base_grid.loc[ms]).values
    if st.button("Reset to model", icon=":material/restart_alt:", key="fm_reset"):
        st.session_state["fm_resets"] = resets + 1
        st.rerun()
    tables = {name: pm.run(s, typ, start, n, name == pm.WITH, grid, base_grid)
              for name in (pm.WITHOUT, pm.WITH)}

    # ── Headline ──
    with top:
        _section(f"Payback — investment {_money(target)} {cur} in {s['payback_months']} months",
                 "Needed to pay back on time: everything else as set (grid changes not "
                 "included).")
        for col, (name, t) in zip(st.columns(2), tables.items()):
            pb = pm.payback_month(t, target)
            ok = pb is not None and pb <= s["payback_months"]
            mem_need = pm.needed(s, typ, start, target, name == pm.WITH, "members")
            amv_need = pm.needed(s, typ, start, target, name == pm.WITH, "amv")
            when = f"Month {pb} ({pnl.period_name(t.index[pb - 1])})" if pb else \
                f"Not within {n} months"
            col.markdown(f'<div class="section-header" style="margin-top:0">{name}</div>',
                         unsafe_allow_html=True)
            a_, b_ = col.columns(2)
            a_.markdown(_card("Payback", ("✅ " if ok else "❌ ") + when), unsafe_allow_html=True)
            at = t["Cumulative"].iloc[min(s["payback_months"], n) - 1]
            b_.markdown(_card(f"Profit after {s['payback_months']} months", _money_short(at)),
                        unsafe_allow_html=True)
            a_.markdown(_card("Mature members needed",
                              "–" if mem_need is None else f"{mem_need:,.0f}"),
                        unsafe_allow_html=True)
            b_.markdown(_card("Price needed $ / week",
                              "–" if amv_need is None else f"${amv_need:,.2f}"),
                        unsafe_allow_html=True)
        _model_chart(tables, target, s["payback_months"], helpers, cur)

    # ── Monthly P&L for a financial year ──
    _section("P&L by month")
    a, b = st.columns(2)
    fy = a.selectbox("Financial year", fys, format_func=pf.fy_label, key="fm_fy")
    scen = b.radio("Scenario", [pm.WITHOUT, pm.WITH], horizontal=True, key="fm_scenario")
    t = tables[scen]
    ms = [p for p in months if pf.fy_start(p) == fy]
    cols = [(pnl.period_name(p), f"Month {int(t.at[p, 'Model month'])}", t.loc[p].to_dict())
            for p in ms]
    tot = t.loc[ms, pnl.ROW_ORDER + [pnl.GROSS_MARGIN, "EBITDA", pm.OWNER, "Profit"]] \
        .sum().to_dict()
    tot.update(Members=t.loc[ms, "Members"].mean(), Cumulative=t.at[ms[-1], "Cumulative"])
    tot["AMV"] = tot["Total Revenue"] / tot["Members"] / pnl.WEEKS_PER_MONTH / len(ms) \
        if tot["Members"] else None
    cols.append(("Total", f"{len(ms)} months", tot))
    st.markdown(side_by_side_html(cols, rows=MODEL_TABLE_ROWS), unsafe_allow_html=True)
    st.caption("Total: amounts added up over the year; members and price are the average per "
               "month; cumulative profit at the end of the year.")
    years = []
    for name, tt in tables.items():
        y = pm.fy_totals(tt)
        y.insert(0, "Scenario", name)
        y.insert(1, "Financial year", [pf.fy_label(f)[:9] for f in y.index])
        years.append(y)
    st.dataframe(pd.concat(years), hide_index=True, use_container_width=True,
                 column_config={c_: st.column_config.NumberColumn(format="%,.0f")
                                for c_ in ["Total Revenue", pnl.GROSS_MARGIN, "EBITDA", pm.OWNER,
                                           "Profit", "Members (end)", "Cumulative"]})

    # ── Benchmarks ──
    _section("Benchmark % for each line")
    values = t.loc[ms, pnl.ROW_ORDER + [pnl.GROSS_MARGIN, "EBITDA", "Members"]].mean()
    values["Country"] = country
    _forecast_benchmarks(values, typ["peers"], f"{typ['centres']} {country} centres", min_n=1,
                         label="Model")

    # ── Actual vs model ──
    _section("Actual vs model, by centre age",
             "Typical actual centre-month against the model's average for the same ages.")
    info = location_info(b_pnl, vl)
    weeks = pd.to_numeric(vl.set_index(LOC)["Weeks old"], errors="coerce").dropna().to_dict() \
        if "Weeks old" in vl.columns and LOC in vl.columns else {}
    actual = pm.actual_by_age(pnl.location_months(b_pnl), members, weeks,
                              info["Country"].to_dict(), country, pd.Timestamp.now())
    if actual.empty:
        st.caption(f"No {country} centres older than 18 weeks with P&L and Weeks old.")
    else:
        cmp = pm.compare(actual, t)
        cmp["Centres"] = [f"{n_}" + (" ⚠" if pnl.small_sample(n_) else "") for n_ in cmp["Centres"]]
        money = {c_: st.column_config.NumberColumn(format="%,.0f") for c_ in cmp.columns
                 if c_.startswith(("Members", "Revenue"))}
        pcts = {c_: st.column_config.NumberColumn(format="%.1f%%") for c_ in cmp.columns
                if "%" in c_}
        st.dataframe(cmp, hide_index=True, use_container_width=True,
                     column_config={**money, **pcts})
        _model_actual_chart(actual, t, helpers)
    st.info("A model is an estimate from typical figures, not a promise. Check it against what "
            "you know: site, rent, staffing and local demand.")
    with st.expander("How this is worked out"):
        st.markdown(_md(HOW_MODEL))
