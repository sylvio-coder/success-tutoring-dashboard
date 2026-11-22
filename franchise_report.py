"""Franchise Model report (admins only): a new centre's 5-year plan from the uploaded
assumptions workbook, for one model (AU National or a state) and opening month, with Worst /
Base / Best. Sections: Summary, Investor Metrics, Your operating plan, Road to profitability,
Profit and cash by year, Risks, Assumptions used. Calculations are in franchise_plan.py."""
import html
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import franchise_inputs as fin
import franchise_plan as fp

TEAL, RED, GREEN, TEXT, GREY = "#1f9a9a", "#c8322f", "#2e8540", "#1c2a30", "#5d6d73"
CASE_COLOURS = {"Worst": RED, "Base": TEAL, "Best": GREEN}
DISCLAIMER = ("This plan is an estimate from typical figures and stated assumptions, not a "
              "promise or prediction of earnings. Results vary by site, rent, staffing and "
              "local demand. Prospective owners should seek independent financial, legal and "
              "accounting advice.")

CSS = """<style>
.fm-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin:4px 0 8px}
.fm-card{background:#fff;border:1px solid #dde5e6;border-top:4px solid var(--c);border-radius:10px;padding:14px 16px}
.fm-card .nm{font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--c)}
.fm-card .tag{float:right;font-size:12px;font-weight:700;border-radius:12px;padding:2px 9px;color:var(--t);background:var(--b)}
.fm-card .big{font-size:26px;font-weight:700;color:#1c2a30;margin-top:6px}
.fm-card .sub{font-size:12.5px;color:#5d6d73}
.fm-card dl{display:grid;grid-template-columns:1fr auto;gap:4px 10px;margin:12px 0 0;font-size:13px}
.fm-card dt{color:#5d6d73} .fm-card dd{margin:0;font-weight:600;color:#1c2a30;text-align:right}
.fm-tbl{width:100%;border-collapse:collapse;font-size:13px;background:#fff;border:1px solid #dde5e6}
.fm-tbl th{background:#1c2a30;color:#fff;font-weight:600;padding:6px 10px;text-align:right}
.fm-tbl th:first-child,.fm-tbl td:first-child{text-align:left}
.fm-tbl td{padding:4px 10px;border-bottom:1px solid #eef1f3;text-align:right}
.fm-tbl tr.grp td{background:#f2f8f8;color:#1f9a9a;font-weight:700;text-align:left}
.fm-tbl td.neg{color:#c8322f}
.fm-wrap{overflow-x:auto;margin-bottom:8px}
.fm-wrap .fm-tbl td:first-child,.fm-wrap .fm-tbl th:first-child{position:sticky;left:0;z-index:1;min-width:230px}
.fm-wrap .fm-tbl td:first-child{background:#fff}
.fm-road{list-style:none;margin:0;padding:0 0 0 14px;border-left:2px solid #dde5e6}
.fm-road li{position:relative;padding:0 0 9px 12px}
.fm-road li:before{content:"";position:absolute;left:-21px;top:4px;width:10px;height:10px;border-radius:50%;background:#1f9a9a;border:2px solid #fff}
.fm-road li.goal:before{background:#2e8540} .fm-road li.never:before{background:#c8322f}
.fm-road .m{font-size:12px;color:#5d6d73} .fm-road .e{font-size:14px;font-weight:600;color:#1c2a30}
.fm-warn{background:#fff4e5;border-left:4px solid #e0a100;padding:8px 12px;margin:6px 0;font-size:13.5px;color:#1c2a30;border-radius:4px}
@media print{
 [data-testid="stSidebar"],[data-testid="stHeader"],[data-testid="stToolbar"],.no-print,
 [data-testid="stExpander"] summary svg{display:none!important}
 .fm-cards{grid-template-columns:repeat(3,1fr)}
 .fm-wrap{overflow:visible}
}
</style>"""


# ── Formatting ────────────────────────────────────────────────────────────────
def money(v, short=False):
    if v is None or pd.isna(v):
        return "–"
    if short and abs(v) >= 1000:
        s = f"${abs(v) / 1000:,.0f}k"
    else:
        s = f"${abs(v):,.0f}"
    return f"-{s}" if v < -0.5 else s


def month_label(start, m):
    if m is None:
        return "Not within 5 years"
    d = fp.week_date(start, int((m - 1) * fp.WPM) + 1)      # the month's first week
    return f"Month {m} · from {d.day} {d:%b %Y}"


def section(title, caption=None):
    st.markdown(f'<div class="section-header">{html.escape(title)}</div>', unsafe_allow_html=True)
    if caption:
        st.caption(caption)


def table_html(head, rows, wrap=False):
    """rows: [(label, [cells], kind)] with kind 'grp' (group heading) or 'row'; cells are
    (text, negative?)."""
    th = "".join(f"<th>{html.escape(h)}</th>" for h in head)
    body = []
    for label, cells, kind in rows:
        if kind == "grp":
            body.append(f'<tr class="grp"><td colspan="{len(head)}">{html.escape(label)}</td></tr>')
            continue
        tds = "".join(f'<td class="{"neg" if neg else ""}">{html.escape(txt)}</td>'
                      for txt, neg in cells)
        body.append(f"<tr><td>{html.escape(label)}</td>{tds}</tr>")
    t = f'<table class="fm-tbl"><tr>{th}</tr>{"".join(body)}</table>'
    return (f'<div class="fm-wrap">{t}</div>' if wrap else t).replace("$", "&#36;")


# ── Running the model (cached) ────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def compute(rows, model, start, case, extras=False):
    """Run one case. extras: also thresholds, sensitivity and the opening-month table."""
    inputs = fin.model_inputs(rows, model)
    s = fp.settings(inputs, case)
    t = fp.run(s, start)
    out = dict(s=s, t=t, start=start, summary=fp.summary(t, s), cash=fp.cash_flow(t, s)[0],
               years=fp.year_totals(t), checkpoints=fp.checkpoints(t),
               milestones=fp.milestones(t, s))
    if extras:
        out["thresholds"] = fp.thresholds(s, start)
        out["sensitivity"] = fp.sensitivity(s, start)
        opens = []
        for i in range(12):
            y, m = divmod(start.month - 1 + i, 12)
            d = date(start.year + y, m + 1, 1)
            tt = fp.run(s, d)
            opens.append({"Opening month": f"{d:%B %Y}", "Payback": fp.payback_month(tt, s),
                          "5-year profit": float(tt["Profit"].sum())})
        out["openings"] = pd.DataFrame(opens)
    return out


# ── Sections ──────────────────────────────────────────────────────────────────
def verdict(sm):
    pay, target = sm["payback"], sm["payback_target"]
    if pay is None:
        return "Not paid back in 5 years", RED, "#fbe9e8"
    if pay <= target:
        return "On target", GREEN, "#e7f3ea"
    return f"{pay - target} months late", RED, "#fbe9e8"


def summary_cards(res):
    cards = []
    for case, r in res.items():
        sm, t = r["summary"], r["t"]
        tag, tc, tb = verdict(sm)
        big = "Never" if sm["payback"] is None else f"Month {sm['payback']}"
        sub = ("payback (target month " + str(sm["payback_target"]) + ")") if sm["payback"] is None \
            else f"{t.loc[sm['payback'], 'Date']:%B %Y} · target month {sm['payback_target']}"
        rows = [("Total cash needed", money(sm["total_cash"])),
                ("5-year profit (after salary)", money(sm["profit_5y"])),
                ("Owner's salary, 5 years", money(sm["owner_5y"])),
                ("<b>Owner's total return, 5 years</b>",
                 f"<b>{money(sm['owner_5y'] + sm['profit_5y'])}</b>"),
                ("Average owner's return a year",
                 money((sm["owner_5y"] + sm["profit_5y"]) / 5)),
                ("First month in profit", "–" if sm["first_profit"] is None
                 else f"Month {sm['first_profit']}"),
                ("Yearly return (IRR)", "–" if sm["irr"] is None else f"{sm['irr'] * 100:.0f}%"),
                ("Members after 5 years", f"{t['Members end'].iloc[-1]:,.0f}")]
        esc = lambda v: v if v.startswith("<b>") else html.escape(v)
        dl = "".join(f"<dt>{esc(a)}</dt><dd>{esc(b)}</dd>" for a, b in rows)
        cards.append(f'<div class="fm-card" style="--c:{CASE_COLOURS[case]};--t:{tc};--b:{tb}">'
                     f'<span class="tag">{html.escape(tag)}</span><div class="nm">{case}</div>'
                     f'<div class="big">{html.escape(big)}</div><div class="sub">{html.escape(sub)}'
                     f'</div><dl>{dl}</dl></div>')
    st.markdown(f'<div class="fm-cards">{"".join(cards)}</div>'.replace("$", "&#36;"),
                unsafe_allow_html=True)


def investor_metrics(r):
    s, t, sm, th = r["s"], r["t"], r["summary"], r["thresholds"]
    cp = r["checkpoints"]
    rows = [(f"Month {int(x['By month'])}", [(f"{x['Members']:,.0f}", False),
                                             (money(x["Revenue a month"]), False),
                                             (f"{x['Tutor wages % of fees']:.0f}%" if pd.notna(
                                                 x["Tutor wages % of fees"]) else "–", False),
                                             (money(x["Profit so far"]), x["Profit so far"] < 0)],
             "row") for _, x in cp.iterrows()]
    st.markdown("**Checkpoints: are you on plan?**")
    st.markdown(table_html(["By", "Members", "Revenue a month", "Tutor wages (% of fees)",
                            "Profit so far (after salary)"], rows), unsafe_allow_html=True)
    target = sm["payback_target"]
    st.markdown(f"**To recover the investment within {target} months, you need any ONE of:**")
    need = []
    k = th["members_scale"]
    if k:
        tg = ", ".join(f"{x * k:,.0f}" for x in s["targets"])
        need.append(("More members", f"{(k - 1) * 100:+.0f}% on plan: {s['t12'] * k:,.0f} at week 12; "
                                     f"year ends {tg}"))
    else:
        need.append(("More members", "not enough on its own"))
    need.append(("Standard price", "not enough on its own" if th["price"] is None else
                 f"at least ${th['price']:,.2f} a week (now ${s['standard']:,.2f})"))
    if s["owner_on"]:
        need.append(("Owner's salary", "not enough on its own" if th["owner_wage"] is None else
                     f"at most {money(th['owner_wage'])} a year (now {money(s['owner_wage'])})"))
    need.append(("Rent", "not enough on its own" if th["rent"] is None else
                 f"at most {money(th['rent'])} a month (now {money(s['rent'])})"))
    if sm["payback"] is not None and sm["payback"] <= target:
        st.caption(f"Already on target: payback in month {sm['payback']}. The figures below "
                   "are how far each could slip before payback passes the target.")
    st.markdown(table_html(["Change (everything else as planned)", "Needed"],
                           [(a, [(b, False)], "row") for a, b in need]), unsafe_allow_html=True)
    warns = []
    if sm["ads_gap"] > 100:
        y1 = t.loc[4:15, "Advertising needed"].mean()
        warns.append(f"Advertising: the plan budgets about {money(t.loc[4:15, 'Advertising planned'].mean())} "
                     f"a month; reaching the member targets is likely to need about {money(y1)} a "
                     "month (enquiries needed × cost per lead).")
    if sm["full_month"]:
        warns.append(f"Capacity: the centre is full in month {sm['full_month']} "
                     f"({sm['capacity']:,.0f} members). Plan more tables or slots before then.")
    for w in warns:
        st.markdown(f'<div class="fm-warn">⚠ {html.escape(w)}</div>'.replace("$", "&#36;"),
                    unsafe_allow_html=True)


def operating_plan(r):
    s, t = r["s"], r["t"]
    def yr(col, how="sum"):
        parts = [t.loc[(y - 1) * 12 + 1:y * 12, col] for y in range(1, 6)]
        return [p_.iloc[-1] if how == "last" else p_.agg(how) for p_ in parts]
    pct = lambda num, den: [n / d * 100 if d else None for n, d in zip(yr(num), yr(den))]
    head = ["Plan by year", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]
    f0 = lambda vals: [(f"{v:,.0f}", False) for v in vals]
    f1 = lambda vals: [(f"{v:,.1f}", False) for v in vals]
    fm = lambda vals: [(money(v), v < -0.5) for v in vals]
    fp_ = lambda vals: [("–" if v is None else f"{v:.0f}%", False) for v in vals]
    fd = lambda vals: [(f"${v:,.2f}", False) for v in vals]
    rows = [("1. Members", [], "grp"),
            ("Members at year end", f0(yr("Members end", "last")), "row"),
            ("New members a month", f1([v / 12 for v in yr("New members")]), "row"),
            ("Cancelled a month", f1([v / 12 for v in yr("Cancelled")]), "row"),
            ("Cancelled % a month", [(f"{c * 100:.1f}%", False) for c in s["cancel"]], "row"),
            ("2. Enquiries and marketing", [], "grp"),
            ("Enquiries needed a month", f0([v / 12 for v in yr("Enquiries needed")]), "row"),
            ("Student Assessments a month", f0([v / 12 * s["enquiry_to_assessment"]
                                               for v in yr("Enquiries needed")]), "row"),
            ("Advertising needed a month", fm([v / 12 for v in yr("Advertising needed")]), "row"),
            ("Advertising planned a month", fm([v / 12 for v in yr("Advertising planned")]), "row"),
            ("3. Price", [], "grp"),
            ("Average price banked a week", fd(yr("Price banked", "mean")), "row"),
            ("Revenue per member a week", fd([r_ / max(m, 1) / 52 for r_, m in
                                              zip(yr("Revenue"), yr("Members end", "mean"))]), "row"),
            ("4. Study hubs", [], "grp"),
            ("Study hubs a week", f0(yr("Study hubs", "mean")), "row"),
            ("Students per hub (plan)", f1([s["fill"][3 + y] if y else sum(s["fill"][:4]) / 4
                                            for y in range(5)]), "row"),
            ("Hubs after 7 pm a week", f0(yr("Hubs after 7 pm", "mean")), "row"),
            ("Tables needed at peak", f0(yr("Tables at peak", "max")), "row"),
            ("Capacity used (year end)", fp_(yr("Capacity used %", "last")), "row"),
            ("5. Costs (% of revenue)", [], "grp"),
            ("Tutor wages", fp_(pct("Tutor wages", "Revenue")), "row"),
            ("Rent", fp_(pct("Rent", "Revenue")), "row"),
            ("Royalties", fp_(pct("Royalties", "Revenue")), "row"),
            ("Marketing", fp_(pct("Marketing", "Revenue")), "row"),
            ("Manager", fp_(pct("Manager", "Revenue")), "row"),
            ("Running costs and CRM", fp_([(a + b) / c * 100 if c else None for a, b, c in
                                          zip(yr("Running costs"), yr("CRM"), yr("Revenue"))]),
             "row"),
            ("6. Owner and cash", [], "grp"),
            ("Owner's salary", fm(yr("Owner's wage")), "row"),
            ("Profit after salary", fm(yr("Profit")), "row"),
            ("Cash after loan and tax*", fm([r["cash"].loc[(y - 1) * 12 + 1:y * 12, "Cash"].sum()
                                            for y in range(1, 6)]), "row")]
    st.markdown(table_html(head, rows, wrap=True), unsafe_allow_html=True)
    sm = r["summary"]
    pay = "not within 5 years" if sm["payback"] is None else f"month {sm['payback']}"
    irr = "–" if sm["irr"] is None else f"{sm['irr'] * 100:.0f}%"
    st.caption(f"Total cash needed {money(sm['total_cash'])}: investment {money(sm['investment'])}"
               f" + bond {money(sm['bond'])} + working capital {money(sm['working_capital'])} "
               f"(peak cash need plus {s['wc_buffer'] * 100:.0f}%). Return: payback {pay}, "
               f"NPV {money(sm['npv'])}, IRR {irr} (with a Year 5 sale at "
               f"{s['sale_multiple']:g}× yearly profit). * Year 5 cash includes the bond back and "
               "the sale value.".replace("$", "&#36;"))


def road(r, helpers):
    t = r["t"]
    items = []
    for m, label in r["milestones"]:
        cls = "never" if m is None else "goal" if label == "Investment paid back" else ""
        items.append(f'<li class="{cls}"><div class="m">{html.escape(month_label(r["start"], m))}</div>'
                     f'<div class="e">{html.escape(label)}</div></li>')
    st.markdown(f'<ol class="fm-road">{"".join(items)}</ol>', unsafe_allow_html=True)
    costs = t["Revenue"] - t["Profit"]
    x = list(t.index)
    names = [f"{d:%b %Y}" for d in t["Date"]]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x, y=t["Profit"], name="Profit (after salary)", customdata=names,
                         marker_color=[GREEN if v >= 0 else RED for v in t["Profit"]],
                         hovertemplate="Month %{x} (%{customdata})<br>Profit %{y:$,.0f}<extra></extra>"))
    fig.add_trace(go.Scatter(x=x, y=t["Revenue"], name="Revenue", mode="lines", customdata=names,
                             line=dict(color=TEAL, width=3),
                             hovertemplate="Month %{x} (%{customdata})<br>Revenue %{y:$,.0f}<extra></extra>"))
    fig.add_trace(go.Scatter(x=x, y=costs, name="Costs (incl. salary)", mode="lines",
                             customdata=names, line=dict(color=GREY, width=2, dash="dot"),
                             hovertemplate="Month %{x} (%{customdata})<br>Costs %{y:$,.0f}<extra></extra>"))
    fig.update_layout(**helpers["std_layout"]("Revenue, costs and profit each month", "", height=360))
    fig.update_xaxes(title="Months since the presale starts")
    helpers["show_chart"](fig)


def profit_and_cash(r):
    y, c = r["years"], r["cash"]
    head = ["", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "5 years"]
    tot = lambda col: list(y[col]) + [y[col].sum()]
    cash = [c.loc[(i - 1) * 12 + 1:i * 12, "Cash"].sum() for i in range(1, 6)]
    rows = [("Members at year end", [(f"{v:,.0f}", False) for v in y["Members (end)"]] +
             [(f"{y['Members (end)'].iloc[-1]:,.0f}", False)], "row")]
    for label, col in [("Revenue", "Revenue"), ("Tutor wages", "Tutor wages"),
                       ("Royalties", "Royalties"), ("Profit before salary (EBITDA)", "EBITDA"),
                       ("Owner's salary", "Owner's wage"), ("Profit after salary", "Profit")]:
        rows.append((label, [(money(v), v < -0.5) for v in tot(col)], "row"))
    ret = [o + p for o, p in zip(tot("Owner's wage"), tot("Profit"))]
    rows.append(("Owner's total return (salary + profit)", [(money(v), v < -0.5) for v in ret],
                 "row"))
    rows.append(("Cash after loan and tax*", [(money(v), v < -0.5) for v in cash + [sum(cash)]],
                 "row"))
    rows.append(("Profit added up since opening", [(money(v), v < -0.5) for v in y["Cumulative"]] +
                 [(money(y["Cumulative"].iloc[-1]), y["Cumulative"].iloc[-1] < 0)], "row"))
    st.markdown(table_html(head, rows, wrap=True), unsafe_allow_html=True)
    st.caption("* Year 5 includes the bond coming back and selling the business at the end of "
               "Year 5 (sale value from the assumptions).")
    with st.expander("All 60 months"):
        t = r["t"]
        mhead = ["Line"] + [f"{d:%b %y}" for d in t["Date"]]
        lines = [("Members (end)", "Members end", "n"), ("New members", "New members", "n1"),
                 ("Price banked a week", "Price banked", "d"), ("Revenue", "Revenue", "m"),
                 ("Tutor wages", "Tutor wages", "m"), ("Rent", "Rent", "m"),
                 ("Manager", "Manager", "m"), ("Royalties", "Royalties", "m"),
                 ("Marketing", "Marketing", "m"), ("Running costs and CRM", None, "m"),
                 ("Profit before salary", "EBITDA", "m"), ("Owner's salary", "Owner's wage", "m"),
                 ("Profit after salary", "Profit", "m"), ("Profit added up", "Cumulative", "m")]
        mrows = []
        for label, col, kind in lines:
            vals = t["Running costs"] + t["CRM"] if col is None else t[col]
            fmt = {"n": lambda v: f"{v:,.0f}", "n1": lambda v: f"{v:,.1f}",
                   "d": lambda v: f"${v:,.2f}", "m": money}[kind]
            mrows.append((label, [(fmt(v), kind == "m" and v < -0.5) for v in vals], "row"))
        st.markdown(table_html(mhead, mrows, wrap=True), unsafe_allow_html=True)


def risks(r):
    sens = r["sensitivity"]
    pay = lambda v: "Never" if v is None or pd.isna(v) else f"Month {int(v)}"
    rows = [(x["Input"], [(pay(x["Payback (10% worse)"]), False),
                          (money(x["5-year profit (10% worse)"]), x["5-year profit (10% worse)"] < 0),
                          (pay(x["Payback (10% better)"]), False),
                          (money(x["5-year profit (10% better)"]), x["5-year profit (10% better)"] < 0)],
             "row") for _, x in sens.iterrows()]
    st.markdown("**What moves the result most** (each input 10% worse or better, everything else "
                "as planned)")
    st.markdown(table_html(["Input", "Payback (10% worse)", "5-year profit (10% worse)",
                            "Payback (10% better)", "5-year profit (10% better)"], rows),
                unsafe_allow_html=True)
    op = r["openings"]
    st.markdown("**Opening month** (school holidays and seasonality make some openings better)")
    st.markdown(table_html(["Opening month", "Payback", "5-year profit"],
                           [(x["Opening month"], [(pay(x["Payback"]), False),
                                                  (money(x["5-year profit"]), x["5-year profit"] < 0)],
                             "row") for _, x in op.iterrows()]), unsafe_allow_html=True)


def assumptions(rows, model, case):
    inp = fin.model_inputs(rows, model)
    out, last = [], None
    for v, d in inp.items():
        if str(d.get(case, "")).lower() == "calc" or v.startswith("Check only"):
            continue
        if d["Section"] != last:
            out.append((d["Section"] or "Other", [], "grp"))
            last = d["Section"]
        val = d[case]
        txt = f"{val:,.2f}".rstrip("0").rstrip(".") if isinstance(val, float) else str(val)
        out.append((v, [(f"{txt} {d['Unit']}".strip(), False), (str(d["Basis"]), False)], "row"))
    st.markdown(table_html(["Input", f"{case} value", "Basis"], out), unsafe_allow_html=True)


# ── The page ──────────────────────────────────────────────────────────────────
def report_franchise_model(versions, helpers):
    st.markdown('<div class="report-title">Franchise Model</div>', unsafe_allow_html=True)
    st.markdown('<div class="report-subtitle">A new centre\'s 5-year plan from the uploaded '
                'assumptions: payback, cash needed, the road to profit and what an investor '
                'needs to hit.</div>', unsafe_allow_html=True)
    if st.session_state.get("access_level") != "admin":
        st.error("This report is for admins only.")
        return
    ver, rows, when = fin.latest(versions)
    if not ver:
        st.info("No assumptions uploaded yet. Upload the workbook on Admin → Franchise Model "
                "Upload.")
        return
    ready = [m for m, p in fin.readiness(rows).items() if not p]
    if not ready:
        st.warning("None of the uploaded models is ready yet. Check the upload page for what to fix.")
        return
    st.markdown(CSS, unsafe_allow_html=True)
    today = datetime.now(ZoneInfo("Australia/Sydney")).date()
    openings = [date(today.year + (today.month - 1 + i) // 12, (today.month - 1 + i) % 12 + 1, 1)
                for i in range(1, 25)]
    with st.container(border=True):
        a, b, c = st.columns(3)
        model = a.selectbox("Model", ready, key="fr_model",
                            index=ready.index("AU National") if "AU National" in ready else 0)
        start = b.selectbox("Presale starts", openings, format_func=lambda d: f"{d:%B %Y}",
                            key="fr_start")
        case = c.radio("Case for the detail", list(fin.CASES), index=1, horizontal=True,
                       key="fr_case")
        st.caption(f"Assumptions: upload {ver} from {when} · model {model}. Print or save as PDF "
                   "from the browser for an investor copy.")
    with st.spinner("Working out the plan..."):
        res = {cs: compute(rows, model, start, cs, extras=cs == case) for cs in fin.CASES}
    r = res[case]
    section("1. Summary", "Worst, Base and Best side by side. Payback is when profit after the "
            "owner's salary, added up since opening, covers the investment.")
    summary_cards(res)
    section(f"2. Investor Metrics · {case}", "What needs to happen to reach the target return.")
    investor_metrics(r)
    section(f"3. Your operating plan · {case}", "The plan figures to commit to and track each "
            "month.")
    operating_plan(r)
    section(f"4. Road to profitability · {case}")
    road(r, helpers)
    section(f"5. Profit and cash by year · {case}")
    profit_and_cash(r)
    section(f"6. Risks · {case}")
    risks(r)
    section(f"7. Assumptions used · {case}", f"From upload {ver} ({when}), model {model}.")
    assumptions(rows, model, case)
    st.info(DISCLAIMER)
