"""P&L Centre Forecast: calculations, the admin-only page and its menu entry."""
import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

import pnl
import pnl_forecast as pf
import pnl_reports as pr

LOC = pr.LOC
A = "Success Tutoring - Auburn"


def months_data(rows):
    """Centre months: {period: (members, revenue, tutor, rent, royalties, marketing)}."""
    out = pd.DataFrame({p: {"Members": m, "Total Revenue": r, "Tutor Wages + Super": t,
                            "Rent": rent, "Wages - Manager/Admin + Super": 0.0,
                            "Royalties": roy, "Marketing - Agency Fees": 0.0,
                            "Marketing": mk, "OPEX": 1000.0}
                        for p, (m, r, t, rent, roy, mk) in rows.items()}).T
    costs = pf.COST_LINES
    out["EBITDA"] = out["Total Revenue"] - out[costs].sum(axis=1)
    return out


DATA = months_data({"2026-07": (80, 28000, 7000, 4000, 2240, 2000),
                    "2026-08": (80, 29000, 7250, 4200, 2320, 2200),
                    "2026-09": (81, 29026, 7995, 4167, 2322, 2333)})


def test_months_and_financial_years():
    assert pf.add_months("2026-11", 3) == "2027-02"
    assert pf.fy_start("2026-07") == 2026 and pf.fy_start("2026-06") == 2025
    assert pf.fy_months(2026)[0] == "2026-07" and pf.fy_months(2026)[-1] == "2027-06"
    assert pf.fy_label(2026) == "FY2026–27 (Jul 2026 – Jun 2027)"
    assert pf.window("2026-09", "Financial year") == pf.fy_months(2026)
    assert pf.window("2026-06", "Financial year") == pf.fy_months(2026)   # next month's FY
    assert pf.window("2026-09", "Next 12 months") == [pf.add_months("2026-09", i + 1)
                                                      for i in range(12)]


def test_starting_point_is_the_last_three_months():
    s = pf.starting_point(DATA)
    assert s["months"] == ["2026-07", "2026-08", "2026-09"]
    rev = 28000 + 29000 + 29026
    assert s["pct"]["Tutor Wages + Super"] == pytest.approx((7000 + 7250 + 7995) / rev)
    assert s["pct"]["Marketing"] == pytest.approx((2000 + 2200 + 2333) / rev)
    assert s["fixed"]["Rent"] == pytest.approx((4000 + 4200 + 4167) / 3)
    amv = (28000 / 80 + 29000 / 80 + 29026 / 81) / 3 / pnl.WEEKS_PER_MONTH
    assert s["amv"] == pytest.approx(amv)
    one = pf.starting_point(DATA.tail(1))                 # fewer months: what there is
    assert one["months"] == ["2026-09"]


def series(values, start="2025-01"):
    return pd.Series(values, index=[pf.add_months(start, i) for i in range(len(values))],
                     dtype=float)


def test_members_known_months_then_trend_that_fades():
    s = series([100, 102, 104.04, 106.12])               # +2% a month, no seasonal pattern
    f = pf.forecast_members(s, ["2025-03", "2025-04", "2025-05", "2025-06"], {})
    assert f["2025-03"] == 104.04 and f["2025-04"] == 106.12        # membership data used
    assert f["2025-05"] == pytest.approx(106.12 * (1 + 0.02 * pf.DAMP), rel=1e-3)
    assert f["2025-06"] == pytest.approx(f["2025-05"] * (1 + 0.02 * pf.DAMP ** 2), rel=1e-3)
    flat = pf.forecast_members(series([50, 50, 50]), ["2025-05"], {"05": -0.10})
    assert flat["2025-05"] == pytest.approx(45)          # the seasonal dip only
    fast = pf.forecast_members(series([10, 20, 40]), ["2025-04"], {})
    assert fast["2025-04"] == pytest.approx(40 * (1 + pf.MAX_TREND * pf.DAMP))   # capped
    assert pf.forecast_members(pd.Series(dtype=float), ["2025-04"], {}) == {"2025-04": None}


def test_network_pattern_is_typical_change_less_overall():
    rows = {}
    for i in range(6):                                   # six centres, January dips 10%
        loc = f"C{i}"
        vals = [100, 90, 90, 90] if i < 5 else [100, 90, 90, 200]   # one jump left out
        for j, v in enumerate(vals):
            rows[(loc, pf.add_months("2025-12", j))] = v
    members = pd.Series(rows)
    p = pf.network_pattern(members)
    assert p["01"] == pytest.approx(-0.10)
    assert p["02"] == pytest.approx(0.0)
    assert "03" not in p or p["03"] == pytest.approx(0.0)
    small = pd.Series({(l, m): v for (l, m), v in rows.items() if l == "C0"})
    assert pf.network_pattern(small) == {}               # too few centres to trust


def test_own_pattern_needs_two_years():
    assert pf.own_pattern(series([100] * 23)) is None
    s = series([100, 90, 100] * 9)                       # 27 months
    assert pf.own_pattern(s) is not None


def test_defaults_and_fixed_lines_stay_fixed():
    s = pf.starting_point(DATA)
    months = ["2026-10", "2026-11"]
    members = {"2026-10": 80.0, "2026-11": 100.0}
    base = pf.defaults(months, members, s, 2200)
    assert list(base.columns) == pf.INPUTS
    out = pf.compute(base, base, s, 2200)
    # untouched rent stays at its $ even though members change
    assert out.loc["2026-10", "Rent"] == pytest.approx(s["fixed"]["Rent"])
    assert out.loc["2026-11", "Rent"] == pytest.approx(s["fixed"]["Rent"])
    rev = 100 * s["amv"] * pnl.WEEKS_PER_MONTH
    assert out.loc["2026-11", "Total Revenue"] == pytest.approx(rev)
    assert out.loc["2026-11", "Tutor Wages + Super"] == pytest.approx(
        s["pct"]["Tutor Wages + Super"] * rev)
    assert out.loc["2026-11", "Royalties"] == pytest.approx(max(2200, 0.08 * rev))
    costs = out.loc["2026-11", pf.COST_LINES].sum()
    assert out.loc["2026-11", "EBITDA"] == pytest.approx(rev - costs)


def test_changed_line_uses_its_percent_of_revenue():
    s = pf.starting_point(DATA)
    base = pf.defaults(["2026-10"], {"2026-10": 80.0}, s, 2200)
    edited = base.copy()
    edited.loc["2026-10", pf.PCT["Rent"]] = 10.0
    edited.loc["2026-10", pf.PCT["Tutor Wages + Super"]] = 30.0
    edited.loc["2026-10", "Members"] = 90.0
    out = pf.compute(edited, base, s, 2200).loc["2026-10"]
    rev = 90 * s["amv"] * pnl.WEEKS_PER_MONTH
    assert out["Rent"] == pytest.approx(0.10 * rev)
    assert out["Tutor Wages + Super"] == pytest.approx(0.30 * rev)


def test_price_rise_and_royalty_relief():
    s = pf.starting_point(DATA)
    months = ["2026-12", "2027-01", "2027-02"]
    members = {p: 80.0 for p in months}
    base = pf.defaults(months, members, s, 2200, relief_until="2026-12",
                       rise_from="2027-01", rise_pct=5)
    assert base.loc["2026-12", "AMV $ / week"] == pytest.approx(s["amv"])
    assert base.loc["2027-01", "AMV $ / week"] == pytest.approx(s["amv"] * 1.05)
    out = pf.compute(base, base, s, 2200, relief_until="2026-12")
    assert out.loc["2026-12", "Royalties"] == 0                  # relief: $0 is normal
    assert out.loc["2027-01", "Royalties"] >= 2200


def test_break_even_members_matches_the_month():
    s = pf.starting_point(DATA)
    base = pf.defaults(["2026-10"], {"2026-10": 81.0}, s, 2200)
    out = pf.compute(base, base, s, 2200)
    be = out.loc["2026-10", "BE members"]
    # at the break-even members (costs left at their defaults), EBITDA is 0
    at_be_base = pf.defaults(["2026-10"], {"2026-10": be}, s, 2200)
    at_be = pf.compute(at_be_base, at_be_base, s, 2200)
    assert at_be.loc["2026-10", "EBITDA"] == pytest.approx(0, abs=1)
    assert 0 < be < 81


def test_months_table_flags_and_totals():
    s = pf.starting_point(DATA)
    months = pf.fy_months(2026)
    fc_months = [p for p in months if p > "2026-09"]
    members = {p: 81.0 for p in fc_months}
    base = pf.defaults(fc_months, members, s, 2200)
    fc = pf.compute(base, base, s, 2200)
    t = pf.months_table(DATA, fc, months, seen_months={"07", "08", "09", "10"})
    assert list(t.index) == months
    assert (t.loc[["2026-07", "2026-08", "2026-09"], "Basis"] == pf.ACTUAL).all()
    assert t.loc["2026-10", "Basis"] == pf.FORECAST          # has an October in its history
    assert t.loc["2026-11", "Basis"] == pf.ESTIMATE          # never had a November
    tot = pf.totals(t)
    assert tot["Months"] == 12 and tot["Forecast months"] == 9 and tot["Estimate months"] == 8
    assert tot["Revenue"] == pytest.approx(t["Total Revenue"].sum())
    assert tot["End members"] == 81
    assert t.loc["2026-09", "BE members"] > 0                 # actual months too
    v = pf.benchmark_values(t, "Australia")
    assert v["Total Revenue"] == pytest.approx(fc["Total Revenue"].mean())
    gap = pf.months_table(DATA.drop("2026-08"), fc, months, set())
    assert gap.loc["2026-08", "Basis"] == pf.MISSING
    assert pf.totals(gap)["Months"] == 11


# ── The page ──────────────────────────────────────────────────────────────────
NAMES = ["Auburn", "Burwood", "Castle Hill", "Dural", "Epping", "Frenchs Forest", "Gordon"]
P_MONTHS = ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09"]


def _lines(loc, period, amounts):
    return [{"Period": period, "Location": loc, "Country": "Australia", "Currency": "AUD",
             "Status": "Submitted", "Account code": code, "Account": str(code),
             "P&L section": pnl.section_of(code), "Amount": amt, "FX rate to AUD": 1.0}
            for code, amt in amounts.items()]


def _page_data():
    rows, wm = [], []
    for i, n in enumerate(NAMES):
        loc = f"Success Tutoring - {n}"
        members = 70 + 5 * i
        for j, p in enumerate(P_MONTHS):
            rev = (members + j) * 85 * 4.2
            roy = 0 if n == "Auburn" else max(2200, 0.08 * rev)     # Auburn: royalty relief
            rows += _lines(loc, p, {200: rev, 472: 0.27 * rev, 469: 4000 + 100 * i, 467: roy,
                                    400: 0.08 * rev, 453: 3500, 471: 0 if i % 2 else 2500})
        for d in pd.date_range("2025-01-05", "2026-10-25", freq="W-SUN"):
            wm.append({LOC: loc, "Date": d, "# Active members": members + (d.month % 3)})
    vl = pd.DataFrame({LOC: [f"Success Tutoring - {n}" for n in NAMES],
                       "Region": ["New South Wales"] * 7, "GPM": ["A"] * 7,
                       "Country": ["Australia"] * 7, "Age (Months)": [30] * 7,
                       "Royalty Min": [""] * 7})
    return pd.DataFrame(rows), pd.DataFrame(wm), vl


PAGE_PNL, PAGE_WM, PAGE_VL = _page_data()


def run_forecast(level="admin", allowed=()):
    def script(level, allowed):
        import streamlit as st
        import pnl_reports as pr
        from test_pnl_forecast import PAGE_PNL, PAGE_WM, PAGE_VL
        st.session_state["access_level"] = level
        df = PAGE_PNL if level == "admin" else PAGE_PNL[PAGE_PNL["Location"].isin(allowed)]
        wm = PAGE_WM if level == "admin" else PAGE_WM[PAGE_WM[pr.LOC].isin(allowed)]
        pr.report_forecast(df, wm, PAGE_VL, dict(
            age_group=lambda m: "24+ months", show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height)),
            bench=(PAGE_PNL, PAGE_WM))
    return AppTest.from_function(script, args=(level, allowed), default_timeout=60).run()


def test_forecast_page_for_an_admin():
    at = run_forecast()
    assert not at.exception, at.exception
    assert at.selectbox(key="pf_centre").options[0] == "Auburn"
    text = "".join(m.value for m in at.markdown)
    order = ["Settings", "Headline", "Adjust the forecast", "Forecast by month",
             "Benchmark % for each line"]
    pos = [text.index(h) for h in order]
    assert pos == sorted(pos)
    assert "FY2026–27" in text and "Estimate ⚠" in text and "Total" in text
    assert at.selectbox(key="pf_relief_Success Tutoring - Auburn").value == "Whole forecast"
    grid = next(d.value for d in at.dataframe if "Basis" in d.value)      # adjustment grid
    assert list(grid.index) == ["October 2026", "November 2026", "December 2026",
                                "January 2027", "February 2027", "March 2027", "April 2027",
                                "May 2027", "June 2027"]
    assert grid.loc["April 2027", "Basis"] == "Forecast"      # it has an April P&L
    assert grid.loc["October 2026", "Basis"] == "Estimate ⚠"
    assert (grid["Royalties %"] == 0).all()                   # royalty relief
    bench = next(d.value for d in at.dataframe if "Normal range" in d.value)
    lines = list(bench["Line"])
    for line in ["Tutor Wages + Super", "Rent", "Royalties", "Wages - Manager/Admin + Super",
                 "Marketing", "OPEX", "EBITDA"]:
        assert line in lines                         # admins: every line benchmarked
    assert (bench["Compared with"].str.extract(r"(\d+)")[0].astype(int) >= 5).all()
    assert "↳ vs centres that pay a manager" not in lines    # only 3 pay one: left out
    assert len(at.get("plotly_chart")) == 1
    assert at.get("download_button")
    assert any(e.label == "How this is worked out" for e in at.expander)
    at.selectbox(key="pf_view").set_value("Next 12 months").run()
    assert not at.exception, at.exception
    at.selectbox(key="pf_centre").set_value("Gordon").run()
    assert not at.exception, at.exception
    at.selectbox(key="pf_view").set_value("Financial year").run()
    at.selectbox(key="pf_fy").set_value(2027).run()
    assert not at.exception, at.exception
    assert "FY2027–28" in "".join(m.value for m in at.markdown)


def test_forecast_page_price_rise_changes_revenue():
    at = run_forecast()
    before = next(m.value for m in at.markdown if "pnl-wrap" in m.value)
    at.number_input(key="pf_rise_Success Tutoring - Auburn").set_value(10.0).run()
    assert not at.exception, at.exception
    after = next(m.value for m in at.markdown if "pnl-wrap" in m.value)
    assert before != after


def test_gpm_cannot_see_the_forecast():
    at = run_forecast("gpm", ["Success Tutoring - Gordon"])
    assert not at.exception, at.exception
    assert any("admins only" in e.value for e in at.error)
    text = "".join(m.value for m in at.markdown)
    for n in NAMES:
        assert n not in text                         # not even their own centre
    assert not at.dataframe and not at.get("plotly_chart")


def test_menu_key_position_and_admin_only():
    src = open(pr.__file__.replace("pnl_reports.py", "app.py")).read()
    reports = src[src.index("REPORTS=["):src.index("]", src.index("REPORTS=["))]
    assert '"14 · P&L Share",\n    "15 · P&L Centre Forecast",' in reports
    for key in ["1 · Campus Locations", "12 · P&L Summary", "13 · P&L Deep Dive",
                "14 · P&L Share", "11 · AI Outlier Analysis"]:
        assert f'"{key}"' in reports                 # existing keys unchanged
    nav = src[src.index("NAV_GROUPS = ["):src.index("ADMIN_NAV = [")]
    finance = nav[nav.index('("Finance"'):nav.index('("Insights"')]
    assert finance.index('"13 · P&L Deep Dive"') < finance.index('"15 · P&L Centre Forecast"') \
        < finance.index('"14 · P&L Share"')
    assert 'selected_report=="15 · P&L Centre Forecast"' in src and "report_forecast" in src
    # hidden from non-admins after allowed_tabs (so "all" doesn't let a GPM in)
    allowed = src.index("REPORTS = [r for r in REPORTS if report_allowed(")
    lock = src.index('ADMIN_ONLY_REPORTS = {"15 · P&L Centre Forecast", "16 · Franchise Model",')
    assert allowed < lock < src.index("NAV_GROUPS = [")
    block = src[lock:src.index("\n\n", lock)] if "\n\n" in src[lock:] else src[lock:]
    assert 'access_level") != "admin"' in block and "not in ADMIN_ONLY_REPORTS" in block


def test_admin_only_filter_runs_like_the_app():
    """The same lines as app.py, run for a GPM with allowed_tabs = all."""
    src = open(pr.__file__.replace("pnl_reports.py", "app.py")).read()
    start = src.index("def report_allowed(")
    end = src.index("if not REPORTS:", start)
    code = src[start:end]

    class State(dict):
        get = dict.get
    for level, expect in [("gpm", False), ("admin", True)]:
        import types
        st = types.SimpleNamespace(session_state=State(access_level=level, allowed_tabs=["all"]))
        ns = {"st": st, "REPORTS": ["14 · P&L Share", "15 · P&L Centre Forecast",
                                    "16 · Franchise Model"]}
        exec(code.replace("ADMIN_PAGES = [", "_ = [", 1), ns)
        assert ("15 · P&L Centre Forecast" in ns["REPORTS"]) is expect
        assert ("16 · Franchise Model" in ns["REPORTS"]) is expect
        assert "14 · P&L Share" in ns["REPORTS"]


def test_forecast_shows_gross_margin_by_month_benchmark_and_excel():
    import io
    at = run_forecast()
    assert not at.exception, at.exception
    table = next(m.value for m in at.markdown if "pnl-wrap" in m.value)
    assert table.index(">Tutor Wages + Super<") < table.index(">Gross Margin<") < \
        table.index(">Rent<")
    row = table.split(">Gross Margin</td>")[1].split("</tr>")[0]
    assert row.count("73.0%") >= 10          # tutor wages are 27% of revenue in every month
    bench = next(d.value for d in at.dataframe if "Normal range" in d.value)
    lines = list(bench["Line"])
    assert lines.index("Tutor Wages + Super") + 1 == lines.index("Gross Margin")
    assert bench.loc[bench["Line"] == "Gross Margin", "Forecast (avg)"].iloc[0] == "73.0%"
    grid = next(d.value for d in at.dataframe if "Basis" in d.value)
    assert not any("Gross" in c for c in grid.columns)      # nothing to type in
    # Excel: Gross Margin and its % straight after tutor wages.
    t = pd.DataFrame({"Basis": ["Actual"], "Members": [80.0], "AMV": [85.0],
                      "Total Revenue": [1000.0], "Tutor Wages + Super": [270.0],
                      "Rent": [100.0], "Wages - Manager/Admin + Super": [0.0],
                      "Royalties": [80.0], "Marketing - Agency Fees": [0.0],
                      "Marketing": [80.0], "OPEX": [50.0], "EBITDA": [420.0],
                      "BE members": [30.0]}, index=["2026-09"])
    xl = pd.read_excel(io.BytesIO(pr._forecast_excel(t, None, {"Centre": "X"})), "Forecast")
    cols = list(xl.columns)
    assert cols.index("Tutor Wages + Super") + 1 == cols.index("Gross Margin") and \
        cols.index("Gross Margin") + 1 == cols.index("Gross Margin %")
    assert xl.at[0, "Gross Margin"] == 730 and xl.at[0, "Gross Margin %"] == 73.0
    assert "Gross Margin" in pr.HOW_FORECAST


def test_gpm_sees_no_gross_margin_forecast():
    at = run_forecast("gpm", ["Success Tutoring - Gordon"])
    assert not at.exception, at.exception
    assert "Gross Margin" not in "".join(m.value for m in at.markdown)


# ── Franchise Model (report 16) ──────────────────────────────────────────────
import pnl_model as pm

TYP = {"members": 120.0, "amv": 80.0, "marketing_pct": 0.05, "manager": 3500.0,
       "fixed": {"Rent": 4000.0, "Marketing - Agency Fees": 0.0, "OPEX": 1500.0},
       "rates": {"20–39": 40.0, "100–119": 30.0}, "fx": 1.0}
# Simple prices for the arithmetic: everyone pays $80 a week, no GST, all of it collected, and
# nothing rises each July.
FLAT = {**pm.DEFAULTS, "standard_price": 80.0, "founder_prices": (80.0, 80.0, 80.0), "gst": 0.0,
        "collected": 100.0, "price_rise": 0.0, "cost_rise": 0.0, "wage_rise": 0.0}


def test_model_ramp_relief_and_royalty_minimum():
    s = dict(FLAT)
    t = pm.run(s, TYP, "2026-07", 60)
    m4 = 40 + 10 - 0.085 * 40                                  # 10 sign-ups, 8.5% leave
    ends = [14.0, 28.0, 40.0, m4, m4 + 10 - 0.085 * m4]         # end of each month
    assert list(t["Members"].iloc[:5]) == pytest.approx(
        [(a + b) / 2 for a, b in zip([0] + ends, ends)])        # the month's average
    assert list(t["Members end"].iloc[:3]) == pytest.approx([14.0, 28.0, 40.0])  # 40 at wk 12
    assert list(pm.grid_from(t)["Members"].iloc[:3]) == pytest.approx([14.0, 28.0, 40.0])
    assert t["Leavers"].iloc[3] == pytest.approx(3.4) and t["Leavers"].iloc[2] == 0
    assert list(t["Sign-ups"].iloc[:4]) == pytest.approx([14, 14, 12, 10])
    books = t["On books"]                       # levels off where leavers = sign-ups (118)
    assert 110 < books.iloc[-1] < 10 / 0.085 and books.is_monotonic_increasing
    assert t["AMV"].iloc[5] == pytest.approx(80)
    assert pm.relief_share(4, 18) == 1                                 # weeks 12.6–16.8
    assert pm.relief_share(5, 18) == pytest.approx(1.2 / 4.2)          # weeks 16.8–21
    assert pm.relief_share(6, 18) == 0
    assert pm.royalty_min("2026-11", s) == 2266
    assert pm.royalty_min("2027-07", s) == pytest.approx(2266 * 1.03)    # +3% each July
    t = pm.run(s, TYP, "2026-07", 6)
    assert (t["Royalties"].iloc[:4] == 0).all()
    assert t["Royalties"].iloc[4] == pytest.approx(
        (1 - 1.2 / 4.2) * max(2266, 0.08 * t["Total Revenue"].iloc[4]))


def test_model_manager_owner_wage_and_payback():
    s = {**FLAT, "owner_on": True}
    w = pm.run(s, TYP, "2026-07", 48, manager=True)
    wo = pm.run(s, TYP, "2026-07", 48, manager=False)
    hired = w[w["Wages - Manager/Admin + Super"] > 0]
    assert hired["Members"].iloc[0] >= 80 and w.loc[w.index < hired.index[0], "Members"].max() < 80
    assert (wo["Wages - Manager/Admin + Super"] == 0).all()
    assert (w[pm.OWNER] == 5000).all()                                 # 60k / 12, kept
    stop = pm.run({**s, "owner_stops": True}, TYP, "2026-07", 48)
    assert (stop.loc[hired.index, pm.OWNER] == 0).all()
    assert w["Profit"].sum() == pytest.approx((w["EBITDA"] - 5000).sum())
    assert w["Cumulative"].iloc[-1] == pytest.approx(w["Profit"].sum())
    pb_w, pb_wo = pm.payback_month(w, 110000), pm.payback_month(wo, 110000)
    assert pb_wo <= pb_w                                               # manager costs time
    assert w.loc[w["Model month"] == pb_w, "Cumulative"].iloc[0] >= 110000
    need = pm.needed(s, TYP, "2026-07", 110000, True, "signups")
    t = pm.run({**s, "signups": need}, TYP, "2026-07", 42)
    assert pm.payback_month(t, 110000) == 42 or abs(t["Cumulative"].iloc[41] - 110000) < 50
    price = pm.needed(s, TYP, "2026-07", 110000, True, "price")
    t = pm.run({**s, "standard_price": price, "founder_prices": (price,) * 3}, TYP, "2026-07", 42)
    assert pm.payback_month(t, 110000) == 42 or abs(t["Cumulative"].iloc[41] - 110000) < 50
    assert pm.fy_totals(w)["Months"].sum() == 48


def test_model_grid_change_overrides_one_month():
    s = dict(FLAT)
    base = pm.run(s, TYP, "2026-07", 12)
    g0 = pm.grid_from(base)
    g = g0.copy()
    g.loc["2027-01", "Members"] = 100.0
    g.loc["2027-02", "Rent %"] = 10.0
    t = pm.run(s, TYP, "2026-07", 12, True, g, g0)
    assert t.at["2027-01", "Members end"] == 100                 # end of month, as typed
    assert t.at["2027-01", "Members"] == (base.at["2026-12", "Members end"] + 100) / 2
    assert t.at["2027-02", "Members end"] == pytest.approx(100 + 10 - 8.5)  # carries on
    assert t.at["2027-02", "Rent"] == pytest.approx(0.1 * t.at["2027-02", "Total Revenue"])
    assert t.at["2027-03", "Rent"] == 4000


def run_model(level="admin"):
    def script(level):
        import streamlit as st
        import pnl_reports as pr
        from test_pnl_forecast import PAGE_PNL, PAGE_WM, PAGE_VL
        st.session_state["access_level"] = level
        vl = PAGE_VL.assign(**{"Weeks old": [130, 60, 40, 20, 150, 200, 90]})
        pr.report_franchise_model(PAGE_PNL, PAGE_WM, vl, dict(
            age_group=lambda m: "24+ months", show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height)),
            bench=(PAGE_PNL, PAGE_WM))
    return AppTest.from_function(script, args=(level,), default_timeout=60).run()


def test_franchise_model_page_for_an_admin():
    at = run_model()
    assert not at.exception, at.exception
    text = "".join(m.value for m in at.markdown)
    for h in ["Payback", "Road to profitability", "When can the centre afford royalties",
              "Monthly P&amp;L — all 60 months", "Profit each year", "Test the assumptions",
              "Worst, base and best case", "Foundation prices", "Opening month"]:
        assert h in text, h
    labels = [e.label for e in at.expander]
    for h in ["Settings", "Adjust the model month by month", "Benchmark % for each line",
              "Actual vs model, by centre age", "How this is worked out"]:
        assert h in labels, h
    assert "With manager" in text and "Without manager" in text and "Gross Margin" in text
    assert "Month 60" in text and "Joining fees" in text and "First month in profit" in text
    assert "Investment paid back" in text and "Can afford royalties every month" in text
    assert len(at.get("plotly_chart")) == 4              # cash, road, royalties, actual
    tabs = [t_.label for t_ in at.tabs]
    for h in ["Goal & growth", "Price", "Costs", "Royalties", "Manager & owner", "Holidays"]:
        assert h in tabs, h
    cases = next(d_.value for d_ in at.dataframe if "Growth" in d_.value)
    assert list(cases["Case"]) == ["Worst case", "Base case", "Best case"]
    assert next(d_.value for d_ in at.dataframe if "Opening month" in d_.value).shape[0] == 12
    bench = next(d_.value for d_ in at.dataframe if "Normal range" in d_.value)
    assert "Gross Margin" in list(bench["Line"])                       # 7 centres is enough
    cmp = next(d_.value for d_ in at.dataframe if "Centre age" in d_.value)
    assert len(cmp) and cmp["Centres"].str.contains("⚠").all()         # fewer than 8
    at.toggle(key="fm_Australia_owner_on").set_value(True).run()
    assert not at.exception, at.exception
    assert "Owner&#x27;s wage" in "".join(m.value for m in at.markdown) or \
        "Owner's wage" in "".join(m.value for m in at.markdown)
    at.radio(key="fm_scenario").set_value("With manager").run()
    assert "Year 5" in "".join(m.value for m in at.markdown)
    at.radio(key="fm_years").set_value("Financial years").run()
    years = "".join(m.value for m in at.markdown)
    assert "FY20" in years and "Year 5<br>" not in years
    at.number_input(key="fm_Australia_trigger").set_value(60.0).run()
    at.number_input(key="fm_Australia_rent").set_value(2000.0).run()
    assert not at.exception, at.exception


def test_gpm_cannot_see_the_franchise_model():
    at = run_model("gpm")
    assert any("admins only" in e.value for e in at.error)
    assert not at.dataframe and not at.get("plotly_chart")


def test_model_settings_by_country():
    au, ph = pm.settings_for("Australia")
    assert not ph and au["royalty_min"] == 2266 and au["royalty_rise"] == 3.0 \
        and au["relief_weeks"] == 18 and au["target_aud"] == 110000
    nz, ph = pm.settings_for("New Zealand")
    assert ph and nz["royalty_min"] == 2266                    # placeholder until set
    assert pm.run(nz, TYP, "2026-07", 6)["Royalties"].iloc[5] > 0


def test_holiday_months_slow_new_members_and_suspend_members():
    s = {**FLAT, "holiday_months": ["12", "01"], "holiday": {"12": 15.0, "01": 20.0}}
    t = pm.run(s, TYP, "2027-02", 24)           # opens Feb: 14, 28, 40 by April
    f = 12 / (12 - 2 * 0.85)                    # other months make up for the holidays
    assert t.at["2027-05", "Sign-ups"] == pytest.approx(10 * f)
    assert t.at["2027-12", "Sign-ups"] == pytest.approx(10 * f * 0.15)  # 85% fewer
    assert t.loc["2027-05":"2028-04", "Sign-ups"].sum() == pytest.approx(120)  # 10 a month
    nov, dec = t.at["2027-11", "On books"], t.at["2027-12", "On books"]
    assert dec == pytest.approx(nov + 10 * f * 0.15 - 0.085 * nov)
    avg = (nov + dec) / 2
    assert t.at["2027-12", "Suspended"] == pytest.approx(avg * 0.15)
    assert t.at["2027-12", "Members"] == pytest.approx(avg * 0.85)    # active drop by suspended
    assert t.at["2027-12", "Members end"] == pytest.approx(dec * 0.85)
    assert t.at["2028-02", "Suspended"] == 0                          # back the next month
    assert t.at["2027-12", "Suspension fees"] == pytest.approx(avg * 0.15 * 5 * 4.2)
    assert t.at["2027-12", "Tuition"] == pytest.approx(avg * 0.85 * 80 * 4.2)
    assert t.at["2027-12", "Total Revenue"] == pytest.approx(
        avg * 0.85 * 80 * 4.2 + avg * 0.15 * 5 * 4.2 + 10 * f * 0.15 * 100)
    late = pm.run(s, TYP, "2026-11", 6)        # opens Nov: holidays slow the way to 40
    assert list(late["On books"].iloc[:4]) == pytest.approx([14.0, 16.1, 18.2, 32.2])
    assert late["Sign-ups"].iloc[1] == pytest.approx(2.1)            # ramp joiners drop 85%
    assert late["On books"].iloc[4] == 40                            # 40 in March
    hols = pd.DataFrame({"State": ["NSW", "NSW", "VIC", "NZ"], "Year": [2026] * 4,
                         "Start_Week": [51, 15, 51, 28], "End_Week": [5, 16, 4, 29]})
    months, fallback = pm.holiday_months(hols, "Australia")
    assert not fallback and months == ["01", "12"]                # Apr: NSW only (1 of 2)
    assert pm.holiday_months(hols, "Canada") == (pm.FALLBACK_HOLIDAYS, True)
    rows = []
    for i in range(5):
        for d, act, sus in [("2025-11-02", 90, 10), ("2025-12-07", 80, 20), ("2026-01-04", 70, 30),
                            ("2026-02-01", 90, 10), ("2026-03-01", 90, 10), ("2026-04-05", 5, 5)]:
            rows.append({"Success Tutoring - Business name": f"C{i}", "Date": pd.Timestamp(d),
                         "# Active members": act, "# Suspended members": sus})
    wm = pd.DataFrame(rows)
    ex, n = pm.suspension_extra(wm, {f"C{i}" for i in range(5)}, ["12", "01", "04"])
    assert n == 5 and ex == {"12": 10.0, "01": 20.0}   # usual 10%; Apr too small (<20 active)
    assert pm.suspension_extra(wm, {"other"}, ["12"]) == ({}, 0)


def test_payback_card_status_progress_and_needs():
    import pnl_reports as pr
    t = pm.run(dict(FLAT), TYP, "2026-07", 60, manager=True)
    pb = pm.payback_month(t, 110000)
    on = pr._payback_card(pm.WITH, t, 110000, pb, 60, 6.5, 7.0, 70.0, 80.0)
    assert f"Month {pb}" in on and "On target" in on and "#2e8540" in on
    assert "now 7.0" in on and "&#36;80.00" in on and "$" not in on
    assert "Sign-ups a month needed" in on
    late = pr._payback_card(pm.WITH, t, 10 ** 7, 42, 60, None, 7.0, None, 80.0)
    assert "Not paid back" in late and "#c8322f" in late
    slow = pr._payback_card(pm.WITH, t, 110000, pb - 2, 60, 9.0, 7.0, 90.0, 80.0)
    assert "2 months late" in slow and 'width:' in slow


def test_presale_no_revenue_before_billing_week_and_presale_marketing():
    s = dict(FLAT)
    t = pm.run(s, TYP, "2026-07", 6)
    assert t["Tuition"].iloc[0] == 0                               # weeks 0–4.2: presale
    assert t["Total Revenue"].iloc[0] == pytest.approx(14 * 100)   # joining fees only
    assert t["Tuition"].iloc[1] == pytest.approx(                  # weeks 8–8.4 billed
        t["Members"].iloc[1] * 80 * 0.4)
    assert t["Tuition"].iloc[2] == pytest.approx(t["Members"].iloc[2] * 80 * 4.2)
    assert t["Total Revenue"].iloc[2] == pytest.approx(t["Tuition"].iloc[2] + 12 * 100)
    assert t["Tutor Wages + Super"].iloc[0] == 0 and t["Rent"].iloc[0] == 4000
    assert t["Marketing - Agency Fees"].iloc[0] == pytest.approx(1200)
    assert t["Marketing"].iloc[0] == pytest.approx(100 * 7 * 4.2)
    # month 3 = weeks 8.4–12.6: presale ends at week 12 (3.6 of 4.2 weeks)
    share = 3.6 / 4.2
    assert t["Marketing - Agency Fees"].iloc[2] == pytest.approx(1200 * share + 250 * (1 - share))
    assert t["Marketing"].iloc[2] == pytest.approx(100 * 7 * 3.6 + 30 * 7 * 0.6)
    assert t["Marketing"].iloc[3] == pytest.approx(30 * 7 * 4.2)   # flat, not % of revenue
    assert t["Marketing - Agency Fees"].iloc[3] == 250              # typical 0 -> minimum
    big = pm.run(s, {**TYP, "fixed": {**TYP["fixed"], "Marketing - Agency Fees": 400.0}},
                 "2026-07", 6)
    assert big["Marketing - Agency Fees"].iloc[4] == 400            # typical above minimum
    late = pm.run({**s, "billing_week": 13}, TYP, "2026-07", 4)
    assert (late["Tuition"].iloc[:2] == 0).all()
    assert (late[pnl.TUTOR].iloc[:2] == 0).all()                  # no sessions, no tutors
    assert late["Tuition"].iloc[2] == pytest.approx(late["Members"].iloc[2] * 80 * 0.6)


def test_model_results_breakeven_peak_cash_royalty_start_and_years():
    s = dict(pm.DEFAULTS)
    t = pm.run(s, TYP, "2026-07", pm.MODEL_MONTHS)
    assert len(t) == 60
    # tutor wages are a % of total revenue (joining fees included) once sessions start
    m = t.iloc[5]
    assert m[pnl.TUTOR] == pytest.approx(pnl.rate_for(m["Members"], TYP["rates"]) *
                                         m["Total Revenue"])
    assert t["Joining fees"].iloc[5] == 10 * 100                     # 10 sign-ups after week 12
    assert t["Royalties"].iloc[0] == 0 and t["Royalties due"].iloc[0] == 2266  # relief
    be = pm.breakeven_month(t)
    assert (t["Profit"].iloc[:be - 1] <= 0).all() and t["Profit"].iloc[be - 1] > 0
    cash, when = pm.peak_cash(t)
    assert cash == pytest.approx(-t["Cumulative"].min()) and cash > 0
    assert t["Cumulative"].iloc[when - 1] == t["Cumulative"].min()
    rs = pm.royalty_start(t)
    after = t["Profit"] + t["Royalties"] - t["Royalties due"]
    assert (after.iloc[rs - 1:] >= 0).all() and after.iloc[rs - 2] < 0
    assert pm.royalty_start(t.assign(Profit=-1.0)) is None
    out = pm.summary(t, 110000)
    assert out["5-year profit"] == pytest.approx(t["Cumulative"].iloc[-1])
    assert out["Break-even month"] == be and out["Royalty start month"] == rs
    years = pm.year_totals(t)
    assert list(years.index) == [1, 2, 3, 4, 5] and (years["Months"] == 12).all()
    assert years["Profit"].sum() == pytest.approx(t["Profit"].sum())
    assert years.at[2, "Members (end)"] == t["On books"].iloc[23]
    fy = pm.year_totals(t, fy=True)
    assert list(fy.index) == [2026, 2027, 2028, 2029, 2030]          # Jul start: FY = year
    assert pm.peak_cash(t.assign(Cumulative=1.0)) == (0.0, None)


def test_model_opening_months_and_cases():
    s = {**FLAT, "holiday_months": ["12", "01"], "holiday": {"12": 15.0, "01": 20.0}}
    om = pm.opening_months(s, TYP, "2026-07", 110000, count=3)
    assert list(om["Opening month"]) == ["2026-07", "2026-08", "2026-09"]
    assert om.at[0, "Payback (without manager)"] == \
        pm.payback_month(pm.run(s, TYP, "2026-07", 60, False), 110000)
    c = pm.cases(s, TYP, "2026-07", 110000)
    assert list(c["Case"]) == ["Worst case", "Base case", "Best case"]
    assert c.at[2, "Members after 5 years (without manager)"] == pytest.approx(145)  # targets
    best = pm.run({**s, "growth_mode": "targets"}, TYP, "2026-07", 60, False)
    assert [round(best["On books"].iloc[m]) for m in (11, 23, 35, 47)] == [80, 110, 130, 140]
    p = c["5-year profit (with manager)"]
    assert p.iloc[0] < p.iloc[1] < p.iloc[2]
    assert c.at[1, "5-year profit (without manager)"] == pytest.approx(
        pm.run(s, TYP, "2026-07", 60, False)["Profit"].sum())
    assert c.at[0, "Members after 5 years (without manager)"] == pytest.approx(8 / 0.10, abs=2)


def test_model_price_cost_rises_foundation_prices_and_manager_cost():
    s = dict(pm.DEFAULTS)
    # price: first rise the first July at least 12 months after opening
    assert pm.price_rises("2028-06", "2026-12", 3) == 1                   # Dec 27 + Jul 28
    assert pm.price_rises("2028-07", "2026-12", 3) == pytest.approx(1.03)
    assert pm.price_rises("2027-07", "2026-07", 3) == pytest.approx(1.03)  # exactly 12 months
    assert pm.price_rises("2027-06", "2026-07", 3) == 1
    flat = {**FLAT, "price_rise": 3.0, "cost_rise": 3.0}
    t = pm.run(flat, TYP, "2026-12", 24)
    assert t.at["2027-07", "AMV"] == 80 and t.at["2028-07", "AMV"] == pytest.approx(80 * 1.03)
    # rent and OPEX rise 3% each July after opening
    assert t.at["2027-06", "Rent"] == 4000 and t.at["2027-07", "Rent"] == pytest.approx(4120)
    assert t.at["2027-07", "OPEX"] == pytest.approx(1500 * 1.03)
    assert pm.manager_monthly(s) == pytest.approx(32000 * 1.12 / 12)       # ex super + 12%
    # foundation members: first 20 at $50, next 20 at $60, next 10 at $65, then standard
    assert pm.founders_after([0, 0, 0], 0, 14, s, False) == [14, 0, 0]
    assert pm.founders_after([14, 0, 0], 14, 14, s, False) == [20, 8, 0]
    assert pm.founders_after([20, 20, 0], 40, 15, s, False) == [20, 20, 10]   # 5 left over
    left = pm.founders_after([20, 20, 10], 60, 10, s, True)                 # 8.5% leave
    assert left == pytest.approx([20 * 0.915, 20 * 0.915, 10 * 0.915])
    assert pm.list_price([20, 20, 10], 50, s) == pytest.approx((1000 + 1200 + 650) / 50)
    assert pm.list_price([20, 20, 10], 100, {**s, "standard_price": 70.0}) == \
        pytest.approx((1000 + 1200 + 650 + 50 * 70) / 100)
    # GST off and the collected % on
    assert pm.collected_share({**s, "gst": 10.0, "collected": 80.0}) == pytest.approx(0.8 / 1.1)
    t = pm.run({**s, "collected": 80.0, "gst": 10.0}, TYP, "2026-07", 4)
    assert t["AMV"].iloc[0] == pytest.approx(50 * 0.8 / 1.1)              # 14 founders at $50
    assert t["Founders"].iloc[3] == pytest.approx(40 * 0.915 + 10)       # last 10 join as
    #                                                       the first 40 start to leave


def test_model_package_mix_rrp_price_and_collected_pct():
    import membership_composition as mc
    comp = mc.prepare(pd.DataFrame({
        "Week ending": ["2026-11-08"] * 4, "Location": ["A", "A", "B", "B"],
        "Country": ["Australia"] * 4, "Type": ["Group", "Group", "Group", "Group"],
        "Sessions/week": [1, 2, 2, 2], "Members": [2, 6, 2, 3],
        "Weekly price": [40, 70, 70, 5]}))                                 # $5: a freeze, left out
    mix, n = pm.plan_mix(comp, "Australia")
    assert n == 2 and mix == pytest.approx({("Group", 1): 20.0, ("Group", 2): 80.0})
    assert pm.plan_mix(None, "Australia") == (pm.FALLBACK_MIX, 0)
    rrp = mc.rrp_table()[("AU", "all")]                                   # 42 / 74 / 100
    assert pm.rrp_mix_price(rrp, mix) == pytest.approx(0.2 * 42 + 0.8 * 74)
    peers = pd.DataFrame({"Members": [50.0, 100.0, 0.0], "Region": ["X", "X", "X"],
                          "Total Revenue": [50 * 4.2 * 60 + 1000, 100 * 4.2 * 30 + 1000, 1.0]})
    s = dict(pm.DEFAULTS)                                                 # 10 x $100 joining
    pct, n = pm.collected_pct(peers, lambda reg: 66.0, s)                 # $60 after GST
    assert n == 2 and pct == pytest.approx((100 + 50) / 2)                # median of 100%, 50%

def test_holiday_suspensions_are_not_taken_off_twice():
    s = {**pm.DEFAULTS, "gst": 10.0, "suspension_fee": 5.0}
    extra = {"01": 12.0, "12": 12.0}                     # 2% of member-months over a year
    banked = 77.0 / 1.1 * 0.70                           # $49 a week
    pct = pm.holiday_gross_up(70.0, s, 77.0, extra)
    assert pct == pytest.approx(70 / (1 - 0.02 * (1 - 5 / banked)))
    assert pm.holiday_gross_up(70.0, s, 77.0, {}) == 70.0
    assert pm.holiday_gross_up(None, s, 77.0, extra) is None


def test_model_milestones_road_to_profit():
    s = dict(FLAT)
    t = pm.run(s, TYP, "2026-07", 60, manager=True)
    ms = pm.milestones(t, s, 110000)
    labels = [l for _, l in ms]
    months = dict((l, m) for m, l in ms)
    assert months["Sessions and billing start"] == 2                 # week 9 is in month 2
    assert months["Presale campaign ends"] == 3 and months["40 members"] == 3
    assert months["Royalty relief ends"] == 5                         # 18 weeks
    assert months["First month in profit"] == pm.breakeven_month(t)
    assert months["Investment paid back"] == pm.payback_month(t, 110000)
    assert months["Manager hired"] == int(t[t[pnl.MANAGER] > 0]["Model month"].iloc[0])
    known = [m for m, _ in ms if m is not None]
    assert known == sorted(known)
    wo = pm.milestones(pm.run(s, TYP, "2026-07", 60, manager=False), s, 110000)
    assert "Manager hired" not in [l for _, l in wo]
    never = pm.milestones(t, s, 10 ** 9)
    assert never[-1] == (None, "Investment paid back")              # not within 5 years
