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
    lock = src.index('ADMIN_ONLY_REPORTS = {"15 · P&L Centre Forecast"}')
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
        ns = {"st": st, "REPORTS": ["14 · P&L Share", "15 · P&L Centre Forecast"]}
        exec(code.replace("ADMIN_PAGES = [", "_ = [", 1), ns)
        assert ("15 · P&L Centre Forecast" in ns["REPORTS"]) is expect
        assert "14 · P&L Share" in ns["REPORTS"]
