"""P&L summary and break-even: calculations and the report page."""
import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

import pnl
import pnl_reports as pr

LOC = pr.LOC


def lines(loc, period, amounts, currency="AUD", country="Australia", fx=1.0):
    return [{"Period": period, "Location": loc, "Country": country, "Currency": currency,
             "Status": "Submitted", "Account code": code, "Account": str(code),
             "P&L section": pnl.section_of(code), "Amount": amt, "FX rate to AUD": fx}
            for code, amt in amounts.items()]


AUBURN = {200: 12600, 472: 3300, 474: 396, 469: 3500, 467: 0, 402: 250, 400: 900, 453: 2000,
          437: 100, 479: 500}
HOWICK = {200: 8000, 472: 2000, 469: 2000, 467: 2200, 400: 300}
PNL = pd.DataFrame(
    lines("Success Tutoring - Auburn", "2026-07", {**AUBURN, 200: 11600}) +
    lines("Success Tutoring - Auburn", "2026-08", AUBURN) +
    lines("Success Tutoring - Howick", "2026-08", HOWICK, "NZD", "New Zealand", 0.9) +
    lines("Success Tutoring - Burwood", "2026-08", {200: 5000, 472: 1000, 469: 6000}))
VL = pd.DataFrame({LOC: ["Success Tutoring - Auburn", "Success Tutoring - Howick",
                         "Success Tutoring - Burwood"],
                   "Region": ["New South Wales", "Auckland (N)", "New South Wales"],
                   "Stage": ["Growth", "Growth", "Onboarding"], "GPM": ["A", "B", "A"],
                   "Country": ["Australia", "New Zealand", "Australia"],
                   "Age (Months)": [30, 14, 2], "Royalty Min": ["", "$2,000", ""]})
WM = pd.DataFrame([{LOC: loc, "Date": pd.Timestamp(d), "# Active members": n}
                   for loc, n in [("Success Tutoring - Auburn", 50), ("Success Tutoring - Howick", 30),
                                  ("Success Tutoring - Burwood", 20)]
                   for d in ["2026-07-12", "2026-08-02", "2026-08-09"]])
AGE = lambda m: "24+ months" if m and m > 24 else "0–3 months" if m is not None and m <= 3 \
    else "12–24 months"


def test_break_even_matches_the_model_sheet():
    # 50 members x 4.2 wks x $60; tutor 29.33%; fixed 3,500 + 250 + 900 + 2,000; royalty min 2,200
    r, m = pnl.break_even(12600, 12600 * 0.29333, 6650, 2200, 50)
    assert round(m, 1) == 49.7
    col = pnl.model_column(50, 252, 0.29333, {"Rent": 3500, "Marketing - Agency Fees": 250,
                                             "Marketing": 900, "OPEX": 2000}, 2200)
    assert round(col["EBITDA"]) == 54


def test_break_even_when_royalty_is_8_percent():
    # Large fixed costs: at break-even 8% of revenue beats the minimum.
    r, _ = pnl.break_even(10000, 3000, 30000, 2200)
    assert r == pytest.approx(30000 / (1 - 0.3 - 0.08))
    assert pnl.break_even(10000, 9500, 1000, 2200) == (None, None)    # tutor eats it all
    assert pnl.break_even(0, 0, 1000, 2200) == (None, None)


def test_tutor_benchmarks_by_band():
    cent = pd.DataFrame({"Members": [25, 30, 35, 45, 90, 95, 99, 100, 10],
                         "Total Revenue": [100] * 9,
                         "Tutor Wages + Super": [50, 40, 45, 44, 30, 28, 32, 0, 60]})
    b = pnl.tutor_benchmarks(cent)
    assert b.at["20–39", "Centres"] == 3 and b.at["20–39", "Use %"] == 45.0
    assert b.at["80–99", "Use %"] == 30.0 and b.at["80–99", "Source"] == "This band"
    assert b.at["40–59", "Source"] == "Nearest band (20–39)"      # only 1 centre there
    assert b.at["100–119", "Centres"] == 0                        # no tutor wages: left out
    assert b.at["200+", "Use %"] == 30.0


def test_bigger_bands_never_use_a_higher_tutor_pct():
    cent = pd.DataFrame({"Members": [25, 30, 35, 105, 110, 115],
                         "Total Revenue": [100] * 6,
                         "Tutor Wages + Super": [36, 36, 36, 43, 43, 43]})
    b = pnl.tutor_benchmarks(cent)
    assert b.at["100–119", "Benchmark %"] == 43 and b.at["100–119", "Use %"] == 36
    assert b.at["100–119", "Source"].startswith("Capped")


def test_tutor_pct_blends_between_bands():
    rates = {b: 40.0 for b in pnl.membership_bands()}
    rates["40–59"] = 30.0                                  # middle of the band: 50 members
    assert pnl.rate_for(50, rates) == pytest.approx(0.30)
    assert pnl.rate_for(40, rates) == pytest.approx(0.35)  # halfway from 30 (at 40%) to 50
    assert pnl.rate_for(500, rates) == pytest.approx(0.40)


def test_banded_break_even_uses_lower_tutor_pct_when_bigger():
    rates = {b: 50.0 for b in pnl.membership_bands()}
    flat = pnl.break_even_banded(250, 8000, 2200, rates)[1]
    rates.update({"60–79": 40.0, "80–99": 30.0})
    falling = pnl.break_even_banded(250, 8000, 2200, rates)[1]
    assert falling < flat and 60 <= falling < 100
    assert pnl.ebitda_at(falling, 250, rates, 8000, 2200) == pytest.approx(0, abs=0.01)
    # A group of 2 average centres needs twice the members.
    assert pnl.break_even_banded(250, 16000, 4400, rates, centres=2)[1] == \
        pytest.approx(2 * falling, rel=1e-6)


def test_location_months_rows():
    lm = pnl.location_months(PNL).loc[("Success Tutoring - Auburn", "2026-08")]
    assert lm["Total Revenue"] == 12600 and lm["Tutor Wages + Super"] == 3696
    assert lm["OPEX"] == 2000 and lm["Interest"] == 100
    assert lm["EBITDA"] == 12600 - 3696 - 3500 - 250 - 900 - 2000     # 479 left out
    assert lm["Net Profit"] == lm["EBITDA"] - 100


def test_periods_and_bands():
    ps = ["2025-06", "2025-07", "2025-12", "2026-01", "2026-02"]
    assert pr.select_periods(ps, "Single month", "2026-01") == ["2026-01"]
    assert pr.select_periods(ps, "Last 3 months", "2026-02") == ["2025-12", "2026-01", "2026-02"]
    assert pr.select_periods(ps, "Year to date", "2026-02") == ["2026-01", "2026-02"]
    assert pr.select_periods(ps, "Year to date", "2025-12") == ["2025-06", "2025-07", "2025-12"]
    assert pnl.membership_band(0) == "0–19" and pnl.membership_band(199.5) == "180–199"
    assert pnl.membership_band(200) == "200+" and len(pnl.membership_bands()) == 11


def test_centre_table_and_groups():
    cent = pr.centre_table(PNL, pr.members_by_month(WM), VL, ["2026-07", "2026-08"], AGE)
    a = cent.loc["Success Tutoring - Auburn"]
    assert a["Months"] == 2 and a["Total Revenue"] == 12100           # monthly average
    assert a["Members"] == 50 and a["Band"] == "40–59" and a["Age group"] == "24+ months"
    assert a["Royalty Min"] == 2200
    assert cent.loc["Success Tutoring - Howick", "Royalty Min"] == 2000
    rates = {b: 30.0 for b in pnl.membership_bands()}
    cols, left = pr.build_columns(cent, "Country", lambda r: r, rates)
    assert [c[0] for c in cols] == ["Australia (2)", "New Zealand (1)"]
    assert cols[1][1] == "NZD" and left == []
    net = pr.build_columns(cent, "Network", lambda r: r, rates)[0][0]
    assert net[1] == "AUD"                                              # mixed: in AUD
    assert net[2]["Total Revenue"] == pytest.approx(12100 + 5000 + 8000 * 0.9)
    assert net[2]["Royalty Min"] == pytest.approx(2200 + 2200 + 2000 * 0.9)


def run_report(level="admin", allowed=()):
    def script(level, allowed):
        import streamlit as st
        import pnl_reports as pr
        from test_pnl_reports import PNL, VL, WM, AGE
        st.session_state["access_level"] = level

        def filters(df, key_prefix="", panel=None, show_label=True, **k):
            return df
        df = PNL if level == "admin" else PNL[PNL["Location"].isin(allowed)]
        pr.report_pnl(df, WM, VL, dict(
            age_group=AGE, age_order=["0–3 months", "12–24 months", "24+ months", "Unknown"],
            state_of=lambda r: r, show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height),
            report_filters=filters), bench=(PNL, WM))
    return AppTest.from_function(script, args=(level, allowed), default_timeout=30).run()


def test_report_renders_views_and_model():
    at = run_report()
    assert not at.exception, at.exception
    html = "".join(m.value for m in at.markdown)
    assert "Auburn" in html and "Howick" in html and "Break-even members" in html
    at.selectbox(key="pnl_view").set_value("Network").run()
    assert not at.exception, at.exception
    assert "Network (3)" in "".join(m.value for m in at.markdown)
    at.selectbox(key="pnl_view").set_value("Centre").run()
    at.selectbox(key="pnl_model_pick").set_value("Auburn").run()
    assert not at.exception, at.exception
    text = "".join(m.value for m in at.markdown)
    assert "to break even and has" in text


def test_report_for_a_gpm_shows_only_their_centres():
    at = run_report("gpm", ["Success Tutoring - Howick"])
    assert not at.exception, at.exception
    html = "".join(m.value for m in at.markdown)
    assert "Howick" in html and "Auburn" not in html
    bench = at.dataframe[0].value          # typical values come from every centre
    assert bench["Centres"].sum() == 3


def test_report_with_no_pnl_yet():
    def script():
        import pandas as pd
        import pnl_reports as pr
        pr.report_pnl(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {})
    at = AppTest.from_function(script).run()
    assert any("No P&L figures yet" in i.value for i in at.info)


def bench_frame():
    rng = [(f"C{i}", "Australia", 50 + i, 10000, 3500 + 100 * i, 0) for i in range(8)]
    rng += [("Big", "Australia", 55, 10000, 6000, 3000)]
    rows = []
    for name, country, mem, rev, tutor, mgr in rng:
        r = {"Country": country, "Members": mem, "Total Revenue": rev,
             "Tutor Wages + Super": tutor, "Rent": 2000, pnl.MANAGER: mgr, "Royalties": 800,
             "Marketing - Agency Fees": 0, "Marketing": 500, "OPEX": 1000}
        r["EBITDA"] = rev - sum(v for k, v in r.items() if k not in
                                ("Country", "Members", "Total Revenue"))
        rows.append(pd.Series(r, name=name))
    return pd.DataFrame(rows)


def test_peers_compare_and_outliers():
    b = bench_frame()
    peers, desc = pnl.peer_set(b, "Australia", 55, "Big")
    assert "Big" not in peers.index and len(peers) >= pnl.MIN_PEERS
    assert desc.startswith("Australia centres with")
    cmp = pnl.compare_lines(b.loc["Big"], peers, pnl.BENCH_LINES).set_index("Line")
    assert cmp.at["Tutor Wages + Super", "Status"] == "Very high"
    assert cmp.at["Tutor Wages + Super", "Gap $"] == pytest.approx(6000 - 3850)
    assert cmp.at[pnl.MANAGER, "Gap $"] == 3000          # typical is no manager wage
    assert "↳ vs centres that pay a manager" not in cmp.index   # no peer pays one
    assert cmp.at["EBITDA", "Status"].endswith("low")
    assert pnl.is_bad("EBITDA", "Very low") and pnl.is_bad("Rent", "High")
    assert not pnl.is_bad("Rent", "Low")
    assert cmp.at["Average Membership Value", "Status"] == "Typical"


def test_peer_set_widens_then_falls_back():
    b = bench_frame()
    _, desc = pnl.peer_set(b, "Australia", 150, None)          # nobody near 150 members
    assert desc == "all Australia centres"
    _, desc = pnl.peer_set(b, "New Zealand", 55, None)
    assert desc == "all reporting centres"


def test_report_shows_outliers_for_a_centre():
    at = run_report()
    assert not at.exception, at.exception
    heads = "".join(m.value for m in at.markdown)
    assert "Outliers vs similar centres" in heads and "Outlier list" in heads
    at.selectbox(key="pnl_view").set_value("GPM").run()
    assert not at.exception, at.exception
    at.selectbox(key="pnl_view").set_value("Centre").run()
    at.toggle(key="pnl_out_detail").set_value(True).run()
    assert not at.exception, at.exception
