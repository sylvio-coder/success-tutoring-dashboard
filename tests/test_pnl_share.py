"""P&L Share: the anonymity rules, sensitive lines and the report page (admin and GPM)."""
import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

import pnl
import pnl_reports as pr
import pnl_share as ps

LOC = pr.LOC
NAMES = {200: "Sales", 270: "Interest Income", 472: "Wages - Tutor",
         474: "Superannuation - Tutor", 469: "Rent", 471: "Wages - Manager/ Admin",
         478: "Superannuation - Manager/ Admin", 467: "Royalties", 400: "Marketing",
         453: "Office Expenses", 446: "Water", 437: "Interest Expense", 486: "CRM Software"}


def lines(loc, period, amounts, currency="AUD", country="Australia"):
    return [{"Period": period, "Location": loc, "Country": country, "Currency": currency,
             "Status": "Submitted", "Account code": code, "Account": NAMES[code],
             "P&L section": pnl.section_of(code), "Amount": amt, "FX rate to AUD": 1.0}
            for code, amt in amounts.items()]


PEERS = [f"Success Tutoring - Peer {i}" for i in range(8)]
BASE = {200: 20000, 270: 0, 472: 6000, 474: 700, 469: 3000, 471: 0, 478: 0, 467: 2200,
        400: 800, 453: 900, 446: 0, 437: 0, 486: 450}
MONTHS = ["2026-06", "2026-07", "2026-08"]
# Auburn: high tutor wages and rent, a manager, royalty relief, some interest.
AUBURN = {**BASE, 472: 9000, 474: 1000, 469: 6000, 471: 2500, 478: 300, 467: 0, 446: 150,
          437: 1100, 270: 400, 486: 1500}
PNL = pd.DataFrame(
    [r for p in MONTHS for r in lines("Success Tutoring - Auburn", p, AUBURN)] +
    [r for i, loc in enumerate(PEERS) for p in MONTHS
     for r in lines(loc, p, {**BASE, 472: 5600 + 100 * i, 453: 700 + 50 * i,
                             446: 120 if i < 2 else 0})] +
    [r for p in MONTHS for r in lines("Success Tutoring - Howick", p, BASE, "NZD",
                                      "New Zealand")])
ALL = ["Success Tutoring - Auburn", "Success Tutoring - Howick"] + PEERS
VL = pd.DataFrame({LOC: ALL, "Region": ["New South Wales", "Auckland (N)"] + ["Victoria"] * 8,
                   "Stage": "Growth", "GPM": ["A", "B"] + ["C"] * 8,
                   "Country": ["Australia", "New Zealand"] + ["Australia"] * 8,
                   "Age (Months)": 30, "Weeks old": 130, "Royalty Min": ""})
DATES = pd.date_range("2026-04-05", "2026-08-30", freq="W")       # 22 weeks
WM = pd.DataFrame([{LOC: loc, "Date": d, "# Active members": 60 + (i if loc.endswith("Auburn")
                                                                   else 0)}
                   for loc in ALL for i, d in enumerate(DATES)])
REVENUE = pd.DataFrame([{LOC: loc, "Date": d, "# Active Students": 60, "Total Sessions": 30}
                        for loc in ALL for d in DATES])
AGE = lambda m: "24+ months"


def peers_frame(n=8, users=8):
    """n similar centres; the first `users` of them spend on Water."""
    rows = {f"P{i}": {"Country": "Australia", "Members": 60, "Total Revenue": 20000.0,
                      "Wages - Tutor": 5000.0 + 200 * i, "Water": 100.0 if i < users else 0.0,
                      "Rent": 3000.0, "Royalties": 2200.0, "Wages - Manager/ Admin": 0.0,
                      "Interest Expense": 0.0}
            for i in range(n)}
    return pd.DataFrame(rows).T.infer_objects()


def me(**over):
    v = {"Total Revenue": 20000.0, "Members": 60, "Country": "Australia",
         "Wages - Tutor": 8000.0, "Water": 120.0, "Rent": 9000.0, "Royalties": 0.0,
         "Wages - Manager/ Admin": 4000.0, "Interest Expense": 1500.0}
    v.update(over)
    return pd.Series(v)


ACCTS = ["Wages - Tutor", "Water", "Rent", "Royalties", "Wages - Manager/ Admin",
         "Interest Expense"]


# ── Sensitive lines ───────────────────────────────────────────────────────────
def test_sensitive_accounts_found_by_code_and_name():
    names = ps.sensitive_accounts(PNL)
    assert {"Rent", "Royalties", "Wages - Manager/ Admin", "Superannuation - Manager/ Admin",
            "Interest Expense", "CRM Software"} <= names
    assert "Wages - Tutor" not in names and "Interest Income" not in names
    # Odd account names still match by code...
    odd = PNL.assign(Account=PNL["Account"].replace({"Rent": "Premises lease"}))
    assert "Premises lease" in ps.sensitive_accounts(odd)
    # ...and by name, whatever the code.
    assert ps.is_sensitive("Royalty Fees") and ps.is_sensitive("Rent")
    assert ps.is_sensitive("Wages - Manager/Admin + Super") and ps.is_sensitive("Interest")
    assert ps.is_sensitive("CRM Software") and ps.is_sensitive("CRM subscription")
    assert not ps.is_sensitive("Current assets") and not ps.is_sensitive("Wages - Tutor")


def test_sensitive_lines_show_own_figures_but_no_benchmark():
    t = ps.benchmark_lines(me(), peers_frame(), ACCTS).set_index("Line")
    for line in ["Rent", "Royalties", "Wages - Manager/ Admin", "Interest Expense"]:
        r = t.loc[line]
        assert r["Note"] == ps.NOT_BENCHMARKED and not r["Benchmarked"]
        assert pd.notna(r["Actual $"]) and pd.notna(r["Actual %"])
        for c in ["Typical %", "Range low %", "Range high %", "Gap $", "Status", "Centres"]:
            assert r[c] is None or pd.isna(r[c]), (line, c)
    assert t.loc["Wages - Tutor", "Benchmarked"]


def test_sensitive_lines_left_out_of_tiles_totals_and_opportunities():
    t = ps.benchmark_lines(me(Water=100.0), peers_frame(), ACCTS)
    more, less = ps.out_of_step(t)
    shown = set(more["Line"]) | set(less["Line"])
    assert "Wages - Tutor" in shown
    assert not shown & {"Rent", "Royalties", "Wages - Manager/ Admin", "Interest Expense"}
    opps = ps.top_opportunities(t)
    assert [o["Line"] for o in opps] == ["Wages - Tutor"]
    assert opps[0]["Text"].startswith("If wages - tutor matched similar centres, you'd keep "
                                      "about $")
    # The sensitive lines' gaps would be big if they were counted.
    assert more["Gap $"].sum() < 9000 - 3000


def test_grouped_manager_row_and_paying_centres_row_are_not_benchmarked():
    b = pd.DataFrame({f"P{i}": {"Country": "Australia", "Members": 60, "Total Revenue": 10000,
                                "Tutor Wages + Super": 3500, "Rent": 2000,
                                pnl.MANAGER: 1000, "Royalties": 800,
                                "Marketing - Agency Fees": 0, "Marketing": 500, "OPEX": 1000,
                                "EBITDA": 1200} for i in range(6)}).T.infer_objects()
    v = b.iloc[0].copy()
    v[pnl.MANAGER] = 3000
    t = ps.benchmark_lines(v, b, pnl.BENCH_LINES).set_index("Line")
    assert t.loc[pnl.MANAGER, "Note"] == ps.NOT_BENCHMARKED
    assert not any(l.startswith("↳") for l in t.index)        # "vs centres that pay one"
    assert t.loc["Tutor Wages + Super", "Benchmarked"]


# ── Anonymity rules ───────────────────────────────────────────────────────────
def test_needs_five_similar_centres():
    t = ps.benchmark_lines(me(), peers_frame(n=4), ACCTS).set_index("Line")
    assert (t["Note"].drop(["Rent", "Royalties", "Wages - Manager/ Admin",
                            "Interest Expense"]) == ps.NOT_ENOUGH).all()
    assert t["Typical %"].isna().all() and t["Range low %"].isna().all()
    assert ps.top_opportunities(t.reset_index()) == []
    t = ps.benchmark_lines(me(), peers_frame(n=5), ACCTS).set_index("Line")
    assert t.loc["Wages - Tutor", "Benchmarked"] and t.loc["Wages - Tutor", "Centres"] == 5


def test_share_peers_widens_then_gives_up():
    b = pd.DataFrame({f"E{i}": {"Country": "Australia", "Members": 60, "Total Revenue": 10000,
                                "Age (Months)": 30} for i in range(6)}).T.infer_objects()
    dev = pd.DataFrame({f"D{i}": {"Country": "Australia", "Members": 60,
                                  "Total Revenue": 10000, "Age (Months)": 6}
                        for i in range(3)}).T.infer_objects()
    bench = pd.concat([b, dev])
    peers, desc = ps.share_peers(bench, 30, "Australia", 60, "E0")
    assert len(peers) == 5 and desc == "established Australian centres with 60–79 members"
    # A developing centre with only 2 other developing centres: the stage is dropped.
    peers, desc = ps.share_peers(bench, 6, "Australia", 60, "D0")
    assert len(peers) == 8 and "developing" not in desc and "established" not in desc
    assert desc == "Australian centres with 60–79 members"
    # Young centres (18 weeks or younger) are left out; then there aren't enough.
    peers, desc = ps.share_peers(bench, 30, "Australia", 60, "E0",
                                 young={"E1", "E2", "D0", "D1", "D2"})
    assert peers.empty and desc is None


def test_normal_range_is_the_middle_half_not_lowest_to_highest():
    p = peers_frame(n=9)
    t = ps.benchmark_lines(me(), p, ["Wages - Tutor"]).set_index("Line").loc["Wages - Tutor"]
    pct = p["Wages - Tutor"] / p["Total Revenue"] * 100
    assert t["Range low %"] == pytest.approx(pct.quantile(0.25))
    assert t["Range high %"] == pytest.approx(pct.quantile(0.75))
    assert t["Range low %"] > pct.min() and t["Range high %"] < pct.max()
    assert t["Typical %"] == pytest.approx(pct.median())


def test_only_benchmarked_when_most_similar_centres_use_the_account():
    t = ps.benchmark_lines(me(), peers_frame(n=8, users=4), ACCTS).set_index("Line")
    assert t.loc["Water", "Note"] == ps.NO_BENCHMARK and t.loc["Water", "Users"] == 4
    assert pd.isna(t.loc["Water", "Status"])          # not "Well above typical"
    t = ps.benchmark_lines(me(), peers_frame(n=8, users=5), ACCTS).set_index("Line")
    assert t.loc["Water", "Benchmarked"]


def test_quarter_bands_not_ranks():
    peers = [1, 2, 3, 4, 5, 6, 7, 8]
    assert ps.quarter_band(9, peers) == "Top quarter"
    assert ps.quarter_band(5, peers) == "Above the middle"
    assert ps.quarter_band(3, peers) == "Below the middle"
    assert ps.quarter_band(0, peers) == "Bottom quarter"
    assert ps.quarter_band(0, peers, higher_is_better=False) == "Top quarter"
    assert ps.quarter_band(9, peers[:4]) is None                     # under 5 centres


def test_soft_wording():
    assert ps.soft_status("Wages - Tutor", "Very high") == "Worth a look"
    assert ps.soft_status("Average Membership Value", "Very low") == "Worth a look"
    assert ps.soft_status("Average Membership Value", "Very high") == "Well above typical"


def test_typical_by_month_needs_five_centres():
    lm = pnl.location_months(PNL)
    country = pr.location_info(PNL, VL)["Country"].to_dict()
    t = ps.typical_by_month(lm, pr.members_by_month(WM), PEERS, MONTHS, "Australia", country)
    assert t.loc["2026-08", "Revenue"] == 20000 and t.loc["2026-08", "Centres"] == 8
    assert t.loc["2026-08", "EBITDA %"] == pytest.approx(
        (pnl.location_months(PNL).xs("2026-08", level="Period").loc[PEERS, "EBITDA"]
         / 20000 * 100).median())
    t = ps.typical_by_month(lm, None, PEERS[:4], MONTHS, "Australia", country)
    assert t["Revenue"].isna().all() and t["EBITDA %"].isna().all()


def test_checks_before_sharing():
    v = pd.Series({"Total Revenue": 20000.0, "Members": 60, "Tutor Wages + Super": 6000.0,
                   "Rent": 0.0, pnl.MANAGER: 0.0, "Royalties": 0.0, "Marketing": 500.0,
                   "OPEX": 1000.0})
    peers = pd.DataFrame({f"P{i}": {"Total Revenue": 20000.0, "Tutor Wages + Super": 6000.0,
                                    "Rent": 3000.0, pnl.MANAGER: 0.0, "Royalties": 2200.0,
                                    "Marketing": 500.0, "OPEX": 900.0}
                          for i in range(6)}).T
    found = " ".join(ps.checks(v, peers, interest_income=400))
    assert "Rent is $0" in found
    assert "Royalties" not in found                    # royalty relief is normal
    assert "Wages - Manager" not in found              # most centres don't pay one
    assert "small sample" in found and "interest income" in found
    assert any("Fewer than 5" in c for c in ps.checks(v, peers.iloc[:3]))


def test_members_trend_for_one_centre():
    t = ps.members_trend(WM, "Success Tutoring - Auburn", DATES[-1])
    assert not t["too_new"] and t["trend"] > 0 and t["vs"] > 0
    assert len(t["spark"]) == 12


# ── The page ──────────────────────────────────────────────────────────────────
def run_share(level="admin", allowed=()):
    def script(level, allowed):
        import streamlit as st
        import pnl_reports as pr
        from test_pnl_share import PNL, VL, WM, AGE, REVENUE
        df = PNL if level == "admin" else PNL[PNL["Location"].isin(allowed)]
        wm = WM if level == "admin" else WM[WM[pr.LOC].isin(allowed)]
        pr.report_share(df, wm, VL, dict(
            age_group=AGE, age_order=[], state_of=lambda r: r, show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height),
            report_filters=None, revenue=REVENUE, viewer_name="Pat Example"),
            bench=(PNL, WM))
    return AppTest.from_function(script, args=(level, allowed), default_timeout=30).run()


def page_text(at):
    return " ".join([m.value for m in at.markdown] + [c.value for c in at.caption] +
                    [str(c.proto.spec) for c in at.get("plotly_chart")])


def test_share_page_for_an_admin():
    at = run_share()
    assert not at.exception, at.exception
    assert at.selectbox(key="ps_centre").options[:2] == ["Auburn", "Howick"]
    at.selectbox(key="ps_centre").set_value("Auburn").run()
    assert not at.exception, at.exception
    text = page_text(at)
    order = ["Check before sharing", "Confidential", "Headline — August 2026",
             "Accounts costing more than typical", "Accounts costing less than typical",
             "What&#x27;s going well", "Top opportunities", "Scheduling Efficiency",
             "Last 12 months", "Members trend", "Outliers vs similar centres"]
    joined = " ".join(m.value for m in at.markdown)
    pos = [joined.index(h) for h in order]
    assert pos == sorted(pos), list(zip(order, pos))
    assert "Prepared by Pat Example" in text and "data:image/png;base64" in text
    assert "Ctrl + P" in text
    # Sensitive lines in the table: own figure only.
    table = next(m.value for m in at.markdown if '<table class="share-tbl"' in m.value)
    for line in ["Rent", "Royalties", "Wages - Manager/ Admin", "Interest Expense",
                 "CRM Software"]:
        row = table.split(f">{line}</td>")[1].split("</tr>")[0]
        assert "Not benchmarked" in row and "similar centres" not in row
    # Water: only 2 of 8 similar centres use it.
    water = table.split(">Water</td>")[1].split("</tr>")[0]
    assert "No benchmark" in water and "Worth a look" not in water
    tiles = " ".join(m.value for m in at.markdown if "otile" in m.value)
    assert "Wages - Tutor" in tiles and ">Rent<" not in tiles and "Royalties" not in tiles
    assert "CRM" not in tiles                  # Auburn's CRM is 3x typical, but sensitive
    opps = next(m.value for m in at.markdown if "Top opportunities" in m.value)
    assert "tutor" in opps and "rent" not in opps.lower() and "manager" not in opps.lower()
    assert "crm" not in opps.lower()
    # Check box: royalty relief is fine; interest income is pointed out.
    check = " ".join(m.value for m in at.markdown if "check-item" in m.value)
    assert "Royalties" not in check and "interest income" in check
    # Other centres are never named.
    assert "Peer" not in text and "Howick" not in text
    assert "Typical similar centre" in text
    for span in pr.DD_SPANS:
        at.selectbox(key="ps_span").set_value(span).run()
        assert not at.exception, (span, at.exception)
    at.selectbox(key="ps_unit").set_value("$ per member").run()
    assert not at.exception, at.exception
    assert "Revenue per member / week" in page_text(at)


def test_share_page_for_a_gpm_shows_only_their_centres():
    at = run_share("gpm", ["Success Tutoring - Auburn"])
    assert not at.exception, at.exception
    assert at.selectbox(key="ps_centre").options == ["Auburn"]
    text = page_text(at)
    assert "Auburn" in text and "Peer" not in text and "Howick" not in text
    assert "Accounts costing more than typical" in text and "otile" in text  # still compared
    assert "8 centres" in text


def test_share_page_with_too_few_similar_centres():
    def script():
        import streamlit as st
        import pnl_reports as pr
        from test_pnl_share import PNL, VL, WM, AGE, REVENUE, PEERS
        small = PNL[~PNL["Location"].isin(PEERS[3:])]    # 3 Australian + 1 NZ centre left
        pr.report_share(small, WM, VL, dict(
            age_group=AGE, age_order=[], state_of=lambda r: r, show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height),
            report_filters=None, revenue=REVENUE), bench=(small, WM))
    at = AppTest.from_function(script, default_timeout=30).run()
    assert not at.exception, at.exception
    text = page_text(at)
    assert "Not enough similar centres to compare" in text
    assert not any("otile" in m.value and "Wages - Tutor" in m.value for m in at.markdown)
    table = next(m.value for m in at.markdown if '<table class="share-tbl"' in m.value)
    assert "similar centres</td>" not in table


def test_share_page_with_no_pnl_yet():
    def script():
        import pandas as pd
        import pnl_reports as pr
        pr.report_share(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {})
    at = AppTest.from_function(script).run()
    assert any("No P&L figures yet" in i.value for i in at.info)


def test_menu_key_and_position():
    src = open(pr.__file__.replace("pnl_reports.py", "app.py")).read()
    reports = src[src.index("REPORTS=["):src.index("]", src.index("REPORTS=["))]
    assert '"13 · P&L Deep Dive",\n    "14 · P&L Share",' in reports
    for key in ["1 · Campus Locations", "12 · P&L Summary", "13 · P&L Deep Dive",
                "11 · AI Outlier Analysis"]:
        assert f'"{key}"' in reports                     # existing keys unchanged
    nav = src[src.index("NAV_GROUPS = ["):src.index("ADMIN_NAV = [")]
    assert nav.index('"13 · P&L Deep Dive"') < nav.index('"14 · P&L Share"') < \
        nav.index('("Insights"')
    assert 'selected_report=="14 · P&L Share"' in src and "report_share" in src


def test_share_table_pins_gross_margin_first_in_bold():
    at = run_share()
    at.selectbox(key="ps_centre").set_value("Auburn").run()
    assert not at.exception, at.exception
    table = next(m.value for m in at.markdown if '<table class="share-tbl"' in m.value)
    first = table.split("</tr>")[1]                         # first row after the header
    assert first.startswith('<tr class="sub"><td class="l">Gross Margin</td>')
    # Auburn: 20,400 revenue (incl. 400 interest) − 10,000 tutor = 10,400, 51.0%
    assert "&#36;10,400" in first and "51.0%" in first
    assert "8 centres" in first or ">8" in first                     # compared with peers
    assert "Worth a look" in first and 'pill bad' in first           # well below typical
    opps = next(m.value for m in at.markdown if "Top opportunities" in m.value)
    assert "gross margin" not in opps.lower()                        # not a cost to cut
    assert "Gross Margin" in pr.HOW_SHARE


def test_share_table_gross_margin_for_a_gpm():
    at = run_share("gpm", ["Success Tutoring - Auburn"])
    assert not at.exception, at.exception
    table = next(m.value for m in at.markdown if '<table class="share-tbl"' in m.value)
    assert '<tr class="sub"><td class="l">Gross Margin</td>' in table
    text = page_text(at)
    assert "Howick" not in text and "Peer" not in text


def test_last_12_months_shows_gross_margin_not_ebitda():
    lm = pnl.location_months(PNL)
    t = ps.typical_by_month(lm, None, PEERS, MONTHS)
    peer_gm = [(20000 - (5600 + 100 * i) - 700) / 20000 * 100 for i in range(8)]
    assert t.loc["2026-08", "Gross margin %"] == pytest.approx(pd.Series(peer_gm).median())
    at = run_share()
    assert not at.exception, at.exception
    titles = " ".join(str(c.proto.spec) for c in at.get("plotly_chart"))
    assert "Gross margin % of revenue" in titles and "EBITDA % of revenue" not in titles
    assert titles.index("Gross margin % of revenue") < titles.index("OPEX % of revenue")
    # OPEX = every other operating cost: office 700–1,050 + CRM 450 (+ water 120 at two)
    peer_opex = [(700 + 50 * i + 450 + (120 if i < 2 else 0)) / 20000 * 100 for i in range(8)]
    assert t.loc["2026-08", "OPEX %"] == pytest.approx(pd.Series(peer_opex).median())


def _comp(rows):
    """Prepared Membership Composition rows: (location, week ending, sessions, members, price)."""
    return pd.DataFrame([{"Week ending": pd.Timestamp(w), "Location": l, "Country": "AU",
                          "Type": "Group", "Sessions/week": s, "Members": m, "Weekly price": p,
                          "Freeze": p < 10} for l, w, s, m, p in rows])


def test_composition_by_month_uses_the_last_upload_in_the_month():
    comp = _comp([("A", "2026-08-02", 1, 10, 40), ("A", "2026-08-30", 1, 10, 42),
                  ("A", "2026-08-30", 2, 30, 74)])
    cm = ps.composition_by_month(comp, ["2026-08"])
    r = cm.loc[("A", "2026-08")]
    assert r["Members"] == 40                                     # 30 Aug upload only
    assert r["ARPM"] == pytest.approx((10 * 42 + 30 * 74) / 40)
    assert r["Revenue per session"] == pytest.approx((10 * 42 + 30 * 74) / (10 + 60))
    assert ps.composition_by_month(None, ["2026-08"]).empty
    peers = _comp([(f"P{i}", "2026-08-30", 2, 20, 70 + i) for i in range(5)])
    t = ps.typical_composition(ps.composition_by_month(peers, ["2026-08"]),
                               [f"P{i}" for i in range(5)], ["2026-08"])
    assert t.at["2026-08", "ARPM"] == 72 and t.at["2026-08", "Revenue per session"] == 36
    t4 = ps.typical_composition(ps.composition_by_month(peers, ["2026-08"]),
                                [f"P{i}" for i in range(4)], ["2026-08"])
    assert t4.at["2026-08", "ARPM"] is None                       # fewer than 5 centres


def test_last_12_months_has_two_columns_of_charts():
    at = run_share()
    assert not at.exception, at.exception
    specs = [str(c.proto.spec) for c in at.get("plotly_chart")]
    joined = " ".join(specs)
    for t in ["Revenue", "Gross margin % of revenue", "OPEX % of revenue", "Total members",
              "ARPM (per member a week)", "Revenue per session"]:
        assert t in joined, t
    no_data = [s for s in specs if "Data Not Available" in s]
    assert len(no_data) == 2                     # no composition uploads: ARPM and per session
    assert not any("Total members" in s for s in no_data)


def test_composition_charts_show_this_centres_figures():
    def script():
        import streamlit as st
        import pnl_reports as pr
        from test_pnl_share import PNL, VL, WM, AGE, REVENUE, _comp
        comp = _comp([("Success Tutoring - Auburn", "2026-08-30", 2, 60, 74)])
        pr.report_share(PNL, WM, VL, dict(
            age_group=AGE, age_order=[], state_of=lambda r: r, show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height),
            report_filters=None, revenue=REVENUE, viewer_name="Pat", composition=comp),
            bench=(PNL, WM))
    at = AppTest.from_function(script, default_timeout=30).run()
    assert not at.exception, at.exception
    at.selectbox(key="ps_centre").set_value("Auburn").run()
    specs = [str(c.proto.spec) for c in at.get("plotly_chart")]
    arpm = next(s for s in specs if "ARPM (per member a week)" in s)
    assert "Data Not Available" not in arpm and "74" in arpm
    per_s = next(s for s in specs if "Revenue per session" in s)
    assert "37" in per_s
