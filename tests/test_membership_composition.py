"""Membership Composition calculations: reading the export, national RRP, the current
composition table, the ideal composition and the group overview. Made-up data only."""
from datetime import date

import pandas as pd
import pytest

import membership_composition as mc

P = "Success Tutoring - "

CSV = """Date,Centre,Country,Row,Type,No. members,Sessions/week,Membership price,Total value
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,1,1,40.00,40.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,1,2,40.00,40.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,6,2,80.00,480.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,1,4,80.00,80.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,1,4,160.00,160.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,1:1,2,1,300.00,600.00
2026-11-08,Success Tutoring - Alpha,AU,Total,,12,,,1400.00
2026-11-08,Success Tutoring - Beta,AU,Plan,Group,4,2,60.00,240.00
2026-11-08,Success Tutoring - Beta,AU,Plan,Group,1,2,4.97,4.97
2026-11-08,Success Tutoring – Kiwi,NZ,Plan,Group,18,1,40.00,720.00
2026-11-08,Success Tutoring – Kiwi,NZ,Plan,Group,6,2,70.00,420.00
2026-11-08,Success Tutoring – Kiwi,NZ,Plan,Group,1,2,4.97,4.97
2026-11-08,Success Tutoring Tex,US,Plan,Group,2,2,70.00,140.00
2026-11-08,Success Tutoring - Empty,AU,Total,,0,,,0.00
"""


def week_from(csv=CSV):
    x = mc.read_export(csv.encode())
    locs = {n: P + n.replace("Success Tutoring", "").strip(" -–") for n in x["rows"]["Centre"]}
    return mc.prepare(pd.DataFrame(mc.records(x["rows"], locs, x["week_end"])))


W = week_from()
AU = mc.rrp_for(mc.rrp_table(), "AU")


def plans(name):
    return W[W["Location"] == P + name]


# ── Reading the export ────────────────────────────────────────────────────────
def test_read_export_plans_only_and_us_prices_as_weekly():
    x = mc.read_export(CSV.encode())
    assert x["week_end"] == date(2026, 11, 8) and x["empty"] == ["Success Tutoring - Empty"]
    assert len(x["rows"]) == 12
    tex = x["rows"][x["rows"]["Centre"] == "Success Tutoring Tex"].iloc[0]
    assert tex["Weekly price"] == 70.0                           # exported weekly now
    assert mc.looks_monthly(x["rows"]).empty
    old = mc.read_export(CSV.replace("Tex,US,Plan,Group,2,2,70.00,140.00",
                                     "Tex,US,Plan,Group,2,2,280.00,560.00").encode())["rows"]
    assert list(mc.looks_monthly(old)["Centre"]) == ["Success Tutoring Tex"]   # $140 a session


def test_read_export_rejects_other_files():
    with pytest.raises(mc.CompositionError, match="Membership Composition export"):
        mc.read_export(b"Date,Location,# Active members\n2026-11-08,X,5\n")


def test_records_keep_week_as_text_and_skip_unmatched():
    x = mc.read_export(CSV.encode())
    recs = mc.records(x["rows"], {"Success Tutoring - Alpha": P + "Alpha"}, x["week_end"])
    assert len(recs) == 6 and recs[0]["Week ending"] == "'2026-11-08"
    assert list(recs[0]) == mc.HEADERS


def test_composition_scope_alias():
    master = pd.DataFrame({"Location": [P + "Alpha", P + "Kiwi Town"]})
    found, sugg = mc.match_centres(["Success Tutoring – Kiwi", "Success Tutoring - Alpha"],
                                   master, {"Composition: Success Tutoring – Kiwi": P + "Kiwi Town"})
    assert found == {"Success Tutoring – Kiwi": P + "Kiwi Town",
                     "Success Tutoring - Alpha": P + "Alpha"} and sugg == {}


# ── RRP ───────────────────────────────────────────────────────────────────────
def test_national_rrp_except_us_regions():
    t = mc.rrp_table()
    assert mc.rrp_for(t, "Australia", "New South Wales")["mkt"] == (42, 74, 100)
    assert mc.rrp_for(t, "NZ", "Auckland (N)")["mkt"] == (47, 82, 110)
    assert mc.rrp_for(t, "CA", "British Columbia")["mkt"] == (51, 89, 121)
    assert mc.rrp_for(t, "US", "Texas")["mkt"] == (49, 85, 115)
    assert mc.rrp_for(t, "US", "Ohio")["mkt"] == (47, 82, 110)
    assert mc.rrp_for(t, "GB") is None


def test_state_rows_in_the_rrp_tab_are_ignored_for_australia():
    tab = pd.DataFrame([mc._rrp("AU", "All", (42, 74, 100), (36, 63, 85), 90),
                        mc._rrp("AU", "New South Wales", (46, 81, 110), (39, 69, 93), 99)])
    assert mc.rrp_for(mc.rrp_table(tab), "AU", "New South Wales")["mkt"] == (42, 74, 100)


def test_rrp_price_rules():
    assert mc.rrp_price(AU, "Group", 2) == 74
    assert mc.rrp_price(AU, "Group", 4) == pytest.approx(100 / 3 * 4)
    assert mc.rrp_price(AU, "1:1", 2) == 180


# ── Table 1 ───────────────────────────────────────────────────────────────────
def test_current_table_rows_totals_and_variances():
    t = mc.current_table(plans("Alpha"), AU, mc.plan_medians(W))
    rows = t["rows"]
    assert list(rows["Type"]) == ["Group"] * 5 + ["1:1"]            # every visit count shown
    assert list(rows["Sessions/week"]) == [1, 2, 2, 4, 4, 1]
    assert list(rows["Total sessions"]) == [1, 2, 12, 4, 4, 2] and t["sessions"] == 25
    two = rows.iloc[2]
    assert two["Revenue per session"] == 40 and two["RRP variance"] == 6
    assert two["RRP variance %"] == pytest.approx(8.1, abs=0.1)
    # AU 2x prices: Alpha 1 x $40 + 6 x $80, Beta 4 x $60 -> median of 11 = $80
    assert two["Median"] == 80 and two["Median variance"] == 0
    assert rows.iloc[1]["Median variance"] == -40
    assert rows.iloc[3]["RRP variance"] == pytest.approx(80 - 400 / 3)
    assert pd.isna(rows.iloc[0]["Median"])                           # only Alpha has AU 1x
    assert t["members"] == 12 and t["total"] == 1400 and t["arpm"] == pytest.approx(116.67, 0.01)
    assert t["per_session"] == pytest.approx(1400 / 25)
    comp = t["composition"].set_index("Plan")
    assert comp.loc["Group 2×", "Members"] == 7
    assert comp.loc["Group 4×", "% of members"] == pytest.approx(2 / 12 * 100)
    assert list(t["composition"]["Plan"]) == ["Group 1×", "Group 2×", "Group 4×", "1:1"]


def test_under_10_left_out_of_rows_but_in_totals():
    t = mc.current_table(plans("Beta"), AU, mc.plan_medians(W))
    assert len(t["rows"]) == 1 and t["members"] == 5
    assert t["total"] == pytest.approx(244.97) and t["freeze_members"] == 1
    assert t["composition"].set_index("Plan").loc["Group 2×", "Members"] == 5


# ── Table 2 ───────────────────────────────────────────────────────────────────
def test_ideal_table_target_mix_and_own_higher_price():
    t = mc.ideal_table(plans("Alpha"), AU)
    rows = t["rows"].set_index("Plan")
    assert list(rows.index) == ["1×", "2×", "3+", "1:1", mc.NOT_RECOMMENDED]
    assert list(rows["Members"]) == [2, 8, 0, 0, 2]                 # whole, add up to 12
    assert rows.loc["2×", "Members"] == pytest.approx(8.0)          # 10 recommended x 80%
    assert rows.loc["1×", "Membership price"] == 42                 # own $40 < RRP
    assert rows.loc["2×", "Membership price"] == pytest.approx(520 / 14 * 2)  # own > $74
    assert rows.loc["3+", "Membership price"] == 100                # no 3x members: RRP
    assert rows.loc["1:1", "Membership price"] == 300 and rows.loc["1:1", "Own price used"]
    nr = rows.loc[mc.NOT_RECOMMENDED]                               # the two 4x members as-is
    assert nr["Members"] == 2 and nr["Total value"] == 240 and nr["Total sessions"] == 8
    assert t["total"] == pytest.approx(2 * 42 + 8 * 520 / 7 + 240)
    assert t["vs_today"] == pytest.approx(t["total"] - 1400)
    comp = t["composition"].set_index("Plan")
    assert comp.loc["Group 3+", "Current %"] == 0
    assert comp.loc["Group " + mc.NOT_RECOMMENDED, "Current %"] == pytest.approx(2 / 12 * 100)
    assert comp.loc["1:1", "Variance"] == pytest.approx(2 / 12 * 100)      # ideal rounds to 0


def test_ideal_variance_is_current_minus_ideal():
    comp = mc.ideal_table(plans("Kiwi"), mc.rrp_for(mc.rrp_table(), "NZ"))["composition"]
    one = comp.set_index("Plan").loc["Group 1×"]
    assert one["Current %"] == pytest.approx(72)        # ideal 1x: 4 of 25 whole members
    assert one["Ideal %"] == pytest.approx(16) and one["Variance"] == pytest.approx(56)


# ── Group overview ────────────────────────────────────────────────────────────
def test_group_table_sorted_and_single_country_median_blank():
    vl = pd.DataFrame({mc.LOC: [P + "Tex"], "Region": ["Texas"]})
    g = mc.group_table(W, vl, mc.rrp_table())
    assert list(g["% of RRP"]) == sorted(g["% of RRP"])
    tex = g[g["Centre"] == "Tex"].iloc[0]
    assert tex["RRP used"] == "US · Texas" and tex["% of RRP"] == pytest.approx(70 / 85 * 100)
    assert pd.isna(tex["ARPM vs country median"])
    au = g[g["Country"] == "AU"].set_index("Centre")
    assert au.loc["Alpha", "ARPM vs country median"] == pytest.approx(
        1400 / 12 - (1400 / 12 + 244.97 / 5) / 2)
    beta = au.loc["Beta"]
    assert beta["% of RRP"] == pytest.approx(240 / 296 * 100)          # $4.97 left out


def test_composition_columns_mix_1to1_sessions():
    plans = pd.DataFrame({"Country": "AU", "Type": ["1:1", "1:1", "Group"],
                          "Sessions/week": [1, 2, 2], "Members": [3, 1, 4],
                          "Weekly price": [100.0, 180.0, 70.0], "Freeze": False})
    comp = mc.current_table(plans, None, {})["composition"].set_index("Plan")
    one = comp.loc["1:1"]                       # 3 x 1 session + 1 x 2 sessions
    assert one["Members"] == 4 and one["Total sessions"] == 5
    assert one["Total value"] == 480 and one["Average price"] == 120
    assert one["Revenue per session"] == 480 / 5
    assert comp.loc["Group 2×", "Revenue per session"] == 35


def test_target_fill_rises_smoothly_with_members():
    got = {m: mc.target_fill(m) for m in (5, 20, 30, 40, 60, 80, 100, 120, 135, 150, 165,
                                           180, 300)}
    assert got == {5: 3.0, 20: 3.0, 30: 3.3, 40: 3.5, 60: 3.9, 80: 4.3, 100: 4.6, 120: 5.0,
                   135: 5.3, 150: 5.5, 165: 5.8, 180: 6.0, 300: 6.0}


def test_scheduling_hubs_tutor_cost_and_margin():
    r = mc.scheduling(120, 200, 10, 8000, "AU")          # 200 group sessions at 5.0 + 10 x 1:1
    assert r["Target fill"] == 5.0 and r["Hubs needed a week"] == 50
    assert r["Tutor cost"] == 50 * 37.90 and r["Margin a week"] == 8000 - 50 * 37.90
    assert r["Margin per member a week"] == (8000 - 50 * 37.90) / 120
    assert r["Tutor cost %"] == 50 * 37.90 / 8000 * 100
    nz = mc.scheduling(120, 200, 10, 8000, "NZ")         # no NZ tutor rate yet
    assert nz["Hubs needed a week"] == 50 and nz["Tutor cost"] is None
    assert nz["Margin per member a week"] is None


def test_weekly_totals_and_trend_measures():
    comp = pd.DataFrame({"Week ending": pd.to_datetime(["2026-11-08"] * 3),
                         "Location": "X", "Country": "AU", "Type": ["Group", "Group", "1:1"],
                         "Sessions/week": [1, 2, 1], "Members": [2, 6, 2],
                         "Weekly price": [42.0, 74.0, 90.0], "Freeze": False})
    w = mc.weekly_totals(comp)
    r = w.iloc[0]
    assert (r["Members"], r["Sessions"], r["One"], r["Value"]) == (10, 16, 2, 708)
    assert (r["1×"], r["2×"], r["3+"], r["1:1"]) == (2, 6, 0, 2)
    assert r["Margin"] == 708 - (14 / 3.0 + 2) * 37.90
    m = mc.trend_measures(w).iloc[0]
    assert m["Sessions per member"] == 1.6 and m["ARPM"] == 70.8 and m["2×"] == 60


def test_where_the_money_is_adds_up():
    plans = mc.prepare(pd.DataFrame({
        "Week ending": "'2026-11-22", "Location": "X", "Country": "AU",
        "Type": ["Group", "Group", "1:1"], "Sessions/week": [1, 2, 1], "Members": [10, 20, 2],
        "Weekly price": [35.0, 60.0, 90.0]}))
    rrp = mc.rrp_for(mc.rrp_table(None), "AU", None)
    ideal = mc.ideal_table(plans, rrp)
    r = mc.where_the_money_is(plans, rrp, ideal, "AU", run=(30, 32))
    t, a = r["today"], r["ideal"]
    assert r["target"] == 3.3 and abs(r["measured_fill"] - 50 / 28) < 1e-9
    assert r["fill"] == 3.3                                   # target beats measured 1.7
    need = 50 / 3.3 + 2
    assert abs(r["hubs_needed"] - need) < 1e-9 and t["Hubs a week"] == 30
    gains = dict((g[0], g[1]) for g in r["gains"])
    assert gains["Price"] == 10 * (42 - 35) + 20 * (74 - 60)  # up to RRP
    assert abs(gains["Scheduling"] - (30 - need) * 37.90) < 1e-9
    assert abs(t["Margin a week"] + sum(gains.values()) - a["Margin a week"]) < 1e-6
    nz = mc.where_the_money_is(plans, rrp, ideal, "NZ", run=(30, 32))
    assert nz["today"]["Margin a week"] is None and "Scheduling" not in dict(
        (g[0], g[1]) for g in nz["gains"])


def test_4x_plans_not_recommended_no_rrp_uplift_and_left_out_of_pct_rrp():
    assert mc.not_recommended("Group", 4) and mc.not_recommended("Group", 8)
    assert not mc.not_recommended("Group", 3) and not mc.not_recommended("1:1", 4)
    p = plans("Alpha")
    r = mc.where_the_money_is(p, AU, mc.ideal_table(p, AU), "AU")
    price = dict((g[0], g[1]) for g in r["gains"])["Price"]
    assert price == pytest.approx(2 + 34)          # 1x $40->$42, 2x $40->$74; 4x untouched
    assert r["not_recommended"] == {"members": 2, "share": pytest.approx(2 / 12 * 100),
                                    "value": 240, "sessions": 8}
    p2 = mc.ideal_table(p, AU)["rows"].set_index("Plan").loc["2×", "Membership price"]
    pl = r["plans"]                                # 2 members: 8 sessions -> 4 on 2x
    assert pl["sessions"] == 4 and pl["fees"] == pytest.approx(2 * p2 - 240)
    assert pl["net"] == pytest.approx(pl["fees"] + pl["hubs"] * 37.90)
    th = r["theory"]
    assert th["Sessions per member"] == pytest.approx(22 / 12)   # 2 on 1x, 10 on 2x
    assert th["Hubs a week"] == pytest.approx(22 / 3.0)
    g = mc.group_table(W, None, mc.rrp_table()).set_index("Centre")
    assert g.loc["Alpha", "Members on 4×+"] == 2
    paid, at = 40 + 40 + 480 + 600, 42 + 74 + 6 * 74 + 2 * 90 * 1
    assert g.loc["Alpha", "% of RRP"] == pytest.approx(paid / at * 100)


def test_whole_split_adds_up():
    assert mc.whole_split(43, mc.TARGET_MIX) == {"1×": 7, "2×": 34, "3+": 0, "1:1": 2}
    assert mc.whole_split(56, mc.TARGET_MIX) == {"1×": 8, "2×": 45, "3+": 1, "1:1": 2}
    for n in range(0, 200):
        assert sum(mc.whole_split(n, mc.TARGET_MIX).values()) == n
