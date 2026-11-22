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
2026-11-08,Success Tutoring Tex,US,Plan,Group,2,2,280.00,560.00
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
def test_read_export_plans_only_and_us_monthly_to_weekly():
    x = mc.read_export(CSV.encode())
    assert x["week_end"] == date(2026, 11, 8) and x["empty"] == ["Success Tutoring - Empty"]
    assert len(x["rows"]) == 12
    tex = x["rows"][x["rows"]["Centre"] == "Success Tutoring Tex"].iloc[0]
    assert tex["Weekly price"] == 70.0                           # US$280 a month / 4


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
    assert list(rows.index) == ["1×", "2×", "3+", "1:1"]
    assert rows.loc["2×", "Members"] == pytest.approx(9.6)
    assert rows.loc["1×", "Membership price"] == 42                 # own $40 < RRP
    assert rows.loc["2×", "Membership price"] == pytest.approx(520 / 14 * 2)  # own > $74
    assert rows.loc["3+", "Membership price"] == 100                # 4x at $30/session < RRP
    assert rows.loc["1:1", "Membership price"] == 300 and rows.loc["1:1", "Own price used"]
    assert t["total"] == pytest.approx(1.8 * 42 + 9.6 * 520 / 7 + 0.12 * 100 + 0.48 * 300)
    assert t["vs_today"] == pytest.approx(t["total"] - 1400)
    comp = t["composition"].set_index("Plan")
    assert comp.loc["Group 3+", "Current %"] == pytest.approx(2 / 12 * 100)
    assert comp.loc["1:1", "Variance"] == pytest.approx(2 / 12 * 100 - 4)   # current - ideal


def test_ideal_variance_is_current_minus_ideal():
    comp = mc.ideal_table(plans("Kiwi"), mc.rrp_for(mc.rrp_table(), "NZ"))["composition"]
    one = comp.set_index("Plan").loc["Group 1×"]
    assert one["Current %"] == pytest.approx(72) and one["Variance"] == pytest.approx(57)


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
