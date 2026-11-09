"""Membership Composition: reading the export, RRP, mix and gaps, flags, pathway and
the rule-based commentary. Made-up centres and prices only."""
from datetime import date

import pandas as pd
import pytest

import membership_composition as mc

LOC = mc.LOC
P = "Success Tutoring - "

CSV = """Date,Centre,Country,Row,Type,No. members,Sessions/week,Membership price,Total value
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,4,1,40.00,160.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,10,2,81.00,810.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,2,4,100.00,200.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,1:1,1,1,90.00,90.00
2026-11-08,Success Tutoring - Alpha,AU,Plan,Group,1,2,4.97,4.97
2026-11-08,Success Tutoring - Alpha,AU,Total,,18,,,1264.97
2026-11-08,Success Tutoring - Alpha,AU,ARPM,,,,,70.28
2026-11-08,Success Tutoring – Kiwi,NZ,Plan,Group,6,2,60.00,360.00
2026-11-08,Success Tutoring – Kiwi,NZ,Total,,6,,,360.00
2026-11-08,Success Tutoring Tex,US,Plan,Group,2,2,280.00,560.00
2026-11-08,Success Tutoring Tex,US,Total,,2,,,560.00
2026-11-08,Success Tutoring - Empty,AU,Total,,0,,,0.00
"""


def comp_from(csv=CSV, week=None):
    x = mc.read_export(csv.encode())
    locs = {n: P + n.replace("Success Tutoring", "").strip(" -–") for n in x["rows"]["Centre"]}
    return mc.prepare(pd.DataFrame(mc.records(x["rows"], locs, week or x["week_end"])))


# ── Reading the export ────────────────────────────────────────────────────────
def test_read_export_plans_only_and_us_monthly_to_weekly():
    x = mc.read_export(CSV.encode())
    assert x["week_end"] == date(2026, 11, 8)
    rows = x["rows"]
    assert len(rows) == 7 and x["empty"] == ["Success Tutoring - Empty"]
    tex = rows[rows["Centre"] == "Success Tutoring Tex"].iloc[0]
    assert tex["Weekly price"] == 70.0                          # US$280 a month / 4
    assert set(rows["Type"]) == {"Group", "1:1"}


def test_read_export_rejects_other_files():
    with pytest.raises(mc.CompositionError, match="Membership Composition export"):
        mc.read_export(b"Date,Location,# Active members\n2026-11-08,X,5\n")


def test_records_keep_week_as_text_and_skip_unmatched():
    x = mc.read_export(CSV.encode())
    recs = mc.records(x["rows"], {"Success Tutoring - Alpha": P + "Alpha"}, x["week_end"])
    assert len(recs) == 5 and all(r["Location"] == P + "Alpha" for r in recs)
    assert recs[0]["Week ending"] == "'2026-11-08"
    assert list(recs[0]) == mc.HEADERS


def test_prepare_mix_and_freeze():
    comp = comp_from()
    a = comp[comp["Location"] == P + "Alpha"]
    assert sorted(a["Mix"].unique()) == ["1:1", "1×", "2×", "4+×"]
    assert a.loc[a["Freeze"], "Weekly price"].tolist() == [4.97]
    assert mc.prepare(pd.DataFrame()).empty


def test_composition_scope_alias():
    master = pd.DataFrame({"Location": [P + "Alpha", P + "Kiwi Town"]})
    found, sugg = mc.match_centres(["Success Tutoring – Kiwi", "Success Tutoring - Alpha"],
                                   master, {"Composition: Success Tutoring – Kiwi": P + "Kiwi Town"})
    assert found == {"Success Tutoring – Kiwi": P + "Kiwi Town",
                     "Success Tutoring - Alpha": P + "Alpha"} and sugg == {}


# ── RRP ───────────────────────────────────────────────────────────────────────
def test_rrp_by_state_country_and_fallback():
    t = mc.rrp_table()
    assert mc.rrp_for(t, "Australia", "New South Wales")["mkt"] == (46, 81, 110)
    assert mc.rrp_for(t, "AU", "WA")["mkt"] == (44, 77, 104)            # abbreviation
    assert mc.rrp_for(t, "AU", "")["mkt"] == (42, 74, 100)              # national
    assert mc.rrp_for(t, "New Zealand", "Auckland (N)")["mkt"] == (47, 82, 110)
    assert mc.rrp_for(t, "US", "Texas")["mkt"] == (49, 85, 115)
    assert mc.rrp_for(t, "US", "Ohio")["mkt"] == (47, 82, 110)
    assert mc.rrp_for(t, "GB", "") is None


def test_rrp_tab_overrides_defaults():
    tab = pd.DataFrame([mc._rrp("AU", "All", (50, 90, 120), (40, 80, 100), 95)])
    assert mc.rrp_for(mc.rrp_table(tab), "AU", "Victoria")["mkt"] == (50, 90, 120)


def test_rrp_price_rules():
    r = mc.rrp_for(mc.rrp_table(), "AU", "New South Wales")
    assert mc.rrp_price(r, "Group", 2) == 81
    assert mc.rrp_price(r, "Group", 2, foundational=True) == 69
    assert mc.rrp_price(r, "Group", 4) == pytest.approx(110 / 3 * 4)    # 3x rate per session
    assert mc.rrp_price(r, "1:1", 2) == 198                             # per session


# ── Metrics ───────────────────────────────────────────────────────────────────
def test_centre_metrics_mix_prices_and_gap():
    comp = comp_from()
    r = mc.rrp_for(mc.rrp_table(), "AU", "New South Wales")
    m = mc.centre_metrics(comp[comp["Location"] == P + "Alpha"], r)
    assert m["Members"] == 18 and m["n 2×"] == 11 and m["% 2×"] == pytest.approx(11 / 18 * 100)
    assert m["Freeze"] == 1 and m["Avg 2×"] == 81                        # $4.97 left out
    # Below RRP: 4 on 1x ($6 each), 2 on 4x (146.67-100 each), 1:1 ($99-90)
    assert m["Below RRP"] == 7
    assert m["RRP gap"] == pytest.approx(4 * 6 + 2 * (110 / 3 * 4 - 100) + 9)
    assert m["Below foundational"] == 2                                 # 4x below 4 x 31
    assert m["ARPM"] == pytest.approx((160 + 810 + 200 + 90) / 17)


def test_summary_benchmarks_flags_and_fx():
    comp = comp_from()
    vl = pd.DataFrame({LOC: [P + "Alpha", P + "Kiwi", P + "Tex"],
                       "Region": ["New South Wales", "Auckland (N)", "Texas"],
                       "Weeks old": [40, 30, 10], "Age (Months)": [9, 7, 2]})
    s = mc.summary(comp, vl, mc.rrp_table(), age_group=lambda m: "6–12 months",
                   fx={"NZD": 0.9})
    kiwi = s[s["Centre"] == "Kiwi"].iloc[0]
    assert kiwi["RRP gap"] == pytest.approx(6 * 22)
    assert kiwi["RRP gap AUD"] == pytest.approx(6 * 22 * 0.9)
    assert pd.isna(s[s["Centre"] == "Tex"].iloc[0]["RRP gap AUD"])      # no USD rate
    assert bool(s[s["Centre"] == "Tex"].iloc[0]["Young"])
    b = mc.benchmarks(s)
    assert b["6–12 months"]["n"] == 2                                   # young one left out
    f = mc.flags(s)
    assert "#1 biggest RRP gap" in f[P + "Alpha"]


def test_two_sd_outlier_flag():
    rows = [{"Location": f"L{i}", "Young": False, "% of RRP": 90.0 + i % 3, "% 2×": 70.0,
             "% 1×": 20.0, "% below RRP": 50.0, "RRP gap AUD": 10.0} for i in range(10)]
    rows.append({"Location": "Low", "Young": False, "% of RRP": 60.0, "% 2×": 70.0,
                 "% 1×": 20.0, "% below RRP": 50.0, "RRP gap AUD": 0.0})
    f = mc.flags(pd.DataFrame(rows))
    assert any("% of RRP 60%" in x for x in f["Low"])
    assert sum("biggest RRP gap" in x for v in f.values() for x in v) == 5


def test_history_by_week():
    a = comp_from()
    b = comp_from(CSV.replace("2026-11-08", "2026-11-15"))
    h = mc.history(pd.concat([a, b]), P + "Kiwi", mc.rrp_for(mc.rrp_table(), "NZ", ""))
    assert len(h) == 2 and list(h["% 2×"]) == [100.0, 100.0]


# ── Pathway ───────────────────────────────────────────────────────────────────
def test_pathway_july_reviews_capped_and_1to1_4plus_untouched():
    comp = comp_from()
    r = mc.rrp_for(mc.rrp_table(), "AU", "New South Wales")
    p = mc.pathway(comp[comp["Location"] == P + "Alpha"], r, age_months=20,
                   today=date(2026, 11, 9))
    steps = [s["Step"] for s in p["steps"]]
    assert steps[0] == "Price every new member at RRP"
    assert "End foundational pricing" not in steps                     # over 12 months
    reviews = [s for s in p["steps"] if s["Step"].startswith("July")]
    assert reviews[0]["When"] == "Jul 2027"
    assert reviews[0]["Weekly gain"] == pytest.approx(4 * 40 * 0.05)    # 1x: $40 -> $42
    assert reviews[-1]["When"] == "Jul 2029"                           # 42 -> 44.1 -> 46
    assert p["not_moved"] == 3 and p["still_below"] == 0               # 1:1 and 4+ stay


def test_pathway_foundational_end_and_new_2x():
    comp = comp_from()
    r = mc.rrp_for(mc.rrp_table(), "NZ", "")
    p = mc.pathway(comp[comp["Location"] == P + "Kiwi"], r, age_months=4,
                   new_per_week=2, target_2x=50, today=date(2026, 11, 9))
    f = next(s for s in p["steps"] if s["Step"] == "End foundational pricing")
    assert f["When"] == "Jul 2027" and f["Weekly gain"] == pytest.approx(6 * 22)
    assert p["projection"][-1]["% of RRP"] == pytest.approx(100)
    assert p["new_2x"] == 0                                             # already 100% on 2x

    alpha = comp[comp["Location"] == P + "Alpha"]
    p = mc.pathway(alpha, mc.rrp_for(mc.rrp_table(), "AU", "NSW"), age_months=20,
                   new_per_week=2, target_2x=80, today=date(2026, 11, 9))
    # 11 of 18 on 2x; (11 + x) / (18 + x) >= 0.8 -> x = 17
    assert p["new_2x"] == 17 and p["weeks_to_2x"] == 9


def test_next_july():
    assert mc.next_july(date(2026, 11, 9)) == date(2027, 7, 1)
    assert mc.next_july(date(2027, 3, 1)) == date(2027, 7, 1)


# ── Commentary ────────────────────────────────────────────────────────────────
def test_commentary_sentences():
    row = {"Currency": "NZD", "Age group": "3–6 months", "% 2×": 34.0, "% 1×": 62.0,
           "% of RRP": 80.0, "Below RRP": 28, "RRP gap": 300.0, "Below foundational": 3,
           "Avg 1×": 40.0, "n 1×": 18, "RRP 1×": 47, "Avg 2×": 65.0, "n 2×": 10,
           "RRP 2×": 82, "n 4+×": 1, "Freeze": 1}
    text = " ".join(mc.commentary(row, {"3–6 months": {"% 2×": 70.0, "% 1×": 20.0}}))
    assert "34% of members are on 2× a week, 36 points below" in text
    assert "1× share (62%) is high" in text
    assert "$300 NZD a week ($15,600 NZD a year)" in text
    assert "biggest single gap is the 2× plan" in text                 # 17 x 10 > 7 x 18
    assert "phased out" in text and "freeze fee" in text
    assert "foundational" in text


def test_commentary_at_rrp():
    text = " ".join(mc.commentary({"% of RRP": 105.0, "Currency": "AUD"}, {}))
    assert "at or above RRP" in text


def test_new_members_coverage_and_fx():
    wm = pd.DataFrame({LOC: ["A"] * 10, "Date": pd.date_range("2026-09-06", periods=10, freq="W"),
                       "# New members": [0, 0] + [2] * 8, "# Active members": range(10)})
    assert mc.new_members_per_week(wm, "A") == 2
    assert mc.active_members(wm, "A") == 9
    pnl = pd.DataFrame({"Period": ["2026-08", "2026-09"], "Currency": ["NZD", "NZD"],
                        "FX rate to AUD": [0.9, 0.91]})
    assert mc.latest_fx(pnl) == {"NZD": 0.91}
