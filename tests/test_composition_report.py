"""Membership Composition report page: admin group overview and centre deep dive, the
admin-only rule, and its menu entry. Made-up centres and prices only."""
import pandas as pd
from streamlit.testing.v1 import AppTest

import membership_composition as mc

P = "Success Tutoring - "
LOC = mc.LOC


def _plans(week, loc, country, plans):
    return [{"Week ending": "'" + week, "Location": P + loc, "Country": country, "Type": t,
             "Sessions/week": s, "Members": m, "Weekly price": p, "Export name": loc}
            for t, s, m, p in plans]


def _week(week, extra_2x=0):
    rows = []
    # Ten established AU centres around RRP, one well below, one NZ, one young.
    for i in range(10):
        rows += _plans(week, f"Peer {i}", "AU", [("Group", 1, 2, 44), ("Group", 2, 10 + i, 77)])
    rows += _plans(week, "Lowprice", "AU", [("Group", 1, 12, 35), ("Group", 2, 4 + extra_2x, 55),
                                            ("Group", 4, 2, 60), ("Group", 2, 1, 4.97),
                                            ("1:1", 1, 1, 90)])
    rows += _plans(week, "Kiwi", "NZ", [("Group", 2, 6, 60)])
    rows += _plans(week, "Newbie", "AU", [("Group", 2, 5, 65)])
    return rows


COMP = pd.DataFrame(_week("2026-11-01") + _week("2026-11-08", extra_2x=2))
ALL = [P + f"Peer {i}" for i in range(10)] + [P + "Lowprice", P + "Kiwi", P + "Newbie"]
VL = pd.DataFrame({LOC: ALL, "Region": ["Victoria"] * 11 + ["Auckland (N)", "Victoria"],
                   "Country": ["Australia"] * 11 + ["New Zealand", "Australia"],
                   "Weeks old": [60] * 12 + [8], "Age (Months)": [14] * 12 + [2]})
WM = pd.DataFrame([{LOC: loc, "Date": d, "# Active members": 20, "# New members": 2}
                   for loc in ALL for d in pd.date_range("2026-09-06", periods=10, freq="W")])
PNL = pd.DataFrame({"Period": ["2026-09"], "Currency": ["NZD"], "FX rate to AUD": [0.9]})


def age_group(m):
    return "0–3 months" if m <= 3 else "12–24 months"


def run(level="admin", comp=COMP):
    def script(level, comp):
        import pandas as pd
        import streamlit as st
        import composition_report as cr
        from test_composition_report import VL, WM, PNL, age_group
        st.session_state["access_level"] = level
        cr.report_composition(comp, WM, VL, pd.DataFrame(), PNL, dict(
            age_group=age_group, show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height)))
    return AppTest.from_function(script, args=(level, comp), default_timeout=60).run()


def test_group_overview_ranks_by_gap_and_flags():
    at = run()
    assert not at.exception, at.exception
    text = " ".join(m.value for m in at.markdown)
    assert "Biggest gaps and opportunities" in text and "Lowprice" in text
    rank = at.dataframe[0].value
    assert rank.iloc[0]["Centre"] == "Lowprice"                 # biggest RRP gap first
    assert "Newbie" not in set(rank["Centre"])                  # young: separate table
    young = at.dataframe[1].value
    assert list(young["Centre"]) == ["Newbie"]
    low = rank[rank["Centre"] == "Lowprice"].iloc[0]
    assert "⚑" in low["Flag"] and "#1 biggest RRP gap" in low["Flag"]
    assert any("beyond 2 SD" in f for f in rank["Flag"])
    kiwi = rank[rank["Centre"] == "Kiwi"].iloc[0]
    assert kiwi["RRP gap /wk"] == 6 * 22 and kiwi["RRP gap /wk AUD"] == 6 * 22 * 0.9
    assert len(at.get("plotly_chart")) == 2


def test_group_overview_country_filter_and_week():
    at = run()
    at.selectbox(key="comp_country").set_value("NZ").run()
    assert not at.exception, at.exception
    assert list(at.dataframe[0].value["Centre"]) == ["Kiwi"]
    at.selectbox(key="comp_week_pick").set_value("2026-11-01").run()
    assert not at.exception, at.exception


def test_centre_deep_dive_trends_pathway_and_commentary():
    at = run()
    at.radio(key="comp_view").set_value("Centre deep dive").run()
    at.selectbox(key="comp_centre").set_value("Lowprice").run()
    assert not at.exception, at.exception
    text = " ".join(m.value for m in at.markdown)
    assert "What stands out" in text and "Recommended pathway to RRP" in text
    assert "1× share" in text and "phased out" in text and "freeze fee" in text
    charts = at.get("plotly_chart")
    assert len(charts) >= 3                                     # mix, price, pathway
    steps = next(d.value for d in at.dataframe if "Step" in d.value.columns)
    assert steps.iloc[0]["Step"] == "Price every new member at RRP"
    assert "Recruit new members onto 2×" in set(steps["Step"])
    assert not any(s.startswith("Lift") or "1:1" in s for s in steps["Step"])
    assert any("left as they are" in c.value for c in at.caption)
    gaps = next(d.value for d in at.dataframe if "Plan" in d.value.columns)
    assert list(gaps["Gap /wk"]) == sorted(gaps["Gap /wk"], reverse=True)
    assert gaps.set_index("Plan").loc["4+×", "Note"] == "Phase out"
    assert any("⚑" in w.value for w in at.warning)


def test_gpm_cannot_see_the_report():
    at = run("gpm")
    assert not at.exception, at.exception
    assert any("admins only" in e.value for e in at.error)
    assert not at.dataframe and not at.get("plotly_chart")


def test_no_data_yet():
    at = run(comp=pd.DataFrame())
    assert any("No Membership Composition data yet" in i.value for i in at.info)


def test_menu_entry_admin_only_and_upload_page():
    src = open(mc.__file__.replace("membership_composition.py", "app.py")).read()
    reports = src[src.index("REPORTS=["):src.index("]", src.index("REPORTS=["))]
    assert '"17 · Membership Composition"' in reports
    nav = src[src.index("NAV_GROUPS = ["):src.index("ADMIN_NAV = [")]
    membership = nav[nav.index('("Membership"'):nav.index('("Growth"')]
    assert membership.index('"3 · Membership by Age"') < \
        membership.index('"17 · Membership Composition"')
    lock = src[src.index("ADMIN_ONLY_REPORTS = {"):src.index("if not REPORTS:")]
    assert '"17 · Membership Composition"' in lock
    assert 'selected_report=="17 · Membership Composition"' in src
    assert '"⚙ Composition Upload": admin_pages.page_composition_upload' in src
