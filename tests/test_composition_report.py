"""Membership Composition report page: group overview, centre tables, GPMs see only their own centres
and the menu entry. Made-up data only."""
import pandas as pd
from streamlit.testing.v1 import AppTest

import membership_composition as mc
from test_membership_composition import W

P = "Success Tutoring - "
SHEET = pd.DataFrame([{**r, "Week ending": "'" + r["Week ending"].strftime("%Y-%m-%d")}
                      for r in W.drop(columns="Freeze").to_dict("records")])
LATER = SHEET.assign(**{"Week ending": "'2026-11-15"})
COMP = pd.concat([SHEET, LATER])
VL = pd.DataFrame({mc.LOC: [P + "Tex"], "Region": ["Texas"]})


def run(level="admin", comp=COMP, allowed=None):
    def script(level, comp, allowed):
        import pandas as pd
        import streamlit as st
        import composition_report as cr
        from test_composition_report import VL
        st.session_state["access_level"] = level
        cr.report_composition(comp, VL, pd.DataFrame(), allowed=allowed)
    return AppTest.from_function(script, args=(level, comp, allowed),
                                 default_timeout=60).run()


def test_group_overview():
    at = run()
    assert not at.exception, at.exception
    g = at.dataframe[0].value
    assert list(g.columns) == ["Centre", "Country", "RRP used", "Members", "ARPM",
                               "Revenue per session", "% of RRP", "ARPM vs country median"]
    assert list(g["% of RRP"]) == sorted(g["% of RRP"])
    at.selectbox(key="comp_country").set_value("NZ").run()
    assert list(at.dataframe[0].value["Centre"]) == ["Kiwi"]
    at.selectbox(key="comp_week_pick").set_value("2026-11-08").run()
    assert not at.exception, at.exception


def test_centre_tables():
    at = run()
    at.radio(key="comp_view").set_value("Centre deep dive").run()
    at.selectbox(key="comp_centre").set_value("Alpha").run()
    assert not at.exception, at.exception
    html = " ".join(m.value for m in at.markdown)
    assert "Success Tutoring - Alpha" in html and "Ideal composition to achieve RRP" in html
    for h in ["No. members", "Sessions/week", "Membership price", "Total value",
              "Revenue per session", "RRP variance", "Variance from median"]:
        assert f"<th>{h}</th>" in html
    assert "% of members" in html and "Current %" in html and "Ideal %" in html
    comp = html.split("<th>Composition</th>")[1].split("</table>")[0]
    for h in ["Sessions/week", "Average membership price", "Total value", "Revenue per session"]:
        assert f"<th>{h}</th>" in comp
    g2 = comp.split(">Group 2×</td>")[1].split("</tr>")[0]   # 1 at $40 + 6 at $80
    assert all(v in g2 for v in [">7<", ">2<", "$74.29", "$520.00", "$37.14"])
    assert "$1,400.00" in html and "$116.67" in html and "$56.00" in html
    assert "+$6.00 (+8%)" in html and "−$40.00 (−50%)" in html
    assert "−$455" not in html or "vs today" in html
    assert "+167 pts" not in html and "+13 pts" in html              # 1:1: 17% - 4%
    assert "own average price" in html
    assert not at.dataframe and not at.get("plotly_chart")            # tables only


def test_centre_note_for_prices_under_10():
    at = run()
    at.radio(key="comp_view").set_value("Centre deep dive").run()
    at.selectbox(key="comp_centre").set_value("Kiwi").run()
    html = " ".join(m.value for m in at.markdown)
    assert "Members under $10" in html and ": 1 ($4.97 a week)" in html
    assert "$4.97" not in html.split("Members under $10")[0]          # not in the rows


def test_gpm_sees_only_their_own_centres():
    at = run("gpm", allowed=[P + "Alpha"])
    assert not at.exception, at.exception
    g = at.dataframe[0].value
    assert list(g["Centre"]) == ["Alpha"]
    assert g["ARPM vs country median"].isna().all()          # AU has under 5 centres
    assert at.selectbox(key="comp_country").options == ["All countries", "AU"]
    at.radio(key="comp_view").set_value("Centre deep dive").run()
    assert at.selectbox(key="comp_centre").options == ["Alpha"]
    html = " ".join(m.value for m in at.markdown)
    for other in ["Beta", "Kiwi", "Tex"]:
        assert P + other not in html


def test_gpm_with_no_centres_sees_nothing():
    for allowed in ([], None):                               # None: a GPM is never shown all
        at = run("gpm", allowed=allowed)
        assert not at.exception, at.exception
        assert any("None of your centres" in i.value for i in at.info)
        assert not at.dataframe


def test_no_data_yet():
    at = run(comp=pd.DataFrame())
    assert any("No Membership Composition data yet" in i.value for i in at.info)


def test_menu_entry_and_upload_page():
    src = open(mc.__file__.replace("membership_composition.py", "app.py")).read()
    reports = src[src.index("REPORTS=["):src.index("]", src.index("REPORTS=["))]
    assert '"17 · Membership Composition"' in reports
    nav = src[src.index("NAV_GROUPS = ["):src.index("ADMIN_NAV = [")]
    membership = nav[nav.index('("Membership"'):nav.index('("Growth"')]
    assert membership.index('"3 · Membership by Age"') < \
        membership.index('"17 · Membership Composition"')
    lock = src[src.index("ADMIN_ONLY_REPORTS = {"):src.index("if not REPORTS:")]
    assert '"17 · Membership Composition"' not in lock        # GPMs: own centres only
    assert "allowed=None if st.session_state.get(\"access_level\") == \"admin\"" in src
    assert 'selected_report=="17 · Membership Composition"' in src
    assert '"⚙ Composition Upload": admin_pages.page_composition_upload' in src
