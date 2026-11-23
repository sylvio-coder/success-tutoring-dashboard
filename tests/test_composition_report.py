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
    assert list(g.columns) == ["Centre", "Country", "Members", "Total sessions/week",
                               "Sessions per member", "ARPM", "Revenue per session",
                               "Target fill", "Hubs needed a week", "Hubs run a week",
                               "Students per hub", "Extra hubs a week", "Tutor cost %",
                               "Margin per member a week", "% of RRP",
                               "ARPM vs country median", "RRP used", "Check"]
    alpha = g.set_index("Centre").loc["Alpha"]          # 1x1 + 1x2 + 6x2 + 1x4 + 1x4 + 2x1
    assert alpha["Members"] == 12 and alpha["Total sessions/week"] == 25
    assert alpha["Sessions per member"] == 25 / 12
    hubs = 23 / 3.0 + 2                                 # 23 group sessions at 3.0 + 2 x 1:1
    assert alpha["Target fill"] == 3.0 and abs(alpha["Hubs needed a week"] - hubs) < 1e-9
    assert abs(alpha["Margin per member a week"] - (1400 - hubs * 37.90) / 12) < 1e-9
    kiwi = g.set_index("Centre").loc["Kiwi"]            # no NZ tutor rate yet
    assert pd.isna(kiwi["Tutor cost %"]) and pd.isna(kiwi["Margin per member a week"])
    chart = at.get("plotly_chart")[0].proto.spec
    assert "Discount trap" in chart and "Efficient" in chart and "Alpha" in chart
    assert alpha["Check"] == "No sessions-run data" and pd.isna(alpha["Hubs run a week"])
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
    for h in ["No. members", "Total sessions/week", "Membership price", "Total value",
              "Revenue per session", "RRP variance", "Variance from median"]:
        assert f"<th>{h}</th>" in html
    assert "% of members" in html and "Current %" in html and "Ideal %" in html
    heads = ["Plan", "No. members", "Total sessions/week", "Membership price",
             "Revenue per session", "Total value", "RRP variance", "Variance from median"]
    assert "".join(f"<th>{h}</th>" for h in heads) in html
    top = html.split("<th>Plan</th><th>No. members</th>")[1].split("</table>")[0]
    assert ">Group 4×</td>" in top and ">1:1 (1×)</td>" in top
    assert "2.08 sessions" in top and "$116.67 ARPM" in top       # per member: 25 / 12
    assert "<th>Composition</th><th>No. members</th><th>% of members</th><th>Total sessions/week" \
           "</th><th>Average membership price</th><th>Revenue per session</th><th>Total value" \
           "</th>" in html
    ideal_t = html.split("Ideal composition to achieve RRP")[1].split("</table>")[0]
    assert "<th>Plan</th><th>No. members</th><th>Total sessions/week</th><th>Membership price" \
           "</th><th>Revenue per session</th><th>Total value</th>" in ideal_t
    assert ">Group 2×</td>" in ideal_t and "Per member" in ideal_t and " sessions</td>" in ideal_t
    comp = html.split("<th>Composition</th>")[1].split("</table>")[0]
    for h in ["Total sessions/week", "Average membership price", "Total value",
              "Revenue per session"]:
        assert f"<th>{h}</th>" in comp
    g2 = comp.split(">Group 2×</td>")[1].split("</tr>")[0]   # 1 at $40 + 6 at $80
    assert all(v in g2 for v in [">7<", ">14<", "$74.29", "$520.00", "$37.14"])
    total = comp.split(">Total</td>")[1].split("</tr>")[0]
    assert ">12<" in total and ">25<" in total and ">100%<" in total
    ideal = html.split("Ideal composition to achieve RRP")[1]
    assert "<th>Total sessions/week</th>" in ideal
    assert "$1,400.00" in html and "$116.67" in html and "$56.00" in html
    assert "+$6.00 (+8%)" in html and "−$40.00 (−50%)" in html
    assert "−$455" not in html or "vs today" in html
    assert "+167 pts" not in html and "+13 pts" in html              # 1:1: 17% - 4%
    assert "own average price" in html
    assert not at.dataframe                                           # tables, then trends
    keys = [c.key for c in at.get("plotly_chart")]
    assert "trend_c_Members" in keys and len(keys) == 1 + 6 + 3        # chart, trends, pies
    money = at.get("plotly_chart")[0].proto.spec                       # today -> achievable
    assert "achievable" in money and "Achievable" in money
    assert "Today vs achievable: where the money is" in html
    assert "Where the gain comes from" in html and "Price:" in html and "Mix:" in html
    assert "keep it" in html                    # Alpha's mix already beats the ideal
    pies = [c.proto.spec for c in at.get("plotly_chart") if "_mix_" in c.key]
    assert "08 Nov 2026" in pies[0] and "Latest" in pies[1] and "Ideal mix" in pies[2]
    assert all('"type":"pie"' in p.replace(" ", "") for p in pies)
    text = " ".join([m.value for m in at.markdown] + [c.value for c in at.caption])
    assert "Trends · Alpha" in text and "Typical centre" not in str(
        [c.proto.spec for c in at.get("plotly_chart")])                # too few centres


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
    chart = at.get("plotly_chart")[0].proto.spec             # own dot, no medians (< 5)
    assert "Alpha" in chart and "Beta" not in chart and "Discount trap" not in chart
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


def test_group_trends_for_a_gpm_add_up_only_their_centres():
    at = run("gpm", allowed=[P + "Alpha", P + "Beta"])
    assert not at.exception, at.exception
    keys = [c.key for c in at.get("plotly_chart")]
    assert "trend_g_Members" in keys and "trend_g_mix_2" in keys
    members = next(c for c in at.get("plotly_chart") if c.key == "trend_g_Members")
    import base64, json
    import numpy as np
    def values(y):
        return list(np.frombuffer(base64.b64decode(y["bdata"]), dtype=y["dtype"])) \
            if isinstance(y, dict) else list(y)
    ys = [values(t["y"]) for t in json.loads(members.proto.spec)["data"]]
    assert [17.0, 17.0] in ys                                        # Alpha 12 + Beta 5
    text = " ".join([m.value for m in at.markdown] + [c.value for c in at.caption])
    assert "Group trends · AU" in text and "your centres" in text and "All 2 centres" in text


def test_holiday_periods_from_iso_weeks():
    import composition_report as cr
    h = pd.DataFrame({"State": ["NSW", "NSW", "VIC"], "Year": [2026, 2026, 2026],
                      "Start_Week": [27, 51, 27], "End_Week": [28, 5, 28]})
    p = cr.holiday_periods(h, "NSW")
    assert p[0] == (pd.Timestamp("2026-06-29"), pd.Timestamp("2026-07-12"))
    assert p[1] == (pd.Timestamp("2026-12-14"), pd.Timestamp("2027-02-07"))  # into next year
    assert cr.holiday_periods(h, None) == [] and cr.holiday_periods(None, "NSW") == []


def test_group_overview_measured_hubs_from_sessions_run():
    def script():
        import pandas as pd
        import streamlit as st
        import composition_report as cr
        from test_composition_report import COMP, VL, P
        st.session_state["access_level"] = "admin"
        rev = pd.DataFrame({"Success Tutoring - Business name": [P + "Alpha", P + "Beta"],
                            "Date": pd.to_datetime(["2026-11-15", "2026-11-15"]),
                            "Total Sessions": [8, 4], "# Active Students": [12, 9]})
        cr.report_composition(COMP, VL, pd.DataFrame(), revenue=rev)
    at = AppTest.from_function(script, default_timeout=60).run()
    assert not at.exception, at.exception
    g = at.dataframe[0].value.set_index("Centre")
    a = g.loc["Alpha"]                     # 25 sessions, 2 of them 1:1; 8 hubs run
    assert a["Hubs run a week"] == 8 and a["Check"] == ""
    assert abs(a["Students per hub"] - 23 / 6) < 1e-9
    assert abs(a["Extra hubs a week"] - (8 - (23 / 3.0 + 2))) < 1e-9
    b = g.loc["Beta"]                      # 5 members vs 9 CRM students: data check
    assert pd.isna(b["Hubs run a week"]) and b["Check"] == "Members 5 vs CRM students 9"
