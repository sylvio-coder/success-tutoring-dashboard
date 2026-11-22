"""Franchise Model report page (franchise_report.py): sections, cases, models, admin only."""
from streamlit.testing.v1 import AppTest

import franchise_inputs as fin
from test_franchise_plan import APPROVED


def rows_for(*models, broken=None):
    rows = []
    for m in models:
        for v, vals in APPROVED.items():
            rec = {"Model": m, "Section": "S", "Variable": v, "Unit": "", "Basis": "Policy",
                   "Note": ""}
            rec.update({c: str(x) for c, x in zip(fin.CASES, vals)})
            if m == broken and v == "GST":
                rec["Base"] = "?"
            rows.append(rec)
    return rows


def run(versions, level="admin"):
    def script(versions, level):
        import streamlit as st
        import franchise_report
        st.session_state["access_level"] = level
        franchise_report.report_franchise_model(versions, dict(
            show_chart=st.plotly_chart,
            std_layout=lambda t, y="", height=400: dict(title=t, height=height)))
    return AppTest.from_function(script, args=(versions, level), default_timeout=120).run()


def text(at):
    return "".join(m.value for m in at.markdown)


def test_page_shows_every_section_for_an_admin():
    at = run({1: rows_for("AU National", "NSW", broken="NSW")})
    assert not at.exception, at.exception
    t = text(at)
    for h in ["1. Summary", "2. Investor Metrics · Base", "3. Your operating plan · Base",
              "4. Road to profitability · Base", "5. Profit and cash by year · Base",
              "6. Risks · Base", "7. Assumptions used · Base"]:
        assert h in t, h
    for h in ["Worst", "Base", "Best", "Total cash needed", "Checkpoints", "you need any ONE of",
              "Standard price", "Enquiries needed a month", "Study hubs a week",
              "Centre opens", "Investment paid back", "What moves the result most",
              "Opening month", "Profit after salary"]:
        assert h in t, h
    assert at.selectbox(key="fr_model").options == ["AU National"]   # NSW not ready
    assert len(at.get("plotly_chart")) == 1
    assert any("upload 1" in c.value for c in at.caption)
    at.radio(key="fr_case").set_value("Best").run()
    assert not at.exception, at.exception
    assert "2. Investor Metrics · Best" in text(at)


def test_no_upload_and_nothing_ready_messages():
    at = run({})
    assert any("No assumptions uploaded yet" in i.value for i in at.info)
    at = run({1: rows_for("NSW", broken="NSW")})
    assert any("None of the uploaded models is ready" in w.value for w in at.warning)


def test_gpm_cannot_see_the_franchise_model():
    at = run({1: rows_for("AU National")}, level="gpm")
    assert any("admins only" in e.value for e in at.error)
    assert "Summary" not in text(at) and not at.get("plotly_chart")


def test_app_wiring_admin_only():
    src = open(fin.__file__.replace("franchise_inputs.py", "app.py")).read()
    assert 'selected_report=="16 · Franchise Model":        franchise_report.report_franchise_model(' in src
    lock = src[src.index("ADMIN_ONLY_REPORTS = {"):src.index("}", src.index("ADMIN_ONLY_REPORTS = {"))]
    assert '"16 · Franchise Model"' in lock
