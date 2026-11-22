"""Franchise Model inputs: reading the uploaded workbook, checks, comparing uploads, storing
versions in the 'Franchise Model' tab, and the admin upload page."""
import io

import openpyxl
import pytest
from streamlit.testing.v1 import AppTest

import franchise_inputs as fin
from fake_sheet import FakeSpreadsheet

TEXT = {
    "Hire a manager at": ("when sessions start", 80, "never"),
    "Owner's wage included": ("yes", "yes", "yes"),
    "Saturday open": ("no", "yes", "yes"),
    "Owner's wage starts": ("first profitable month",) * 3,
    "Sign-up seasonality (1.0 = average month)":
        (None, "0.76, 1.82, 1.24, 0.72, 1.18, 1.13, 1.19, 1.45, 0.72, 0.77, 0.71, 0.30", None),
    "Extra members suspended in holidays": (None, "Jan 17, Apr 9, Jul 5, Oct 2.5, Dec 15", None),
    "School holiday months": (None, "Jan Apr Jul Oct Dec", None),
    "Package mix: Group 1× / 2× / 3×+ / 1:1": (None, "12 / 73 / 8 / 7", None),
    "Foundation member numbers (tiers)": (None, "20 / 20 / 10", None),
    "Foundation prices (tiers)": (None, "50 / 60 / 65", None),
}


def model_rows(overrides=None):
    """Every required input, as (variable, unit, worst, base, best). Numbers are 1, 2, 3."""
    rows = []
    for v in fin.REQUIRED:
        w, b, best = TEXT.get(v, (1, 2, 3))
        rows.append((v, "unit", w, b, best))
    for v, vals in (overrides or {}).items():
        rows = [(n, u, *vals) if n == v else (n, u, w, b, bb) for n, u, w, b, bb in rows]
    return rows


def workbook(models, extra_tab=True, header=("Variable", "Unit", "Worst", "Base", "Best",
                                             "Basis", "Note")):
    wb = openpyxl.Workbook()
    wb.active.title = "How to use"
    wb.active["A1"] = "Notes only: not a model tab"
    for name, rows in models.items():
        ws = wb.create_sheet(name)
        ws["A1"] = f"{name} · assumptions"
        for j, h in enumerate(header, 1):
            ws.cell(4, j, h)
        ws.cell(5, 1, "A. Section one")
        for i, (v, unit, w, b, best) in enumerate(rows):
            for j, val in enumerate([v, unit, w, b, best, "Policy", ""], 1):
                ws.cell(6 + i, j, val)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_read_workbook_finds_model_tabs_sections_and_values():
    rows, problems = fin.read_workbook(workbook({"AU National": model_rows(),
                                                 "NSW": model_rows()}))
    assert not problems
    assert {r["Model"] for r in rows} == {"AU National", "NSW"}     # How to use ignored
    assert len(rows) == 2 * len(fin.REQUIRED)
    r = next(r for r in rows if r["Model"] == "NSW" and r["Variable"] == "Franchise grant")
    assert r["Section"] == "A. Section one" and (r["Worst"], r["Base"], r["Best"]) == (1, 2, 3)
    assert all(not p for p in fin.readiness(rows).values())          # both ready


def test_read_workbook_problems():
    dup = model_rows() + [("Franchise grant", "$", 1, 2, 3)]
    rows, problems = fin.read_workbook(workbook({"AU National": dup}))
    assert ("AU National", "Franchise grant", "appears twice") in problems
    _, problems = fin.read_workbook(workbook({"AU National": model_rows()},
                                             header=("Variable", "Unit", "Base")))
    assert problems and "missing column(s): Worst, Best" in problems[0][2]
    with pytest.raises(fin.InputsError):
        fin.read_workbook(b"not a workbook")
    wb = openpyxl.Workbook()
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(fin.InputsError, match="No model tabs"):
        fin.read_workbook(buf.getvalue())


def test_readiness_flags_missing_not_set_and_bad_values():
    bad = model_rows({"Rent (including outgoings)": (4000, "?", 3000),
                      "GST": (10, "ten", 10),
                      "Hire a manager at": ("soon", 80, "never"),
                      "Saturday open": ("no", "maybe", "yes"),
                      "Foundation prices (tiers)": (None, "50 / 60", None),
                      "School holiday months": (None, "Jan Smarch", None),
                      "Extra members suspended in holidays": (None, "January", None)})
    bad = [r for r in bad if r[0] != "Super"]                       # one input missing
    rows, _ = fin.read_workbook(workbook({"NZ": bad}))
    probs = "\n".join(fin.readiness(rows)["NZ"])
    for text in ["Super: missing", "Rent (including outgoings) (Base): not set yet (?)",
                 "GST (Base): should be a number, not 'ten'",
                 "Hire a manager at (Worst): should be a number of members",
                 "Saturday open (Base): should be yes or no",
                 "Foundation prices (tiers) (Base): should be 3 numbers",
                 "School holiday months (Base): should be month names",
                 "Extra members suspended in holidays (Base): should look like"]:
        assert text in probs, text


def test_parsers():
    assert fin.parse_numbers("50 / 60 / 65") == [50, 60, 65]
    assert fin.parse_numbers("0.76, 1.82") == [0.76, 1.82]
    assert fin.parse_numbers("50 / x") is None
    assert fin.parse_month_values("Jan 17, Apr 9, Oct 2.5") == {"01": 17, "04": 9, "10": 2.5}
    assert fin.parse_month_values("Jan seventeen") is None
    assert fin.parse_months("Jan Apr Jul Oct Dec") == ["01", "04", "07", "10", "12"]
    assert fin.parse_months("–") == [] and fin.parse_months("Jan Foo") is None
    assert fin.to_number("110,000") == 110000 and fin.to_number("$70") == 70
    assert fin.to_number("'4.75") == 4.75 and fin.to_number("yes") is None


def test_changes_between_uploads():
    old, _ = fin.read_workbook(workbook({"NSW": model_rows()}))
    new, _ = fin.read_workbook(workbook({
        "NSW": [r for r in model_rows({"Rent (including outgoings)": (4000, 4500, 3000)})
                if r[0] != "Renewal fee"] + [("Brand new input", "$", 1, 2, 3)]}))
    ch = fin.changes(old, new)
    assert ("NSW", "Rent (including outgoings)", "Worst", 1, 4000) in ch
    assert ("NSW", "Rent (including outgoings)", "Base", 2, 4500) in ch
    assert ("NSW", "Brand new input", "added", "", 2) in ch
    assert ("NSW", "Renewal fee", "removed", 2, "") in ch
    assert fin.changes(old, old) == []


def test_save_keeps_two_versions_and_reads_back_text_and_numbers():
    ss = FakeSpreadsheet({"Weekly Membership": [["x"]]})
    rows, _ = fin.read_workbook(workbook({"NSW": model_rows()}))
    assert fin.save_version(ss, rows, "22/11/2026 10:00 AM") == 1
    assert ss.worksheet(fin.TAB).get_all_values()[0] == fin.HEADERS
    rows2, _ = fin.read_workbook(workbook({"NSW": model_rows({"GST": (10, 15, 10)})}))
    assert fin.save_version(ss, rows2, "23/11/2026") == 2
    assert fin.save_version(ss, rows2, "24/11/2026") == 3
    versions = fin.read_versions(ss)
    assert sorted(versions) == [2, 3]                               # only two kept
    v, cur, when = fin.latest(versions)
    assert v == 3 and when == "24/11/2026"
    inp = fin.model_inputs(cur, "NSW")
    assert inp["GST"]["Base"] == 15 and inp["GST"]["Worst"] == 10
    assert inp["Foundation member numbers (tiers)"]["Base"] == "20 / 20 / 10"   # not a date
    assert inp["Foundation member numbers (tiers)"]["Best"] == "20 / 20 / 10"   # blank = Base
    assert inp["Hire a manager at"]["Worst"] == "when sessions start"
    assert inp["Hire a manager at"]["Base"] == 80
    assert all(not p for p in fin.readiness(cur).values())


def test_admin_menu_and_dispatch():
    src = open(fin.__file__.replace("franchise_inputs.py", "app.py")).read()
    pages = src[src.index("ADMIN_PAGES = ["):src.index("\n", src.index("ADMIN_PAGES = ["))]
    assert '"⚙ Franchise Model Upload"' in pages and 'access_level") == "admin"' in pages
    assert '"⚙ Franchise Model Upload": admin_pages.page_franchise_upload' in src
    nav = src[src.index("ADMIN_NAV = ["):src.index("]", src.index("ADMIN_NAV = ["))]
    assert '"⚙ Franchise Model Upload"' in nav


def run_upload(spreadsheet, data, click=None):
    def script(spreadsheet, data, click):
        import streamlit as st
        import admin_pages
        class Upload:
            name = "model.xlsx"

            def getvalue(self):
                return data
        if data is not None:
            st.file_uploader = lambda *a, **k: Upload()
        admin_pages.page_franchise_upload(spreadsheet, None, None, lambda: None)

    at = AppTest.from_function(script, args=(spreadsheet, data, click), default_timeout=60)
    at.run()
    if click:
        at.button(key=click).click().run()
    return at


def test_upload_page_preview_save_changes_and_restore():
    ss = FakeSpreadsheet({"Weekly Membership": [["x"]]})
    at = run_upload(ss, None)
    assert not at.exception, at.exception
    assert any("No model uploaded yet" in c.value for c in at.caption)
    first = workbook({"AU National": model_rows(), "New Zealand": [("GST", "%", 15, 15, 15)]})
    at = run_upload(ss, first)
    assert not at.exception, at.exception
    status = next(d.value for d in at.dataframe if "Status" in d.value)
    assert dict(zip(status["Model"], status["Status"]))["AU National"] == "Ready"
    assert dict(zip(status["Model"], status["Status"]))["New Zealand"].startswith("Not ready")
    assert any("First upload" in c.value for c in at.caption)
    assert not ss.tabs.get(fin.TAB)                                 # nothing written yet
    at = run_upload(ss, first, click="fm_save")
    assert not at.exception, at.exception
    assert any("Saved as upload 1" in s.value for s in at.success)
    second = workbook({"AU National": model_rows({"GST": (10, 12, 10)})})
    at = run_upload(ss, second)
    ch = next(d.value for d in at.dataframe if "Changed" in d.value)
    assert ("AU National", "GST", "Base", "2", "12") in [tuple(r) for r in ch.values]
    at = run_upload(ss, second, click="fm_save")
    assert any("Saved as upload 2" in s.value for s in at.success)
    at = run_upload(ss, None, click="fm_restore")
    assert any("Restored as upload 3" in s.value for s in at.success)
    _, cur, _ = fin.latest(fin.read_versions(ss))
    assert fin.model_inputs(cur, "AU National")["GST"]["Base"] == 2  # back to upload 1


def test_upload_page_blocks_save_when_the_file_has_problems():
    ss = FakeSpreadsheet({"Weekly Membership": [["x"]]})
    dup = model_rows() + [("GST", "%", 1, 2, 3)]
    at = run_upload(ss, workbook({"AU National": dup}))
    assert any("appears twice" in e.value for e in at.error)
    assert at.button(key="fm_save").disabled
