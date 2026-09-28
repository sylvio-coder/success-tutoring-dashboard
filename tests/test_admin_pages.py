"""Runs the admin pages in Streamlit's AppTest against an in-memory sheet."""
import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from fake_sheet import FakeSpreadsheet
from test_ingest import HAPANA_MEMBERS, HAPANA_REV_AU, HAPANA_REV_NZ, INHOUSE_MEMBERS, \
    INHOUSE_REVENUE

VLOOKUP = [
    ["Location", "Stage", "Status", "GPM", "Country", "Region", "Location Start",
     "Months old", "Onboarding Week", "Onboarding Members"],
] + [[f"Success Tutoring - {n}", "Growth", "Trading", "Sara Walliar", c, r, "1-Jan-2024",
      f"=DATEDIF(G{row},TODAY(),\"M\")", "", ""]
     for row, (n, c, r) in enumerate([
         ("Auburn", "Australia", "New South Wales"),
         ("Epping", "Australia", "New South Wales"),
         ("Epping VIC", "Australia", "Victoria"),
         ("Belmont", "Australia", "Victoria"),
         ("Belmont WA", "Australia", "Western Australia"),
         ("Howick", "New Zealand", "Auckland"),
         ("Burwood", "Australia", "New South Wales"),
         ("Landsdale", "Australia", "Western Australia"),
     ], start=2)]

WM = [["Success Tutoring - Business name", "Date - Week/Year", "Date", "# Active members",
       "# Suspended members", "# Cancelled members", "# New members", "Stage", "GPM",
       "Country"],
      ["Success Tutoring - Auburn", "37/2026", "13/9/2026", 80, 9, 1, 1,
       "=VLOOKUP(A2,Vlookup!A:J,2,0)", "=VLOOKUP(A2,Vlookup!A:J,4,0)",
       "=VLOOKUP(A2,Vlookup!A:J,5,0)"]]

REV = [["Location", "Date - Week/Year", "Date", "# Active Students", "Total Sessions",
        "Gross Revenue", "Net Revenue", "Student Visits", "Revenue per Session",
        "Revenue per Student", "Country"],
       ["Success Tutoring - Howick", "37/2026", "13/9/2026", 19, 17, "$1,320", "$1,200",
        37.2, "$70.59", "$63.16", "New Zealand"],
       ["Success Tutoring - Auburn", "37/2026", "13/9/2026", 80, 30, "$5,500", "$5,000",
        156.8, "=G3/E3", "$62.50", "Australia"]]


class Upload:
    def __init__(self, name, data):
        self.name, self._data = name, data

    def getvalue(self):
        return self._data


FILES = [Upload("Weekly Membership.xlsx", HAPANA_MEMBERS),
         Upload("SA - Revenue Location Summary (16).xlsx", HAPANA_REV_AU),
         Upload("SA - Revenue Location Summary (17).xlsx", HAPANA_REV_NZ),
         Upload("members-report-2026-09-14-to-2026-09-20.xlsx", INHOUSE_MEMBERS),
         Upload("revenue-report-2026-09-14-to-2026-09-20.xlsx", INHOUSE_REVENUE)]


def frames(extra_wm=()):
    df_wm = pd.DataFrame([{"Success Tutoring - Business name": "Success Tutoring - Auburn",
                           "Date - Week/Year": "37/2026", "Date": pd.Timestamp(2026, 9, 13),
                           "# Active members": 80, "Country": "Australia"}, *extra_wm])
    df_rv = pd.DataFrame([{"Success Tutoring - Business name": "Success Tutoring - Auburn",
                           "Date - Week/Year": "37/2026", "Date": pd.Timestamp(2026, 9, 13),
                           "Net Revenue": 5000.0, "Country": "Australia"}])
    return df_wm, df_rv


def run_page(page, spreadsheet, files=None, frames=frames):
    def script(page, spreadsheet, files, frames):
        import streamlit as st
        import admin_pages
        if files is not None:
            st.file_uploader = lambda *a, **k: files
        df_wm, df_rv = frames()
        getattr(admin_pages, page)(spreadsheet, df_wm, df_rv, lambda: None)

    at = AppTest.from_function(script, args=(page, spreadsheet, files, frames),
                               default_timeout=30)
    return at.run()


@pytest.fixture
def ss():
    return FakeSpreadsheet({"Vlookup": VLOOKUP, "Weekly Membership": WM, "Revenue": REV})


def test_upload_flags_new_location_and_blocks_write(ss):
    at = run_page("page_weekly_upload", ss, FILES)
    assert not at.exception
    labels = [b.label for b in at.button]
    confirm = next(b for b in at.button if b.label.startswith("✅ Confirm"))
    assert "38/2026" in confirm.label
    assert confirm.disabled   # Torquay is unknown
    assert any("Torquay" in w.value for w in at.warning)


def test_upload_writes_week_after_leaving_out_unknown(ss):
    at = run_page("page_weekly_upload", ss, FILES)
    at.selectbox(key="map_Success Tutoring - Torquay").set_value("Leave out this week").run()
    assert not at.exception
    confirm = next(b for b in at.button if b.label.startswith("✅ Confirm"))
    assert not confirm.disabled
    confirm.click().run()
    assert not at.exception, at.exception
    assert any("Week 38/2026 written" in s.value for s in at.success)

    wm = ss.worksheet("Weekly Membership").cells
    new = [r for r in wm[1:] if r[1] == "38/2026"]
    names = sorted(r[0] for r in new)
    assert "Success Tutoring - Belmont WA" in names and "Success Tutoring - Belmont" in names
    assert not any("Sandbox" in n for n in names)
    auburn = next(r for r in wm if r[0] == "Success Tutoring - Auburn" and r[1] == "38/2026")
    row = wm.index(auburn) + 1
    assert auburn[7] == f"=VLOOKUP(A{row},Vlookup!A:J,2,0)"

    rev = ss.worksheet("Revenue").cells
    howick = next(r for r in rev if r[0] == "Success Tutoring - Howick" and r[1] == "38/2026")
    assert howick[6] == round(1346 / 1.15, 2)
    auburn = next(r for r in rev if r[0] == "Success Tutoring - Auburn" and r[1] == "38/2026")
    row = rev.index(auburn) + 1
    assert auburn[8] == f"=G{row}/E{row}"     # calculated column keeps the sheet formula

    # Uploading the same week again replaces rather than duplicates.
    at.run()
    next(b for b in at.button if b.label.startswith("✅ Confirm")).click().run()
    wm = ss.worksheet("Weekly Membership").cells
    assert sum(1 for r in wm if r[1] == "38/2026") == len(new)


def torquay_frames():
    from test_admin_pages import frames
    return frames([{"Success Tutoring - Business name": "Success Tutoring - Torquay",
                    "Date - Week/Year": "37/2026", "Date": pd.Timestamp(2026, 9, 13),
                    "# Active members": 10, "Country": ""}])


def test_locations_page_adds_new_location(ss):
    at = run_page("page_locations", ss, frames=torquay_frames)
    assert not at.exception, at.exception
    for col, v in [("Stage", "Growth"), ("Status", "Trading"), ("GPM", "Sara Walliar"),
                   ("Country", "Australia"), ("Region", "Victoria")]:
        at.selectbox(key=f"{col}_Success Tutoring - Torquay").set_value(v)
    next(b for b in at.button if b.label == "Add to master list").click().run()
    assert not at.exception, at.exception
    vl = ss.worksheet("Vlookup").cells
    # Start date defaults to the first week it had members.
    assert vl[-1][:7] == ["Success Tutoring - Torquay", "Growth", "Trading", "Sara Walliar",
                          "Australia", "Victoria", "13-Sep-2026"]
    assert vl[-1][7] == f"=DATEDIF(G{len(vl)},TODAY(),\"M\")"


def test_locations_page_nz_gst_fix(ss):
    at = run_page("page_locations", ss)
    assert not at.exception, at.exception
    assert any("NZ revenue rows" in w.value for w in at.warning)
    next(b for b in at.button if b.label.startswith("Correct NZ net revenue")).click().run()
    assert not at.exception, at.exception
    howick = ss.worksheet("Revenue").cells[1]
    assert howick[6] == round(1320 / 1.15, 2)
    assert howick[8] == round(round(1320 / 1.15, 2) / 17, 2)
    assert ss.worksheet("Revenue").cells[2][6] == "$5,000"   # AU untouched


def test_upload_resolves_same_name_clash(ss):
    from test_ingest import RAW_INHOUSE_MEMBERS
    files = [Upload("Weekly Membership.xlsx", HAPANA_MEMBERS),
             Upload("members-report-2026-09-14-to-2026-09-20.xlsx", RAW_INHOUSE_MEMBERS)]
    at = run_page("page_weekly_upload", ss, files)
    assert not at.exception, at.exception
    confirm = next(b for b in at.button if b.label.startswith("✅ Confirm"))
    assert confirm.disabled
    at.selectbox(key="clash_In-house: Success Tutoring - Belmont").set_value(
        "Success Tutoring - Belmont WA").run()
    assert not at.exception, at.exception
    assert not any("more than once" in e.value for e in at.error)
    next(b for b in at.button if b.label.startswith("✅ Confirm")).click().run()
    assert not at.exception, at.exception
    aliases = ss.worksheet("Location Aliases").get_all_records()
    assert aliases == [{"Alias": "In-house: Success Tutoring - Belmont",
                        "Location": "Success Tutoring - Belmont WA"}]
    wm = ss.worksheet("Weekly Membership").cells
    wa = next(r for r in wm if r[0] == "Success Tutoring - Belmont WA" and r[1] == "38/2026")
    vic = next(r for r in wm if r[0] == "Success Tutoring - Belmont" and r[1] == "38/2026")
    assert (wa[3], vic[3]) == (28, 17)


def test_upload_lists_old_rows_it_will_not_replace(ss):
    def frames_with_old_row():
        df_wm, df_rv = frames()
        old = pd.DataFrame([{"Success Tutoring - Business name": "Success Tutoring - Belmont (WA)",
                             "Date - Week/Year": "38/2026", "Date": pd.Timestamp(2026, 9, 20),
                             "Net Revenue": 90.0, "Country": "Australia"}])
        return df_wm, pd.concat([df_rv, old], ignore_index=True)

    at = run_page("page_weekly_upload", ss, FILES, frames=frames_with_old_row)
    assert not at.exception, at.exception
    w = next(w.value for w in at.warning if "not in this upload" in w.value)
    assert w.startswith("Revenue already has week 38/2026") and "Belmont (WA)" in w
