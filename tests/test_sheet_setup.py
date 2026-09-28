import sheet_setup
import pytest
from fake_sheet import FakeSpreadsheet


def make():
    return FakeSpreadsheet({
        "Weekly Membership": [
            ["Success Tutoring - Business name", "Date - Week/Year", "Date",
             "# Active members", "Stage", "Age (Months)", "Onboarding week",
             "Onboarding Members"],
            ["Success Tutoring - Auburn", "38/2026", "20/9/2026", 83,
             "=VLOOKUP(A2,Vlookup!A:J,2,0)", "=VLOOKUP(A2,Vlookup!A:J,8,0)",
             "=VLOOKUP(A2,Vlookup!A:J,9,0)", "=VLOOKUP(A2,Vlookup!A:J,10,0)"],
        ],
        "Revenue": [
            ["Location", "Date - Week/Year", "Gross Revenue", "Net Revenue", "Region",
             "Age (Months)"],
            ["Success Tutoring - Auburn", "38/2026", 5970, 5427.27,
             "=VLOOKUP(A2,Vlookup!A:J,6,0)", "=VLOOKUP(A2,Vlookup!A:J,8,0)"],
        ],
        "Vlookup": [
            ["Location", "Stage", "Status", "GPM", "Country", "Region", "Location Start",
             "Months old", "Onboarding Week", "Onboarding Members"],
            ["Success Tutoring - Auburn", "Growth", "Trading", "Sara Walliar", "Australia",
             "New South Wales", "1-Jan-2024", '=DATEDIF(G2,TODAY(),"M")', 12, 60],
            ["Success Tutoring - Epping", "Growth", "Trading", "Selina Tan", "Australia",
             "New South Wales", "1-Jan-2024", '=DATEDIF(G3,TODAY(),"M")', "", ""],
        ],
        "Membership Churn": [["Date", "Month/Year"], ["x", "y"]],
    })


def test_switch_to_weeks_old():
    ss = make()
    done = sheet_setup.switch_to_weeks_old(ss)
    assert ss.worksheet("Weekly Membership").row_values(1) == [
        "Success Tutoring - Business name", "Date - Week/Year", "Date", "# Active members",
        "Stage"]
    assert ss.worksheet("Revenue").row_values(1) == [
        "Location", "Date - Week/Year", "Gross Revenue", "Net Revenue", "Region"]
    vl = ss.worksheet("Vlookup").cells
    assert vl[0] == ["Location", "Stage", "Status", "GPM", "Country", "Region",
                     "Location Start", "Weeks old"]
    assert vl[1][7] == sheet_setup.weeks_old_formula(2, "A", "C", "D")
    assert vl[2][7].startswith('=IF(COUNTIFS(') and "$A3" in vl[2][7]
    assert any(d.startswith("Weekly Membership: removed Age (Months)") for d in done)
    # Running it again is harmless.
    sheet_setup.switch_to_weeks_old(ss)
    assert ss.worksheet("Vlookup").cells[0][-1] == "Weeks old"


def test_weeks_old_formula_text():
    f = sheet_setup.weeks_old_formula(5, "A", "C", "D")
    assert f == ("=IF(COUNTIFS('Weekly Membership'!$A:$A,$A5,'Weekly Membership'!$D:$D,\">0\")"
                 "=0,\"\",INT((MAX('Weekly Membership'!$C:$C)-MINIFS('Weekly Membership'!$C:$C,"
                 "'Weekly Membership'!$A:$A,$A5,'Weekly Membership'!$D:$D,\">0\"))/7)+1)")


def test_apply_formatting():
    ss = make()
    ss.worksheet("Vlookup").tables = [{"tableId": "t1", "rowsProperties": {
        "headerColorStyle": {"rgbColor": {"red": 0.2, "green": 0.4, "blue": 0.3}}}}]
    ss.worksheet("Revenue").banded_ranges = [{"bandedRangeId": 7}]
    report = dict(sheet_setup.apply_formatting(ss))
    assert report["Vlookup"].startswith("font (table")
    assert report["Membership Churn"].startswith("font only")
    rev = ss.worksheet("Revenue")
    assert [b["bandedRangeId"] for b in rev.banded_ranges] != [7]          # replaced
    assert rev.banded_ranges[0]["rowProperties"]["headerColor"] == \
        {"red": 0.2, "green": 0.4, "blue": 0.3}                             # table's colour
    assert rev.frozen_rows == 1
    assert ss.worksheet("Vlookup").banded_ranges == []                    # table untouched
    fmts = [r["repeatCell"] for r in ss.format_requests if "repeatCell" in r
            and "numberFormat" in r["repeatCell"]["cell"]["userEnteredFormat"]]
    money = [f for f in fmts if f["range"]["sheetId"] == rev.id and
             f["range"]["startColumnIndex"] == 2]
    assert money[0]["cell"]["userEnteredFormat"]["numberFormat"]["pattern"] == "$#,##0"


def test_broken_formulas():
    ss = make()
    ss.worksheet("Revenue").cells[1][4] = "#REF!"
    assert sheet_setup.broken_formulas(ss) == {"Revenue": 1}


def test_formatting_keeps_existing_unlisted_bands():
    ss = make()
    ss.worksheet("Weekly Membership").hidden_banding = True
    report = dict(sheet_setup.apply_formatting(ss))
    assert "kept its existing alternating rows" in report["Weekly Membership"]
    assert "alternating rows" in report["Revenue"]
    assert ss.worksheet("Weekly Membership").frozen_rows == 1     # rest still applied
    assert ss.worksheet("Revenue").banded_ranges                  # other tabs unaffected


def test_location_start_formula_text():
    f = sheet_setup.location_start_formula(5, "A", "C", "D")
    assert f == ("=IF(COUNTIFS('Weekly Membership'!$A:$A,$A5,'Weekly Membership'!$D:$D,\">0\")"
                 "=0,\"\",MINIFS('Weekly Membership'!$C:$C,'Weekly Membership'!$A:$A,$A5,"
                 "'Weekly Membership'!$D:$D,\">0\"))")
    assert sheet_setup.weeks_from_start_formula(5, "G", "C") == \
        "=IF($G5=\"\",\"\",INT((MAX('Weekly Membership'!$C:$C)-$G5)/7)+1)"
    closed = sheet_setup.weeks_from_start_formula(5, "G", "C", "A", "D", "C")
    assert closed == ("=IF($G5=\"\",\"\",INT((IF($C5=\"Closed\",MAXIFS('Weekly Membership'!$C:$C,"
                      "'Weekly Membership'!$A:$A,$A5,'Weekly Membership'!$D:$D,\">0\"),"
                      "MAX('Weekly Membership'!$C:$C))-$G5)/7)+1)")


def test_link_location_start():
    ss = make()
    sheet_setup.switch_to_weeks_old(ss)
    assert not sheet_setup.start_is_formula(ss.worksheet("Vlookup"),
                                            ss.worksheet("Vlookup").row_values(1))
    done = sheet_setup.link_location_start(ss)
    vl = ss.worksheet("Vlookup").cells
    assert vl[0] == ["Location", "Stage", "Status", "GPM", "Country", "Region",
                     "Location Start", "Weeks old"]           # headers unchanged
    for row in (2, 3):
        assert vl[row - 1][6] == sheet_setup.location_start_formula(row, "A", "C", "D")
        assert vl[row - 1][7] == sheet_setup.weeks_from_start_formula(row, "G", "C", "A", "D", "C")
    fmt = [r["repeatCell"] for r in ss.format_requests if "repeatCell" in r][-1]
    assert fmt["range"]["startColumnIndex"] == 6
    assert fmt["cell"]["userEnteredFormat"]["numberFormat"]["type"] == "DATE"
    assert any("Location Start is now the first week" in d for d in done)
    assert sheet_setup.start_is_formula(ss.worksheet("Vlookup"), ss.worksheet("Vlookup").row_values(1))
    # Running it again is harmless.
    sheet_setup.link_location_start(ss)
    assert ss.worksheet("Vlookup").cells[1][6] == sheet_setup.location_start_formula(2, "A", "C", "D")


def test_link_location_start_needs_weeks_old():
    ss = make()        # still has 'Months old'
    with pytest.raises(ValueError, match="Weeks old"):
        sheet_setup.link_location_start(ss)


def test_repair_text_dates():
    ss = make()
    wm = ss.worksheet("Weekly Membership")
    date_i = wm.row_values(1).index("Date")
    wm.cells[1][date_i] = "20/9/2026"          # stored as text by an earlier upload
    done = sheet_setup.repair_text_dates(ss)
    assert wm.cells[1][date_i] == "2026-09-20"
    assert any(d.startswith("Weekly Membership: 1 date") for d in done)
    assert sheet_setup.repair_text_dates(ss) == [] or all("0 date" not in d for d in done)


def test_sheet_date_is_iso():
    import datetime
    import ingest
    assert ingest.sheet_date(datetime.date(2026, 10, 4)) == "2026-10-04"
