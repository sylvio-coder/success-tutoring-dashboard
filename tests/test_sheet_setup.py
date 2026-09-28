import sheet_setup
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
