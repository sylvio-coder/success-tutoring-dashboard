import sheet_writer as sw
from fake_sheet import FakeSpreadsheet

WM_HEADERS = ["Success Tutoring - Business name", "Date - Week/Year", "Date",
              "# Active members", "Stage", "GPM"]


def make():
    return FakeSpreadsheet({"Weekly Membership": [
        WM_HEADERS,
        ["Success Tutoring - Auburn", "36/2026", "6/9/2026", 80,
         "=VLOOKUP(A2,Vlookup!A:J,2,0)", "=VLOOKUP(A2,Vlookup!A:J,4,0)"],
        ["Success Tutoring - Auburn", "37/2026", "13/9/2026", 82,
         "=VLOOKUP(A3,Vlookup!A:J,2,0)", "=VLOOKUP(A3,Vlookup!A:J,4,0)"],
    ]})


def rec(name, week, active):
    return {"Success Tutoring - Business name": name, "Date - Week/Year": week,
            "Date": "20/9/2026", "# Active members": active}


def test_append_copies_formulas_down():
    ss = make()
    ws = ss.worksheet("Weekly Membership")
    n = sw.append_records(ws, [rec("Success Tutoring - Auburn", "38/2026", 83),
                               rec("Success Tutoring - Epping", "38/2026", 100)])
    assert n == 2
    assert ws.cells[3][:4] == ["Success Tutoring - Auburn", "38/2026", "20/9/2026", 83]
    assert ws.cells[3][4] == "=VLOOKUP(A4,Vlookup!A:J,2,0)"
    assert ws.cells[4][5] == "=VLOOKUP(A5,Vlookup!A:J,4,0)"
    assert ws.formats_copied == [(3, 5)]


def test_replace_week_removes_only_that_week_and_those_locations():
    ss = make()
    ws = ss.worksheet("Weekly Membership")
    sw.append_records(ws, [rec("Success Tutoring - Auburn", "38/2026", 1),
                           rec("Success Tutoring - Epping", "38/2026", 2)])
    deleted, added = sw.replace_week(ws, [rec("Success Tutoring - Auburn", "38/2026", 83)],
                                     WM_HEADERS[0], "38/2026")
    assert (deleted, added) == (1, 1)
    rows = ws.get_all_values()[1:]
    assert [(r[0], r[1], r[3]) for r in rows] == [
        ("Success Tutoring - Auburn", "36/2026", "80"),
        ("Success Tutoring - Auburn", "37/2026", "82"),
        ("Success Tutoring - Epping", "38/2026", "2"),
        ("Success Tutoring - Auburn", "38/2026", "83"),
    ]
    assert ws.cells[4][4] == "=VLOOKUP(A5,Vlookup!A:J,2,0)"


def test_facts_never_replaced_by_formula():
    ss = FakeSpreadsheet({"Revenue": [
        ["Location", "Gross Revenue", "Net Revenue", "Revenue per Session", "Total Sessions"],
        ["A", 110, "=B2/1.1", "=C2/E2", 10],
    ]})
    ws = ss.worksheet("Revenue")
    sw.append_records(ws, [{"Location": "B", "Gross Revenue": 115, "Net Revenue": 100,
                            "Revenue per Session": 9.99, "Total Sessions": 10}],
                      prefer_formula=["Revenue per Session"])
    assert ws.cells[2] == ["B", 115, 100, "=C3/E3", 10]


def test_week_label_kept_as_text():
    assert sw._text_safe("1/2027") == "'1/2027"
    assert sw._text_safe("20/9/2026") == "20/9/2026"


def test_delete_rows_merges_runs():
    ss = FakeSpreadsheet({"T": [["h"]] + [[str(i)] for i in range(2, 10)]})
    ws = ss.worksheet("T")
    sw.delete_rows(ws, [3, 4, 5, 8])
    assert [r[0] for r in ws.get_all_values()] == ["h", "2", "6", "7", "9"]


def test_aliases_tab_created_on_first_write():
    ss = FakeSpreadsheet({})
    assert sw.read_aliases(ss) == {}
    sw.add_aliases(ss, [("ST Burwood", "Success Tutoring - Burwood")])
    assert sw.read_aliases(ss) == {"ST Burwood": "Success Tutoring - Burwood"}
