"""One-off Sheet maintenance run from the admin Sheet Setup page.

- apply_formatting: one font, a consistent header row, alternating row
  colours and number formats on every tab.
- switch_to_weeks_old: replace the Vlookup "Months old" column with a
  "Weeks old" formula (week 1 = first week with members) and remove the
  onboarding columns from Vlookup, Weekly Membership and Revenue.
"""
from gspread.utils import rowcol_to_a1

FONT = {"fontFamily": "Arial", "fontSize": 10}
HEADER_BG = {"red": 0.169, "green": 0.369, "blue": 0.310}      # dark green
BAND_BG = {"red": 0.945, "green": 0.961, "blue": 0.953}        # pale grey-green
WHITE = {"red": 1, "green": 1, "blue": 1}

# Tabs whose layout isn't a single header row + data (merged headers).
FONT_ONLY_TABS = {"Membership Churn"}

REMOVED_COLUMNS = {"Onboarding Week", "Onboarding week", "Onboarding Members",
                   "Age (Months)"}

NUMBER_FORMATS = [
    (("Gross Revenue", "Net Revenue"), "CURRENCY", "$#,##0"),
    (("Revenue per Session", "Revenue per Student"), "CURRENCY", "$#,##0.00"),
    (("Student Visits",), "NUMBER", "0.0"),
    (("Sessions per Student", "Student per Session", "Sessions per Student Visit",
      "Student Visits per Session"), "NUMBER", "0.00"),
    (("Date",), "DATE", "d/m/yyyy"),
    (("Location Start",), "DATE", "d-mmm-yyyy"),
    (("Date - Week/Year",), "TEXT", "@"),
]
INTEGER_HEADERS = ("Total Sessions", "Weeks old", "Months old", "Active Members",
                   "Suspended Members", "Cancelled Members", "New Members", "Year",
                   "Start_Week", "End_Week")


def _number_format(header):
    for names, kind, pattern in NUMBER_FORMATS:
        if header in names:
            return {"type": kind, "pattern": pattern}
    if header.startswith("#") or header in INTEGER_HEADERS:
        return {"type": "NUMBER", "pattern": "0"}
    return None


def _rng(sheet_id, r0, r1, c0, c1):
    return {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
            "startColumnIndex": c0, "endColumnIndex": c1}


def _table_colours(meta):
    """Reuse the colours of an existing Sheets table, so every tab matches it."""
    for sheet in meta["sheets"]:
        for t in sheet.get("tables", []):
            rp = t.get("rowsProperties", {})
            h = rp.get("headerColorStyle", {}).get("rgbColor")
            b = rp.get("secondBandColorStyle", {}).get("rgbColor")
            if h:
                return h, b or BAND_BG
    return HEADER_BG, BAND_BG


def _metadata(spreadsheet):
    try:
        return spreadsheet.fetch_sheet_metadata(
            params={"fields": "sheets(properties,bandedRanges,tables)"})
    except Exception:
        return spreadsheet.fetch_sheet_metadata()


def _already_banded(err):
    return "alternating background colors" in str(err)


def apply_formatting(spreadsheet):
    """Returns a list of (tab, what was done).

    Each tab is formatted in its own request, so one tab can't block the
    rest. A tab that already has alternating colours the API doesn't list
    (e.g. a Sheets table) keeps them and gets everything else.
    """
    meta = _metadata(spreadsheet)
    header_bg, band_bg = _table_colours(meta)
    report = []
    for sheet in meta["sheets"]:
        requests = []
        p = sheet["properties"]
        sid, title = p["sheetId"], p["title"]
        grid = p.get("gridProperties", {})
        nrows, ncols = grid.get("rowCount", 1000), grid.get("columnCount", 26)
        is_table = bool(sheet.get("tables"))

        requests.append({"repeatCell": {
            "range": _rng(sid, 0, nrows, 0, ncols),
            "cell": {"userEnteredFormat": {"textFormat": FONT,
                                           "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat.textFormat.fontFamily,"
                      "userEnteredFormat.textFormat.fontSize,"
                      "userEnteredFormat.verticalAlignment"}})
        if title in FONT_ONLY_TABS:
            spreadsheet.batch_update({"requests": requests})
            report.append((title, "font only (merged header layout)"))
            continue
        if is_table:
            # A Sheets table styles its own header, bands and column types.
            requests.append({"autoResizeDimensions": {"dimensions": {
                "sheetId": sid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": ncols}}})
            spreadsheet.batch_update({"requests": requests})
            report.append((title, "font (table keeps its own styling)"))
            continue

        headers = spreadsheet.worksheet(title).row_values(1)
        width = max(len(headers), 1)
        requests += [
            {"repeatCell": {
                "range": _rng(sid, 0, 1, 0, width),
                "cell": {"userEnteredFormat": {
                    "backgroundColor": header_bg, "wrapStrategy": "WRAP",
                    "textFormat": {**FONT, "bold": True, "foregroundColor": WHITE}}},
                "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.wrapStrategy,"
                          "userEnteredFormat.textFormat"}},
            {"updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 32}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 1,
                          "endIndex": nrows},
                "properties": {"pixelSize": 21}, "fields": "pixelSize"}},
        ]
        banding = [{"deleteBanding": {"bandedRangeId": b["bandedRangeId"]}}
                   for b in sheet.get("bandedRanges", [])]
        banding.append({"addBanding": {"bandedRange": {
            "range": _rng(sid, 0, nrows, 0, width),
            "rowProperties": {"headerColor": header_bg, "firstBandColor": WHITE,
                              "secondBandColor": band_bg}}}})
        formatted = 0
        for i, h in enumerate(headers):
            fmt = _number_format(h.strip())
            if fmt:
                formatted += 1
                requests.append({"repeatCell": {
                    "range": _rng(sid, 1, nrows, i, i + 1),
                    "cell": {"userEnteredFormat": {"numberFormat": fmt}},
                    "fields": "userEnteredFormat.numberFormat"}})
        requests.append({"autoResizeDimensions": {"dimensions": {
            "sheetId": sid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": width}}})
        try:
            spreadsheet.batch_update({"requests": requests + banding})
            bands = "alternating rows"
        except Exception as e:
            if not _already_banded(e):
                raise
            spreadsheet.batch_update({"requests": requests})
            bands = "kept its existing alternating rows"
        report.append((title, f"font, header row, {bands}, {formatted} number formats"))
    return report


def weeks_old_formula(row, name_col, date_col, active_col, tab="Weekly Membership"):
    """Week 1 = the first week the location had active members; blank if none yet."""
    t = f"'{tab}'"
    names, dates, active = (f"{t}!${c}:${c}" for c in (name_col, date_col, active_col))
    return (f'=IF(COUNTIFS({names},$A{row},{active},">0")=0,"",'
            f'INT((MAX({dates})-MINIFS({dates},{names},$A{row},{active},">0"))/7)+1)')


def _col_letter(i):
    return rowcol_to_a1(1, i + 1)[:-1]


def switch_to_weeks_old(spreadsheet):
    """Returns a list of change descriptions."""
    done = []
    wm = spreadsheet.worksheet("Weekly Membership")
    wm_headers = wm.row_values(1)
    name_col = _col_letter(wm_headers.index("Success Tutoring - Business name"))
    date_col = _col_letter(wm_headers.index("Date"))
    active_col = _col_letter(wm_headers.index("# Active members"))

    # Other tabs first: their columns look up these Vlookup columns by position.
    for title in ("Weekly Membership", "Revenue", "Vlookup"):
        ws = spreadsheet.worksheet(title)
        headers = ws.row_values(1)
        drop = sorted((i for i, h in enumerate(headers) if h.strip() in REMOVED_COLUMNS),
                      reverse=True)
        if drop:
            spreadsheet.batch_update({"requests": [
                {"deleteDimension": {"range": {"sheetId": ws.id, "dimension": "COLUMNS",
                                               "startIndex": i, "endIndex": i + 1}}}
                for i in drop]})
            done.append(f"{title}: removed " + ", ".join(headers[i] for i in sorted(drop)))

    vl = spreadsheet.worksheet("Vlookup")
    values = vl.get_all_values()
    headers = values[0]
    col = next((i for i, h in enumerate(headers) if h.strip() in ("Months old", "Weeks old")),
               None)
    if col is None:
        raise ValueError("Vlookup has no 'Months old' or 'Weeks old' column.")
    last = max((i + 1 for i, r in enumerate(values) if r and str(r[0]).strip()), default=1)
    letter = _col_letter(col)
    formulas = [[weeks_old_formula(r, name_col, date_col, active_col)] for r in range(2, last + 1)]
    data = [{"range": f"{letter}1", "values": [["Weeks old"]]}]
    if formulas:
        data.append({"range": f"{letter}2:{letter}{last}", "values": formulas})
    vl.batch_update(data, value_input_option="USER_ENTERED")
    done.append(f"Vlookup: '{headers[col]}' is now 'Weeks old' (formula on {len(formulas)} "
                "locations; week 1 = first week with members)")
    return done


def broken_formulas(spreadsheet):
    """{tab: number of #REF! cells} for tabs that have any."""
    out = {}
    for sheet in spreadsheet.fetch_sheet_metadata()["sheets"]:
        title = sheet["properties"]["title"]
        n = sum(str(v).startswith("#REF") for row in spreadsheet.worksheet(title).get_all_values()
                for v in row)
        if n:
            out[title] = n
    return out
