"""Writes to the Google Sheet for the admin pages.

New rows are written by column header, so the tabs' column order doesn't
matter. Formatting and any per-row formulas (e.g. the VLOOKUPs that fill
Stage/GPM/Region) are copied down from the last existing row, the same as
dragging them down by hand.
"""
import re

from gspread.utils import rowcol_to_a1

ALIASES_TAB = "Location Aliases"
ALIASES_HEADERS = ["Alias", "Location"]


def _is_row_formula(value):
    v = str(value or "")
    return v.startswith("=") and "ARRAYFORMULA" not in v.upper()


def _grid_range(ws, row0, row1, col0, col1):
    """0-based, end-exclusive GridRange."""
    return {"sheetId": ws.id, "startRowIndex": row0, "endRowIndex": row1,
            "startColumnIndex": col0, "endColumnIndex": col1}


def _text_safe(value):
    """Stop Sheets reading a week label like '1/2027' or a month like '2026-08' as a date."""
    if isinstance(value, str) and re.fullmatch(r"\d{1,2}/\d{4}|\d{4}-\d{2}", value):
        return "'" + value
    return value


def _grid_rows(ws):
    meta = ws.spreadsheet.fetch_sheet_metadata()
    for sheet in meta["sheets"]:
        if sheet["properties"]["sheetId"] == ws.id:
            return sheet["properties"]["gridProperties"]["rowCount"]
    return ws.row_count


def read_tab(ws):
    """Returns (headers, rows) with rows as lists of formatted strings."""
    values = ws.get_all_values()
    if not values:
        return [], []
    return values[0], values[1:]


def delete_rows(ws, row_numbers):
    """Delete sheet rows (1-based), merging consecutive runs."""
    rows = sorted(set(row_numbers), reverse=True)
    if not rows:
        return
    runs, start, end = [], rows[0], rows[0]
    for r in rows[1:]:
        if r == start - 1:
            start = r
        else:
            runs.append((start, end))
            start = end = r
    runs.append((start, end))
    ws.spreadsheet.batch_update({"requests": [
        {"deleteDimension": {"range": {"sheetId": ws.id, "dimension": "ROWS",
                                       "startIndex": s - 1, "endIndex": e}}}
        for s, e in runs]})


def append_records(ws, records, prefer_formula=()):
    """Append dict records below the last row, matched to the header row.

    Columns not in the records get the last row's formula (if any). Columns
    named in prefer_formula also use the formula when there is one, so
    calculated columns keep following the sheet's own logic; every other
    record value is written as a value.
    """
    if not records:
        return 0
    values = ws.get_all_values()
    headers = values[0]
    last = len(values)                       # 1-based row number of last row
    ncols = len(headers)

    template = None
    if last >= 2:
        template = ws.get(f"A{last}:{rowcol_to_a1(last, ncols)}",
                          value_render_option="FORMULA")
        template = (template[0] if template else []) + [""] * ncols
    has_formula = {i for i in range(ncols) if template and _is_row_formula(template[i])}
    fill = [i for i in sorted(has_formula)
            if headers[i] in prefer_formula or not any(headers[i] in r for r in records)]

    rows = []
    for rec in records:
        row = []
        for i, h in enumerate(headers):
            if i in fill or rec.get(h) is None:
                row.append("")
            else:
                row.append(_text_safe(rec[h]))
        rows.append(row)

    first_new = last + 1
    last_new = last + len(rows)
    # ws.row_count is cached when the worksheet is opened and goes stale
    # after rows are deleted, so ask the API for the real size.
    if _grid_rows(ws) < last_new:
        ws.resize(rows=last_new)

    requests = []
    if last >= 2:
        requests.append({"copyPaste": {
            "source": _grid_range(ws, last - 1, last, 0, ncols),
            "destination": _grid_range(ws, first_new - 1, last_new, 0, ncols),
            "pasteType": "PASTE_FORMAT"}})
    if requests:
        ws.spreadsheet.batch_update({"requests": requests})

    # Write values, leaving formula columns to the copy below.
    data = []
    for c in range(ncols):
        if c in fill:
            continue
        col = rowcol_to_a1(first_new, c + 1)
        end = rowcol_to_a1(last_new, c + 1)
        data.append({"range": f"{col}:{end}", "values": [[r[c]] for r in rows]})
    ws.batch_update(data, value_input_option="USER_ENTERED")

    if fill:
        ws.spreadsheet.batch_update({"requests": [{"copyPaste": {
            "source": _grid_range(ws, last - 1, last, c, c + 1),
            "destination": _grid_range(ws, first_new - 1, last_new, c, c + 1),
            "pasteType": "PASTE_FORMULA"}} for c in fill]})
    return len(rows)


def replace_week(ws, records, name_header, week_label, prefer_formula=(),
                 key_header="Date - Week/Year"):
    """Replace this week's (or month's, with key_header="Period") rows for these locations.

    New rows are added first and the old ones removed afterwards, so a
    failure part-way leaves the week in twice rather than missing.
    """
    headers, rows = read_tab(ws)
    ni, wi = headers.index(name_header), headers.index(key_header)
    names = {r[name_header] for r in records}
    old = [i + 2 for i, r in enumerate(rows)
           if len(r) > max(ni, wi) and r[wi].strip() == week_label and r[ni] in names]
    added = append_records(ws, records, prefer_formula)
    delete_rows(ws, old)          # rows above the new ones: numbers still valid
    return len(old), added


def update_cells(ws, updates):
    """updates: list of (row, col, value), 1-based."""
    if not updates:
        return
    ws.batch_update([{"range": rowcol_to_a1(r, c), "values": [[_text_safe(v)]]}
                     for r, c, v in updates], value_input_option="USER_ENTERED")


def get_or_add_tab(spreadsheet, title, headers):
    """The tab, created with this header row if it doesn't exist yet."""
    import gspread
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title, rows=1000, cols=len(headers))
        ws.update([headers], "A1")
        return ws


def get_aliases_ws(spreadsheet, create=True):
    import gspread
    try:
        return spreadsheet.worksheet(ALIASES_TAB)
    except gspread.WorksheetNotFound:
        if not create:
            return None
        ws = spreadsheet.add_worksheet(ALIASES_TAB, rows=200, cols=len(ALIASES_HEADERS))
        ws.update([ALIASES_HEADERS], "A1")
        return ws


def read_aliases(spreadsheet):
    ws = get_aliases_ws(spreadsheet, create=False)
    if ws is None:
        return {}
    return {r["Alias"]: r["Location"] for r in ws.get_all_records()
            if str(r.get("Alias", "")).strip() and str(r.get("Location", "")).strip()}


def add_aliases(spreadsheet, pairs):
    """pairs: list of (alias, location)."""
    if not pairs:
        return
    ws = get_aliases_ws(spreadsheet)
    ws.append_rows([[a, l] for a, l in pairs], value_input_option="RAW")
