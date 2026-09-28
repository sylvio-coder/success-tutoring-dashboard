"""In-memory stand-in for the parts of gspread the admin pages use."""
import re

import gspread
from gspread.utils import a1_to_rowcol


def _shift_formula(f, delta):
    # Relative refs like F12 -> F13; absolute ($F$12) left alone.
    return re.sub(r"(?<![$A-Z])([A-Z]{1,2})(\d+)\b",
                  lambda m: f"{m.group(1)}{int(m.group(2)) + delta}", f)


class FakeWorksheet:
    def __init__(self, spreadsheet, title, rows, sheet_id):
        self.spreadsheet = spreadsheet
        self.title = title
        self.id = sheet_id
        self.cells = [list(r) for r in rows]        # raw input (formulas kept)
        self.grid_rows = max(len(rows), 1)       # real grid size, like a full sheet
        self.row_count = self.grid_rows          # gspread's cached copy (can go stale)
        self.formats_copied = []
        self.banded_ranges = []                  # [{"bandedRangeId": ...}]
        self.tables = []                         # Sheets tables (styled by Sheets)
        self.frozen_rows = 0

    # values as displayed: formulas shown as "<f>"
    def _display(self, v):
        return "<f>" if str(v).startswith("=") else ("" if v is None else str(v))

    def _trim(self):
        while self.cells and not any(str(c) for c in self.cells[-1]):
            self.cells.pop()

    def get_all_values(self):
        self._trim()
        w = max((len(r) for r in self.cells), default=0)
        return [[self._display(c) for c in r] + [""] * (w - len(r)) for r in self.cells]

    def row_values(self, row):
        vals = self.get_all_values()
        return [v for v in vals[row - 1]] if row <= len(vals) else []

    def get_all_records(self):
        vals = self.get_all_values()
        return [dict(zip(vals[0], r)) for r in vals[1:]]

    def get(self, rng=None, value_render_option=None):
        self._trim()
        if rng is None:
            return [list(map(str, r)) for r in self.cells]
        a, b = rng.split(":")
        (r0, c0), (r1, c1) = a1_to_rowcol(a), a1_to_rowcol(b)
        return [[str(x) for x in (self.cells[r][c0 - 1:c1] if r < len(self.cells) else [])]
                for r in range(r0 - 1, r1)]

    def _check(self, r):
        if r > self.grid_rows:
            raise gspread.exceptions.GSpreadException(
                f"Range exceeds grid limits. Max rows: {self.grid_rows}")

    def _set(self, r, c, v):
        while len(self.cells) < r:
            self.cells.append([])
        row = self.cells[r - 1]
        while len(row) < c:
            row.append("")
        if isinstance(v, str) and v.startswith("'"):
            v = v[1:]
        row[c - 1] = v

    def batch_update(self, data, value_input_option=None):
        for d in data:
            rng = d["range"]
            start = rng.split(":")[0]
            r0, c0 = a1_to_rowcol(start)
            self._check(r0 + len(d["values"]) - 1)
            for i, vals in enumerate(d["values"]):
                for j, v in enumerate(vals):
                    self._set(r0 + i, c0 + j, v)

    def update(self, values, rng="A1"):
        self.batch_update([{"range": rng, "values": values}])

    def add_rows(self, n):
        self.resize(rows=self.row_count + n)

    def resize(self, rows=None, cols=None):
        if rows is not None:
            self.grid_rows = self.row_count = rows

    def append_rows(self, values, value_input_option=None):
        self._trim()
        start = len(self.cells) + 1
        self.grid_rows = max(self.grid_rows, start + len(values) - 1)
        for i, r in enumerate(values):
            for j, v in enumerate(r):
                self._set(start + i, j + 1, v)


class FakeSpreadsheet:
    FORMAT_REQUESTS = ("repeatCell", "updateSheetProperties", "updateDimensionProperties",
                       "autoResizeDimensions")

    def __init__(self, tabs, title="Weekly Membership - App (TEST)"):
        self.title = title
        self.format_requests = []
        self._next_band = 100
        self.tabs = {}
        for i, (title, rows) in enumerate(tabs.items()):
            self.tabs[title] = FakeWorksheet(self, title, rows, i + 1)

    def worksheet(self, title):
        if title not in self.tabs:
            raise gspread.WorksheetNotFound(title)
        return self.tabs[title]

    def add_worksheet(self, title, rows, cols):
        ws = FakeWorksheet(self, title, [], len(self.tabs) + 1)
        self.tabs[title] = ws
        return ws

    def fetch_sheet_metadata(self, params=None):
        return {"sheets": [{"properties": {"sheetId": w.id, "title": w.title,
                                           "gridProperties": {"rowCount": w.grid_rows,
                                                              "columnCount": 26}},
                            "bandedRanges": list(w.banded_ranges),
                            "tables": list(w.tables)}
                           for w in self.tabs.values()]}

    def _by_id(self, sid):
        return next(w for w in self.tabs.values() if w.id == sid)

    def batch_update(self, body):
        for req in body["requests"]:
            if "deleteDimension" in req and req["deleteDimension"]["range"]["dimension"] == "COLUMNS":
                r = req["deleteDimension"]["range"]
                ws = self._by_id(r["sheetId"])
                for row in ws.cells:
                    del row[r["startIndex"]:r["endIndex"]]
            elif any(k in req for k in self.FORMAT_REQUESTS):
                self.format_requests.append(req)
                if "updateSheetProperties" in req:
                    p = req["updateSheetProperties"]["properties"]
                    self._by_id(p["sheetId"]).frozen_rows = p["gridProperties"]["frozenRowCount"]
            elif "deleteBanding" in req:
                bid = req["deleteBanding"]["bandedRangeId"]
                for w in self.tabs.values():
                    w.banded_ranges = [b for b in w.banded_ranges if b["bandedRangeId"] != bid]
            elif "addBanding" in req:
                br = req["addBanding"]["bandedRange"]
                self._next_band += 1
                self._by_id(br["range"]["sheetId"]).banded_ranges.append(
                    {"bandedRangeId": self._next_band, **br})
            elif "deleteDimension" in req:
                r = req["deleteDimension"]["range"]
                ws = self._by_id(r["sheetId"])
                del ws.cells[r["startIndex"]:r["endIndex"]]
                n = r["endIndex"] - r["startIndex"]
                ws.grid_rows -= n                        # row_count is NOT updated
                for row in ws.cells[r["startIndex"]:]:   # Sheets re-points same-row refs
                    for i, c in enumerate(row):
                        if str(c).startswith("="):
                            row[i] = _shift_formula(str(c), -n)
            elif "copyPaste" in req:
                cp = req["copyPaste"]
                src, dst = cp["source"], cp["destination"]
                ws = self._by_id(src["sheetId"])
                ws._check(dst["endRowIndex"])
                if cp["pasteType"] == "PASTE_FORMAT":
                    ws.formats_copied.append((dst["startRowIndex"], dst["endRowIndex"]))
                    continue
                srow = src["startRowIndex"]
                for r in range(dst["startRowIndex"], dst["endRowIndex"]):
                    for c in range(src["startColumnIndex"], src["endColumnIndex"]):
                        f = ws.cells[srow][c] if c < len(ws.cells[srow]) else ""
                        ws._set(r + 1, c + 1, _shift_formula(str(f), r - srow))
            else:
                raise NotImplementedError(req)
