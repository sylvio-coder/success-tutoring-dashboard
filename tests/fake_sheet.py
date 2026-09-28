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
        self.row_count = max(len(rows), 1) + 5
        self.formats_copied = []

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
            for i, vals in enumerate(d["values"]):
                for j, v in enumerate(vals):
                    self._set(r0 + i, c0 + j, v)

    def update(self, values, rng="A1"):
        self.batch_update([{"range": rng, "values": values}])

    def add_rows(self, n):
        self.row_count += n

    def append_rows(self, values, value_input_option=None):
        self._trim()
        start = len(self.cells) + 1
        for i, r in enumerate(values):
            for j, v in enumerate(r):
                self._set(start + i, j + 1, v)


class FakeSpreadsheet:
    def __init__(self, tabs):
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

    def _by_id(self, sid):
        return next(w for w in self.tabs.values() if w.id == sid)

    def batch_update(self, body):
        for req in body["requests"]:
            if "deleteDimension" in req:
                r = req["deleteDimension"]["range"]
                ws = self._by_id(r["sheetId"])
                del ws.cells[r["startIndex"]:r["endIndex"]]
                n = r["endIndex"] - r["startIndex"]
                for row in ws.cells[r["startIndex"]:]:   # Sheets re-points same-row refs
                    for i, c in enumerate(row):
                        if str(c).startswith("="):
                            row[i] = _shift_formula(str(c), -n)
            elif "copyPaste" in req:
                cp = req["copyPaste"]
                src, dst = cp["source"], cp["destination"]
                ws = self._by_id(src["sheetId"])
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
