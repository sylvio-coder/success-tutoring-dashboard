"""Franchise Model inputs: the assumptions workbook (Worst / Base / Best for AU National and
each state) that an admin uploads, checked, compared with the last upload and stored in the
'Franchise Model' tab. Calculations only; the upload page is in admin_pages.py.

Workbook layout (one tab per model, e.g. 'AU National', 'NSW'): a header row with
Variable / Unit / Worst / Base / Best / Basis / Note; a row with only a name in column A is a
section heading. Tabs without that header row (How to use, guides) are ignored.
"""
import io
import re

import openpyxl

TAB = "Franchise Model"
HEADERS = ["Version", "Uploaded", "Model", "Section", "Variable", "Unit", "Worst", "Base",
           "Best", "Basis", "Note"]
CASES = ("Worst", "Base", "Best")
KEEP_VERSIONS = 2                       # the current upload and the one before it
NOT_SET = "?"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Inputs the model reads, and the kind of value each must hold. Every model tab needs all of
# them before it can be used ("ready"). 'number': a number (Worst / Best may be blank =
# same as Base). The other kinds are checked as described in check_value().
NUMBER = "number"
KINDS = {name: NUMBER for name in [
    "Franchise grant", "Success Pack", "Reservation deposit", "Fit-out and signage",
    "Legal and accounting (company set-up)", "Legal costs for the lease",
    "Presale marketing budget (weeks 1–12)", "Rental bond", "Working capital buffer",
    "Owner's cash share of total cash needed", "Loan share of funding", "Loan interest rate",
    "Loan term", "Exchange rate (1 unit of local currency in AUD)", "Payback target",
    "Required return (discount rate)", "Business sale value at Year 5", "Company tax rate",
    "Fit-out depreciation", "Owner's wage", "Franchise term", "Renewal fee",
    "Lease starts before week 1", "First sign-ups", "Centre opens (sessions and billing start)",
    "Presale campaign ends", "Model length", "Members at week 12",
    *[f"Members at end of Year {y}" for y in range(1, 6)],
    *[f"Year {y}: cancelled" for y in range(1, 6)],
    "Study hub tables in the centre", "Standard price (members after the first 50)",
    "Foundation members leaving", "Collected after GST (allowance for uncollected)", "GST",
    "Price rise each July (from 12 months)", "Joining fee", "Suspension fee",
    "Other revenue (holiday programs, assessments)", "Student visits per member",
    "1:1 share of visits", "Students per study hub: Year 1, Q1 (weeks 1–13)",
    "Students per study hub: Year 1, Q2", "Students per study hub: Year 1, Q3",
    "Students per study hub: Year 1, Q4",
    *[f"Students per study hub: Year {y}" for y in range(2, 6)],
    "Most students per study hub", "Study hub length", "Tutors per study hub",
    "Target students per study hub (planning)", "Merge study hubs below", "Weekdays open",
    "Peak time slots: weekday before 7 pm", "Weekday slots after 7 pm (busier centres)",
    "Peak time slots: Saturday", "Share of study hubs on Saturday", "Minimum paid shift",
    "Tutor hourly rate: weekday (casual, level 2, age 20)",
    "Tutor hourly rate: Saturday and after 7 pm (casual)", "Super",
    "Manager salary (ex super)", "Front desk / admin", "Wage rise each July",
    "Rent (including outgoings)", "Rent-free months", "Rent rise each July", "Lease term",
    "Agency fee (at least)", "Advertising", "Marketing fund levy",
    "Enquiry → Student Assessment", "Student Assessment → member",
    "Cost per lead: presale (weeks 1–12)", "Cost per lead: after presale", "Royalty",
    "Royalty minimum (FY2026–27)", "Royalty minimum rise each July", "CRM software (Hapana)",
    "Learning materials, printing and stationery", "Other software and subscriptions",
    "Office expenses", "Insurance", "Energy and water", "Internet and phone",
    "Accounting, bookkeeping and consulting", "Cleaning", "Repairs and maintenance",
    "Workers compensation", "Bank fees", "Staff expenses and uniforms",
    "Running costs rise each July", "Replacement spending (computers, furniture)"]}
KINDS.update({
    "Hire a manager at": "manager",                 # members, 'when sessions start', 'never'
    "Owner's wage included": "yes_no",
    "Saturday open": "yes_no",
    "Owner's wage starts": ("first profitable month", "when sessions start"),
    "Sign-up seasonality (1.0 = average month)": ("numbers", 12),
    "Extra members suspended in holidays": "month_values",
    "School holiday months": "months",
    "Package mix: Group 1× / 2× / 3×+ / 1:1": ("numbers", 4),
    "Foundation member numbers (tiers)": ("numbers", 3),
    "Foundation prices (tiers)": ("numbers", 3),
})
REQUIRED = list(KINDS)


class InputsError(Exception):
    pass


# ── Reading values ────────────────────────────────────────────────────────────
def to_number(value):
    """A number from a cell (int/float, or text like '110,000', '$70', '4.75%'); else None."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value or "").strip().lstrip("'").replace(",", "").replace("$", "").rstrip("%")
    try:
        return float(s)
    except ValueError:
        return None


def parse_numbers(text):
    """'50 / 60 / 65' or '0.76, 1.82, …' -> [50.0, 60.0, 65.0]; None if any part isn't one."""
    parts = [p for p in re.split(r"[/,]", str(text or "")) if p.strip()]
    nums = [to_number(p) for p in parts]
    return nums if parts and all(n is not None for n in nums) else None


def parse_month_values(text):
    """'Jan 17, Apr 9' -> {'01': 17.0, '04': 9.0}; None if it can't be read."""
    out = {}
    for part in [p.strip() for p in str(text or "").split(",") if p.strip()]:
        m = re.fullmatch(r"([A-Za-z]{3})\w*\s+(-?[\d.]+)%?", part)
        if not m or m.group(1).title() not in MONTHS:
            return None
        out[f"{MONTHS.index(m.group(1).title()) + 1:02d}"] = float(m.group(2))
    return out


def parse_months(text):
    """'Jan Apr Jul' -> ['01', '04', '07']; [] for '–' or blank; None if a word isn't a month."""
    words = re.split(r"[\s,]+", str(text or "").strip())
    words = [w for w in words if w and w not in ("–", "-")]
    if any(w[:3].title() not in MONTHS for w in words):
        return None
    return [f"{MONTHS.index(w[:3].title()) + 1:02d}" for w in words]


def check_value(kind, value, case):
    """A problem with this value for its kind, or None. Blank Worst / Best = same as Base."""
    text = str(value if value is not None else "").strip()
    if text == NOT_SET:
        return "not set yet (?)"
    if text == "" or text.lower() == "calc":
        return "is blank" if case == "Base" else None
    if kind == NUMBER:
        return None if to_number(value) is not None else f"should be a number, not '{text}'"
    if kind == "manager":
        return None if to_number(value) is not None or text.lower() in (
            "when sessions start", "never") else \
            "should be a number of members, 'when sessions start' or 'never'"
    if kind == "yes_no":
        return None if text.lower() in ("yes", "no") else "should be yes or no"
    if isinstance(kind, tuple) and kind[0] == "numbers":
        nums = parse_numbers(text)
        return None if nums and len(nums) == kind[1] else \
            f"should be {kind[1]} numbers separated by / or ,"
    if isinstance(kind, tuple):
        return None if text.lower() in kind else "should be " + " or ".join(f"'{k}'" for k in kind)
    if kind == "month_values":
        return None if parse_month_values(text) is not None else \
            "should look like 'Jan 17, Apr 9' (month and %)"
    if kind == "months":
        return None if text.lower() == "by state" or parse_months(text) is not None else \
            "should be month names, e.g. 'Jan Apr Jul Oct Dec'"
    return None


# ── The workbook ──────────────────────────────────────────────────────────────
def _header(ws):
    for r in range(1, min(ws.max_row, 10) + 1):
        if str(ws.cell(r, 1).value or "").strip() == "Variable":
            return r, {str(ws.cell(r, c).value or "").strip(): c
                       for c in range(1, ws.max_column + 1)}
    return None, None


def read_workbook(data):
    """Rows (dicts with HEADERS[2:]) for every model tab, and problems: [(model, variable,
    message)]. Raises InputsError if the file can't be read or has no model tabs."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        raise InputsError(f"This file can't be read as an Excel workbook ({e}).") from e
    rows, problems = [], []
    for ws in wb.worksheets:
        hr, cols = _header(ws)
        if not hr:
            continue
        missing = [h for h in ("Variable", *CASES) if h not in cols]
        if missing:
            problems.append((ws.title, "", "missing column(s): " + ", ".join(missing)))
            continue
        section, seen = "", set()
        for r in range(hr + 1, ws.max_row + 1):
            name = str(ws.cell(r, 1).value or "").strip()
            if not name:
                continue
            get = lambda h: ws.cell(r, cols[h]).value if h in cols else None
            if all(get(h) in (None, "") for h in ("Unit", *CASES)):
                section = name
                continue
            if name in seen:
                problems.append((ws.title, name, "appears twice"))
                continue
            seen.add(name)
            rec = {"Model": ws.title, "Section": section, "Variable": name}
            for h in ("Unit", *CASES, "Basis", "Note"):
                v = get(h)
                rec[h] = "" if v is None else (v if isinstance(v, (int, float)) else str(v).strip())
            rows.append(rec)
    if not rows and not problems:
        raise InputsError("No model tabs found. Each model tab needs a header row with "
                          "Variable, Unit, Worst, Base and Best.")
    return rows, problems


def readiness(rows):
    """{model: [problems]} for every model: missing inputs and values that can't be used.
    A model with no problems is ready for the report."""
    out = {}
    for model in dict.fromkeys(r["Model"] for r in rows):
        mine = {r["Variable"]: r for r in rows if r["Model"] == model}
        probs = [f"{v}: missing" for v in REQUIRED if v not in mine]
        for v in REQUIRED:
            if v not in mine:
                continue
            for case in CASES:
                p = check_value(KINDS[v], mine[v][case], case)
                if p:
                    probs.append(f"{v} ({case}): {p}")
        out[model] = probs
    return out


# ── Stored versions ───────────────────────────────────────────────────────────
def from_tab(headers, values):
    """The stored rows (strings from the tab) -> {version: [rows]}."""
    if not headers or "Version" not in headers:
        return {}
    out = {}
    for v in values:
        rec = {h: (v[i].lstrip("'") if i < len(v) else "") for i, h in enumerate(headers)}
        ver = to_number(rec.get("Version"))
        if ver is None or not rec.get("Model"):
            continue
        out.setdefault(int(ver), []).append(rec)
    return out


def latest(versions):
    """(version number, rows, uploaded) of the latest version, or (0, [], '')."""
    if not versions:
        return 0, [], ""
    v = max(versions)
    return v, versions[v], versions[v][0].get("Uploaded", "")


def _same(a, b):
    na, nb = to_number(a), to_number(b)
    if na is not None and nb is not None:
        return abs(na - nb) < 1e-9
    return str(a if a is not None else "").strip() == str(b if b is not None else "").strip()


def changes(old_rows, new_rows):
    """What changed from the last upload: [(model, variable, what, old, new)]."""
    key = lambda r: (r["Model"], r["Variable"])
    old, new = {key(r): r for r in old_rows}, {key(r): r for r in new_rows}
    out = []
    for k, r in new.items():
        if k not in old:
            out.append((*k, "added", "", r.get("Base", "")))
            continue
        for case in CASES:
            if not _same(old[k].get(case), r.get(case)):
                out.append((*k, case, old[k].get(case, ""), r.get(case, "")))
    for k, r in old.items():
        if k not in new:
            out.append((*k, "removed", r.get("Base", ""), ""))
    return out


def records(rows, version, uploaded):
    """Rows ready to write to the tab. Text is stored as text (a leading ') so Sheets never
    turns '20 / 20 / 10' or '1–8' into a date or a formula."""
    out = []
    for r in rows:
        rec = {"Version": version, "Uploaded": "'" + str(uploaded)}
        for h in HEADERS[2:]:
            v = r.get(h, "")
            rec[h] = v if isinstance(v, (int, float)) and not isinstance(v, bool) else \
                ("'" + str(v) if str(v) != "" else "")
        out.append(rec)
    return out


def model_inputs(rows, model):
    """{variable: {'Worst', 'Base', 'Best', 'Unit', 'Section', 'Basis', 'Note'}} for one model
    (blank Worst / Best = Base). Numbers come back as floats, text as text."""
    out = {}
    for r in rows:
        if r["Model"] != model:
            continue
        vals = {}
        for case in CASES:
            v = r.get(case, "")
            v = v.lstrip("'") if isinstance(v, str) else v
            n = to_number(v)
            vals[case] = n if n is not None and KINDS.get(r["Variable"]) in (NUMBER, "manager") \
                else v
        for case in ("Worst", "Best"):
            if vals[case] in ("", None):
                vals[case] = vals["Base"]
        out[r["Variable"]] = {**vals, **{h: r.get(h, "") for h in
                                         ("Unit", "Section", "Basis", "Note")}}
    return out


# ── Reading and writing the tab ───────────────────────────────────────────────
def read_versions(spreadsheet):
    """{version: [rows]} stored in the tab ({} if the tab doesn't exist yet)."""
    import gspread
    import sheet_writer as sw
    try:
        ws = spreadsheet.worksheet(TAB)
    except gspread.WorksheetNotFound:
        return {}
    return from_tab(*sw.read_tab(ws))


def save_version(spreadsheet, rows, uploaded):
    """Write these rows as a new version and keep only the latest KEEP_VERSIONS versions.
    New rows go in first and old ones are removed after, so a failure part-way leaves an
    extra version rather than none. Returns the new version number."""
    import sheet_writer as sw
    ws = sw.get_or_add_tab(spreadsheet, TAB, HEADERS)
    versions = from_tab(*sw.read_tab(ws))
    new = (max(versions) if versions else 0) + 1
    sw.append_records(ws, records(rows, new, uploaded))
    keep = set(sorted(versions)[-(KEEP_VERSIONS - 1):] + [new]) if KEEP_VERSIONS > 1 else {new}
    headers, values = sw.read_tab(ws)
    vi = headers.index("Version")
    sw.delete_rows(ws, [i + 2 for i, r in enumerate(values)
                        if len(r) > vi and to_number(r[vi]) not in keep])
    return new
