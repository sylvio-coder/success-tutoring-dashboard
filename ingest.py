"""Weekly CRM export ingestion.

Parses the weekly exports from Hapana and the in-house CRM, matches every
location to the master location list (Vlookup tab), and produces the rows to
write to the Weekly Membership and Revenue tabs.

Pure pandas: nothing here talks to Streamlit or Google, so it can be tested
on its own.
"""
import difflib
import io
import re
from datetime import date, datetime, timedelta

import pandas as pd

PREFIX = "Success Tutoring - "
EXCLUDED_NAME_WORDS = ("sandbox",)

# Hapana reports gross revenue including GST; net = gross / divisor.
# In-house CRM reports net revenue itself, which is used as given.
GST_DIVISOR = {"australia": 1.10, "new zealand": 1.15}

STUDENT_VISITS_PER_STUDENT = 1.96

# Hapana names two sites "Epping" and relies on the Region column to tell them
# apart; the master list calls the Victorian one "Epping VIC".
AU_STATES = {
    "new south wales": "NSW", "victoria": "VIC", "queensland": "QLD",
    "western australia": "WA", "south australia": "SA", "tasmania": "TAS",
    "australian capital territory": "ACT", "northern territory": "NT",
}


def state_abbr(region):
    """'Victoria' or 'VIC' -> 'VIC'; None for anything else."""
    r = str(region or "").strip().lower()
    if r in AU_STATES:
        return AU_STATES[r]
    return r.upper() if r.upper() in AU_STATES.values() else None

HAPANA_MEMBERS = "hapana_members"
HAPANA_REVENUE = "hapana_revenue"
INHOUSE_MEMBERS = "inhouse_members"
INHOUSE_REVENUE = "inhouse_revenue"

CRM_OF = {HAPANA_MEMBERS: "Hapana", HAPANA_REVENUE: "Hapana",
          INHOUSE_MEMBERS: "In-house", INHOUSE_REVENUE: "In-house"}

KIND_LABELS = {
    HAPANA_MEMBERS: "Hapana – members",
    HAPANA_REVENUE: "Hapana – revenue",
    INHOUSE_MEMBERS: "In-house – members",
    INHOUSE_REVENUE: "In-house – revenue",
}

# Header cells (normalised) that identify each export.
SIGNATURES = {
    HAPANA_MEMBERS: {"date - week/year", "# active members", "# new members"},
    HAPANA_REVENUE: {"date - week/year", "total sessions"},
    INHOUSE_MEMBERS: {"location name", "active members", "new members"},
    INHOUSE_REVENUE: {"location name", "sessions", "gross revenue"},
}

# Column name (normalised) -> standard field, per export kind.
COLUMNS = {
    HAPANA_MEMBERS: {
        "# active members": "active", "# suspended members": "suspended",
        "# cancelled members": "cancelled", "# new members": "new",
    },
    HAPANA_REVENUE: {
        "# active members - date range end": "active", "# active members": "active",
        "total sessions": "sessions",
        "average gross revenue per location": "gross", "gross revenue": "gross",
    },
    INHOUSE_MEMBERS: {
        "active members": "active", "suspended members": "suspended",
        "cancelled members": "cancelled", "new members": "new",
    },
    INHOUSE_REVENUE: {
        "sessions": "sessions", "gross revenue": "gross", "net revenue (ex tax)": "net",
    },
}


class ExportError(ValueError):
    pass


# ── Small helpers ─────────────────────────────────────────────────────────────
def norm_header(value):
    s = "" if value is None else str(value)
    s = s.replace("\\", "").replace("\n", " ").strip().lower()
    return re.sub(r"\s+", " ", s)


def location_key(name):
    """Comparison key for a location name: case-insensitive, prefix removed."""
    s = re.sub(r"[\u2010-\u2015\u2212]", "-", str(name or ""))   # en/em dashes -> "-"
    s = re.sub(r"\s+", " ", s).strip().lower()
    for p in ("success tutoring - ", "success tutoring -", "success tutoring "):
        if s.startswith(p):
            s = s[len(p):]
            break
    return s.strip()


def parse_number(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return 0.0 if pd.isna(value) else float(value)
    s = str(value).strip().replace("$", "").replace(",", "").replace(" ", "")
    if s in ("", "-", "nan", "None"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return -n if neg else n


def parse_date(value):
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date() if hasattr(value, "date") else value
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def week_label(d):
    """Week label as stored in the sheet: Sunday-start week number, e.g. '38/2026'."""
    return f"{int(d.strftime('%U'))}/{d.year}"


def week_label_to_date(label):
    """'38/2026' -> the Sunday that starts that week (2026-09-20)."""
    m = re.fullmatch(r"\s*(\d{1,2})/(\d{4})\s*", str(label))
    if not m:
        return None
    week, year = int(m.group(1)), int(m.group(2))
    jan1 = date(year, 1, 1)
    first_sunday = jan1 + timedelta(days=(6 - jan1.weekday()) % 7)
    return first_sunday + timedelta(weeks=week - 1)


def sheet_date(d):
    return f"{d.day}/{d.month}/{d.year}"


def is_excluded(name):
    n = str(name).lower()
    return any(w in n for w in EXCLUDED_NAME_WORDS)


# ── Reading one export ────────────────────────────────────────────────────────
def _read_grids(filename, data):
    """All tabs of the file as grids (Hapana puts 'Export info' on the first
    tab and the data on the second)."""
    if filename.lower().endswith(".csv"):
        return [pd.read_csv(io.BytesIO(data), header=None, dtype=str,
                            keep_default_na=False, skip_blank_lines=False,
                            on_bad_lines="skip", engine="python")]
    return list(pd.read_excel(io.BytesIO(data), header=None, dtype=object,
                              sheet_name=None).values())


def _find_header(grid):
    for i in range(min(len(grid), 40)):
        cells = {norm_header(v) for v in grid.iloc[i].tolist()}
        for kind, sig in SIGNATURES.items():
            if sig <= cells:
                if kind == HAPANA_REVENUE and not any("gross revenue" in c for c in cells):
                    continue
                if kind == HAPANA_MEMBERS and "total sessions" in cells:
                    continue
                return i, kind
    return None, None


def _date_range_from_cells(cells):
    """Find 'dd/mm/yyyy - dd/mm/yyyy' in the export-info rows above the header."""
    for c in cells:
        m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s*-\s*(\d{1,2}/\d{1,2}/\d{4})", str(c))
        if m:
            return parse_date(m.group(1)), parse_date(m.group(2))
    return None


def _date_range_from_filename(filename):
    m = re.search(r"(\d{4}-\d{2}-\d{2})-to-(\d{4}-\d{2}-\d{2})", filename)
    if m:
        return parse_date(m.group(1)), parse_date(m.group(2))
    return None


def read_export(filename, data):
    """Parse one uploaded export.

    Returns dict(kind, rows, week_end, warnings) where rows is a DataFrame with
    columns 'name' plus the standard fields for that kind.
    """
    try:
        grids = _read_grids(filename, data)
    except Exception as e:
        raise ExportError(f"{filename}: could not be read ({e})")
    grid, hdr, kind = None, None, None
    for g in grids:
        hdr, kind = _find_header(g)
        if kind is not None:
            grid = g
            break
    if kind is None:
        raise ExportError(f"{filename}: not recognised as a Hapana or in-house "
                          "members/revenue export")

    headers = [norm_header(v) for v in grid.iloc[hdr].tolist()]
    body = grid.iloc[hdr + 1:].reset_index(drop=True)

    region_idx = headers.index("region") if "region" in headers else None
    name_idx = next((i for i, h in enumerate(headers)
                     if h in ("location name", "business name")
                     or h.endswith("business name")), 0)
    week_idx = headers.index("date - week/year") if "date - week/year" in headers else None
    field_idx = {}
    for i, h in enumerate(headers):
        f = COLUMNS[kind].get(h)
        # In-house foundation members count as active members.
        if kind == INHOUSE_MEMBERS and "foundation" in h:
            f = "foundation"
        if f and f not in field_idx:
            field_idx[f] = i
    warnings = []

    def cell(vals, i):
        if i is None or i >= len(vals):
            return ""
        v = vals[i]
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        return re.sub(r"\s+", " ", str(v)).strip()

    # Hapana blanks a cell that repeats the row above: Region within a group,
    # and the Business name of the second "Epping" when two sit together.
    records, weeks = [], set()
    region = last_name = ""
    for _, r in body.iterrows():
        vals = r.tolist()
        region = cell(vals, region_idx) or region
        name = cell(vals, name_idx)
        if not name and last_name and cell(vals, week_idx):
            name = last_name
        if not name:
            continue
        last_name = name
        if norm_header(name).startswith(("total", "grand total")):
            continue
        rec = {"name": name, "region": region}
        for f, i in field_idx.items():
            rec[f] = parse_number(vals[i]) if i < len(vals) else 0.0
        if "foundation" in rec:
            rec["active"] = rec.get("active", 0.0) + rec.pop("foundation")
        if week_idx is not None and week_idx < len(vals) and vals[week_idx] is not None:
            w = str(vals[week_idx]).strip()
            if w and w.lower() != "nan":
                weeks.add(w)
        records.append(rec)

    rows = pd.DataFrame(records)
    if rows.empty:
        raise ExportError(f"{filename}: no location rows found")

    # Work out the week this export covers.
    week_end = None
    info_cells = [v for row in grid.iloc[:hdr].values.tolist() for v in row] + \
        [v for g in grids if g is not grid for row in g.head(40).values.tolist() for v in row]
    rng = _date_range_from_cells(info_cells) \
        or _date_range_from_filename(filename)
    if rng and all(rng):
        start, end = rng
        if (end - start).days > 7:
            raise ExportError(f"{filename}: covers {start:%d/%m/%Y}–{end:%d/%m/%Y}, "
                              "which is longer than one week")
        week_end = end
    if len(weeks) > 1:
        raise ExportError(f"{filename}: contains more than one week ({', '.join(sorted(weeks))})")
    if weeks:
        from_label = week_label_to_date(next(iter(weeks)))
        if week_end is None:
            week_end = from_label
        elif from_label and from_label != week_end:
            warnings.append(f"{filename}: date range ends {week_end:%d/%m/%Y} but the "
                            f"week column says {next(iter(weeks))}")

    return {"filename": filename, "kind": kind, "rows": rows,
            "week_end": week_end, "warnings": warnings}


# ── Name matching ─────────────────────────────────────────────────────────────
class LocationMatcher:
    def __init__(self, master_names, aliases=None, regions=None):
        self.master, self.region = {}, {}
        for i, n in enumerate(master_names):
            if str(n).strip():
                self.master.setdefault(location_key(n), str(n).strip())
                if regions is not None:
                    self.region.setdefault(location_key(n), str(regions[i] or "").strip().lower())
        self.aliases = {}
        for alias, target in (aliases or {}).items():
            if str(alias).strip() and str(target).strip():
                self.aliases[alias_key(alias)] = str(target).strip()

    def resolve(self, raw, region="", crm=None):
        """Master name for raw, or None. A mapping saved for this CRM only
        (e.g. 'In-house: Belmont') wins. With an Australian region, 'Epping' in
        Victoria resolves to 'Epping VIC' rather than the NSW 'Epping'."""
        if crm:
            scoped = self.aliases.get(alias_key(scoped_alias(crm, raw)))
            if scoped:
                return scoped
        abbr = state_abbr(region)
        if not abbr:
            k = location_key(raw)
            return self.aliases.get(k) or self.master.get(k)
        for cand in (f"{raw} {abbr}", raw):
            k = location_key(cand)
            if k in self.aliases:
                return self.aliases[k]
            loc = self.master.get(k)
            # Only a different known state rules a match out.
            if loc and state_abbr(self.region.get(k)) in (None, abbr):
                return loc
        return None

    def unknown_label(self, raw, region=""):
        """Name to show for an unmatched row: 'Epping VIC' when 'Epping' exists
        but in another state."""
        abbr = state_abbr(region)
        if abbr and location_key(raw) in self.master:
            return f"{raw} {abbr}"
        return raw

    def suggestions(self, raw, n=3):
        k = location_key(split_scope(raw)[1])
        keys = difflib.get_close_matches(k, list(self.master), n=n, cutoff=0.6)
        # "Belmont WA" -> also suggest "Belmont" (prefix/containment matches)
        for mk in self.master:
            if mk not in keys and (k.startswith(mk + " ") or mk.startswith(k + " ")):
                keys.append(mk)
        return [self.master[x] for x in keys[:n]]


def scoped_alias(crm, raw):
    """Alias text that applies to one CRM's exports only."""
    return f"{crm}: {raw}"


def split_scope(alias):
    m = re.match(r"\s*(hapana|in-house)\s*:\s*(.*)$", str(alias), re.I)
    return (m.group(1).lower(), m.group(2)) if m else (None, str(alias))


def alias_key(alias):
    crm, name = split_scope(alias)
    return f"{crm}:{location_key(name)}" if crm else location_key(name)


def canonical_new_name(raw):
    s = re.sub(r"\s+", " ", split_scope(raw)[1]).strip()
    return s if s.lower().startswith("success tutoring") else PREFIX + s


# ── Combining the exports for one week ────────────────────────────────────────
def combine(exports, master, aliases=None):
    """Combine parsed exports into rows for the sheet.

    exports: list of read_export() results.
    master:  DataFrame of the Vlookup tab (needs 'Location' and 'Country').
    aliases: {alias name: master location name}.

    Returns dict with membership / revenue DataFrames (keyed by master
    location name), unknown names, issues and the detected week.
    """
    matcher = LocationMatcher(master["Location"].tolist(), aliases,
                              master["Region"].tolist() if "Region" in master.columns else None)
    country_of = {location_key(r["Location"]): str(r.get("Country", "")).strip()
                  for _, r in master.iterrows()}
    status_of = {r["Location"]: str(r.get("Status", "")).strip().lower()
                 for _, r in master.iterrows()}

    # The in-house CRM lists sites with no members and no revenue (e.g. a
    # placeholder "Epping", Head Office): leave those out.
    inhouse = {}
    for e in exports:
        if CRM_OF[e["kind"]] != "In-house":
            continue
        for _, r in e["rows"].iterrows():
            t = inhouse.setdefault(location_key(r["name"]), {"name": r["name"], "any": 0.0})
            for f in ("active", "gross", "net"):
                t["any"] += abs(float(r.get(f, 0) or 0))
    empty_inhouse = {k: t["name"] for k, t in inhouse.items() if t["any"] == 0}

    # Hapana's members export separates same-named sites by Region, but its
    # revenue export has no Region and adds them together into one row.
    shared = {}
    for e in exports:
        if e["kind"] == HAPANA_MEMBERS:
            for _, r in e["rows"].iterrows():
                if r.get("region"):
                    shared.setdefault(location_key(r["name"]), set()).add(r["region"])
    shared = {k: v for k, v in shared.items() if len(v) > 1}

    errors, warnings, info = [], [], []
    for e in exports:
        warnings.extend(e["warnings"])

    weeks = {e["week_end"] for e in exports if e["week_end"]}
    week_end = None
    if len(weeks) > 1:
        errors.append("The files are for different weeks: " +
                      ", ".join(f"{w:%d/%m/%Y}" for w in sorted(weeks)))
    elif weeks:
        week_end = next(iter(weeks))

    kinds = {e["kind"] for e in exports}
    for k in (HAPANA_MEMBERS, HAPANA_REVENUE, INHOUSE_MEMBERS, INHOUSE_REVENUE):
        if k not in kinds:
            warnings.append(f"No {KIND_LABELS[k]} file uploaded.")

    unknown = {}      # raw name -> {sources, suggestions}
    excluded = set()
    members, revenue = {}, {}   # location -> record
    seen = {}                   # (table, location) -> (source, crm, raw)
    clashes = {}                # (crm, raw) to remap -> clash details

    def take(table, store, loc, rec, source, raw, crm):
        key = (table, loc)
        if key in seen:
            s0, c0, r0 = seen[key]
            errors.append(f"{loc} appears more than once in the {table} data "
                          f"({s0} '{r0}' and {source} '{raw}'). "
                          "Two different sites have the same name.")
            # Ask about the in-house row (Hapana rows carry a region).
            pick = (c0, r0) if c0 == "In-house" and crm != "In-house" else (crm, raw)
            other = (crm, raw) if pick == (c0, r0) else (c0, r0)
            clashes.setdefault(pick, {"crm": pick[0], "name": pick[1], "location": loc,
                                      "other": f"{other[0]} '{other[1]}'",
                                      "suggestions": matcher.suggestions(pick[1])})
            return
        seen[key] = (source, crm, raw)
        store[loc] = rec

    for e in exports:
        source, crm = KIND_LABELS[e["kind"]], CRM_OF[e["kind"]]
        for _, r in e["rows"].iterrows():
            raw = r["name"]
            if is_excluded(raw):
                excluded.add(raw)
                continue
            if crm == "In-house" and location_key(raw) in empty_inhouse:
                continue
            if e["kind"] == HAPANA_REVENUE and not r.get("region") and \
                    location_key(raw) in shared:
                sites = sorted(filter(None, (matcher.resolve(raw, reg, crm)
                                             for reg in shared[location_key(raw)])))
                warnings.append(
                    f"Hapana revenue has one '{raw}' row for {len(sites)} sites "
                    f"({', '.join(sites)}): Hapana adds same-named sites together in that "
                    "report, so this row is left out and those sites get no revenue this "
                    "week. To fix it at the source, give the sites different names in Hapana "
                    "or add Region to the revenue report.")
                continue
            region = r.get("region", "") or ""
            loc = matcher.resolve(raw, region, crm)
            if loc is None:
                raw = matcher.unknown_label(raw, region)
                u = unknown.setdefault(raw, {"sources": set(),
                                             "suggestions": matcher.suggestions(raw),
                                             "active": 0.0})
                u["sources"].add(source)
                u["active"] = max(u["active"], float(r.get("active", 0) or 0))
                continue
            if e["kind"] in (HAPANA_MEMBERS, INHOUSE_MEMBERS):
                take("membership", members, loc, {
                    "active": r.get("active", 0.0), "suspended": r.get("suspended", 0.0),
                    "cancelled": r.get("cancelled", 0.0), "new": r.get("new", 0.0),
                    "source": source}, source, raw, crm)
            else:
                rec = {"active": r.get("active"), "sessions": r.get("sessions", 0.0),
                       "gross": r.get("gross", 0.0), "source": source}
                if e["kind"] == INHOUSE_REVENUE:
                    rec["net"] = r.get("net", 0.0)
                    rec["active"] = None
                else:
                    country = country_of.get(location_key(loc), "")
                    div = GST_DIVISOR.get(country.lower())
                    if div is None:
                        errors.append(f"{loc}: no GST rule for country "
                                      f"'{country or 'blank'}' — set its Country on the "
                                      "Locations page.")
                        rec["net"] = 0.0
                    else:
                        rec["net"] = round(rec["gross"] / div, 2)
                take("revenue", revenue, loc, rec, source, raw, crm)

    # In-house revenue has no member count: take it from the members export.
    for loc, rec in revenue.items():
        if rec["active"] is None or pd.isna(rec["active"]):
            rec["active"] = members.get(loc, {}).get("active", 0.0)

    if excluded:
        info.append("Excluded test account(s): " + ", ".join(sorted(excluded)))
    if empty_inhouse:
        names = sorted(empty_inhouse.values())
        info.append("Left out in-house rows with no members and no revenue: " + ", ".join(names))

    # Trading sites missing from both CRMs get a 0-member row so the gap shows
    # in the charts, and are flagged: the partner may not be using the system.
    missing = sorted(loc for loc, s_ in status_of.items()
                     if s_ == "trading" and loc not in members)
    if missing:
        if {HAPANA_MEMBERS, INHOUSE_MEMBERS} <= kinds:
            for loc in missing:
                members[loc] = {"active": 0.0, "suspended": 0.0, "cancelled": 0.0,
                                "new": 0.0, "source": "Not in either CRM (follow up)"}
            warnings.append("**Follow up:** these Trading locations have no data from either "
                            "CRM this week and will be written with 0 members. The franchise "
                            "partner may not be recording at location level: "
                            + ", ".join(missing) + ". If a site has closed, change its Status "
                            "on the Locations page.")
        else:
            warnings.append("**Follow up:** these Trading locations have no membership row: "
                            + ", ".join(missing) + ". 0-member rows are only added when both "
                            "the Hapana and in-house members files are uploaded.")
    no_members = sorted(set(revenue) - set(members))
    if no_members:
        info.append("Revenue but no membership row (normal for presale sites): "
                    + ", ".join(no_members))
    no_revenue = sorted(set(members) - set(revenue))
    if no_revenue:
        info.append("Membership but no revenue row: " + ", ".join(no_revenue))

    mem_df = pd.DataFrame([{"Location": k, **v} for k, v in sorted(members.items())])
    rev_df = pd.DataFrame([{"Location": k, **v} for k, v in sorted(revenue.items())])
    if not rev_df.empty:
        rev_df = add_revenue_metrics(rev_df)

    for raw, u in unknown.items():
        u["sources"] = sorted(u["sources"])

    return {"week_end": week_end, "membership": mem_df, "revenue": rev_df,
            "unknown": unknown, "clashes": list(clashes.values()),
            "errors": errors, "warnings": warnings, "info": info}


def _div(a, b):
    return round(a / b, 2) if b else 0.0


def add_revenue_metrics(df):
    df = df.copy()
    df["visits"] = (df["active"] * STUDENT_VISITS_PER_STUDENT).round(1)
    df["rev_per_session"] = [_div(n, s) for n, s in zip(df["net"], df["sessions"])]
    df["rev_per_student"] = [_div(n, a) for n, a in zip(df["net"], df["active"])]
    df["sessions_per_student"] = [_div(s, a) for s, a in zip(df["sessions"], df["active"])]
    df["students_per_session"] = [_div(a, s) for a, s in zip(df["active"], df["sessions"])]
    df["sessions_per_visit"] = [_div(s, v) for s, v in zip(df["sessions"], df["visits"])]
    df["visits_per_session"] = [_div(v, s) for v, s in zip(df["visits"], df["sessions"])]
    return df


# ── Sheet rows ────────────────────────────────────────────────────────────────
def membership_records(mem_df, week_end):
    wl, ds = week_label(week_end), sheet_date(week_end)
    return [{
        "Success Tutoring - Business name": r["Location"],
        "Date - Week/Year": wl, "Date": ds,
        "# Active members": r["active"], "# Suspended members": r["suspended"],
        "# Cancelled members": r["cancelled"], "# New members": r["new"],
    } for _, r in mem_df.iterrows()]


def revenue_records(rev_df, week_end):
    wl, ds = week_label(week_end), sheet_date(week_end)
    return [{
        "Location": r["Location"], "Date - Week/Year": wl, "Date": ds,
        "# Active Students": r["active"], "Total Sessions": r["sessions"],
        "Gross Revenue": r["gross"], "Net Revenue": r["net"],
        "Student Visits": r["visits"],
        "Revenue per Session": r["rev_per_session"],
        "Revenue per Student": r["rev_per_student"],
        "Sessions per Student": r["sessions_per_student"],
        "Student per Session": r["students_per_session"],
        "Sessions per Student Visit": r["sessions_per_visit"],
        "Student Visits per Session": r["visits_per_session"],
    } for _, r in rev_df.iterrows()]


# ── Checks on historical data ─────────────────────────────────────────────────
def first_active_week(df_wm, name_col="Success Tutoring - Business name"):
    """location key -> first Date with # Active members > 0."""
    d = df_wm[pd.to_numeric(df_wm["# Active members"], errors="coerce").fillna(0) > 0]
    d = d.assign(_k=d[name_col].map(location_key))
    return d.groupby("_k")["Date"].min().to_dict()


def duplicate_weeks(df, name_col, week_col="Date - Week/Year"):
    """Location/week pairs that appear more than once."""
    g = df.groupby([name_col, week_col]).size().reset_index(name="Rows")
    return g[g["Rows"] > 1].sort_values([week_col, name_col])


def nz_gst_candidates(df_rv, tolerance=1.0, min_gross=20.0):
    """Revenue rows for NZ locations whose net was calculated as gross / 1.1.

    df_rv must hold raw sheet values plus a '_row' column with the sheet row
    number. Returns those rows with 'New Net' = gross / 1.15.
    """
    d = df_rv.copy()
    d["_gross"] = d["Gross Revenue"].map(parse_number)
    d["_net"] = d["Net Revenue"].map(parse_number)
    is_nz = d["Country"].astype(str).str.strip().str.lower() == "new zealand"
    wrong = (d["_gross"] >= min_gross) & ((d["_net"] - d["_gross"] / 1.10).abs() <= tolerance)
    right = (d["_net"] - d["_gross"] / 1.15).abs() <= tolerance
    d = d[is_nz & wrong & ~right].copy()
    d["New Net"] = (d["_gross"] / 1.15).round(2)
    return d
