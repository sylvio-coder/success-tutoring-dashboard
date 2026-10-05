"""Monthly P&L workbook: reading it, checking it, and rows for the Sheet's P&L tab.

The workbook has one column per centre in its own currency ("P&L (native
currency)"), a long "Line items" tab with the same amounts, and a "Summary" tab
with each centre's country, currency, status and rate to AUD. Amounts are stored
in each centre's own currency; AUD is only worked out for mixed totals.
"""
import io
import re
from datetime import date

import pandas as pd
from openpyxl import load_workbook

import ingest

PNL_TAB = "P&L"
PNL_HEADERS = ["Period", "Location", "Country", "Currency", "Status", "Account code",
               "Account", "P&L section", "Amount", "FX rate to AUD"]
CRM = "P&L"                     # alias scope: names matched for P&L uploads only

REVENUE_CODES = {200, 260, 270}
TUTOR_CODES = {472, 474}
EXCLUDED_CODES = {479}          # Unrelated Expenses: memo only, not in profit
INTEREST_CODES = {437}

# Rows of the summary / break-even layout and the accounts each one adds up.
SUMMARY_ROWS = [
    ("Total Revenue", REVENUE_CODES),
    ("Tutor Wages + Super", TUTOR_CODES),
    ("Rent", {469}),
    ("Wages - Manager/Admin + Super", {471, 478}),
    ("Royalties", {467}),
    ("Marketing - Agency Fees", {402}),
    ("Marketing", {400}),
    ("OPEX", None),             # every other operating cost
]

MONTHS = {m.lower(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July", "August",
     "September", "October", "November", "December"], start=1)}
LOW_REVENUE = 1000


class PnlError(Exception):
    pass


def period_label(d):
    return f"{d.year:04d}-{d.month:02d}"


def period_name(period):
    y, m = map(int, period.split("-"))
    return date(y, m, 1).strftime("%B %Y")


def _period_from_text(text):
    s = str(text or "")
    m = re.search(r"\b(20\d\d)-(0[1-9]|1[0-2])\b", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"\b(" + "|".join(MONTHS) + r")\s+(20\d\d)\b", s, re.I)
    if m:
        return f"{m.group(2)}-{MONTHS[m.group(1).lower()]:02d}"
    return None


def _code(value):
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


def section_of(code, given=""):
    if code in REVENUE_CODES:
        return "Revenue"
    if code in TUTOR_CODES:
        return "Cost of Sales"
    if code in INTEREST_CODES:
        return "Interest, Tax & Depreciation"
    if code in EXCLUDED_CODES:
        return "Memo (excluded from Net Profit)"
    return str(given).strip() or "Operating Costs"


def _summary(wb):
    """{centre: {country, currency, status, fx}} plus the period, from 'Summary'."""
    info, period = {}, None
    if "Summary" not in wb.sheetnames:
        return info, period
    rows = list(wb["Summary"].iter_rows(values_only=True))
    for r in rows[:3]:
        period = period or _period_from_text(r[0] if r else "")
    hdr = next((i for i, r in enumerate(rows)
                if r and [ingest.norm_header(c) for c in r[:2]] == ["centre", "country"]), None)
    if hdr is None:
        return info, period
    cols = [ingest.norm_header(c) for c in rows[hdr]]
    fx_i = next((i for i, c in enumerate(cols) if c.startswith("fx rate")), None)
    for r in rows[hdr + 1:]:
        name = str(r[0] or "").strip()
        if not name or name.lower().startswith("network total"):
            continue
        get = lambda h: str(r[cols.index(h)] or "").strip() if h in cols else ""
        fx = r[fx_i] if fx_i is not None and fx_i < len(r) else None
        info.setdefault(name, []).append({
            "country": get("country"), "currency": get("currency"),
            "status": get("status"),
            "fx": float(fx) if isinstance(fx, (int, float)) and fx else None})
    return info, period


def _read_native(ws):
    rows = list(ws.iter_rows(values_only=True))
    hdr = next((i for i, r in enumerate(rows)
                if r and [ingest.norm_header(c) for c in r[:2]] == ["code", "account"]), None)
    if hdr is None:
        return None, None
    period = next((p for r in rows[:hdr] for p in [_period_from_text(r[0] if r else "")] if p),
                  None)
    centres = []
    for j, h in enumerate(rows[hdr][2:], start=2):
        m = re.match(r"\s*(.+?)\s*\(([A-Z]{3})\)\s*$", str(h or ""))
        if m:
            centres.append((j, m.group(1), m.group(2)))
        elif str(h or "").strip():
            centres.append((j, str(h).strip(), ""))
    out, section = [], ""
    for r in rows[hdr + 1:]:
        code = _code(r[0])
        if code is None:
            label = str(r[1] or "").strip() if len(r) > 1 else ""
            if label and all(v in (None, "") for v in r[2:]):
                section = label
            continue
        for j, name, cur in centres:
            out.append({"Centre": name, "Currency": cur, "Account code": code,
                        "Account": str(r[1] or "").strip(),
                        "P&L section": section_of(code, section),
                        "Amount": ingest.parse_number(r[j] if j < len(r) else None)})
    return pd.DataFrame(out), period


def _read_line_items(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None, None
    cols = [ingest.norm_header(c) for c in rows[0]]
    need = ["centre", "account code", "account", "amount"]
    if not all(c in cols for c in need):
        return None, None
    ix = {c: cols.index(c) for c in cols}
    out, periods = [], set()
    for r in rows[1:]:
        code = _code(r[ix["account code"]])
        name = str(r[ix["centre"]] or "").strip()
        if code is None or not name:
            continue
        if "period" in ix and r[ix["period"]]:
            periods.add(_period_from_text(str(r[ix["period"]])))
        out.append({"Centre": name,
                    "Currency": str(r[ix["currency"]] or "").strip() if "currency" in ix else "",
                    "Account code": code, "Account": str(r[ix["account"]] or "").strip(),
                    "P&L section": section_of(code, r[ix["p&l section"]]
                                              if "p&l section" in ix else ""),
                    "Amount": ingest.parse_number(r[ix["amount"]])})
    periods.discard(None)
    if len(periods) > 1:
        raise PnlError("The Line items tab has more than one month: " + ", ".join(sorted(periods)))
    return pd.DataFrame(out), (periods.pop() if periods else None)


def read_pnl(filename, data):
    """Read a monthly P&L workbook.

    Returns {"period", "lines" (one row per centre and account, own currency),
    "centres" (one row per declared centre), "pending" (names not yet declared),
    "source" (tab read)}.
    """
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    except Exception as e:
        raise PnlError(f"Could not open {filename} as an Excel file ({e}).")
    info, period = _summary(wb)

    lines, source = None, None
    order = sorted(wb.sheetnames, key=lambda n: (0 if "native" in n.lower() else
                                                 1 if n.lower() == "line items" else 2))
    for name in order:
        if name == "Summary" or "aud" in name.lower() and "native" not in name.lower():
            continue
        reader = _read_line_items if name.lower() == "line items" else _read_native
        df, p = reader(wb[name])
        if df is not None and not df.empty:
            lines, source, period = df, name, p or period
            break
    if lines is None:
        raise PnlError("No P&L found. Expected a 'P&L (native currency)' tab with Code and "
                       "Account columns and one column per centre, or a 'Line items' tab.")
    period = period or _period_from_text(filename)

    dupes = sorted(lines.groupby(["Centre", "Account code"]).size().loc[lambda s: s > 1]
                   .index.get_level_values(0).unique())
    if dupes:
        raise PnlError("These centres appear more than once in the file: " + ", ".join(dupes))

    centres = []
    for name, g in lines.groupby("Centre", sort=True):
        meta = (info.get(name) or [{}])[0]
        centres.append({"Centre": name,
                        "Country": meta.get("country", ""),
                        "Currency": meta.get("currency") or g["Currency"].iloc[0],
                        "Status": meta.get("status") or "Submitted",
                        "FX rate to AUD": 1.0 if (meta.get("currency") or g["Currency"].iloc[0])
                        == "AUD" else meta.get("fx")})
    centres = pd.DataFrame(centres)
    totals = summarise(lines)
    centres = centres.merge(totals, on="Centre", how="left")
    declared = set(centres["Centre"])
    pending = sorted({n for n, metas in info.items()
                      if n not in declared and any(m["status"].lower() == "pending" for m in metas)})
    return {"period": period, "lines": lines, "centres": centres, "pending": pending,
            "source": source}


def summarise(lines):
    """Revenue, expenses, EBITDA and net profit per centre (own currency)."""
    df = lines[~lines["Account code"].isin(EXCLUDED_CODES)]
    rev = df[df["Account code"].isin(REVENUE_CODES)].groupby("Centre")["Amount"].sum()
    opex = df[~df["Account code"].isin(REVENUE_CODES | INTEREST_CODES)] \
        .groupby("Centre")["Amount"].sum()
    interest = df[df["Account code"].isin(INTEREST_CODES)].groupby("Centre")["Amount"].sum()
    out = pd.DataFrame({"Revenue": rev, "Expenses": opex, "Interest": interest}).fillna(0.0)
    out["EBITDA"] = out["Revenue"] - out["Expenses"]
    out["Net Profit"] = out["EBITDA"] - out["Interest"]
    return out.drop(columns="Interest").rename_axis("Centre").reset_index()


def _blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def match_centres(centres, master, aliases):
    """Add a 'Location' column (master name or None) and 'suggestions'."""
    matcher = ingest.LocationMatcher(master["Location"].tolist(), aliases)
    out = centres.copy()
    out["Location"] = [matcher.resolve(n, crm=CRM) for n in out["Centre"]]
    out["suggestions"] = [matcher.suggestions(n) if _blank(loc) else []
                          for n, loc in zip(out["Centre"], out["Location"])]
    return out


def checks(lines, centres, master):
    """Things to look at before trusting the month. List of {Location, Check, Detail}."""
    out = []
    master_country = {r["Location"]: str(r.get("Country", "")).strip()
                      for _, r in master.iterrows()}
    amounts = lines.pivot_table(index="Centre", columns="Account code", values="Amount",
                                aggfunc="sum", fill_value=0.0)
    for _, c in centres.iterrows():
        name = c["Centre"]
        loc = c["Centre"] if _blank(c.get("Location")) else c["Location"]
        a = amounts.loc[name] if name in amounts.index else pd.Series(dtype=float)
        get = lambda codes: float(sum(a.get(k, 0.0) for k in codes))
        rev = get(REVENUE_CODES)
        if rev <= 0:
            out.append((loc, "No revenue", "Total revenue is 0 or less"))
        elif rev < LOW_REVENUE:
            out.append((loc, "Very low revenue", f"Total revenue {rev:,.0f}"))
        if rev > 0 and get(TUTOR_CODES) == 0:
            out.append((loc, "No tutor wages", "Wages - Tutor and super are 0"))
        if rev > 0 and get({467}) == 0:
            out.append((loc, "No royalties", "Royalties are 0"))
        neg = lines[(lines["Centre"] == name) & (lines["Amount"] < 0) &
                    ~lines["Account code"].isin(REVENUE_CODES)]
        if not neg.empty:
            out.append((loc, "Negative costs (credits)", ", ".join(
                f"{r['Account']} {r['Amount']:,.0f}" for _, r in neg.iterrows())))
        if c["Currency"] not in ("", "AUD") and _blank(c.get("FX rate to AUD")):
            out.append((loc, "No exchange rate",
                        f"{c['Currency']} with no rate to AUD: left out of AUD totals"))
        mc = master_country.get(loc)
        if mc and not _blank(c.get("Country")) and mc.lower() != str(c["Country"]).lower():
            out.append((loc, "Country differs",
                        f"File says {c['Country']}, location list says {mc}"))
    return pd.DataFrame(out, columns=["Location", "Check", "Detail"])


def records(lines, centres, period):
    """Rows for the Sheet's P&L tab (matched centres only)."""
    meta = centres.set_index("Centre")
    out = []
    for _, r in lines.iterrows():
        c = meta.loc[r["Centre"]]
        if _blank(c.get("Location")):
            continue
        fx = c.get("FX rate to AUD")
        out.append({"Period": period, "Location": c["Location"], "Country": c["Country"],
                    "Currency": c["Currency"], "Status": c["Status"],
                    "Account code": int(r["Account code"]), "Account": r["Account"],
                    "P&L section": r["P&L section"], "Amount": float(r["Amount"]),
                    "FX rate to AUD": "" if _blank(fx) else float(fx)})
    return out


# ── Summary and break-even ────────────────────────────────────────────────────
WEEKS_PER_MONTH = 4.2
ROYALTY_RATE = 0.08
DEFAULT_ROYALTY_MIN = 2200.0
FIXED_ROWS = ["Rent", "Wages - Manager/Admin + Super", "Marketing - Agency Fees", "Marketing",
              "OPEX"]
ROW_ORDER = [r for r, _ in SUMMARY_ROWS]


def row_of(code):
    """Summary row an account adds into (None = left out of profit)."""
    if code in EXCLUDED_CODES:
        return None
    if code in INTEREST_CODES:
        return "Interest"
    for name, codes in SUMMARY_ROWS:
        if codes and code in codes:
            return name
    return "OPEX"


def location_months(lines):
    """One row per (Location, Period) with the summary rows, EBITDA, Interest, Net Profit.
    lines: the Sheet's P&L tab (Location, Period, Account code, Amount)."""
    df = lines.assign(Row=lines["Account code"].map(row_of),
                      Amount=lines["Amount"].astype(float)).dropna(subset=["Row"])
    out = df.pivot_table(index=["Location", "Period"], columns="Row", values="Amount",
                         aggfunc="sum", fill_value=0.0)
    for c in ROW_ORDER + ["Interest"]:
        if c not in out.columns:
            out[c] = 0.0
    costs = [r for r in ROW_ORDER if r != "Total Revenue"]
    out["EBITDA"] = out["Total Revenue"] - out[costs].sum(axis=1)
    out["Net Profit"] = out["EBITDA"] - out["Interest"]
    return out[ROW_ORDER + ["EBITDA", "Interest", "Net Profit"]]


def royalty(revenue, royalty_min, rate=ROYALTY_RATE):
    return max(royalty_min, rate * revenue)


def break_even(revenue, tutor, fixed, royalty_min, members=0, rate=ROYALTY_RATE):
    """Monthly revenue and members where EBITDA is 0.

    Tutor wages move with revenue (at this centre's tutor % of revenue); rent,
    manager wages, marketing and OPEX stay fixed; royalties are the higher of the
    minimum or 8% of revenue. Returns (revenue, members); None where it can't be
    worked out (no revenue, or tutor costs leave nothing to cover the rest).
    """
    if revenue <= 0:
        return None, None
    margin = 1 - tutor / revenue
    r = None
    if margin > 0 and rate * (fixed + royalty_min) / margin <= royalty_min:
        r = (fixed + royalty_min) / margin
    elif margin - rate > 0:
        r = fixed / (margin - rate)
    if r is None:
        return None, None
    per_member = revenue / members if members > 0 else 0
    return r, (r / per_member if per_member > 0 else None)


def model_column(members, per_member, tutor_pct, fixed_lines, royalty_min):
    """Summary rows for a centre with this many members (all monthly)."""
    rev = members * per_member
    col = {"Total Revenue": rev, "Tutor Wages + Super": tutor_pct * rev,
           "Royalties": royalty(rev, royalty_min)}
    col.update({k: fixed_lines.get(k, 0.0) for k in FIXED_ROWS})
    col["EBITDA"] = rev - sum(v for k, v in col.items() if k != "Total Revenue")
    return col


def membership_band(members, step=20, top=200):
    if members is None or pd.isna(members):
        return "No members data"
    if members >= top:
        return f"{top}+"
    lo = int(members // step) * step
    return f"{lo}–{lo + step - 1}"


def membership_bands(step=20, top=200):
    return [f"{lo}–{lo + step - 1}" for lo in range(0, top, step)] + [f"{top}+"]


# ── Tutor % of revenue by members band ────────────────────────────────────────
MIN_BENCH_CENTRES = 3


def tutor_benchmarks(cent, min_n=MIN_BENCH_CENTRES):
    """Typical (median) tutor wages + super as % of revenue for each members band.

    cent: one row per centre with Members, Total Revenue and Tutor Wages + Super.
    Centres with no revenue, no members or no tutor wages are left out. A band with
    fewer than min_n centres borrows the nearest band that has enough (else the
    median of all centres). Returns a DataFrame indexed by band:
    Centres, Benchmark % (its own median, if any), Use %, Source.
    """
    bands = membership_bands()
    ok = cent[(cent["Total Revenue"] > 0) & (cent["Tutor Wages + Super"] > 0) &
              cent["Members"].notna() & (cent["Members"] > 0)]
    pct = ok["Tutor Wages + Super"] / ok["Total Revenue"] * 100
    band = ok["Members"].map(membership_band)
    out = pd.DataFrame(index=bands)
    out["Centres"] = band.value_counts().reindex(bands).fillna(0).astype(int)
    out["Benchmark %"] = pct.groupby(band).median().reindex(bands)
    enough = [i for i, b in enumerate(bands) if out.at[b, "Centres"] >= min_n]
    overall = float(pct.median()) if len(pct) else 35.0
    use, source = [], []
    for i, b in enumerate(bands):
        if i in enough:
            use.append(out.at[b, "Benchmark %"])
            source.append("This band")
        elif enough:
            j = min(enough, key=lambda k: (abs(k - i), k))
            use.append(out.at[bands[j], "Benchmark %"])
            source.append(f"Nearest band ({bands[j]})")
        else:
            use.append(overall)
            source.append("All centres" if len(pct) else "Default")
    # Tutor costs don't rise as a share of revenue as centres grow: cap each band at
    # the band below it.
    for i in range(1, len(use)):
        if use[i] > use[i - 1]:
            use[i] = use[i - 1]
            source[i] = f"Capped at {bands[i - 1]}"
    out["Use %"] = [round(float(u), 1) for u in use]
    out["Source"] = source
    return out


def _band_mid(band, step=20, top=200):
    return top + step / 2 if band.endswith("+") else int(band.split("–")[0]) + step / 2


def rate_for(members, rates):
    """Tutor % (0–1) for this many members: each band's % applies at the middle of the
    band (10, 30, 50 …) and blends smoothly in between, so there are no jumps at band
    edges. rates: {band: percent}."""
    pts = sorted((_band_mid(b), r) for b, r in rates.items())
    if not pts:
        return 0.0
    if members <= pts[0][0]:
        return pts[0][1] / 100
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        if members <= x1:
            return (y0 + (y1 - y0) * (members - x0) / (x1 - x0)) / 100
    return pts[-1][1] / 100


def ebitda_at(members, per_member, rates, fixed, royalty_min):
    rev = members * per_member
    return rev * (1 - rate_for(members, rates)) - fixed - royalty(rev, royalty_min)


def break_even_banded(per_member, fixed, royalty_min, rates, centres=1, max_members=600):
    """Members (and monthly revenue) from which EBITDA stays at 0 or above, with tutor %
    from the members band. A group is treated as `centres` average centres sharing its
    costs. Returns (revenue, members), or (None, None) if not reached."""
    if per_member <= 0 or centres <= 0:
        return None, None
    f, rm = fixed / centres, royalty_min / centres
    step = 0.25
    n = int(max_members / step)
    last_neg = None
    for i in range(n + 1):
        if ebitda_at(i * step, per_member, rates, f, rm) < 0:
            last_neg = i
    if last_neg is None:
        return 0.0, 0.0
    if last_neg == n:
        return None, None
    lo, hi = last_neg * step, (last_neg + 1) * step
    for _ in range(30):
        mid = (lo + hi) / 2
        if ebitda_at(mid, per_member, rates, f, rm) >= 0:
            hi = mid
        else:
            lo = mid
    return hi * per_member * centres, hi * centres


# ── Benchmarks and outliers ───────────────────────────────────────────────────
BENCH_LINES = ROW_ORDER[1:] + ["EBITDA"]
HIGHER_IS_BETTER = {"EBITDA", "Total Revenue", "Average Membership Value"}
MANAGER = "Wages - Manager/Admin + Super"
MIN_PEERS = 5                   # ranges from fewer centres jump around too much


def peer_set(bench, country=None, members=None, exclude=None, min_n=MIN_PEERS):
    """Centres to compare with: same country and members band, widened to nearby bands
    until there are min_n, then the whole country, then every centre. Centres with very
    low revenue are left out (their percentages aren't meaningful).
    Returns (peers DataFrame, description)."""
    pool = bench[bench["Total Revenue"] >= LOW_REVENUE]
    if exclude is not None:
        pool = pool[~pool.index.isin([exclude] if isinstance(exclude, str) else exclude)]
    where = pool
    if country:
        where = pool[pool["Country"].astype(str) == str(country)]
    bands = membership_bands()
    if members is not None and not pd.isna(members) and members > 0:
        i = bands.index(membership_band(members))
        idx = where["Members"].map(lambda m: bands.index(membership_band(m))
                                   if m is not None and not pd.isna(m) and m > 0 else None)
        for w in range(len(bands)):
            sel = where[idx.notna() & ((idx - i).abs() <= w)]
            if len(sel) >= min_n:
                lo, hi = bands[max(i - w, 0)], bands[min(i + w, len(bands) - 1)]
                span = lo if lo == hi else f"{lo.split('–')[0]}–{hi.split('–')[-1].rstrip('+')}" \
                    + ("+" if hi.endswith("+") else "")
                return sel, f"{country + ' c' if country else 'C'}entres with {span} members"
            if w >= 2:
                break
    if country and len(where) >= min_n:
        return where, f"all {country} centres"
    return pool, "all reporting centres"


def _status(v, q1, q3):
    iqr = q3 - q1
    if v > q3 + 1.5 * iqr and v > q3:
        return "Very high"
    if v > q3:
        return "High"
    if v < q1 - 1.5 * iqr and v < q1:
        return "Very low"
    if v < q1:
        return "Low"
    return "Typical"


def compare_lines(values, peers, lines):
    """Each line as % of revenue against the peers' typical (median) and normal range
    (middle half). Gap $ = what the difference from typical is worth each month:
    positive means it costs more (or earns less) than typical."""
    rev = values["Total Revenue"]
    rows = []
    pp = peers[peers["Total Revenue"] > 0]
    for line in lines:
        if line not in values or line not in pp:
            continue
        pct = pp[line] / pp["Total Revenue"] * 100
        actual = values[line] / rev * 100 if rev > 0 else None
        med, q1, q3 = pct.median(), pct.quantile(0.25), pct.quantile(0.75)
        good_up = line in HIGHER_IS_BETTER
        gap = None if actual is None else (med - actual if good_up else actual - med) / 100 * rev
        rows.append({"Line": line, "Actual $": values[line], "Actual %": actual,
                     "Typical %": med, "Range low %": q1, "Range high %": q3,
                     "Centres": len(pct), "Gap $": gap,
                     "Status": _status(actual, q1, q3) if actual is not None else ""})
    # Manager wages: also against only the centres that pay a manager.
    if MANAGER in lines and values.get(MANAGER, 0) > 0 and rev > 0 and MANAGER in pp:
        payers = pp[pp[MANAGER] > 0]
        if len(payers):
            pct = payers[MANAGER] / payers["Total Revenue"] * 100
            actual = values[MANAGER] / rev * 100
            med = pct.median()
            rows.append({"Line": "↳ vs centres that pay a manager", "Actual $": values[MANAGER],
                         "Actual %": actual, "Typical %": med,
                         "Range low %": pct.quantile(0.25), "Range high %": pct.quantile(0.75),
                         "Centres": len(pct), "Gap $": (actual - med) / 100 * rev,
                         "Status": _status(actual, pct.quantile(0.25), pct.quantile(0.75))})
    # Revenue side: average membership value against the peers'.
    mem = values.get("Members")
    if mem and not pd.isna(mem) and mem > 0 and "Members" in pp:
        ok = pp[pp["Members"].fillna(0) > 0]
        # Dollar amounts only compare within one currency (country).
        country = values.get("Country")
        if isinstance(country, str) and "Country" in ok:
            ok = ok[ok["Country"].astype(str) == country]
        if len(ok):
            amv = ok["Total Revenue"] / ok["Members"] / WEEKS_PER_MONTH
            a = rev / mem / WEEKS_PER_MONTH
            rows.append({"Line": "Average Membership Value", "Actual $": a, "Actual %": None,
                         "Typical %": None, "Typical $": amv.median(),
                         "Range low $": amv.quantile(0.25), "Range high $": amv.quantile(0.75),
                         "Centres": len(amv),
                         "Gap $": (amv.median() - a) * mem * WEEKS_PER_MONTH,
                         "Status": _status(a, amv.quantile(0.25), amv.quantile(0.75))})
    return pd.DataFrame(rows)


def is_bad(line, status):
    if status in ("", "Typical"):
        return False
    up = status.lower().endswith("high")
    return (not up) if line in HIGHER_IS_BETTER else up


def account_table(lines, periods):
    """Monthly average per location and account (expense accounts only)."""
    d = lines[lines["Period"].isin(periods) &
              ~lines["Account code"].isin(REVENUE_CODES | EXCLUDED_CODES)]
    d = d.assign(Amount=d["Amount"].astype(float))
    if d.empty:
        return pd.DataFrame()
    per = d.pivot_table(index=["Location", "Period"], columns="Account", values="Amount",
                        aggfunc="sum", fill_value=0.0)
    return per.groupby(level="Location").mean()
