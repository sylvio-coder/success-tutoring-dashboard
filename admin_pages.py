"""Admin pages: weekly CRM upload and the master location list."""
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import ingest
import sheet_writer as sw

WM_NAME = "Success Tutoring - Business name"
VLOOKUP_TAB = "Vlookup"
EDITABLE = ["Stage", "Status", "GPM", "Country", "Region"]
REVENUE_CALCULATED = [
    "Student Visits", "Revenue per Session", "Revenue per Student",
    "Sessions per Student", "Student per Session", "Sessions per Student Visit",
    "Student Visits per Session",
]
WRITE_HINT = ("If this is a permission error, share the Google Sheet with the "
              "dashboard's service account as an **Editor**.")


def _subtitle(text):
    st.markdown(f'<div class="report-subtitle">{text}</div>', unsafe_allow_html=True)


def _load_master(spreadsheet):
    ws = spreadsheet.worksheet(VLOOKUP_TAB)
    headers, rows = sw.read_tab(ws)
    df = pd.DataFrame([r + [""] * (len(headers) - len(r)) for r in rows], columns=headers)
    df["_row"] = range(2, len(df) + 2)
    df = df[df["Location"].astype(str).str.strip() != ""]
    return ws, headers, df


def _pending():
    return st.session_state.setdefault("pending_new_locations", {})


def _options(df, col, extra=()):
    vals = {str(v).strip() for v in df.get(col, pd.Series(dtype=str)).tolist()} | set(extra)
    return sorted(v for v in vals if v)


# ── Weekly upload ─────────────────────────────────────────────────────────────
def page_weekly_upload(spreadsheet, df_wm, df_rv, clear_cache):
    st.markdown('<div class="report-title">Weekly Upload</div>', unsafe_allow_html=True)
    _subtitle("Drop in this week's CRM exports as downloaded: Hapana members and "
              "revenue (AU and NZ), in-house members and revenue.")

    files = st.file_uploader("CRM exports", type=["xlsx", "csv"],
                             accept_multiple_files=True, key="upload_files")
    if not files:
        st.info("Upload the weekly export files to see a preview. Nothing is written "
                "to the Sheet until you confirm.")
        return

    exports, file_rows = [], []
    for f in files:
        try:
            e = ingest.read_export(f.name, f.getvalue())
            exports.append(e)
            file_rows.append({"File": f.name, "Detected as": ingest.KIND_LABELS[e["kind"]],
                              "Locations": len(e["rows"]),
                              "Week ending": f"{e['week_end']:%d/%m/%Y}" if e["week_end"] else "?"})
        except ingest.ExportError as err:
            file_rows.append({"File": f.name, "Detected as": "❌ " + str(err),
                              "Locations": 0, "Week ending": ""})
    st.dataframe(pd.DataFrame(file_rows), hide_index=True, use_container_width=True)
    if not exports:
        return

    try:
        _, _, master = _load_master(spreadsheet)
        aliases = sw.read_aliases(spreadsheet)
    except Exception as e:
        st.error(f"Could not read the master location list: {e}")
        return

    res = ingest.combine(exports, master, aliases)

    # Unknown names: map to an existing location, set up as new, or leave out.
    extra_aliases, left_out = {}, set()
    if res["unknown"]:
        st.markdown("#### Locations not in the master list")
        st.caption("Pick the existing location each name refers to, or set it up as a "
                   "new location. A mapping you choose is remembered for future uploads.")
        all_names = master["Location"].tolist()
        NEW, SKIP, CHOOSE = "➕ New location (set up on the Locations page)", \
            "Leave out this week", "— choose —"
        for raw, u in sorted(res["unknown"].items()):
            opts = [CHOOSE, NEW, SKIP] + u["suggestions"] + \
                [n for n in all_names if n not in u["suggestions"]]
            c1, c2 = st.columns([2, 3])
            c1.markdown(f"**{raw}**  \n<span style='font-size:0.8em'>{', '.join(u['sources'])}"
                        f" · {u['active']:.0f} active</span>", unsafe_allow_html=True)
            choice = c2.selectbox("Maps to", opts, key=f"map_{raw}", label_visibility="collapsed")
            if choice == NEW:
                _pending()[raw] = {"week_end": res["week_end"], "sources": u["sources"]}
            else:
                _pending().pop(raw, None)
            if choice == SKIP:
                left_out.add(raw)
            elif choice not in (CHOOSE, NEW):
                extra_aliases[raw] = choice
        if extra_aliases:
            res = ingest.combine(exports, master, {**aliases, **extra_aliases})

    unresolved = [n for n in res["unknown"] if n not in left_out]
    for msg in res["errors"]:
        st.error(msg)
    for msg in res["warnings"]:
        st.warning(msg)
    for msg in res["info"]:
        st.info(msg)
    if unresolved:
        new = [n for n in unresolved if n in _pending()]
        if new:
            st.warning("Set up these new locations on the **Locations** page, then come "
                       "back here: " + ", ".join(new))
        other = [n for n in unresolved if n not in _pending()]
        if other:
            st.warning("Choose what these names map to: " + ", ".join(other))

    # Week
    default_week = res["week_end"] or (date.today() - timedelta(days=(date.today().weekday() + 1) % 7))
    week_end = st.date_input("Week ending (Sunday)", value=default_week, format="DD/MM/YYYY")
    if week_end.weekday() != 6:
        st.warning("That date isn't a Sunday. The sheet stores each week by its Sunday.")
    wl = ingest.week_label(week_end)
    existing = int((df_wm["Date - Week/Year"].astype(str).str.strip() == wl).sum()) \
        if "Date - Week/Year" in df_wm.columns else 0
    if existing:
        st.info(f"The sheet already has {existing} membership rows for week {wl}. Rows for "
                "the locations in this upload will be replaced, not duplicated.")

    mem, rev = res["membership"], res["revenue"]
    if mem.empty and rev.empty:
        return
    _preview(mem, rev, master, df_wm, df_rv, week_end)

    blocked = bool(res["errors"] or unresolved)
    if st.button(f"✅ Confirm and write week {wl} to the Sheet", type="primary",
                 disabled=blocked):
        try:
            with st.spinner("Writing to the Sheet..."):
                if extra_aliases:
                    sw.add_aliases(spreadsheet, list(extra_aliases.items()))
                m_del, m_add = (0, 0)
                if not mem.empty:
                    m_del, m_add = sw.replace_week(
                        spreadsheet.worksheet("Weekly Membership"),
                        ingest.membership_records(mem, week_end), WM_NAME, wl)
                r_del, r_add = (0, 0)
                if not rev.empty:
                    r_del, r_add = sw.replace_week(
                        spreadsheet.worksheet("Revenue"),
                        ingest.revenue_records(rev, week_end), "Location", wl,
                        prefer_formula=REVENUE_CALCULATED)
            clear_cache()
            st.success(f"Week {wl} written: {m_add} membership rows and {r_add} revenue "
                       f"rows ({m_del + r_del} earlier rows for this week replaced).")
        except Exception as e:
            st.error(f"Writing to the Sheet failed: {e}\n\n{WRITE_HINT}")
    elif blocked:
        st.caption("Resolve the items above to enable writing.")


def _preview(mem, rev, master, df_wm, df_rv, week_end):
    country = dict(zip(master["Location"], master["Country"])) if "Country" in master.columns else {}
    st.markdown("#### Preview")

    prev_wm = df_wm[df_wm["Date"].dt.date < week_end]
    prev_date = prev_wm["Date"].max() if not prev_wm.empty else None
    cols = st.columns(2)

    with cols[0]:
        st.markdown("**Membership by country**")
        cur = mem.assign(Country=mem["Location"].map(country)).groupby("Country")[
            ["active", "new", "cancelled"]].sum()
        cur.columns = ["Active", "New", "Cancelled"]
        if prev_date is not None:
            p = df_wm[df_wm["Date"] == prev_date].groupby("Country")["# Active members"].sum()
            cur["Active last week"] = p.reindex(cur.index).fillna(0)
            cur["Change"] = cur["Active"] - cur["Active last week"]
        st.dataframe(cur.astype(int), use_container_width=True)

    with cols[1]:
        st.markdown("**Revenue by country**")
        if not rev.empty:
            cur = rev.assign(Country=rev["Location"].map(country)).groupby("Country")[
                ["gross", "net"]].sum()
            cur.columns = ["Gross", "Net"]
            if prev_date is not None and not df_rv.empty and "Net Revenue" in df_rv.columns:
                p = df_rv[df_rv["Date"] == prev_date].groupby("Country")["Net Revenue"].sum()
                cur["Net last week"] = p.reindex(cur.index).fillna(0)
            st.dataframe(cur.round(0), use_container_width=True)

    # Big week-on-week moves at single locations are worth a second look.
    if prev_date is not None:
        p = df_wm[df_wm["Date"] == prev_date].set_index(WM_NAME)["# Active members"]
        m = mem.set_index("Location")["active"]
        both = pd.DataFrame({"Last week": p.reindex(m.index), "This week": m}).dropna()
        both = both[(both["Last week"] >= 10) &
                    ((both["This week"] - both["Last week"]).abs() / both["Last week"] > 0.3)]
        if not both.empty:
            st.warning("Active members changed by more than 30% at: " +
                       ", ".join(f"{n} ({a:.0f} → {b:.0f})"
                                 for n, (a, b) in both.iterrows()))

    status = master["Status"] if "Status" in master.columns else pd.Series("", index=master.index)
    trading = master[status.astype(str).str.lower() == "trading"]["Location"]
    missing = sorted(set(trading) - set(mem["Location"]))
    if missing:
        st.warning("Trading locations with no membership row in this upload: " +
                   ", ".join(missing))

    with st.expander(f"Membership rows ({len(mem)})"):
        st.dataframe(mem, hide_index=True, use_container_width=True)
    with st.expander(f"Revenue rows ({len(rev)})"):
        st.dataframe(rev, hide_index=True, use_container_width=True)


# ── Locations (master list) ───────────────────────────────────────────────────
def page_locations(spreadsheet, df_wm, df_rv, clear_cache):
    st.markdown('<div class="report-title">Locations</div>', unsafe_allow_html=True)
    _subtitle("Master location list (Vlookup tab): set GPM, Stage, Status, Country and "
              "Region here. New locations found in uploads appear at the top.")
    try:
        vl_ws, vl_headers, master = _load_master(spreadsheet)
        aliases = sw.read_aliases(spreadsheet)
    except Exception as e:
        st.error(f"Could not read the Vlookup tab: {e}")
        return

    matcher = ingest.LocationMatcher(master["Location"].tolist(), aliases)
    first_week = ingest.first_active_week(df_wm) if not df_wm.empty else {}

    _new_locations(spreadsheet, vl_ws, master, matcher, first_week, df_wm, df_rv, clear_cache)
    _master_editor(vl_ws, vl_headers, master, df_wm, clear_cache)

    with st.expander(f"Name matching list ({len(aliases)})"):
        st.caption("Other spellings used in CRM exports, and the location each one means. "
                   f"Stored in the '{sw.ALIASES_TAB}' tab.")
        if aliases:
            st.dataframe(pd.DataFrame(sorted(aliases.items()), columns=["Name in export", "Location"]),
                         hide_index=True, use_container_width=True)

    _data_checks(spreadsheet, df_wm, df_rv, clear_cache)


def _new_locations(spreadsheet, vl_ws, master, matcher, first_week, df_wm, df_rv, clear_cache):
    pending = _pending()
    # Names already in the sheet data that the master list doesn't know.
    in_data = {}
    for df, col in ((df_wm, WM_NAME), (df_rv, WM_NAME)):
        if df.empty or col not in df.columns:
            continue
        for name, last in df.groupby(col)["Date"].max().items():
            if str(name).strip() and matcher.resolve(name) is None and not ingest.is_excluded(name):
                in_data[name] = max(last, in_data.get(name, last))

    todo = [n for n in pending if matcher.resolve(n) is None] + \
        [n for n in sorted(in_data) if n not in pending]
    for n in [n for n in pending if matcher.resolve(n) is not None]:
        pending.pop(n)
    if not todo:
        return

    st.markdown("#### Locations needing setup")
    for raw in todo:
        from_upload = raw in pending
        label = f"🆕 {raw}" + ("  ·  from this week's upload" if from_upload
                              else f"  ·  in the sheet data, last seen {in_data[raw]:%d/%m/%Y}")
        with st.expander(label, expanded=from_upload):
            if not from_upload:
                st.caption("This name appears in Weekly Membership/Revenue but not in the "
                           "master list, so those rows have no GPM, Stage or Region.")
            _new_location_form(spreadsheet, vl_ws, master, matcher, raw, first_week,
                               pending.get(raw, {}).get("week_end"), clear_cache)


def _new_location_form(spreadsheet, vl_ws, master, matcher, raw, first_week, upload_week,
                       clear_cache):
    k = ingest.location_key(raw)
    mode = st.radio("This is", ["A new location", "Another name for an existing location"],
                    key=f"mode_{raw}", horizontal=True)
    if mode.startswith("Another"):
        sugg = matcher.suggestions(raw)
        opts = sugg + [n for n in master["Location"] if n not in sugg]
        target = st.selectbox("Existing location", opts, key=f"alias_{raw}")
        if st.button("Save mapping", key=f"save_alias_{raw}"):
            try:
                sw.add_aliases(spreadsheet, [(raw, target)])
                clear_cache()
                st.success(f"'{raw}' will be read as {target} in future uploads.")
                st.rerun()
            except Exception as e:
                st.error(f"{e}\n\n{WRITE_HINT}")
        return

    start = first_week.get(k)
    start = start.date() if start is not None else (upload_week or date.today())
    with st.form(f"new_{raw}"):
        name = st.text_input("Location name", ingest.canonical_new_name(raw))
        c1, c2, c3 = st.columns(3)
        vals = {}
        for i, col in enumerate(EDITABLE):
            opts = _options(master, col)
            vals[col] = [c1, c2, c3][i % 3].selectbox(col, opts, key=f"{col}_{raw}",
                                                      index=None, placeholder="Choose…")
        other = st.text_input("Or type a new value, e.g. GPM=Jane Citizen",
                              key=f"other_{raw}", help="Use Field=Value for a value not in the lists.")
        started = st.date_input("Location start (first week with members)", start,
                                format="DD/MM/YYYY", key=f"start_{raw}")
        if st.form_submit_button("Add to master list", type="primary"):
            if "=" in other:
                f, v = [x.strip() for x in other.split("=", 1)]
                if f in vals and v:
                    vals[f] = v
            missing = [c for c in EDITABLE if not vals.get(c)]
            if missing:
                st.error("Please choose: " + ", ".join(missing))
                return
            if ingest.location_key(name) in matcher.master:
                st.error(f"{name} is already in the master list.")
                return
            rec = {"Location": name.strip(), **vals,
                   "Location Start": f"{started.day}-{started:%b-%Y}"}
            try:
                sw.append_records(vl_ws, [rec])
                if ingest.location_key(name) != k:
                    sw.add_aliases(spreadsheet, [(raw, name.strip())])
                _pending().pop(raw, None)
                clear_cache()
                st.success(f"{name} added to the master list.")
                st.rerun()
            except Exception as e:
                st.error(f"{e}\n\n{WRITE_HINT}")


def _master_editor(vl_ws, vl_headers, master, df_wm, clear_cache):
    st.markdown("#### Master location list")
    view = master.copy()
    if not df_wm.empty:
        latest = df_wm[df_wm["Date"] == df_wm["Date"].max()].set_index(WM_NAME)["# Active members"]
        last_seen = df_wm[df_wm["# Active members"] > 0].groupby(WM_NAME)["Date"].max()
        view["Active members"] = view["Location"].map(latest).fillna(0).astype(int)
        view["Last week with members"] = view["Location"].map(last_seen).dt.strftime("%d/%m/%Y")

    f1, f2, f3 = st.columns(3)
    gpm = f1.multiselect("GPM", _options(master, "GPM"), key="loc_gpm")
    stage = f2.multiselect("Stage", _options(master, "Stage"), key="loc_stage")
    search = f3.text_input("Search", key="loc_search")
    if gpm:
        view = view[view["GPM"].isin(gpm)]
    if stage:
        view = view[view["Stage"].isin(stage)]
    if search:
        view = view[view["Location"].str.contains(search, case=False, regex=False)]

    show = ["Location"] + [c for c in EDITABLE + ["Location Start"] if c in view.columns] + \
        [c for c in ["Months old", "Onboarding Week", "Active members",
                     "Last week with members"] if c in view.columns]
    cfg = {c: st.column_config.SelectboxColumn(c, options=_options(master, c))
           for c in EDITABLE if c in view.columns}
    cfg["Location"] = st.column_config.TextColumn("Location", disabled=True)
    for c in ["Months old", "Onboarding Week", "Active members", "Last week with members"]:
        cfg[c] = st.column_config.Column(c, disabled=True)
    st.caption("Edit cells directly, then save. To add a value that isn't in a list "
               "(e.g. a new GPM), use a location's setup form or type it in the Sheet once.")
    edited = st.data_editor(view[show], column_config=cfg, hide_index=True,
                            use_container_width=True, key="master_editor")

    changes = []
    orig = view[show]
    for idx in edited.index:
        for c in [c for c in EDITABLE + ["Location Start"] if c in show]:
            if str(edited.at[idx, c]) != str(orig.at[idx, c]):
                changes.append((int(master.at[idx, "_row"]), vl_headers.index(c) + 1,
                                edited.at[idx, c], master.at[idx, "Location"], c))
    if changes:
        st.info(f"{len(changes)} unsaved change(s): " +
                "; ".join(f"{loc} {c} → {v}" for _, _, v, loc, c in changes[:8]) +
                (" …" if len(changes) > 8 else ""))
        if st.button("💾 Save changes", type="primary"):
            try:
                sw.update_cells(vl_ws, [(r, c, v) for r, c, v, _, _ in changes])
                clear_cache()
                st.success("Saved.")
                st.rerun()
            except Exception as e:
                st.error(f"{e}\n\n{WRITE_HINT}")


def _data_checks(spreadsheet, df_wm, df_rv, clear_cache):
    st.markdown("#### Data checks")

    dups = []
    for label, df in (("Weekly Membership", df_wm), ("Revenue", df_rv)):
        if not df.empty and "Date - Week/Year" in df.columns:
            d = ingest.duplicate_weeks(df, WM_NAME)
            dups.extend({"Tab": label, "Location": r[WM_NAME], "Week": r["Date - Week/Year"],
                         "Rows": r["Rows"]} for _, r in d.iterrows())
    if dups:
        st.warning(f"{len(dups)} location/week pairs appear more than once. This usually "
                   "means two sites (e.g. Belmont and Belmont WA, or Epping and Epping VIC) "
                   "were pasted under one name. Correct the name on the wrong row in the Sheet.")
        st.dataframe(pd.DataFrame(dups), hide_index=True, use_container_width=True)
    else:
        st.success("No location appears twice in the same week.")

    _nz_gst_fix(spreadsheet, clear_cache)


def _nz_gst_fix(spreadsheet, clear_cache):
    ws = spreadsheet.worksheet("Revenue")
    headers, rows = sw.read_tab(ws)
    need = ["Location", "Gross Revenue", "Net Revenue", "Country"]
    if not all(h in headers for h in need):
        return
    df = pd.DataFrame([r + [""] * (len(headers) - len(r)) for r in rows], columns=headers)
    df["_row"] = range(2, len(df) + 2)
    cand = ingest.nz_gst_candidates(df)
    if cand.empty:
        st.success("No NZ revenue rows use the 10% GST rate.")
        return

    st.warning(f"{len(cand)} NZ revenue rows have net revenue calculated as gross ÷ 1.1. "
               "NZ GST is 15%, so net should be gross ÷ 1.15.")
    summary = cand.groupby("Location").agg(
        Rows=("_row", "size"), Current_net=("_net", "sum"), Corrected_net=("New Net", "sum"))
    summary["Difference"] = summary["Corrected_net"] - summary["Current_net"]
    st.dataframe(summary.round(0), use_container_width=True)
    locs = st.multiselect("Locations to correct", summary.index.tolist(),
                          default=summary.index.tolist(), key="gst_locs")
    if st.button(f"Correct NZ net revenue for {len(locs)} location(s)", disabled=not locs):
        todo = cand[cand["Location"].isin(locs)]
        net_col = headers.index("Net Revenue") + 1
        updates = [(int(r["_row"]), net_col, r["New Net"]) for _, r in todo.iterrows()]
        # Per-row values (not formulas) that depend on net need recalculating too.
        formulas = ws.get(value_render_option="FORMULA")
        for h, denom in (("Revenue per Session", "Total Sessions"),
                         ("Revenue per Student", "# Active Students")):
            if h not in headers or denom not in headers:
                continue
            c, dc = headers.index(h), headers.index(denom)
            for _, r in todo.iterrows():
                cell = formulas[r["_row"] - 1][c] if c < len(formulas[r["_row"] - 1]) else ""
                if not str(cell).startswith("="):
                    d = ingest.parse_number(r[headers[dc]])
                    updates.append((int(r["_row"]), c + 1, round(r["New Net"] / d, 2) if d else 0))
        try:
            sw.update_cells(ws, updates)
            clear_cache()
            st.success(f"Corrected {len(todo)} rows.")
            st.rerun()
        except Exception as e:
            st.error(f"{e}\n\n{WRITE_HINT}")
