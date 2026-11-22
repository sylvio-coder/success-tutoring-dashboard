"""Franchise Model engine: a new centre's first 5 years, week by week, from the uploaded
assumptions (franchise_inputs.py) for one model (e.g. 'NSW') and one case (Worst / Base /
Best). Calculations only; the report page is in pnl_reports.py.

Week 1 = the presale campaign starts. Members follow the week-12 and year-end targets (each
year's growth spread over the months by the sign-up seasonality); new members needed = growth
plus cancellations. Revenue = tuition (active members x price banked) + suspension fees +
joining fees + other revenue. Tutor wages come from study hubs. Owner's wage starts the month
after the first profitable month. Amounts are in the model's own currency, excluding GST.
"""
import math
from datetime import date, timedelta

import pandas as pd

import franchise_inputs as fin

WEEKS = 260
WPM = 52 / 12                                   # weeks in a month
CASES = fin.CASES


FLOWS = ["New members", "Cancelled", "Tuition", "Suspension fees", "Joining fees",
         "Other revenue", "Revenue", "Tutor wages", "Manager", "Admin", "Rent", "Running costs",
         "CRM", "Marketing", "Royalties", "Royalties due", "EBITDA", "Owner's wage", "Study hubs",
         "Hubs after 7 pm", "Enquiries needed", "Advertising needed", "Advertising planned"]


# ── Settings for one case ─────────────────────────────────────────────────────
def settings(inputs, case):
    """Plain settings for one case from model_inputs() (variable -> case values)."""
    g = lambda name: inputs[name][case]
    num = lambda name: float(g(name))
    yes = lambda name: str(g(name)).strip().lower() == "yes"
    mgr = g("Hire a manager at")
    mgr_at = 0.0 if str(mgr).strip().lower() == "when sessions start" else \
        None if str(mgr).strip().lower() == "never" else float(mgr)
    seas = fin.parse_numbers(g("Sign-up seasonality (1.0 = average month)"))
    tiers = fin.parse_numbers(g("Foundation member numbers (tiers)"))
    tier_prices = fin.parse_numbers(g("Foundation prices (tiers)"))
    running = ["Other software and subscriptions", "Office expenses", "Insurance",
               "Energy and water", "Internet and phone", "Accounting, bookkeeping and consulting",
               "Cleaning", "Repairs and maintenance", "Workers compensation", "Bank fees",
               "Staff expenses and uniforms"]
    s = dict(
        investment_items={k: num(k) for k in [
            "Franchise grant", "Success Pack", "Reservation deposit", "Fit-out and signage",
            "Legal and accounting (company set-up)", "Legal costs for the lease",
            "Presale marketing budget (weeks 1–12)"]},
        bond_months=num("Rental bond"), wc_buffer=num("Working capital buffer") / 100,
        loan_share=num("Loan share of funding") / 100, loan_rate=num("Loan interest rate") / 100,
        loan_years=num("Loan term"), payback_target=num("Payback target"),
        discount=num("Required return (discount rate)") / 100,
        sale_multiple=num("Business sale value at Year 5"), tax=num("Company tax rate") / 100,
        depreciation_years=num("Fit-out depreciation"), owner_on=yes("Owner's wage included"),
        owner_wage=num("Owner's wage"), lease_before=num("Lease starts before week 1"),
        first_signups=num("First sign-ups"), open_week=num("Centre opens (sessions and billing start)"),
        presale_weeks=num("Presale campaign ends"), t12=num("Members at week 12"),
        targets=[num(f"Members at end of Year {y}") for y in range(1, 6)],
        cancel=[num(f"Year {y}: cancelled") / 100 for y in range(1, 6)],
        seasonality=seas, suspend=fin.parse_month_values(g("Extra members suspended in holidays")),
        tables=num("Study hub tables in the centre"),
        standard=num("Standard price (members after the first 50)"),
        tiers=tiers, tier_prices=tier_prices,
        founder_leave=num("Foundation members leaving") / 100,
        collected=num("Collected after GST (allowance for uncollected)") / 100,
        gst=num("GST") / 100, price_rise=num("Price rise each July (from 12 months)") / 100,
        joining_fee=num("Joining fee"), suspension_fee=num("Suspension fee"),
        other_revenue=num("Other revenue (holiday programs, assessments)"),
        visits=num("Student visits per member"), one_share=num("1:1 share of visits") / 100,
        fill=[num("Students per study hub: Year 1, Q1 (weeks 1–13)"),
              num("Students per study hub: Year 1, Q2"), num("Students per study hub: Year 1, Q3"),
              num("Students per study hub: Year 1, Q4")] +
        [num(f"Students per study hub: Year {y}") for y in range(2, 6)],
        max_fill=num("Most students per study hub"), hub_hours=num("Study hub length"),
        tutors_per_hub=num("Tutors per study hub"), weekdays=num("Weekdays open"),
        saturday=yes("Saturday open"), early_slots=num("Peak time slots: weekday before 7 pm"),
        late_slots=num("Weekday slots after 7 pm (busier centres)"),
        sat_slots=num("Peak time slots: Saturday"),
        sat_share=num("Share of study hubs on Saturday") / 100,
        min_shift=num("Minimum paid shift"),
        rate=num("Tutor hourly rate: weekday (casual, level 2, age 20)"),
        rate_penalty=num("Tutor hourly rate: Saturday and after 7 pm (casual)"),
        super_=num("Super") / 100, manager_at=mgr_at,
        manager_salary=num("Manager salary (ex super)"), admin=num("Front desk / admin"),
        wage_rise=num("Wage rise each July") / 100, rent=num("Rent (including outgoings)"),
        rent_free=num("Rent-free months"), rent_rise=num("Rent rise each July") / 100,
        agency=num("Agency fee (at least)"), ads_day=num("Advertising"),
        levy=num("Marketing fund levy") / 100,
        enquiry_to_assessment=num("Enquiry → Student Assessment") / 100,
        assessment_to_member=num("Student Assessment → member") / 100,
        lead_presale=num("Cost per lead: presale (weeks 1–12)"),
        lead_after=num("Cost per lead: after presale"),
        royalty=num("Royalty") / 100, royalty_min=num("Royalty minimum (FY2026–27)"),
        royalty_rise=num("Royalty minimum rise each July") / 100,
        crm=num("CRM software (Hapana)"),
        materials=num("Learning materials, printing and stationery"),
        running=sum(num(k) for k in running),
        cost_rise=num("Running costs rise each July") / 100,
        replacement=num("Replacement spending (computers, furniture)"),
        renewal_fee=num("Renewal fee"),
    )
    s["investment"] = sum(s["investment_items"].values())
    return s


# ── Helpers ───────────────────────────────────────────────────────────────────
def week_date(start, w):
    """Calendar date of week w (week 1 starts on the 1st of the opening month)."""
    return start + timedelta(days=7 * (w - 1))


def julys_between(a, b):
    """Number of 1 Julys after date a, up to and including date b."""
    n = 0
    for y in range(a.year, b.year + 1):
        j = date(y, 7, 1)
        if a < j <= b:
            n += 1
    return n


def fy_start(d):
    return d.year if d.month >= 7 else d.year - 1


def model_month(w):
    return int((w - 1) // WPM) + 1


def member_path(s, start):
    """Members on the books at the end of each week: from the first sign-ups to the week-12
    target in a straight line, then each year's growth spread over its weeks by the sign-up
    seasonality (more in February, less in December)."""
    path, prev_target = [0.0] * (WEEKS + 1), s["t12"]
    first = max(int(s["first_signups"]), 1)
    for w in range(1, 13):
        path[w] = s["t12"] * max(w - first + 1, 0) / (12 - first + 1) if w >= first else 0.0
    for y in range(5):
        a, b = 13 if y == 0 else 52 * y + 1, 52 * (y + 1)
        weights = [s["seasonality"][week_date(start, w).month - 1] for w in range(a, b + 1)]
        total, done = sum(weights), 0.0
        for i, w in enumerate(range(a, b + 1)):
            done += weights[i]
            path[w] = prev_target + (s["targets"][y] - prev_target) * done / total
        prev_target = s["targets"][y]
    return path


def students_per_hub(s, w):
    if w <= 52:
        return min(s["fill"][min((w - 1) // 13, 3)], s["max_fill"])
    return min(s["fill"][min(3 + (w - 1) // 52, 7)], s["max_fill"])


def capacity(s):
    """Most members the centre can hold: hub slots a week x most students ÷ visits."""
    slots = s["tables"] * (s["weekdays"] * (s["early_slots"] + s["late_slots"]) +
                           (s["sat_slots"] if s["saturday"] else 0))
    return slots * s["max_fill"] / s["visits"] / (1 - s["one_share"]) if s["visits"] else 0.0


def tutor_week(s, active, w, wage_k):
    """(tutor wages this week, study hubs, hubs after 7 pm, peak tables)."""
    visits = active * s["visits"]
    group, one = visits * (1 - s["one_share"]), visits * s["one_share"]
    fill = students_per_hub(s, w)
    hubs = group / fill if fill else 0.0
    sat_share = s["sat_share"] if s["saturday"] else 0.0
    sat = min(hubs * sat_share, s["tables"] * s["sat_slots"]) if s["saturday"] else 0.0
    days = max(s["weekdays"], 1)
    wd = (hubs - sat) / days                                    # hubs a weekday
    early = min(wd, s["tables"] * s["early_slots"])
    late = max(wd - early, 0.0)
    one_wd = one * (1 - sat_share) / days
    one_sat = one * sat_share
    t_early = math.ceil(early / s["early_slots"] - 1e-9) if early > 0 else 0
    t_sat = math.ceil(sat / s["sat_slots"] - 1e-9) if sat > 0 else 0
    early_paid = max(early * s["tutors_per_hub"] * s["hub_hours"] + one_wd * s["hub_hours"],
                     t_early * s["min_shift"])
    late_paid = late * s["tutors_per_hub"] * s["hub_hours"]
    sat_paid = max(sat * s["tutors_per_hub"] * s["hub_hours"] + one_sat * s["hub_hours"],
                   t_sat * s["min_shift"]) if s["saturday"] else 0.0
    cost = (days * (early_paid * s["rate"] + late_paid * s["rate_penalty"]) +
            sat_paid * s["rate_penalty"]) * (1 + s["super_"]) * wage_k
    return cost, hubs, late * days, max(t_early, t_sat)


# ── The model ─────────────────────────────────────────────────────────────────
def run(s, start, overrides=None):
    """Week by week for 5 years, then grouped into 60 months. start: date of week 1 (the 1st
    of the opening month). Returns a DataFrame indexed by model month (1–60)."""
    s = {**s, **(overrides or {})}
    path = member_path(s, start)
    open_w = int(s["open_week"])
    royalty_w = max(9, open_w)                                 # later of week 8 or opening
    open_date = week_date(start, open_w)
    lease_start = start - timedelta(days=30.4 * s["lease_before"])
    rent_from = lease_start + timedelta(days=30.4 * s["rent_free"])
    founders = [0.0] * len(s["tiers"])
    joined, owner_on, hired = 0.0, False, False
    months, cur = [], None
    for w in range(1, WEEKS + 1):
        d = week_date(start, w)
        m = model_month(w)
        books, prev = path[w], path[w - 1]
        y = min((w - 1) // 52, 4)
        canc = prev * s["cancel"][y] / WPM if w > 12 else 0.0
        new = max(books - prev + canc, 0.0)
        # foundation members fill the tiers in order and leave at their own rate
        out, lo = [], 0.0
        for r, size in zip(founders, s["tiers"]):
            stay = 1 - (s["founder_leave"] / WPM if w > 12 else 0.0)
            out.append(r * stay + max(0.0, min(joined + new, lo + size) - max(joined, lo)))
            lo += size
        founders, joined = out, joined + new
        f = min(sum(founders), books)
        scale = f / sum(founders) if sum(founders) else 0.0
        listed = ((sum(r * scale * p for r, p in zip(founders, s["tier_prices"])) +
                   (books - f) * s["standard"]) / books) if books else s["standard"]
        rises = julys_between(open_date + timedelta(days=364), d) + \
            (1 if d >= open_date + timedelta(days=364) and
             (open_date + timedelta(days=364)).month == 7 else 0) if d > open_date else 0
        price = listed / (1 + s["gst"]) * s["collected"] * (1 + s["price_rise"]) ** rises
        sus_pct = s["suspend"].get(f"{d.month:02d}", 0.0) / 100
        billing = w >= open_w
        active = books * (1 - sus_pct)
        tuition = active * price if billing else 0.0
        sus_fees = books * sus_pct * s["suspension_fee"] if billing else 0.0
        joining = new * s["joining_fee"] / (1 + s["gst"])
        other = s["other_revenue"] / WPM if billing else 0.0
        revenue = tuition + sus_fees + joining + other
        wage_k = (1 + s["wage_rise"]) ** julys_between(start, d)
        cost_k = (1 + s["cost_rise"]) ** julys_between(start, d)
        tutors, hubs, late_hubs, tables = tutor_week(s, active, w, wage_k) if billing else \
            (0.0, 0.0, 0.0, 0)
        if billing and s["manager_at"] is not None and books >= s["manager_at"]:
            hired = True
        manager = s["manager_salary"] * (1 + s["super_"]) / 52 * wage_k if hired else 0.0
        admin = s["admin"] / WPM * wage_k if billing else 0.0
        rent = s["rent"] / WPM * (1 + s["rent_rise"]) ** julys_between(lease_start, d) \
            if d >= rent_from else 0.0
        running = (s["running"] + s["materials"] * books) / WPM * cost_k
        crm = s["crm"] / WPM * cost_k
        after_presale = w > s["presale_weeks"]
        marketing = ((s["agency"] / WPM + s["ads_day"] * 7) if after_presale else 0.0) + \
            s["levy"] * revenue                      # presale campaign is in the investment
        rmin = s["royalty_min"] / WPM * (1 + s["royalty_rise"]) ** (fy_start(d) - 2026)
        royalty_due = max(s["royalty"] * revenue, rmin)
        royalties = royalty_due if w >= royalty_w else 0.0
        ebitda = revenue - tutors - manager - admin - rent - running - crm - marketing - royalties
        conv = s["enquiry_to_assessment"] * s["assessment_to_member"]
        enquiries = new / conv if conv else 0.0
        ads_needed = enquiries * (s["lead_after"] if after_presale else s["lead_presale"])
        if cur is None or cur["Model month"] != m:
            if cur is not None:
                months.append(cur)
                if not owner_on and s["owner_on"] and cur["EBITDA"] > 0 and \
                        cur["Model month"] >= model_month(open_w):
                    owner_on = True                    # first profitable month: wage from next
            cur = {"Model month": m, "Date": date(d.year, d.month, 1), "Members end": 0.0,
                   "Active end": 0.0, "Suspended": 0.0, "New members": 0.0, "Cancelled": 0.0,
                   "Founders": 0.0, "Price banked": 0.0, "Tuition": 0.0,
                   "Suspension fees": 0.0, "Joining fees": 0.0, "Other revenue": 0.0,
                   "Revenue": 0.0, "Tutor wages": 0.0, "Manager": 0.0, "Admin": 0.0,
                   "Rent": 0.0, "Running costs": 0.0, "CRM": 0.0, "Marketing": 0.0,
                   "Royalties": 0.0, "Royalties due": 0.0, "EBITDA": 0.0, "Owner's wage": 0.0,
                   "Study hubs": 0.0, "Hubs after 7 pm": 0.0, "Tables at peak": 0,
                   "Enquiries needed": 0.0, "Advertising needed": 0.0, "Advertising planned": 0.0,
                   "_weeks": 0}
        owner = s["owner_wage"] / 52 * wage_k if owner_on else 0.0
        for k, v in [("New members", new), ("Cancelled", canc), ("Tuition", tuition),
                     ("Suspension fees", sus_fees), ("Joining fees", joining),
                     ("Other revenue", other), ("Revenue", revenue), ("Tutor wages", tutors),
                     ("Manager", manager), ("Admin", admin), ("Rent", rent),
                     ("Running costs", running), ("CRM", crm), ("Marketing", marketing),
                     ("Royalties", royalties), ("Royalties due", royalty_due),
                     ("EBITDA", ebitda), ("Owner's wage", owner), ("Study hubs", hubs),
                     ("Hubs after 7 pm", late_hubs), ("Enquiries needed", enquiries),
                     ("Advertising needed", ads_needed),
                     ("Advertising planned", (s["ads_day"] * 7) if after_presale else 0.0)]:
            cur[k] += v
        cur.update({"Members end": books, "Active end": active, "Suspended": books * sus_pct,
                    "Founders": f, "Price banked": price,
                    "Tables at peak": max(cur["Tables at peak"], tables)})
        cur["_weeks"] += 1
    months.append(cur)
    t = pd.DataFrame(months).set_index("Model month")
    # a model month holds 4 or 5 weeks: show every month as an average 4.33-week month
    k = WPM / t["_weeks"]
    for col in FLOWS:
        t[col] = t[col] * k
    t["Study hubs"] = t["Study hubs"] / WPM               # a week, on average
    t["Hubs after 7 pm"] = t["Hubs after 7 pm"] / WPM
    t = t.drop(columns="_weeks")
    t["Profit"] = t["EBITDA"] - t["Owner's wage"]
    t["Cumulative"] = t["Profit"].cumsum()
    t["Tutor wages %"] = (t["Tutor wages"] / t["Tuition"].where(t["Tuition"] > 0) * 100)
    t["Capacity used %"] = t["Members end"] / capacity(s) * 100 if capacity(s) else None
    return t


# ── Investor results ──────────────────────────────────────────────────────────
def _irr(flows):
    """Monthly IRR by bisection (None if there's no sign change)."""
    npv = lambda r: sum(c / (1 + r) ** i for i, c in enumerate(flows))
    lo, hi = -0.99, 1.0
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if npv(lo) * npv(mid) > 0 else (lo, mid)
    return (lo + hi) / 2


def cash_flow(t, s):
    """Monthly cash: profit, less loan repayments and tax, plus the bond back and the sale
    value at the end. Month 0 = the money put in (investment + bond)."""
    invest, bond = s["investment"], s["rent"] * s["bond_months"]
    peak = max(-t["Cumulative"].min(), 0.0)
    working = peak * (1 + s["wc_buffer"])
    total = invest + bond + working
    loan = total * s["loan_share"]
    n = int(s["loan_years"] * 12)
    r = s["loan_rate"] / 12
    pay = (loan * r / (1 - (1 + r) ** -n) if r else loan / n) if loan and n else 0.0
    dep = s["investment_items"]["Fit-out and signage"] / (s["depreciation_years"] * 12) \
        if s["depreciation_years"] else 0.0
    rows, losses, balance = [], 0.0, loan
    for m, row in t.iterrows():
        interest = balance * r if balance > 0 else 0.0
        principal = min(pay - interest, balance) if balance > 0 else 0.0
        balance -= principal
        taxable = row["Profit"] - interest - dep
        tax = 0.0
        if m % 12 == 0:                                  # yearly, losses carried forward
            year = t.loc[m - 11:m]
            ytax = year["Profit"].sum() - interest * 12 - dep * 12 + losses
            tax = max(ytax, 0.0) * s["tax"]
            losses = min(ytax, 0.0)
        rows.append({"Model month": m, "Loan repayment": interest + principal if loan else 0.0,
                     "Tax": tax, "Cash": row["Profit"] - (interest + principal if loan else 0.0)
                     - tax})
    c = pd.DataFrame(rows).set_index("Model month")
    year5 = t.loc[49:60, "Profit"].sum()
    sale = max(year5, 0.0) * s["sale_multiple"]
    c.loc[c.index[-1], "Cash"] += bond + sale - balance
    c["Cumulative cash"] = c["Cash"].cumsum() - (total - loan)
    flows = [-(total - loan)] + list(c["Cash"])
    mirr = _irr(flows)
    dr = (1 + s["discount"]) ** (1 / 12) - 1
    npv = sum(v / (1 + dr) ** i for i, v in enumerate(flows))
    return c, dict(investment=invest, bond=bond, working_capital=working, total_cash=total,
                   loan=loan, own_cash=total - loan, sale_value=sale, npv=npv,
                   irr=(1 + mirr) ** 12 - 1 if mirr is not None and mirr > -0.5 else None)


def payback_month(t, s):
    hit = t[t["Cumulative"] >= s["investment"]]
    return int(hit.index[0]) if len(hit) else None


def royalty_start(t):
    after = t["Profit"] + t["Royalties"] - t["Royalties due"]
    ok = (after >= -0.5)[::-1].cummin()[::-1]
    hit = t[ok & (t["Revenue"] > 0)]
    return int(hit.index[0]) if len(hit) else None


def summary(t, s):
    c, money = cash_flow(t, s)
    first_profit = t[(t["EBITDA"] > 0) & (t.index >= model_month(int(s["open_week"])))]
    owner_start = t[t["Owner's wage"] > 0]
    cap = capacity(s)
    over_cap = t[t["Members end"] > cap] if cap else t.iloc[0:0]
    return dict(
        payback=payback_month(t, s), payback_target=int(s["payback_target"]),
        first_profit=int(first_profit.index[0]) if len(first_profit) else None,
        owner_start=int(owner_start.index[0]) if len(owner_start) else None,
        peak_cash=max(-t["Cumulative"].min(), 0.0),
        profit_5y=float(t["Profit"].sum()), owner_5y=float(t["Owner's wage"].sum()),
        royalty_start=royalty_start(t), capacity=cap,
        full_month=int(over_cap.index[0]) if len(over_cap) else None,
        ads_gap=float((t["Advertising needed"] - t["Advertising planned"]).loc[4:].mean()),
        **money)


def milestones(t, s):
    """[(model month or None, label)] in order."""
    first = lambda mask: int(t[mask].index[0]) if mask.any() else None
    smry = summary(t, s)
    ev = [(model_month(int(s["open_week"])), "Centre opens: sessions and billing start"),
          (model_month(int(s["presale_weeks"])), "Presale campaign ends"),
          (first(t["Members end"] >= s["t12"] - 1e-6), f"{s['t12']:g} members"),
          (model_month(max(9, int(s["open_week"]))), "Royalties start"),
          (first(t["Rent"] > 0), "Rent starts"),
          (smry["first_profit"], "First month in profit"),
          (smry["owner_start"], "Owner's salary starts"),
          (first(t["Manager"] > 0), "Manager hired"),
          (smry["royalty_start"], "Can afford royalties every month"),
          (smry["full_month"], "Centre full (capacity)"),
          (smry["payback"], "Investment paid back")]
    ev = [e for e in ev if not (e[0] is None and e[1] in ("Manager hired", "Owner's salary starts",
                                                        "Centre full (capacity)"))]
    return sorted(ev, key=lambda e: (e[0] is None, e[0] or 0))


def year_totals(t):
    y = t.assign(Year=(t.index - 1) // 12 + 1).groupby("Year")
    out = y[["Revenue", "Tutor wages", "Royalties", "EBITDA", "Owner's wage", "Profit"]].sum()
    out["Members (end)"] = y["Members end"].last()
    out["New members a month"] = y["New members"].sum() / 12
    out["Cumulative"] = y["Cumulative"].last()
    return out


def checkpoints(t, months=(3, 6, 12, 24, 36)):
    rows = []
    for m in months:
        r = t.loc[m]
        rows.append({"By month": m, "Members": r["Members end"], "Revenue a month": r["Revenue"],
                     "Tutor wages % of fees": r["Tutor wages %"],
                     "Profit so far": t.loc[:m, "Profit"].sum()})
    return pd.DataFrame(rows)


# ── What you need to hit ──────────────────────────────────────────────────────
def _solve(make, lo, hi, s, start, up=True, n=30):
    """The smallest (up) or largest (not up) value for which payback is within the target."""
    ok = lambda v: (payback_month(run(make(v), start), make(v)) or 999) <= s["payback_target"]
    if not ok(hi if up else lo):
        return None
    for _ in range(n):
        mid = (lo + hi) / 2
        if ok(mid) == up:
            hi = mid
        else:
            lo = mid
    return hi if up else lo


def thresholds(s, start):
    """For payback within the target, one change at a time (None = not reachable)."""
    k = _solve(lambda v: {**s, "t12": s["t12"] * v, "targets": [x * v for x in s["targets"]]},
               0.3, 3.0, s, start)
    return dict(
        members_scale=k,
        price=_solve(lambda v: {**s, "standard": v}, 0, 300, s, start),
        owner_wage=_solve(lambda v: {**s, "owner_wage": v}, 0, s["owner_wage"], s, start,
                          up=False) if s["owner_on"] else None,
        rent=_solve(lambda v: {**s, "rent": v}, 0, s["rent"] * 3, s, start, up=False),
        fill_offset=_solve(lambda v: {**s, "fill": [f + v for f in s["fill"]]}, -3, 3, s, start),
    )


# (Cancelled % is left out: members follow the targets, so it changes the new members and
# advertising needed, not profit.)
SENSITIVITY = [("Standard price", "standard"), ("Member targets", "targets"),
               ("Students per study hub", "fill"), ("Rent", "rent"), ("Collected %", "collected"),
               ("Tutor hourly rate", "rate"), ("Owner's salary", "owner_wage")]


def sensitivity(s, start, step=0.10):
    """Payback and 5-year profit with each input 10% worse and 10% better."""
    rows = []
    worse_up = {"rent", "rate", "owner_wage"}            # higher is worse for these
    for label, key in SENSITIVITY:
        r = {"Input": label}
        for name, k in (("10% worse", -1), ("10% better", 1)):
            f = 1 + step * (k if key not in worse_up else -k)
            v = s[key]
            s2 = {**s, key: [x * f for x in v] if isinstance(v, list) else v * f}
            if key == "targets":
                s2["t12"] = s["t12"] * f
            t = run(s2, start)
            r[f"Payback ({name})"] = payback_month(t, s2)
            r[f"5-year profit ({name})"] = float(t["Profit"].sum())
        rows.append(r)
    return pd.DataFrame(rows)
