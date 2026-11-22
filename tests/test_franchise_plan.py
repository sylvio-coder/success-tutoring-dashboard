"""Franchise Model engine (franchise_plan.py): members, revenue, study hub tutor costs, owner's
salary, rent, royalties, investor results, thresholds and sensitivity."""
import datetime as dt

import pytest

import franchise_inputs as fin
import franchise_plan as fp

APPROVED = {  # AU National as approved 22-11-2026: (Worst, Base, Best)
    'Franchise grant': (50000, 50000, 50000),
    'Success Pack': (15000, 15000, 15000),
    'Reservation deposit': (3000, 3000, 3000),
    'Fit-out and signage': (20000, 15000, 13000),
    'Legal and accounting (company set-up)': (5000, 3000, 1000),
    'Legal costs for the lease': (2000, 2000, 2000),
    'Presale marketing budget (weeks 1–12)': (15000, 15000, 12000),
    'Rental bond': (4, 3, 2),
    'Working capital buffer': (20, 20, 20),
    "Owner's cash share of total cash needed": (50, 100, 100),
    'Loan share of funding': (50, 0, 0),
    'Loan interest rate': (11, 9, 7),
    'Loan term': (5, 5, 5),
    'Exchange rate (1 unit of local currency in AUD)': (1, 1, 1),
    'Payback target': (48, 48, 48),
    'Required return (discount rate)': (20, 15, 12),
    'Business sale value at Year 5': (2, 2.5, 3),
    'Company tax rate': (25, 25, 25),
    'Fit-out depreciation': (7, 7, 7),
    "Owner's wage": (60000, 60000, 60000),
    'Franchise term': (5, 5, 5),
    'Renewal fee': (3000, 3000, 3000),
    'Lease starts before week 1': (2, 1, 0),
    'First sign-ups': (3, 2, 1),
    'Centre opens (sessions and billing start)': (11, 9, 7),
    'Presale campaign ends': (12, 12, 12),
    'Model length': (60, 60, 60),
    'Members at week 12': (30, 40, 50),
    'Members at end of Year 1': (50, 80, 90),
    'Members at end of Year 2': (60, 110, 120),
    'Members at end of Year 3': (70, 130, 150),
    'Members at end of Year 4': (80, 140, 165),
    'Members at end of Year 5': (90, 145, 175),
    'Year 1: cancelled': (10, 8, 6),
    'Year 2: cancelled': (10, 8, 6),
    'Year 3: cancelled': (10, 8, 6),
    'Year 4: cancelled': (10, 8, 6),
    'Year 5: cancelled': (10, 8, 6),
    'Study hub tables in the centre': (4, 4, 4),
    'Standard price (members after the first 50)': (70, 70, 70),
    'Foundation members leaving': (8.5, 4, 3),
    'Collected after GST (allowance for uncollected)': (90, 95, 98),
    'GST': (10, 10, 10),
    'Price rise each July (from 12 months)': (0, 3, 4),
    'Joining fee': (100, 100, 100),
    'Suspension fee': (5, 5, 5),
    'Other revenue (holiday programs, assessments)': (0, 0, 0),
    'Student visits per member': (2, 2, 2),
    '1:1 share of visits': (4, 4, 4),
    'Students per study hub: Year 1, Q1 (weeks 1–13)': (2.5, 3, 4),
    'Students per study hub: Year 1, Q2': (3, 4, 4.5),
    'Students per study hub: Year 1, Q3': (3.5, 4.5, 5),
    'Students per study hub: Year 1, Q4': (4, 5, 5.5),
    'Students per study hub: Year 2': (4.5, 5.5, 6),
    'Students per study hub: Year 3': (5, 6, 6),
    'Students per study hub: Year 4': (5, 6, 6),
    'Students per study hub: Year 5': (5, 6, 6),
    'Most students per study hub': (6, 6, 6),
    'Study hub length': (1, 1, 1),
    'Tutors per study hub': (1.2, 1.2, 1.2),
    'Target students per study hub (planning)': (4.5, 5.5, 6),
    'Merge study hubs below': (2, 3, 3),
    'Weekdays open': (4, 4, 5),
    'Peak time slots: weekday before 7 pm': (3, 3, 3),
    'Weekday slots after 7 pm (busier centres)': (1, 1, 1),
    'Peak time slots: Saturday': (4, 4, 4),
    'Share of study hubs on Saturday': (0, 20, 15),
    'Minimum paid shift': (2, 2, 2),
    'Tutor hourly rate: weekday (casual, level 2, age 20)': (40, 33.85, 33.85),
    'Tutor hourly rate: Saturday and after 7 pm (casual)': (45, 39.27, 39.27),
    'Super': (12, 12, 12),
    'Manager salary (ex super)': (40000, 32000, 32000),
    'Front desk / admin': (0, 0, 0),
    'Wage rise each July': (4.75, 4.75, 4.75),
    'Rent (including outgoings)': (4222, 3936, 2894),
    'Rent-free months': (0, 2, 3),
    'Rent rise each July': (4, 3, 3),
    'Lease term': (5, 5, 5),
    'Agency fee (at least)': (1200, 250, 250),
    'Advertising': (40, 30, 30),
    'Marketing fund levy': (0, 0, 0),
    'Enquiry → Student Assessment': (15, 20, 30),
    'Student Assessment → member': (50, 80, 85),
    'Cost per lead: presale (weeks 1–12)': (55, 45, 35),
    'Cost per lead: after presale': (66, 54, 42),
    'Royalty': (8, 8, 8),
    'Royalty minimum (FY2026–27)': (2266, 2266, 2266),
    'Royalty minimum rise each July': (3, 3, 3),
    'CRM software (Hapana)': (445, 250, 250),
    'Learning materials, printing and stationery': (7, 6, 5),
    'Other software and subscriptions': (300, 249, 200),
    'Office expenses': (290, 238, 190),
    'Insurance': (180, 145, 90),
    'Energy and water': (400, 152, 120),
    'Internet and phone': (160, 135, 110),
    'Accounting, bookkeeping and consulting': (300, 153, 120),
    'Cleaning': (150, 70, 50),
    'Repairs and maintenance': (120, 83, 50),
    'Workers compensation': (100, 78, 60),
    'Bank fees': (90, 72, 50),
    'Staff expenses and uniforms': (100, 79, 50),
    'Running costs rise each July': (4, 3, 3),
    'Replacement spending (computers, furniture)': (0, 0, 0),
    'Hire a manager at': ('when sessions start', 80, 'never'),
    "Owner's wage included": ('yes', 'yes', 'yes'),
    'Saturday open': ('no', 'yes', 'yes'),
    "Owner's wage starts": ('first profitable month', 'first profitable month', 'first profitable month'),
    'Sign-up seasonality (1.0 = average month)': ('0.76, 1.82, 1.24, 0.72, 1.18, 1.13, 1.19, 1.45, 0.72, 0.77, 0.71, 0.30', '0.76, 1.82, 1.24, 0.72, 1.18, 1.13, 1.19, 1.45, 0.72, 0.77, 0.71, 0.30', '0.76, 1.82, 1.24, 0.72, 1.18, 1.13, 1.19, 1.45, 0.72, 0.77, 0.71, 0.30'),
    'Extra members suspended in holidays': ('Jan 17, Apr 9, Jul 5, Oct 2.5, Dec 15', 'Jan 17, Apr 9, Jul 5, Oct 2.5, Dec 15', 'Jan 17, Apr 9, Jul 5, Oct 2.5, Dec 15'),
    'School holiday months': ('by state', 'by state', 'by state'),
    'Package mix: Group 1× / 2× / 3×+ / 1:1': ('12 / 73 / 8 / 7', '12 / 73 / 8 / 7', '12 / 73 / 8 / 7'),
    'Foundation member numbers (tiers)': ('20 / 20 / 10', '20 / 20 / 10', '20 / 20 / 10'),
    'Foundation prices (tiers)': ('50 / 60 / 65', '50 / 60 / 65', '50 / 60 / 65'),
}

START = dt.date(2026, 12, 1)


def inputs(**changes):
    """model_inputs()-style dict from APPROVED, with some Base values changed."""
    out = {v: dict(zip(fin.CASES, vals)) for v, vals in APPROVED.items()}
    for v, val in changes.items():
        out[v]["Base"] = val
    return out


def base(**changes):
    return fp.settings(inputs(**changes), "Base")


def test_settings_read_every_case():
    w, b, best = (fp.settings(inputs(), c) for c in fin.CASES)
    assert b["investment"] == 50000 + 15000 + 3000 + 15000 + 3000 + 2000 + 15000
    assert w["manager_at"] == 0 and b["manager_at"] == 80 and best["manager_at"] is None
    assert b["tiers"] == [20, 20, 10] and b["tier_prices"] == [50, 60, 65]
    assert b["suspend"] == {"01": 17, "04": 9, "07": 5, "10": 2.5, "12": 15}
    assert not w["saturday"] and b["saturday"] and b["owner_on"]
    assert b["cancel"] == [0.08] * 5 and w["cancel"] == [0.10] * 5
    assert len(b["seasonality"]) == 12 and len(b["fill"]) == 8


def test_member_targets_and_seasonality():
    s = base()
    p = fp.member_path(s, START)
    assert p[1] == 0 and p[12] == pytest.approx(40)            # first sign-ups week 2
    assert p[2] > 0
    assert p[52] == pytest.approx(80) and p[104] == pytest.approx(110)
    assert p[260] == pytest.approx(145)
    grow = {}                                                  # Year 2 growth by month
    for w in range(53, 105):
        m = fp.week_date(START, w).month
        grow[m] = grow.get(m, 0) + p[w] - p[w - 1]
    assert grow[2] > 3 * grow[12]                              # February 1.82 vs December 0.30


def test_new_members_are_growth_plus_cancellations():
    s = base()
    t = fp.run(s, START)
    y2 = t.loc[13:24]
    assert y2["New members"].sum() == pytest.approx(
        (110 - 80) + y2["Cancelled"].sum(), rel=0.02)
    assert t.loc[1:2, "Cancelled"].sum() == 0                  # none before week 13


def test_revenue_price_joining_and_suspensions():
    s = base(**{"Foundation prices (tiers)": "70 / 70 / 70",
                "Extra members suspended in holidays": "Jan 10"})
    t = fp.run(s, START)
    assert t.loc[1, "Tuition"] == 0                            # presale: opens week 9
    assert t.loc[1, "Joining fees"] > 0                        # joining fees from week 2
    banked = 70 / 1.1 * 0.95
    assert t.loc[12, "Price banked"] == pytest.approx(banked)
    assert t.loc[24, "Price banked"] == pytest.approx(banked * 1.03)   # July 2028: first rise
    jan = t[[d.month == 1 and d.year == 2028 for d in t["Date"]]].iloc[0]
    assert jan["Active end"] == pytest.approx(jan["Members end"] * 0.90)
    assert jan["Suspension fees"] > 0
    assert t["Joining fees"].sum() == pytest.approx(t["New members"].sum() * 100 / 1.1,
                                                    rel=0.02)


def test_study_hub_tutor_cost_minimum_shift_saturday_and_after_7pm():
    s = base()
    cost, hubs, late, tables = fp.tutor_week(s, 0.0, 30, 1.0)
    assert cost == 0 and hubs == 0
    small = fp.tutor_week(s, 10, 30, 1.0)[0]                  # a few students: 2-hour minimums
    assert small == pytest.approx((4 * 2 * 33.85 + 2 * 39.27) * 1.12)
    cost, hubs, late, tables = fp.tutor_week(s, 120, 300, 1.0)   # Year 5: 6 a hub
    group = 120 * 2 * 0.96
    assert hubs == pytest.approx(group / 6)
    assert late == 0 and tables == 3                           # 7.7 hubs a weekday, 3 slots
    busy = fp.tutor_week({**s, "tables": 1}, 120, 300, 1.0)
    assert busy[2] > 0                                         # earlier slots full: after 7 pm
    closed = fp.tutor_week({**s, "saturday": False}, 120, 300, 1.0)
    assert closed[0] < fp.tutor_week(s, 120, 300, 1.0)[0]      # no Saturday penalty hours


def test_manager_owner_rent_and_royalties():
    s = base()
    t = fp.run(s, START)
    hired = t[t["Manager"] > 0].index[0]
    assert t.loc[hired, "Members end"] >= 80 and t.loc[hired - 1, "Members end"] < 80
    assert fp.run({**s, "manager_at": None}, START)["Manager"].sum() == 0
    assert fp.run({**s, "manager_at": 0.0}, START).loc[3, "Manager"] > 0   # opens in month 3
    first = fp.summary(t, s)["first_profit"]                   # from opening (month 3)
    assert first >= 3 and t.loc[first, "EBITDA"] > 0
    assert t.loc[first, "Owner's wage"] == 0 and t.loc[first + 1, "Owner's wage"] > 0
    assert fp.run({**s, "owner_on": False}, START)["Owner's wage"].sum() == 0
    # lease starts 1 month before week 1 with 2 months free: rent from month 2
    assert t.loc[1, "Rent"] < s["rent"] and t.loc[3, "Rent"] == pytest.approx(s["rent"], rel=0.01)
    assert t.loc[1, "Royalties"] == 0 and t.loc[3, "Royalties"] > 0      # from week 9
    late = fp.run({**s, "open_week": 14.0}, START)             # opens in month 4
    assert late.loc[3, "Royalties"] == 0 and late.loc[4, "Royalties"] > 0   # later of 8 / opening


def test_capacity_investor_results_and_milestones():
    s = base()
    assert fp.capacity(s) == pytest.approx(4 * (4 * 4 + 4) * 6 / 2 / 0.96)
    t = fp.run(s, START)
    smry = fp.summary(t, s)
    c, money = fp.cash_flow(t, s)
    assert money["total_cash"] == pytest.approx(
        money["investment"] + money["bond"] + money["working_capital"])
    assert money["working_capital"] == pytest.approx(smry["peak_cash"] * 1.2)
    assert money["bond"] == 3 * 3936
    assert smry["profit_5y"] == pytest.approx(t["Profit"].sum())
    pay = fp.payback_month(t, s)
    assert pay is None or t.loc[pay, "Cumulative"] >= s["investment"]
    ms = fp.milestones(t, s)
    known = [m for m, _ in ms if m is not None]
    assert known == sorted(known) and ms[0][1].startswith("Centre opens")
    loan = fp.cash_flow(t, {**s, "loan_share": 0.5})[1]
    assert loan["loan"] == pytest.approx(loan["total_cash"] * 0.5)
    y = fp.year_totals(t)
    assert list(y["Members (end)"].round()) == [80, 110, 130, 140, 145]
    cp = fp.checkpoints(t)
    assert list(cp["By month"]) == [3, 6, 12, 24, 36]


def test_thresholds_and_sensitivity():
    s = base(**{"Payback target": 60})
    th = fp.thresholds(s, START)
    p = th["price"]
    assert p is not None
    assert (fp.payback_month(fp.run({**s, "standard": p}, START), s) or 999) <= 60
    assert (fp.payback_month(fp.run({**s, "standard": p - 2}, START), s) or 999) > 60
    sens = fp.sensitivity(s, START)
    assert list(sens["Input"])[:2] == ["Standard price", "Member targets"]
    row = sens.set_index("Input").loc["Standard price"]
    assert row["5-year profit (10% worse)"] < row["5-year profit (10% better)"]


def test_worst_base_best_order():
    res = {c: fp.summary(fp.run(fp.settings(inputs(), c), START), fp.settings(inputs(), c))
           for c in fin.CASES}
    assert res["Worst"]["profit_5y"] < res["Base"]["profit_5y"] < res["Best"]["profit_5y"]
    assert res["Best"]["payback"] is not None
