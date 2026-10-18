"""Rolling membership trend vs average."""
import pandas as pd
import pytest

import membership_trend as mt

LOC = mt.LOC


def weekly(series, start="2026-06-07"):
    """{centre: (state, [members per week])} -> Weekly Membership rows."""
    dates = pd.date_range(start, periods=max(len(v) for _, v in series.values()), freq="7D")
    rows = []
    for loc, (state, values) in series.items():
        pad = len(dates) - len(values)            # newer centres start later
        for d, v in zip(dates[pad:], values):
            rows.append({LOC: loc, "Date": d, "# Active members": v, "State": state})
    return pd.DataFrame(rows), dates


FLAT = [100] * 20
GROW = [100] * 16 + [110] * 4          # last 4 weeks 10% up on the 4 before
FALL = [100] * 16 + [90] * 4


def test_categories_and_bounds():
    assert mt.category(5.1) == "Significantly above average trend"
    assert mt.category(5) == "Above average trend"
    assert mt.category(2) == "About average" and mt.category(-2) == "About average"
    assert mt.category(-5) == "Below average trend"
    assert mt.category(-5.1) == "Significantly below average trend"
    assert mt.STATUS["About average"] == "Orange"


def test_trend_vs_network_average():
    df, dates = weekly({"A": ("NSW", GROW), "B": ("NSW", FLAT), "C": ("VIC", FALL),
                        "D": ("VIC", FLAT)})
    rows, too_new, s = mt.trend_table(df, dates[-1])
    assert s["network trend"] == pytest.approx(0.0)       # 410 vs 400... (+10, 0, −10, 0)
    assert rows.loc["A", "Trend %"] == pytest.approx(10)
    assert rows.loc["A", "Category"] == "Significantly above average trend"
    assert rows.loc["C", "Status"] == "Red" and rows.loc["B", "Status"] == "Orange"
    assert rows.loc["A", "Change in members"] == 10 and rows.loc["A", "Members now"] == 110
    assert len(rows.loc["A", "Last 12 weeks"]) == 12
    assert s["counts"]["Significantly above average trend"] == 1 and too_new.empty


def test_state_basis_compares_within_state():
    df, dates = weekly({"A": ("NSW", GROW), "B": ("NSW", GROW), "C": ("VIC", FALL),
                        "D": ("VIC", FALL)})
    rows, _, _ = mt.trend_table(df, dates[-1], "State")
    assert rows["vs average"].abs().max() == pytest.approx(0)     # everyone = their state
    rows, _, _ = mt.trend_table(df, dates[-1], "All centres")
    assert rows.loc["A", "Category"] == "Significantly above average trend"


def test_new_and_small_centres():
    df, dates = weekly({"A": ("NSW", FLAT), "B": ("NSW", FLAT), "New": ("NSW", [5] * 17),
                        "Tiny": ("NSW", [10] * 20)})
    rows, too_new, _ = mt.trend_table(df, dates[-1])
    assert "New" not in rows.index and too_new.loc["New", "Weeks trading"] == 17
    assert bool(rows.loc["Tiny", "Small centre"]) and not bool(rows.loc["A", "Small centre"])


def test_moved_group_since_last_week():
    df, dates = weekly({"A": ("NSW", [100] * 19 + [140]), "B": ("NSW", FLAT),
                        "C": ("NSW", FLAT)})
    rows, _, _ = mt.trend_table(df, dates[-1])
    assert rows.loc["A", "Moved"] == "▲ up"
    rows, _, _ = mt.trend_table(df, dates[-2])                      # as at the week before
    assert rows.loc["A", "Moved"] == ""


def test_not_enough_weeks():
    df, dates = weekly({"A": ("NSW", [100] * 5)})
    rows, _, s = mt.trend_table(df, dates[-1])
    assert rows.empty and s["weeks"] == 5
