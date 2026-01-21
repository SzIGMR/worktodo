from datetime import date

import pytest

from skillrunner.io_utils import parse_german_number
from skillrunner_budgetmatch.skill import (
    Booking,
    Budget,
    parse_month,
    run_forecast,
    run_optimization,
)


def test_parse_german_number() -> None:
    assert parse_german_number("1.640,47") == pytest.approx(1640.47)
    assert parse_german_number(" 7.452,00 € ") == pytest.approx(7452.0)


def test_parse_month() -> None:
    assert parse_month("2025/10") == date(2025, 10, 1)


def test_optimization_heuristic_allocates() -> None:
    bookings = [
        Booking(person="Anna", month=date(2025, 4, 1), amount_eur=100.0),
        Booking(person="Anna", month=date(2025, 5, 1), amount_eur=50.0),
        Booking(person="Ben", month=date(2025, 4, 1), amount_eur=80.0),
    ]
    budgets = [
        Budget(project="A", budget_eur=160.0),
        Budget(project="B", budget_eur=200.0),
    ]
    result = run_optimization(
        bookings=bookings,
        budgets=budgets,
        fy_start_year=2025,
        min_project_rest=1.0,
        person_max_projects=2,
        person_project_penalty=10.0,
        monthly_smoothness_penalty=0.1,
        unallocated_person_penalty=0.01,
        solver="heuristic",
        time_limit_seconds=5,
    )
    total_allocated = sum(
        row["amount_eur"]
        for row in result.allocation_rows
        if row["assigned_project"] != "UNALLOCATED"
    )
    assert total_allocated <= 358.0
    assert total_allocated > 0


def test_forecast_summary() -> None:
    bookings = [
        Booking(person="Anna", month=date(2025, 4, 1), amount_eur=100.0),
        Booking(person="Anna", month=date(2025, 5, 1), amount_eur=120.0),
        Booking(person="Anna", month=date(2025, 6, 1), amount_eur=140.0),
    ]
    budgets = [Budget(project="A", budget_eur=1000.0)]
    result = run_forecast(bookings=bookings, budgets=budgets, fy_start_year=2025)
    forecast_rows = result.diagnostics["forecast_rows"]
    assert forecast_rows[0]["actuals_to_date"] == pytest.approx(360.0)
