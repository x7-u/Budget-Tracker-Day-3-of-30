"""Day 3. Tests for pure variance maths and RAG classification."""
from __future__ import annotations

from budget_schema import BudgetRow
from variance import (
    DEFAULT_TOLERANCE,
    aggregate_by,
    classify_row,
    compute_all,
    headline_stats,
    top_n_by_variance,
)


def _row(**overrides) -> BudgetRow:
    base = dict(
        period="2026-04",
        cost_centre="Engineering",
        category="Salaries",
        line_type="cost",
        budget=1000.0,
        actual=1000.0,
        tolerance_pct=None,
        notes=None,
    )
    base.update(overrides)
    return BudgetRow(**base)


# ---- Per-row classification -------------------------------------------

class TestClassifyRow:
    def test_zero_variance_is_green(self):
        v = classify_row(_row(actual=1000.0))
        assert v.rag == "green"
        assert v.is_adverse is False
        assert v.variance == 0.0
        assert v.variance_pct == 0.0

    def test_cost_overspend_inside_tolerance_is_green(self):
        v = classify_row(_row(actual=1040.0))   # +4%
        assert v.rag == "green"
        assert v.is_adverse is True

    def test_cost_overspend_just_inside_tolerance_is_green(self):
        v = classify_row(_row(actual=1049.0))   # +4.9%
        assert v.rag == "green"

    def test_cost_overspend_at_tolerance_boundary_is_amber(self):
        v = classify_row(_row(actual=1050.0))   # +5.0%
        assert v.rag == "amber"

    def test_cost_overspend_inside_amber_band_is_amber(self):
        v = classify_row(_row(actual=1149.0))   # +14.9%
        assert v.rag == "amber"

    def test_cost_overspend_at_red_boundary_is_red(self):
        v = classify_row(_row(actual=1150.0))   # +15.0%
        assert v.rag == "red"

    def test_cost_underspend_is_favourable(self):
        v = classify_row(_row(actual=850.0))   # 15% under
        assert v.rag == "favourable"
        assert v.is_adverse is False

    def test_revenue_overperform_is_favourable(self):
        v = classify_row(_row(line_type="revenue", actual=1100.0))
        assert v.rag == "favourable"
        assert v.is_adverse is False

    def test_revenue_underperform_outside_tolerance_is_red(self):
        v = classify_row(_row(line_type="revenue", actual=800.0))   # -20%
        assert v.rag == "red"
        assert v.is_adverse is True

    def test_revenue_underperform_inside_tolerance_is_green(self):
        v = classify_row(_row(line_type="revenue", actual=970.0))   # -3%
        assert v.rag == "green"
        assert v.is_adverse is True

    def test_zero_budget_returns_na(self):
        v = classify_row(_row(budget=0.0, actual=500.0))
        assert v.rag == "na"
        assert v.variance_pct is None
        assert v.is_adverse is False

    def test_zero_budget_zero_actual_is_na(self):
        v = classify_row(_row(budget=0.0, actual=0.0))
        assert v.rag == "na"

    def test_per_row_tolerance_override(self):
        # 10% overspend with default 5% tol = amber.
        # Same overspend with 15% tol = green.
        v_default = classify_row(_row(actual=1100.0))
        v_relaxed = classify_row(_row(actual=1100.0, tolerance_pct=0.15))
        assert v_default.rag == "amber"
        assert v_relaxed.rag == "green"
        assert v_relaxed.tolerance_pct == 0.15

    def test_default_tolerance_constant(self):
        # Sanity check the published default has not drifted.
        assert DEFAULT_TOLERANCE == 0.05


# ---- Aggregation -------------------------------------------------------

class TestAggregateBy:
    def _setup(self) -> list:
        rows = [
            _row(cost_centre="Eng",  category="Salaries", budget=100, actual=110),
            _row(cost_centre="Eng",  category="Software", budget=50,  actual=80),
            _row(cost_centre="Mktg", category="Travel",   budget=20,  actual=18),
        ]
        return compute_all(rows)

    def test_aggregate_by_cost_centre(self):
        aggs = aggregate_by(self._setup(), dim="cost_centre")
        assert {a.key for a in aggs} == {"Eng", "Mktg"}
        eng = next(a for a in aggs if a.key == "Eng")
        assert eng.budget == 150
        assert eng.actual == 190
        assert eng.variance == 40

    def test_aggregate_by_category(self):
        aggs = aggregate_by(self._setup(), dim="category")
        assert {a.key for a in aggs} == {"Salaries", "Software", "Travel"}

    def test_aggregate_orders_by_absolute_variance(self):
        # Eng total variance is 40; Mktg total variance is -2. Eng comes first.
        aggs = aggregate_by(self._setup(), dim="cost_centre")
        assert aggs[0].key == "Eng"

    def test_aggregate_handles_zero_total_budget(self):
        rows = compute_all([_row(budget=0.0, actual=10.0)])
        aggs = aggregate_by(rows, dim="cost_centre")
        assert aggs[0].variance_pct is None


# ---- Headline stats ----------------------------------------------------

class TestHeadlineStats:
    def test_totals(self):
        rows = compute_all([
            _row(actual=1100),                                   # adverse, amber (+10%)
            _row(actual=900),                                    # favourable (cost underspend)
            _row(line_type="revenue", actual=1100),              # favourable (revenue beat)
            _row(budget=0, actual=50),                           # na
        ])
        h = headline_stats(rows)
        assert h.total_budget == 1000 + 1000 + 1000 + 0
        assert h.total_actual == 1100 + 900 + 1100 + 50
        assert h.total_variance == h.total_actual - h.total_budget
        assert h.row_count == 4
        assert h.adverse_count == 1
        assert h.favourable_count == 2
        assert h.na_count == 1

    def test_period_count(self):
        rows = compute_all([
            _row(period="2026-Q1"),
            _row(period="2026-Q1", category="Software"),
            _row(period="2026-Q2"),
        ])
        assert headline_stats(rows).period_count == 2


# ---- Top-N selection --------------------------------------------------

class TestTopN:
    def test_top_adverse_orders_by_absolute_variance(self):
        rows = compute_all([
            _row(category="A", budget=100, actual=130),  # +30
            _row(category="B", budget=100, actual=120),  # +20
            _row(category="C", budget=100, actual=160),  # +60
            _row(category="D", budget=100, actual=80),   # favourable, ignored
        ])
        top = top_n_by_variance(rows, n=2, side="adverse")
        assert [r.category for r in top] == ["C", "A"]

    def test_top_favourable_excludes_adverse(self):
        rows = compute_all([
            _row(category="A", budget=100, actual=130),  # adverse
            _row(category="B", budget=100, actual=70),   # favourable, -30
            _row(category="C", budget=100, actual=60),   # favourable, -40
        ])
        top = top_n_by_variance(rows, n=2, side="favourable")
        assert [r.category for r in top] == ["C", "B"]

    def test_top_n_caps_below_requested_size(self):
        rows = compute_all([_row(actual=1100)])  # adverse, +10%
        assert len(top_n_by_variance(rows, n=10, side="adverse")) == 1
