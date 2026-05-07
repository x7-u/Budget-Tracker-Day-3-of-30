"""Day 3. Pure variance maths and direction-aware RAG classification.

A budget row is converted into a VarianceRow with five possible RAG bands:

* ``green``      adverse but within tolerance (small overspend / small revenue miss)
* ``amber``      adverse and between 1 and 3 multiples of tolerance
* ``red``        adverse and beyond 3 multiples of tolerance
* ``favourable`` cost underspend or revenue beat. Rendered blue, never green;
                 large favourable variances still warrant attention.
* ``na``         budget is zero; variance % is undefined.

Default tolerance is 5%. A per-row ``tolerance_pct`` on the input overrides
it on a per-line basis, useful for naturally noisy categories such as travel.
"""
from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from typing import Literal

from budget_schema import BudgetRow

DEFAULT_TOLERANCE = 0.05  # 5%
RAG_VALUES = ("green", "amber", "red", "favourable", "na")
Dim = Literal["cost_centre", "category", "period"]


# ---- Result types ------------------------------------------------------

@dataclass
class VarianceRow:
    period: str
    cost_centre: str
    category: str
    line_type: str
    budget: float
    actual: float
    variance: float                 # actual - budget
    variance_pct: float | None      # variance / budget, None if budget == 0
    tolerance_pct: float            # effective tolerance applied
    is_adverse: bool
    rag: str
    notes: str | None = None
    forecast: float | None = None              # carried through from input
    forecast_variance: float | None = None     # actual - forecast (when forecast present)
    forecast_variance_pct: float | None = None
    parent: str | None = None                  # parent cost centre for hierarchy
    z_score: float | None = None               # populated by anomaly flagger
    z_mean: float | None = None                # historical mean of variance_pct for this key
    z_stdev: float | None = None               # historical stdev
    z_n: int | None = None                     # historical observation count


@dataclass
class Aggregate:
    key: str
    budget: float
    actual: float
    variance: float
    variance_pct: float | None
    rag_counts: dict[str, int]      # one entry per RAG band
    row_count: int


@dataclass
class HeadlineStats:
    total_budget: float
    total_actual: float
    total_variance: float
    total_variance_pct: float | None
    rag_counts: dict[str, int]
    adverse_count: int
    favourable_count: int
    na_count: int
    row_count: int
    period_count: int


# ---- Per-row classification -------------------------------------------

def classify_row(row: BudgetRow, default_tol: float = DEFAULT_TOLERANCE) -> VarianceRow:
    tol = row.tolerance_pct if row.tolerance_pct is not None else default_tol
    variance = row.actual - row.budget

    forecast_variance, forecast_variance_pct = _forecast_variance(row)

    if row.budget == 0:
        return VarianceRow(
            period=row.period,
            cost_centre=row.cost_centre,
            category=row.category,
            line_type=row.line_type,
            budget=row.budget,
            actual=row.actual,
            variance=variance,
            variance_pct=None,
            tolerance_pct=tol,
            is_adverse=False,
            rag="na",
            notes=row.notes,
            forecast=row.forecast,
            forecast_variance=forecast_variance,
            forecast_variance_pct=forecast_variance_pct,
            parent=row.parent,
        )

    variance_pct = variance / row.budget
    is_adverse = (
        (row.line_type == "cost" and variance > 0)
        or (row.line_type == "revenue" and variance < 0)
    )

    # Round both abs_pct and the upper boundary so floats at the band
    # edges land where the spec says (5% becomes amber, 15% becomes red).
    abs_pct = round(abs(variance_pct), 6)
    amber_upper = round(3 * tol, 6)

    if variance == 0:
        rag = "green"
    elif not is_adverse:
        rag = "favourable"
    elif abs_pct < tol:
        rag = "green"
    elif abs_pct < amber_upper:
        rag = "amber"
    else:
        rag = "red"

    return VarianceRow(
        period=row.period,
        cost_centre=row.cost_centre,
        category=row.category,
        line_type=row.line_type,
        budget=row.budget,
        actual=row.actual,
        variance=variance,
        variance_pct=variance_pct,
        tolerance_pct=tol,
        is_adverse=is_adverse,
        rag=rag,
        notes=row.notes,
        forecast=row.forecast,
        forecast_variance=forecast_variance,
        forecast_variance_pct=forecast_variance_pct,
        parent=row.parent,
    )


def _forecast_variance(row: BudgetRow) -> tuple[float | None, float | None]:
    if row.forecast is None:
        return None, None
    fv = row.actual - row.forecast
    fpct = (fv / row.forecast) if row.forecast else None
    return fv, fpct


def compute_all(
    rows: Iterable[BudgetRow],
    default_tol: float = DEFAULT_TOLERANCE,
) -> list[VarianceRow]:
    return [classify_row(r, default_tol=default_tol) for r in rows]


# ---- Aggregations -----------------------------------------------------

def aggregate_by_parent(rows: list[VarianceRow]) -> list[Aggregate]:
    """Roll rows up to their parent cost centre. Skip rows with no parent.

    Returns an empty list when no row in the input declares a parent. Callers
    should treat that as 'no hierarchy in this workbook' and hide the panel.
    """
    has_parent = [r for r in rows if r.parent]
    if not has_parent:
        return []
    groups: dict[str, list[VarianceRow]] = {}
    order: list[str] = []
    for r in has_parent:
        key = r.parent or ""
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)

    aggs: list[Aggregate] = []
    for key in order:
        bucket = groups[key]
        budget = sum(r.budget for r in bucket)
        actual = sum(r.actual for r in bucket)
        variance = actual - budget
        variance_pct = (variance / budget) if budget else None
        aggs.append(Aggregate(
            key=key,
            budget=budget,
            actual=actual,
            variance=variance,
            variance_pct=variance_pct,
            rag_counts=_count_rags(bucket),
            row_count=len(bucket),
        ))
    aggs.sort(key=lambda a: abs(a.variance), reverse=True)
    return aggs


def aggregate_by(rows: list[VarianceRow], dim: Dim) -> list[Aggregate]:
    """Group rows by one of cost_centre / category / period and total them.

    Returns a list ordered by absolute variance descending so the most
    material entries surface first.
    """
    groups: dict[str, list[VarianceRow]] = {}
    order: list[str] = []
    for r in rows:
        key = getattr(r, dim)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)

    aggs: list[Aggregate] = []
    for key in order:
        bucket = groups[key]
        budget = sum(r.budget for r in bucket)
        actual = sum(r.actual for r in bucket)
        variance = actual - budget
        variance_pct = (variance / budget) if budget else None
        counts = _count_rags(bucket)
        aggs.append(Aggregate(
            key=key,
            budget=budget,
            actual=actual,
            variance=variance,
            variance_pct=variance_pct,
            rag_counts=counts,
            row_count=len(bucket),
        ))

    aggs.sort(key=lambda a: abs(a.variance), reverse=True)
    return aggs


def headline_stats(rows: list[VarianceRow]) -> HeadlineStats:
    total_budget = sum(r.budget for r in rows)
    total_actual = sum(r.actual for r in rows)
    total_variance = total_actual - total_budget
    total_variance_pct = (total_variance / total_budget) if total_budget else None

    counts = _count_rags(rows)
    adverse = sum(1 for r in rows if r.is_adverse)
    favourable = sum(1 for r in rows if r.rag == "favourable")
    na = sum(1 for r in rows if r.rag == "na")
    period_count = len({r.period for r in rows})

    return HeadlineStats(
        total_budget=total_budget,
        total_actual=total_actual,
        total_variance=total_variance,
        total_variance_pct=total_variance_pct,
        rag_counts=counts,
        adverse_count=adverse,
        favourable_count=favourable,
        na_count=na,
        row_count=len(rows),
        period_count=period_count,
    )


def top_n_by_variance(
    rows: list[VarianceRow],
    n: int,
    *,
    side: Literal["adverse", "favourable"],
) -> list[VarianceRow]:
    """Top N rows by absolute variance amount on the requested side."""
    if side == "adverse":
        candidates = [r for r in rows if r.is_adverse]
    else:
        candidates = [r for r in rows if r.rag == "favourable"]
    candidates.sort(key=lambda r: abs(r.variance), reverse=True)
    return candidates[:n]


# ---- Helpers ----------------------------------------------------------

def _count_rags(rows: list[VarianceRow]) -> dict[str, int]:
    counts = {k: 0 for k in RAG_VALUES}
    for r in rows:
        counts[r.rag] = counts.get(r.rag, 0) + 1
    return counts
