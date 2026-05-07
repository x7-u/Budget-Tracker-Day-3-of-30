"""Day 3. Tests for the flat CSV writer."""
from __future__ import annotations

from pathlib import Path

from budget_schema import BudgetMetadata, BudgetRow
from csv_writer import COLUMNS, write_csv
from pipeline import AnalysisResult, Commentary
from variance import (
    aggregate_by,
    compute_all,
    headline_stats,
)


def _row(**overrides) -> BudgetRow:
    base = dict(
        period="2026-04",
        cost_centre="Engineering",
        category="Salaries",
        line_type="cost",
        budget=1000.0,
        actual=1100.0,
        tolerance_pct=None,
        notes=None,
    )
    base.update(overrides)
    return BudgetRow(**base)


def _result(rows: list[BudgetRow], source: str = "unit-test.xlsx") -> AnalysisResult:
    var_rows = compute_all(rows)
    return AnalysisResult(
        metadata=BudgetMetadata(company="TestCo", currency="GBP",
                                fiscal_year="FY26", period_label="April 2026"),
        rows=var_rows,
        by_cost_centre=aggregate_by(var_rows, dim="cost_centre"),
        by_category=aggregate_by(var_rows, dim="category"),
        by_period=aggregate_by(var_rows, dim="period"),
        headline=headline_stats(var_rows),
        commentary=Commentary(skipped=True),
        warnings=[],
        source_filename=source,
        elapsed_ms=0,
    )


def test_header_row_matches_columns_constant(tmp_path: Path) -> None:
    out = write_csv(_result([_row()]), tmp_path / "out.csv")
    text = out.read_text(encoding="utf-8-sig")
    first_line = text.splitlines()[0]
    assert first_line.split(",") == list(COLUMNS)


def test_line_type_preserved(tmp_path: Path) -> None:
    out = write_csv(_result([_row(line_type="revenue")]), tmp_path / "out.csv")
    text = out.read_text(encoding="utf-8-sig")
    assert "revenue" in text


def test_zero_budget_row_has_blank_variance_pct(tmp_path: Path) -> None:
    out = write_csv(_result([_row(budget=0.0, actual=50.0)]), tmp_path / "out.csv")
    text = out.read_text(encoding="utf-8-sig")
    rows = text.splitlines()
    data_line = rows[1].split(",")
    pct_index = list(COLUMNS).index("variance_pct")
    assert data_line[pct_index] == ""


def test_encoding_is_utf8_with_bom(tmp_path: Path) -> None:
    out = write_csv(_result([_row()]), tmp_path / "out.csv")
    raw = out.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")


def test_source_filename_repeated_on_every_row(tmp_path: Path) -> None:
    out = write_csv(_result([_row(category="A"), _row(category="B")], source="myfile.xlsx"),
                    tmp_path / "out.csv")
    text = out.read_text(encoding="utf-8-sig")
    assert text.count("myfile.xlsx") == 2


def test_is_adverse_serialised_as_lowercase_string(tmp_path: Path) -> None:
    out = write_csv(_result([_row(actual=1300)]), tmp_path / "out.csv")
    text = out.read_text(encoding="utf-8-sig")
    assert ",true," in text or text.endswith("true")
