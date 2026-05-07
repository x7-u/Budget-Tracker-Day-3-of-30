"""Day 3. Flat CSV output, one row per variance row.

The CSV is the bookkeeping-friendly representation. Drops cleanly into
Power Query, pandas, or a Xero or QuickBooks import.
"""
from __future__ import annotations

import csv
from pathlib import Path

COLUMNS = (
    "period",
    "cost_centre",
    "parent",
    "category",
    "line_type",
    "budget",
    "actual",
    "forecast",
    "variance",
    "variance_pct",
    "forecast_variance",
    "forecast_variance_pct",
    "z_score",
    "rag",
    "tolerance_pct",
    "is_adverse",
    "notes",
    "source_filename",
)


def write_csv(result, out_path: Path | str) -> Path:
    """Write the analysis rows to a flat UTF-8-with-BOM CSV.

    The BOM is what makes Excel pick UTF-8 automatically when you double-click.
    """
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    with out.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        for r in result.rows:
            writer.writerow({
                "period":               r.period,
                "cost_centre":          r.cost_centre,
                "parent":               r.parent or "",
                "category":             r.category,
                "line_type":            r.line_type,
                "budget":               r.budget,
                "actual":               r.actual,
                "forecast":             "" if r.forecast is None else r.forecast,
                "variance":             r.variance,
                "variance_pct":         "" if r.variance_pct is None else r.variance_pct,
                "forecast_variance":    "" if r.forecast_variance is None else r.forecast_variance,
                "forecast_variance_pct": "" if r.forecast_variance_pct is None else r.forecast_variance_pct,
                "z_score":              "" if r.z_score is None else r.z_score,
                "rag":                  r.rag,
                "tolerance_pct":        r.tolerance_pct,
                "is_adverse":           "true" if r.is_adverse else "false",
                "notes":                r.notes or "",
                "source_filename":      result.source_filename or "",
            })
    return out
