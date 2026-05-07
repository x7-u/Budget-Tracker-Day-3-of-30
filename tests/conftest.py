"""Day 3 test isolation.

Each day folder defines modules with the same basenames (excel_writer,
pipeline, csv_writer, server, main, etc.). pytest collects them all in
one process, so without a conftest the second day to run inherits cached
imports from the first.

This conftest evicts the conflicting names from sys.modules and prepends
Day 3's folder to sys.path so 'from variance import ...' resolves to
Day 3's version.
"""
from __future__ import annotations

import sys
from pathlib import Path

DAY_ROOT = Path(__file__).resolve().parent.parent
PROJECT_ROOT = DAY_ROOT.parent

_CONFLICTING = {
    "excel_writer", "pipeline", "csv_writer", "ratios", "sectors",
    "pdf_loader", "invoice_schema", "main", "server", "ledger",
    "variance", "budget_schema", "cost_log", "pptx_writer", "pdf_writer",
    "history_store", "comparison", "run_cache", "power_bi",
    # Cross-day collisions (so later days' modules don't bleed in).
    "news_schema", "aggregation", "sentiment", "chart",
    "corrections", "live_fetcher", "industries",
    "pulse_pptx", "pulse_pdf",
    "cashflow_schema", "cashflow_maths", "scenario",
    # Day 6 modules.
    "cvp_schema", "cvp_maths", "cvp_chart", "cvp_excel", "cvp_csv",
    "break_pptx", "break_pdf", "monte_carlo", "benchmarks",
    # Day 7 modules.
    "transcript_schema", "analysis", "hedge_phrases", "claims",
    "edgar", "brief_chart",
}


def _evict_and_set_path() -> None:
    for name in list(_CONFLICTING):
        sys.modules.pop(name, None)
    for p in (str(DAY_ROOT), str(PROJECT_ROOT)):
        if p in sys.path:
            sys.path.remove(p)
    sys.path.insert(0, str(DAY_ROOT))
    sys.path.insert(0, str(PROJECT_ROOT))


_evict_and_set_path()


def pytest_collectstart(collector):
    p = getattr(collector, "path", None) or getattr(collector, "fspath", None)
    if p is None:
        return
    if str(DAY_ROOT) in str(p):
        _evict_and_set_path()


import pytest


@pytest.fixture(autouse=True)
def _ensure_day_path():
    """Re-evict before every Day 3 test, so cross-day pollution from later
    days does not leak in during the full-suite run."""
    _evict_and_set_path()
    yield


for p in (str(DAY_ROOT), str(PROJECT_ROOT)):
    if p in sys.path:
        sys.path.remove(p)
sys.path.insert(0, str(DAY_ROOT))
sys.path.insert(0, str(PROJECT_ROOT))
