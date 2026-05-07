"""Day 3. Orchestrator: parse, classify, narrate.

Glues budget_schema (input) and variance (maths) together, then calls
Claude to draft a controller-voice management commentary on the variance
data. The output AnalysisResult carries everything the writers and the
web UI need: rows, aggregates, headline stats, the commentary, plus
cost stats from the API call.
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from budget_schema import BudgetMetadata, parse_inputs
from history_store import HistoryStore, score_rows
from variance import (
    Aggregate,
    HeadlineStats,
    VarianceRow,
    aggregate_by,
    aggregate_by_parent,
    compute_all,
    headline_stats,
    top_n_by_variance,
)

from shared.config import CLAUDE_MODEL_FAST
from shared.llm_client import ask_claude_json_with_stats

# ---- Result dataclasses ----------------------------------------------

@dataclass
class Commentary:
    headline: str = ""
    summary: str = ""
    adverse_drivers: list[str] = field(default_factory=list)
    favourable_drivers: list[str] = field(default_factory=list)
    actions: list[str] = field(default_factory=list)
    cost_usd: float = 0.0
    input_tokens: int = 0
    output_tokens: int = 0
    model: str = ""
    skipped: bool = False
    error: str | None = None


@dataclass
class AnalysisResult:
    metadata: BudgetMetadata
    rows: list[VarianceRow]
    by_cost_centre: list[Aggregate]
    by_category: list[Aggregate]
    by_period: list[Aggregate]
    headline: HeadlineStats
    commentary: Commentary
    warnings: list[str]
    source_filename: str
    elapsed_ms: int = 0
    by_parent: list[Aggregate] = field(default_factory=list)  # empty when no rows declare a parent


# ---- AI commentary ---------------------------------------------------

SYSTEM_PROMPT = (
    "Senior financial controller drafting the monthly management commentary "
    "for the board pack. Write in first-person plural ('we', 'our'). Cite "
    "specific figures with the GBP symbol. Attribute drivers to named cost "
    "centres or categories. No hedging, no consultancy filler, material "
    "items first."
)


def get_ai_commentary(
    *,
    metadata: BudgetMetadata,
    rows: list[VarianceRow],
    by_cost_centre: list[Aggregate],
    by_category: list[Aggregate],
    by_period: list[Aggregate],
    headline: HeadlineStats,
    model: str | None = None,
    api_key: str | None = None,
    skip_ai: bool = False,
) -> Commentary:
    if skip_ai:
        return Commentary(skipped=True)

    user_prompt = _build_user_prompt(
        metadata=metadata, rows=rows,
        by_cost_centre=by_cost_centre, by_category=by_category, by_period=by_period,
        headline=headline,
    )

    system_payload = [{
        "type": "text",
        "text": SYSTEM_PROMPT + "\n\nSchema for your reply:\n" + _SCHEMA_HINT,
        "cache_control": {"type": "ephemeral"},
    }]

    try:
        data, stats = ask_claude_json_with_stats(
            user_prompt,
            system=system_payload,
            max_tokens=700,
            model=(model or CLAUDE_MODEL_FAST),
            api_key=api_key,
        )
    except Exception as e:
        return Commentary(
            headline=f"[AI commentary unavailable: {_scrub(e)}]",
            error=_scrub(e),
            model=model or CLAUDE_MODEL_FAST,
        )

    return Commentary(
        headline=str(data.get("headline") or "")[:280],
        summary=str(data.get("summary") or ""),
        adverse_drivers=[str(x) for x in (data.get("adverse_drivers") or [])][:5],
        favourable_drivers=[str(x) for x in (data.get("favourable_drivers") or [])][:5],
        actions=[str(x) for x in (data.get("actions") or [])][:5],
        cost_usd=stats.cost_usd,
        input_tokens=stats.input_tokens,
        output_tokens=stats.output_tokens,
        model=stats.model,
    )


_SCHEMA_HINT = (
    '{\n'
    '  "headline": "1 sentence verdict, cite the headline variance figure",\n'
    '  "summary":  "2 to 3 sentences in controller voice",\n'
    '  "adverse_drivers":    ["short bullet", "short bullet"],\n'
    '  "favourable_drivers": ["short bullet", "short bullet"],\n'
    '  "actions":            ["short imperative", "short imperative"]\n'
    '}'
)


def _build_user_prompt(
    *,
    metadata: BudgetMetadata,
    rows: list[VarianceRow],
    by_cost_centre: list[Aggregate],
    by_category: list[Aggregate],
    by_period: list[Aggregate],
    headline: HeadlineStats,
) -> str:
    lines: list[str] = []
    lines.append(
        f"Company: {metadata.company}  |  "
        f"Period: {metadata.period_label}  |  "
        f"Currency: {metadata.currency}"
    )
    lines.append("")

    lines.append("Headline:")
    lines.append("total_budget|total_actual|total_variance|variance_pct|red|amber|green|favourable|na")
    lines.append("|".join([
        _fmt_amount(headline.total_budget),
        _fmt_amount(headline.total_actual),
        _fmt_amount(headline.total_variance),
        _fmt_pct(headline.total_variance_pct),
        str(headline.rag_counts.get("red", 0)),
        str(headline.rag_counts.get("amber", 0)),
        str(headline.rag_counts.get("green", 0)),
        str(headline.rag_counts.get("favourable", 0)),
        str(headline.rag_counts.get("na", 0)),
    ]))
    lines.append("")

    lines.append("Top adverse (top 5 by absolute variance GBP):")
    lines.append("cost_centre|category|line_type|budget|actual|variance|variance_pct|rag")
    for r in top_n_by_variance(rows, n=5, side="adverse"):
        lines.append(_fmt_row_compact(r))
    lines.append("")

    lines.append("Top favourable (top 5 by absolute variance GBP):")
    lines.append("cost_centre|category|line_type|budget|actual|variance|variance_pct|rag")
    for r in top_n_by_variance(rows, n=5, side="favourable"):
        lines.append(_fmt_row_compact(r))
    lines.append("")

    lines.append("By cost centre:")
    lines.append("cost_centre|budget|actual|variance|variance_pct|rag_red|rag_amber|rag_green|rag_fav|rag_na")
    for a in by_cost_centre:
        lines.append(_fmt_agg_compact(a))
    lines.append("")

    lines.append("By category:")
    lines.append("category|budget|actual|variance|variance_pct|rag_red|rag_amber|rag_green|rag_fav|rag_na")
    for a in by_category:
        lines.append(_fmt_agg_compact(a))
    lines.append("")

    if len(by_period) > 1:
        lines.append("By period:")
        lines.append("period|budget|actual|variance|variance_pct|rag_red|rag_amber|rag_green|rag_fav|rag_na")
        for a in by_period:
            lines.append(_fmt_agg_compact(a))
        lines.append("")

    return "\n".join(lines)


def _fmt_amount(v: float) -> str:
    return f"{v:.2f}"


def _fmt_pct(v: float | None) -> str:
    if v is None:
        return "n/a"
    return f"{v * 100:.2f}%"


def _fmt_row_compact(r: VarianceRow) -> str:
    return "|".join([
        r.cost_centre, r.category, r.line_type,
        _fmt_amount(r.budget), _fmt_amount(r.actual),
        _fmt_amount(r.variance), _fmt_pct(r.variance_pct),
        r.rag,
    ])


def _fmt_agg_compact(a: Aggregate) -> str:
    return "|".join([
        a.key,
        _fmt_amount(a.budget), _fmt_amount(a.actual),
        _fmt_amount(a.variance), _fmt_pct(a.variance_pct),
        str(a.rag_counts.get("red", 0)),
        str(a.rag_counts.get("amber", 0)),
        str(a.rag_counts.get("green", 0)),
        str(a.rag_counts.get("favourable", 0)),
        str(a.rag_counts.get("na", 0)),
    ])


# ---- Top-level analyse() --------------------------------------------

def analyse(
    *,
    file_bytes: bytes | None = None,
    path: Path | str | None = None,
    source_filename: str = "",
    model: str | None = None,
    api_key: str | None = None,
    skip_ai: bool = False,
    history: HistoryStore | None = None,
    record_history: bool = True,
) -> AnalysisResult:
    """Run the full pipeline. Returns an AnalysisResult.

    When ``history`` is provided, score rows against the per-company history
    and (when ``record_history`` is true) append this run's variances to it.
    Comparison mode passes ``record_history=False`` to keep the history
    isolated to genuine analyses.
    """
    import time as _time
    started = _time.time()

    parsed = parse_inputs(file_bytes=file_bytes, path=path, source_filename=source_filename)
    rows = compute_all(parsed.rows)

    if history is not None:
        try:
            buckets = history.keys_with_enough_history(company=parsed.metadata.company)
            score_rows(rows, buckets)
        except Exception:
            # History is best-effort. A corrupted store should not break the run.
            pass

    by_cost_centre = aggregate_by(rows, dim="cost_centre")
    by_category = aggregate_by(rows, dim="category")
    by_period = aggregate_by(rows, dim="period")
    by_parent = aggregate_by_parent(rows)
    headline = headline_stats(rows)

    commentary = get_ai_commentary(
        metadata=parsed.metadata,
        rows=rows,
        by_cost_centre=by_cost_centre,
        by_category=by_category,
        by_period=by_period,
        headline=headline,
        model=model,
        api_key=api_key,
        skip_ai=skip_ai,
    )

    if history is not None and record_history:
        try:
            history.append_run(
                company=parsed.metadata.company,
                period_label=parsed.metadata.period_label,
                rows=rows,
            )
        except Exception:
            pass

    elapsed_ms = int((_time.time() - started) * 1000)

    return AnalysisResult(
        metadata=parsed.metadata,
        rows=rows,
        by_cost_centre=by_cost_centre,
        by_category=by_category,
        by_period=by_period,
        by_parent=by_parent,
        headline=headline,
        commentary=commentary,
        warnings=parsed.warnings,
        source_filename=source_filename,
        elapsed_ms=elapsed_ms,
    )


# ---- Result serialisation for the JSON API --------------------------

def to_dict(result: AnalysisResult) -> dict[str, Any]:
    return {
        "metadata": {
            "company": result.metadata.company,
            "currency": result.metadata.currency,
            "fiscal_year": result.metadata.fiscal_year,
            "period_label": result.metadata.period_label,
        },
        "rows": [_row_to_dict(r) for r in result.rows],
        "by_cost_centre": [_agg_to_dict(a) for a in result.by_cost_centre],
        "by_category":    [_agg_to_dict(a) for a in result.by_category],
        "by_period":      [_agg_to_dict(a) for a in result.by_period],
        "by_parent":      [_agg_to_dict(a) for a in result.by_parent],
        "headline":  _headline_to_dict(result.headline),
        "commentary": _commentary_to_dict(result.commentary),
        "warnings": list(result.warnings),
        "source_filename": result.source_filename,
        "elapsed_ms": result.elapsed_ms,
        "cost_usd": round(result.commentary.cost_usd, 6),
    }


def _row_to_dict(r: VarianceRow) -> dict[str, Any]:
    return {
        "period": r.period,
        "cost_centre": r.cost_centre,
        "category": r.category,
        "line_type": r.line_type,
        "budget": r.budget,
        "actual": r.actual,
        "variance": r.variance,
        "variance_pct": r.variance_pct,
        "tolerance_pct": r.tolerance_pct,
        "is_adverse": r.is_adverse,
        "rag": r.rag,
        "notes": r.notes,
        "forecast": r.forecast,
        "forecast_variance": r.forecast_variance,
        "forecast_variance_pct": r.forecast_variance_pct,
        "parent": r.parent,
        "z_score": r.z_score,
        "z_mean": r.z_mean,
        "z_stdev": r.z_stdev,
        "z_n": r.z_n,
    }


def _agg_to_dict(a: Aggregate) -> dict[str, Any]:
    return {
        "key": a.key,
        "budget": a.budget,
        "actual": a.actual,
        "variance": a.variance,
        "variance_pct": a.variance_pct,
        "rag_counts": dict(a.rag_counts),
        "row_count": a.row_count,
    }


def _headline_to_dict(h: HeadlineStats) -> dict[str, Any]:
    return {
        "total_budget": h.total_budget,
        "total_actual": h.total_actual,
        "total_variance": h.total_variance,
        "total_variance_pct": h.total_variance_pct,
        "rag_counts": dict(h.rag_counts),
        "adverse_count": h.adverse_count,
        "favourable_count": h.favourable_count,
        "na_count": h.na_count,
        "row_count": h.row_count,
        "period_count": h.period_count,
    }


def _commentary_to_dict(c: Commentary) -> dict[str, Any]:
    return {
        "headline": c.headline,
        "summary": c.summary,
        "adverse_drivers": list(c.adverse_drivers),
        "favourable_drivers": list(c.favourable_drivers),
        "actions": list(c.actions),
        "cost_usd": round(c.cost_usd, 6),
        "input_tokens": c.input_tokens,
        "output_tokens": c.output_tokens,
        "model": c.model,
        "skipped": c.skipped,
        "error": c.error,
    }


# ---- Helpers --------------------------------------------------------

def _scrub(e: Exception) -> str:
    """Strip absolute paths and API key fragments before showing to the UI."""
    msg = f"{type(e).__name__}: {e}"
    msg = re.sub(r"/[^\s'\"]+|[A-Z]:\\[^\s'\"]+", "<path>", msg)
    msg = re.sub(r"sk-ant-api\d+-[A-Za-z0-9_\-]+", "sk-ant-api03-***", msg)
    return msg[:300]
