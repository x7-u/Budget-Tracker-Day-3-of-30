"""Day 3. CLI entry for the budget vs actual tracker.

Usage:
    python main.py --input sample_data/sample_general.xlsx
    python main.py --input ./my_budget.xlsx --no-ai
    python main.py --input ./my_budget.xlsx --model claude-sonnet-4-6

Outputs land in ./outputs/. Skipping AI is cheap and offline.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
for p in (str(HERE), str(ROOT)):
    if p not in sys.path:
        sys.path.insert(0, p)

from csv_writer import write_csv  # noqa: E402
from excel_writer import write_workbook  # noqa: E402
from pipeline import analyse  # noqa: E402

OUTPUTS = HERE / "outputs"


def main() -> int:
    parser = argparse.ArgumentParser(description="Budget vs Actual Tracker (Day 3)")
    parser.add_argument("--input", "-i", required=True, help="Path to a budget vs actual .xlsx workbook")
    parser.add_argument("--no-ai", action="store_true", help="Skip the Claude commentary call")
    parser.add_argument("--model", default=None,
                        help="Override the default model (e.g. claude-sonnet-4-6)")
    parser.add_argument("--api-key", default=None,
                        help="Override the ANTHROPIC_API_KEY for this run only (not persisted)")
    args = parser.parse_args()

    in_path = Path(args.input).expanduser().resolve()
    if not in_path.is_file():
        print(f"Input not found: {in_path}", file=sys.stderr)
        return 2
    if in_path.suffix.lower() != ".xlsx":
        print(f"Input must be an .xlsx workbook (got {in_path.suffix}).", file=sys.stderr)
        return 2

    try:
        result = analyse(
            path=in_path,
            source_filename=in_path.name,
            model=args.model,
            api_key=args.api_key,
            skip_ai=args.no_ai,
        )
    except ValueError as e:
        print(f"Validation error: {e}", file=sys.stderr)
        return 1

    xlsx_path = write_workbook(result, OUTPUTS)
    csv_path = write_csv(result, OUTPUTS / "budget_variance.csv")

    print()
    print(f"  Company: {result.metadata.company}")
    print(f"  Period:  {result.metadata.period_label}")
    print(f"  Rows:    {len(result.rows)}")
    print(f"  Budget:  GBP {result.headline.total_budget:,.2f}")
    print(f"  Actual:  GBP {result.headline.total_actual:,.2f}")
    print(f"  Var:     GBP {result.headline.total_variance:,.2f} "
          f"({_pct(result.headline.total_variance_pct)})")
    print(f"  Counts:  {result.headline.rag_counts}")
    if result.commentary.skipped:
        print("  AI:      skipped")
    elif result.commentary.error:
        print(f"  AI:      error. {result.commentary.error}")
    else:
        print(f"  AI:      ${result.commentary.cost_usd:.4f}, "
              f"{result.commentary.input_tokens} in / "
              f"{result.commentary.output_tokens} out, {result.commentary.model}")
        print(f"  Headline: {result.commentary.headline}")
    if result.warnings:
        print(f"  Warnings: {len(result.warnings)}")
        for w in result.warnings[:5]:
            print(f"    - {w}")
        if len(result.warnings) > 5:
            print(f"    ...and {len(result.warnings) - 5} more")
    print()
    print(f"  Wrote: {xlsx_path.name}")
    print(f"         {csv_path.name}")
    return 0


def _pct(v: float | None) -> str:
    if v is None:
        return "n/a"
    return f"{v * 100:.2f}%"


if __name__ == "__main__":
    raise SystemExit(main())
