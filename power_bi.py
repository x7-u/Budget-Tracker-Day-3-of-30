"""Day 3. Power BI scaffolding.

Power BI's binary ``.pbix`` format isn't practical to generate without
the desktop tooling. The lightweight, honest version is to ship a
Power Query M script that the controller pastes into Power BI Desktop
(Get data > Blank query > Advanced editor) and a tiny README that
explains the import workflow.

The generated script reads every ``budget_variance_*.xlsx`` file in
the ``outputs/`` folder, picks the most recent, and projects the
Detail sheet into a clean tabular query. Because it points at a
folder rather than a single workbook, refreshing inside Power BI
just picks up the latest run automatically.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

POWER_QUERY_TEMPLATE = """\
// Day 03 Budget Workbench. Power Query M template.
// Paste this into Power BI Desktop:
//   Home tab > Get data > Blank query > Advanced editor.
// Replace the OutputsFolder path below if your repo lives elsewhere.

let
    OutputsFolder = "{outputs_folder}",
    AllFiles      = Folder.Files(OutputsFolder),
    BudgetFiles   = Table.SelectRows(
        AllFiles,
        each Text.StartsWith([Name], "budget_variance_") and Text.EndsWith([Name], ".xlsx")
    ),
    Sorted        = Table.Sort(BudgetFiles, {{{{"Date modified", Order.Descending}}}}),
    LatestRow     = Table.First(Sorted),
    LatestPath    = LatestRow[Folder Path] & LatestRow[Name],
    BookContents  = Excel.Workbook(File.Contents(LatestPath), null, true),
    DetailSheet   = BookContents{{[Item="Detail", Kind="Sheet"]}}[Data],
    Promoted      = Table.PromoteHeaders(DetailSheet, [PromoteAllScalars=true]),
    Typed         = Table.TransformColumnTypes(Promoted, {{
        {{"Period",      type text}},
        {{"Cost Centre", type text}},
        {{"Category",    type text}},
        {{"Line type",   type text}},
        {{"Budget",      type number}},
        {{"Actual",      type number}},
        {{"Variance",    type number}},
        {{"Variance %",  type number}},
        {{"Tolerance %", type number}},
        {{"Adverse",     type text}},
        {{"RAG",         type text}},
        {{"Notes",       type text}}
    }}),
    AddedRunStamp = Table.AddColumn(Typed, "Run timestamp", each LatestRow[Date modified], type datetime),
    AddedSource   = Table.AddColumn(AddedRunStamp, "Source workbook", each LatestRow[Name], type text)
in
    AddedSource
"""

POWER_BI_README = """\
# Power BI import for Day 03 Budget Workbench

Power BI Desktop is a Microsoft application; the file format ``.pbix`` cannot
be authored on the command line in a way that round-trips through Power BI's
own engine. Instead this folder ships a Power Query M script that you paste
into Power BI Desktop. After the first paste the workflow is one click to
refresh.

## One-time setup

1. Open Power BI Desktop.
2. Home tab. Get data. Blank query. Open Advanced editor.
3. Paste the contents of ``power_query.m`` from this folder. Edit the
   ``OutputsFolder`` line near the top so it points at your local copy of
   ``day-03-budget-tracker/outputs``.
4. Done. Click Apply to load the latest Detail rows.

The query reads every ``budget_variance_*.xlsx`` in the outputs folder and
picks the most recent. Each refresh re-discovers the latest pack, so you
don't need to re-paste when you re-run the analyser.

## What you get

A single table called ``Detail`` with the columns Period, Cost Centre,
Category, Line type, Budget, Actual, Variance, Variance %, Tolerance %,
Adverse, RAG, Notes plus the run timestamp and source workbook name.

Build slicers on Cost Centre, Period, Category and RAG. Use the Variance
column for measures. The Adverse column is text (Yes / No) so it works as a
slicer or a filter directly.

## Why a script and not a .pbix

The binary ``.pbix`` format embeds Power BI's data model, visuals and
connection metadata, and is not stable across versions. Hand-rolling one
from Python would be brittle and would break the moment you opened it.
A Power Query script is portable, version-controllable as text, and gives
you the same end result with one paste.
"""


def write_power_bi_assets(out_dir: Path | str) -> Path:
    """Write the M script and the import guide. Returns the script path."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    outputs_folder = (out_dir.parent / "outputs").resolve()
    # Power Query expects double backslashes inside a quoted string for Windows paths.
    safe_path = str(outputs_folder).replace("\\", "\\\\")

    script_path = out_dir / "power_query.m"
    readme_path = out_dir / "README.md"

    rendered = POWER_QUERY_TEMPLATE.format(outputs_folder=safe_path)
    # Stamp the generation time at the top so the user knows when it was last refreshed.
    header = f"// Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} by power_bi.py\n"
    script_path.write_text(header + rendered, encoding="utf-8")
    readme_path.write_text(POWER_BI_README, encoding="utf-8")
    return script_path


def main() -> None:
    here = Path(__file__).resolve().parent
    out = write_power_bi_assets(here / "power_bi")
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
