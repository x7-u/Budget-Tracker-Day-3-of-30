// Day 03 Budget Workbench. Power Query M template.
// Paste this into Power BI Desktop:
//   Home tab > Get data > Blank query > Advanced editor.
// Replace the OutputsFolder path below with the absolute path of the
// outputs/ folder inside your local clone of the repo.

let
    OutputsFolder = "C:\\path\\to\\day-03-budget-tracker\\outputs",
    AllFiles      = Folder.Files(OutputsFolder),
    BudgetFiles   = Table.SelectRows(
        AllFiles,
        each Text.StartsWith([Name], "budget_variance_") and Text.EndsWith([Name], ".xlsx")
    ),
    Sorted        = Table.Sort(BudgetFiles, {{"Date modified", Order.Descending}}),
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
