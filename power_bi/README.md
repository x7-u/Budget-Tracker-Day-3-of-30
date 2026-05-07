# Power BI integration

`power_query.m` is a Power Query M template you can paste into
Power BI Desktop (Home > Get data > Blank query > Advanced editor).
Replace the `OutputsFolder` path with the absolute path to the
`outputs/` folder inside your local clone, then click Done.

The query selects the most recent `budget_variance_*.xlsx` file in
that folder, opens its `Detail` sheet, types the columns, and adds
the run timestamp + source workbook columns so you can refresh the
Power BI report after every model run.
