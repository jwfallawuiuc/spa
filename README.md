# Subaward spreadsheet parser

Small .NET **9** tool that reads every `.xlsx` in the **`Spreadsheets`** folder at the repository root, lists subrecipients under **G. Other Direct Costs** (lines that begin with **Subaward:**), then prints a combined total per subrecipient across all files.

## Prerequisites

- [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0) (or a newer SDK that can build `net9.0` projects).

If you only have a **newer** runtime (for example .NET 10) and not the .NET 9 runtime, the console app and tests still run because the projects set `RollForward` to `LatestMajor`. For the intended environment, install the .NET 9 runtime so roll-forward is unnecessary.

## How to run

From the repository root (the folder that contains `Spreadsheets` and `SubawardParser.sln`):

```powershell
dotnet restore
dotnet run --project src/SubawardParser/SubawardParser.csproj
```

The program looks for `.\Spreadsheets` using the current directory first, then walks upward from the application folder. To use another folder:

```powershell
dotnet run --project src/SubawardParser/SubawardParser.csproj -- --folder "D:\MyBudgets"
```

## How to run tests

```powershell
dotnet test SubawardParser.sln
```

Sample workbooks under `Spreadsheets\` are copied into the test output so tests pass on a clean clone without extra steps.

## Publishing to GitHub (for reviewers)

1. Create a new **public** empty repository on GitHub.
2. From this folder:

   ```powershell
   git init
   git add .
   git commit -m "Add subaward budget parser and tests"
   git branch -M main
   git remote add origin https://github.com/<your-account>/<your-repo>.git
   git push -u origin main
   ```

Reviewers clone the repo, install .NET 9 (or newer per above), and run the same `dotnet run` / `dotnet test` commands.

---

## Assumptions (ambiguous requirements)

1. **Workbook layout:** The **G. Other Direct Costs** block is identified when column **A** is `G.` and column **B** contains **Other Direct Costs**, or when a cell in the first columns contains text starting with **G.** and including **Other Direct Costs** (excluding rows that start with **Total**).
2. **Section end:** Parsing stops when column **A** matches a single letter plus dot in the range **H.** through **Z.** (next major budget section). If that never appears within 250 rows, scanning stops at that row limit.
3. **Subaward label:** A subaward line is detected when a cell contains **Subaward:** (case-insensitive). The subrecipient name is the text after that prefix in the same cell, plus any following **text** cells on the same row until the first **numeric** cell (amount columns).
4. **Amount on the row:** The tool prefers the **rightmost** numeric value on the row that is not treated as a **rate** (values strictly between 0 and 1 are skipped, to avoid F&A rate columns). If no such rightmost value is greater than zero, it **sums** all remaining non-rate numbers with absolute value at least 1 on that row (helps when the rightmost budget column is blank but period columns are filled).
5. **Formulas:** Workbooks are opened with **RecalculateAllFormulas** so ClosedXML can evaluate formula cells where possible.
6. **File lock:** Files are opened with **read-sharing** so Windows can still read a workbook that is open in Excel (best effort).
7. **File types:** Only `.xlsx` in the top level of `Spreadsheets` is processed; temporary Excel files like `~$*.xlsx` are ignored.
8. **Worksheets:** Every **visible** worksheet in each workbook is scanned (not only the first sheet).

## Questions we would have asked on a real project

1. Should **subrecipient names** be normalized to a canonical registry (DUNS, UEI, legal name) when the same organization appears with different spellings?
2. For multi-year columns, should the “subaward amount” be the **sum of budget periods**, a specific **sponsor year**, or a **total** column only? (The template can make any of these plausible.)
3. Are subawards ever split across **multiple rows** for one recipient, and should those be merged?
4. Should **hidden sheets** or **scenario / what-if** versions of the workbook be included or explicitly excluded?
5. Is **currency and rounding** always USD to two decimals, or should we respect the workbook locale and number formats?
6. Who owns the **official template version** so we can regression-test layout changes?
