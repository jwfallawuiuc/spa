using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace SubawardParser.Core;

public static class SubawardSpreadsheetParser
{
    private static readonly Regex SectionLetterRegex = new(@"^[H-Z]\.$", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    /// <summary>
    /// Parses every visible worksheet in the workbook and returns combined subaward lines.
    /// </summary>
    /// <param name="includeCostShare">
    /// When the section header row defines both <c>Sponsor</c> and <c>Cost Share</c> columns,
    /// adds the cost-share cell to the line amount; otherwise only the <c>Sponsor</c> column is used when present.
    /// </param>
    public static SpreadsheetSubawardResult ParseWorkbook(
        string filePath,
        bool includeCostShare = false,
        bool includeExempt = false)
    {
        var fileName = Path.GetFileName(filePath);
        var list = new List<SubawardEntry>();
        using var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var workbook = new XLWorkbook(stream, new LoadOptions { RecalculateAllFormulas = true });
        foreach (var worksheet in workbook.Worksheets.Where(w => w.Visibility == XLWorksheetVisibility.Visible))
            list.AddRange(ParseWorksheet(worksheet, includeCostShare, includeExempt));

        return new SpreadsheetSubawardResult(fileName, list);
    }

    public static FolderParseResult ParseFolder(
        string folderPath,
        bool includeCostShare = false,
        bool includeExempt = false)
    {
        if (!Directory.Exists(folderPath))
            throw new DirectoryNotFoundException($"Spreadsheet folder not found: {folderPath}");

        var files = Directory.EnumerateFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Where(p => !Path.GetFileName(p).StartsWith('~'))
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var perFile = new List<SpreadsheetSubawardResult>();
        var totals = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            var result = ParseWorkbook(file, includeCostShare, includeExempt);
            perFile.Add(result);
            foreach (var e in result.Entries)
            {
                totals.TryGetValue(e.SubrecipientName, out var sum);
                totals[e.SubrecipientName] = sum + e.Amount;
            }
        }

        var sortedTotals = totals.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

        return new FolderParseResult(perFile, sortedTotals);
    }

    private static IEnumerable<SubawardEntry> ParseWorksheet(
        IXLWorksheet worksheet,
        bool includeCostShare,
        bool includeExempt)
    {
        var used = worksheet.RangeUsed();
        if (used is null)
            yield break;

        var headerRow = FindGOtherDirectCostsRow(used);
        if (headerRow is null)
            yield break;

        var endRow = FindSectionEndRow(worksheet, headerRow.Value);
        var lastCol = used.LastColumn().ColumnNumber();
        var (sponsorCol, costShareCol) = FindSponsorAndCostShareColumns(worksheet, headerRow.Value, lastCol);

        for (var rowNum = headerRow.Value + 1; rowNum <= endRow; rowNum++)
        {
            var row = worksheet.Row(rowNum);
            if (!TryExtractSubawardName(row, lastCol, out var name))
                continue;

            var amount = GetLineAmount(row, lastCol, sponsorCol, costShareCol, includeCostShare);
            if (includeExempt && rowNum < endRow)
                amount += GetExemptAmountFromNextRow(worksheet.Row(rowNum + 1), lastCol);
            yield return new SubawardEntry(name, amount);
        }
    }

    private static decimal GetExemptAmountFromNextRow(IXLRow row, int lastCol)
    {
        for (var c = 1; c <= lastCol; c++)
        {
            var t = CellText(row.Cell(c));
            if (t.IndexOf("Exempt Subaward Cost", StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            for (var c2 = c + 1; c2 <= lastCol; c2++)
            {
                var amount = TryGetBudgetDecimal(row.Cell(c2));
                if (amount != 0m)
                    return amount;
            }

            return 0m;
        }

        return 0m;
    }

    private static (int? SponsorCol, int? CostShareCol) FindSponsorAndCostShareColumns(
        IXLWorksheet worksheet,
        int headerRow,
        int lastCol)
    {
        int? sponsorCol = null;
        int? costShareCol = null;
        var row = worksheet.Row(headerRow);
        for (var c = 1; c <= lastCol; c++)
        {
            var t = CellText(row.Cell(c)).Trim();
            if (t.Length == 0)
                continue;
            if (sponsorCol is null && t.Equals("Sponsor", StringComparison.OrdinalIgnoreCase))
                sponsorCol = c;
            else if (costShareCol is null && t.Equals("Cost Share", StringComparison.OrdinalIgnoreCase))
                costShareCol = c;
            if (sponsorCol is not null && costShareCol is not null)
                break;
        }

        return (sponsorCol, costShareCol);
    }

    /// <summary>
    /// Finds a cell containing "Subaward:" and builds the name from that cell plus following text cells until a numeric amount column.
    /// </summary>
    private static bool TryExtractSubawardName(IXLRow row, int lastCol, out string name)
    {
        name = string.Empty;
        for (var c = 1; c <= lastCol; c++)
        {
            var cell = row.Cell(c);
            var t = CellText(cell);
            if (t.IndexOf("subaward:", StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            var idx = t.IndexOf("subaward:", StringComparison.OrdinalIgnoreCase);
            var after = t[(idx + "subaward:".Length)..].Trim();
            var sb = new System.Text.StringBuilder();
            sb.Append(after);

            for (var c2 = c + 1; c2 <= lastCol; c2++)
            {
                var cell2 = row.Cell(c2);
                if (cell2.DataType == XLDataType.Number)
                    break;
                var piece = CellText(cell2).Trim();
                if (piece.Length == 0)
                    continue;
                sb.Append(piece);
            }

            name = sb.ToString().Trim();
            return name.Length > 0;
        }

        return false;
    }

    private static int? FindGOtherDirectCostsRow(IXLRange used)
    {
        foreach (var row in used.Rows())
        {
            var rowNumber = row.RowNumber();
            var a = CellText(row.Cell(1)).Trim();
            var b = CellText(row.Cell(2));

            if (a.Equals("G.", StringComparison.OrdinalIgnoreCase)
                && b.Contains("Other Direct Costs", StringComparison.OrdinalIgnoreCase))
                return rowNumber;

            foreach (var cell in row.Cells(1, Math.Min(8, used.LastColumn().ColumnNumber())))
            {
                var t = CellText(cell).Trim();
                if (t.Length == 0)
                    continue;
                if (t.StartsWith("Total", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (t.StartsWith("G.", StringComparison.OrdinalIgnoreCase)
                    && t.Contains("Other Direct Costs", StringComparison.OrdinalIgnoreCase))
                    return rowNumber;
            }
        }

        return null;
    }

    private static int FindSectionEndRow(IXLWorksheet worksheet, int sectionHeaderRow)
    {
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? sectionHeaderRow;
        var limit = Math.Min(lastRow, sectionHeaderRow + 250);

        for (var r = sectionHeaderRow + 1; r <= limit; r++)
        {
            var a = CellText(worksheet.Cell(r, 1)).Trim();
            if (SectionLetterRegex.IsMatch(a))
                return r - 1;
        }

        return limit;
    }

    private static string CellText(IXLCell cell)
    {
        if (cell.IsEmpty())
            return string.Empty;

        return cell.GetString();
    }

    /// <summary>
    /// When the G. Other Direct Costs header row includes a <c>Sponsor</c> column, that cell alone defines the
    /// amount unless both <c>Sponsor</c> and <c>Cost Share</c> exist and <paramref name="includeCostShare"/> is true,
    /// in which case the two cells are added. Otherwise uses the rightmost plausible dollar cell (legacy layout).
    /// </summary>
    private static decimal GetLineAmount(
        IXLRow row,
        int lastCol,
        int? sponsorCol,
        int? costShareCol,
        bool includeCostShare)
    {
        if (sponsorCol is int sc)
        {
            var sponsorAmt = TryGetBudgetDecimal(row.Cell(sc));
            if (includeCostShare && costShareCol is int csc)
                return sponsorAmt + TryGetBudgetDecimal(row.Cell(csc));
            return sponsorAmt;
        }

        decimal? rightmost = null;
        var rightmostCol = 0;
        decimal sum = 0;
        var counted = 0;

        for (var c = 1; c <= lastCol; c++)
        {
            var cell = row.Cell(c);
            if (!cell.TryGetValue(out double d))
                continue;

            if (IsLikelyRateNotDollars(d))
                continue;

            if (Math.Abs(d) < 1d)
                continue;

            var dec = (decimal)d;
            sum += dec;
            counted++;

            if (c >= rightmostCol)
            {
                rightmostCol = c;
                rightmost = dec;
            }
        }

        if (rightmost is > 0m)
            return rightmost.Value;

        // Some rows have no rightmost budget column populated; sum non-rate dollar cells instead.
        if (counted > 0)
            return sum;

        return 0m;
    }

    private static decimal TryGetBudgetDecimal(IXLCell cell)
    {
        if (!cell.TryGetValue(out double d))
            return 0m;
        if (IsLikelyRateNotDollars(d))
            return 0m;
        if (Math.Abs(d) < 1d)
            return 0m;
        return (decimal)d;
    }

    private static bool IsLikelyRateNotDollars(double value)
    {
        if (value == 0d)
            return false;
        if (value > 0d && value < 1d)
            return true;
        return false;
    }
}
