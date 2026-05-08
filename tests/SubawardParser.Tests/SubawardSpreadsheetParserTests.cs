using ClosedXML.Excel;
using SubawardParser.Core;

namespace SubawardParser.Tests;

public class SubawardSpreadsheetParserTests
{
    private static string Example1Path =>
        Path.Combine(AppContext.BaseDirectory, "Spreadsheets", "SubawardBudgetExample1.xlsx");

    [Fact]
    public void SubawardBudgetExample1_contains_expected_four_subrecipients()
    {
        Assert.True(File.Exists(Example1Path), $"Missing test file at {Example1Path}");

        var result = SubawardSpreadsheetParser.ParseWorkbook(Example1Path);
        var names = result.Entries.Select(e => e.SubrecipientName).ToHashSet(StringComparer.OrdinalIgnoreCase);

        Assert.Equal(4, result.Entries.Count);

        foreach (var expected in new[] { "Indiana", "Mayo", "Purdue", "Florida" })
            Assert.Contains(expected, names);
    }

    [Fact]
    public void When_sponsor_and_cost_share_headers_exist_flag_controls_whether_cost_share_is_added()
    {
        var path = Path.Combine(Path.GetTempPath(), $"subaward-sponsor-cs-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Budget");
                ws.Cell(5, 1).Value = "G.";
                ws.Cell(5, 2).Value = "Other Direct Costs";
                ws.Cell(5, 8).Value = "Sponsor";
                ws.Cell(5, 9).Value = "Cost Share";
                ws.Cell(6, 2).Value = "Subaward: TestOrg";
                ws.Cell(6, 8).Value = 100_000m;
                ws.Cell(6, 9).Value = 25_000m;
                ws.Cell(20, 1).Value = "H.";
                wb.SaveAs(path);
            }

            var sponsorOnly = SubawardSpreadsheetParser.ParseWorkbook(path, includeCostShare: false);
            var sponsorAndShare = SubawardSpreadsheetParser.ParseWorkbook(path, includeCostShare: true);

            Assert.Single(sponsorOnly.Entries);
            Assert.Single(sponsorAndShare.Entries);
            Assert.Equal(100_000m, sponsorOnly.Entries[0].Amount);
            Assert.Equal(125_000m, sponsorAndShare.Entries[0].Amount);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void When_only_sponsor_header_exists_flag_does_not_change_amount()
    {
        var path = Path.Combine(Path.GetTempPath(), $"subaward-sponsor-only-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Budget");
                ws.Cell(5, 1).Value = "G.";
                ws.Cell(5, 2).Value = "Other Direct Costs";
                ws.Cell(5, 8).Value = "Sponsor";
                ws.Cell(6, 2).Value = "Subaward: Solo";
                ws.Cell(6, 8).Value = 50_000m;
                ws.Cell(6, 9).Value = 10_000m; // no "Cost Share" header — legacy extra number; sponsor column wins
                ws.Cell(20, 1).Value = "H.";
                wb.SaveAs(path);
            }

            var a = SubawardSpreadsheetParser.ParseWorkbook(path, includeCostShare: false);
            var b = SubawardSpreadsheetParser.ParseWorkbook(path, includeCostShare: true);

            Assert.Equal(50_000m, a.Entries[0].Amount);
            Assert.Equal(50_000m, b.Entries[0].Amount);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Include_exempt_adds_exempt_subaward_cost_from_immediately_following_row()
    {
        var path = Path.Combine(Path.GetTempPath(), $"subaward-exempt-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Budget");
                ws.Cell(5, 1).Value = "G.";
                ws.Cell(5, 2).Value = "Other Direct Costs";
                ws.Cell(5, 8).Value = "Sponsor";
                ws.Cell(6, 2).Value = "Subaward: ExemptOrg";
                ws.Cell(6, 8).Value = 100_000m;
                ws.Cell(7, 2).Value = "Exempt Subaward Cost";
                ws.Cell(7, 8).Value = 25_000m;
                ws.Cell(20, 1).Value = "H.";
                wb.SaveAs(path);
            }

            var withoutExempt = SubawardSpreadsheetParser.ParseWorkbook(path, includeExempt: false);
            var withExempt = SubawardSpreadsheetParser.ParseWorkbook(path, includeExempt: true);

            Assert.Single(withoutExempt.Entries);
            Assert.Single(withExempt.Entries);
            Assert.Equal(100_000m, withoutExempt.Entries[0].Amount);
            Assert.Equal(125_000m, withExempt.Entries[0].Amount);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Fact]
    public void Include_exempt_does_not_add_when_exempt_row_is_not_immediately_below_subrecipient()
    {
        var path = Path.Combine(Path.GetTempPath(), $"subaward-exempt-gap-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Budget");
                ws.Cell(5, 1).Value = "G.";
                ws.Cell(5, 2).Value = "Other Direct Costs";
                ws.Cell(5, 8).Value = "Sponsor";
                ws.Cell(6, 2).Value = "Subaward: GapOrg";
                ws.Cell(6, 8).Value = 100_000m;
                ws.Cell(7, 2).Value = "Other line";
                ws.Cell(8, 2).Value = "Exempt Subaward Cost";
                ws.Cell(8, 8).Value = 25_000m;
                ws.Cell(20, 1).Value = "H.";
                wb.SaveAs(path);
            }

            var withExempt = SubawardSpreadsheetParser.ParseWorkbook(path, includeExempt: true);

            Assert.Single(withExempt.Entries);
            Assert.Equal(100_000m, withExempt.Entries[0].Amount);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}
