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
}
