using System.Globalization;
using SubawardParser.Core;

try
{
    var includeCostShare = args.Any(static a =>
        a.Equals("--include-cost-share", StringComparison.OrdinalIgnoreCase));
    var includeExempt = args.Any(static a =>
        a.Equals("--include-exempt", StringComparison.OrdinalIgnoreCase));
    var folder = SpreadsheetFolderResolver.Resolve(args);
    var result = SubawardSpreadsheetParser.ParseFolder(folder, includeCostShare, includeExempt);
    ConsoleReport.Write(result, folder);
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    return 1;
}

internal static class SpreadsheetFolderResolver
{
    internal static string Resolve(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            if (!args[i].Equals("--folder", StringComparison.OrdinalIgnoreCase))
                continue;
            if (i + 1 >= args.Length)
                throw new ArgumentException("Missing path after --folder.");
            return Path.GetFullPath(args[i + 1]);
        }

        var cwdCandidate = Path.Combine(Directory.GetCurrentDirectory(), "Spreadsheets");
        if (Directory.Exists(cwdCandidate))
            return cwdCandidate;

        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "Spreadsheets");
            if (Directory.Exists(candidate))
                return candidate;
            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not find a \"Spreadsheets\" folder. Run from the repository root, or pass --folder <path>.");
    }
}

internal static class ConsoleReport
{
    private static readonly CultureInfo MoneyCulture = CultureInfo.GetCultureInfo("en-US");

    internal static void Write(FolderParseResult result, string folderPath)
    {
        Console.WriteLine();
        Console.WriteLine("  ========================================");
        Console.WriteLine("   Subaward summary (by spreadsheet)");
        Console.WriteLine("  ========================================");
        Console.WriteLine();
        Console.WriteLine($"  Folder: {folderPath}");
        Console.WriteLine($"  Excel files read: {result.PerFile.Count}");
        Console.WriteLine();

        foreach (var file in result.PerFile)
        {
            Console.WriteLine($"  --- {file.FileName} ---");
            if (file.Entries.Count == 0)
            {
                Console.WriteLine("    (No subaward lines were found under \"G. Other Direct Costs\".)");
            }
            else
            {
                Console.WriteLine("    Subrecipients:");
                var n = 1;
                foreach (var e in file.Entries)
                    Console.WriteLine($"      {n++,3}. {e.SubrecipientName}");
            }

            Console.WriteLine();
        }

        Console.WriteLine("  ========================================");
        Console.WriteLine("   Totals across all spreadsheets");
        Console.WriteLine("  ========================================");
        Console.WriteLine();
        Console.WriteLine("  Each line is one subrecipient and the total subaward");
        Console.WriteLine("  amount we found for them in every file combined.");
        Console.WriteLine();

        if (result.TotalsBySubrecipient.Count == 0)
        {
            Console.WriteLine("    (No subaward amounts to show.)");
        }
        else
        {
            var width = result.TotalsBySubrecipient.Keys.Max(k => k.Length);
            foreach (var (name, total) in result.TotalsBySubrecipient)
            {
                var padded = name.PadRight(width);
                Console.WriteLine($"    {padded}   {total.ToString("C2", MoneyCulture)}");
            }
        }

        Console.WriteLine();
    }
}
