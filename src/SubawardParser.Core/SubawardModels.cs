namespace SubawardParser.Core;

public sealed record SubawardEntry(string SubrecipientName, decimal Amount);

public sealed record SpreadsheetSubawardResult(string FileName, IReadOnlyList<SubawardEntry> Entries);

public sealed record FolderParseResult(
    IReadOnlyList<SpreadsheetSubawardResult> PerFile,
    IReadOnlyDictionary<string, decimal> TotalsBySubrecipient);
