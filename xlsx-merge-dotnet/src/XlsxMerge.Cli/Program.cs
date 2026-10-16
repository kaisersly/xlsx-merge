using System.CommandLine;
using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using XlsxMerge;
using XlsxMerge.IO;

var templateOption = new Option<string>("--template", "Template .xlsx file")
{
    IsRequired = true,
};
templateOption.AddAlias("-t");

var dataOption = new Option<string>("--data", "Data file ([globals] / [rows] sections)")
{
    IsRequired = true,
};
dataOption.AddAlias("-d");

var outOption = new Option<string>("--out", "Output .xlsx file")
{
    IsRequired = true,
};
outOption.AddAlias("-o");

var cultureOption = new Option<string>(
    "--culture",
    () => "fr-FR",
    "Culture used to format numbers in text (e.g. fr-FR, en-US)");

var markerOption = new Option<string>(
    "--marker",
    () => string.Empty,
    "Marker field of the repeatable block, e.g. EMPLOYEE_ID (automatic detection otherwise)");

var blockOption = new Option<string>(
    "--block",
    () => "auto",
    "Repeatable block: auto (detection), 1 (marker row only) or 2 (with style row)");

var numbersOption = new Option<string>(
    "--numbers",
    () => "strict",
    "Numeric string conversion: strict or none");

var delimiterOption = new Option<string>(
    "--delimiter",
    () => ";",
    "Data file delimiter");

var renderCommand = new Command(
    "render",
    "Mail-merges an .xlsx template with a data file (all sheets, headers and footers)")
{
    templateOption, dataOption, outOption, cultureOption, markerOption, blockOption, numbersOption,
    delimiterOption,
};

renderCommand.SetHandler(
    (string template, string data, string output, string culture, string marker, string block,
        string numbers, string delimiter) =>
    {
        try
        {
            var options = new MergeOptions
            {
                NumberCulture = new CultureInfo(culture),
                RepeatBlockRows = ParseBlock(block),
                NumericTextParsing = ParseNumbers(numbers),
                RowMarkerField = marker.Length == 0 ? null : marker,
            };
            var mergeData = CsvMergeDataReader.ReadFile(data, delimiter);
            XlsxMerger.RenderFile(template, output, mergeData, options);
            Console.WriteLine($"Output file: {output}");
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"Error: {exception.Message}");
            Environment.ExitCode = 1;
        }
    },
    templateOption,
    dataOption,
    outOption,
    cultureOption,
    markerOption,
    blockOption,
    numbersOption,
    delimiterOption);

var inspectOption = new Option<string>("--template", "Template .xlsx file")
{
    IsRequired = true,
};
inspectOption.AddAlias("-t");

var inspectCommand = new Command("inspect", "Lists the merge fields of an .xlsx template")
{
    inspectOption,
};

inspectCommand.SetHandler(
    (string template) =>
    {
        try
        {
            Inspect(template);
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"Error: {exception.Message}");
            Environment.ExitCode = 1;
        }
    },
    inspectOption);

var rootCommand = new RootCommand(
    "xlsx-merge — mail-merge of Excel files (.xlsx): template + data set")
{
    renderCommand,
    inspectCommand,
};

return await rootCommand.InvokeAsync(args);

static RepeatBlockRows ParseBlock(string block) => block.ToLowerInvariant() switch
{
    "auto" => RepeatBlockRows.Auto,
    "1" => RepeatBlockRows.One,
    "2" => RepeatBlockRows.Two,
    _ => throw new MergeException(
        $"Invalid value for --block: '{block}' (expected: auto, 1 or 2)."),
};

static NumericTextParsing ParseNumbers(string numbers) => numbers.ToLowerInvariant() switch
{
    "strict" => NumericTextParsing.Strict,
    "none" => NumericTextParsing.None,
    _ => throw new MergeException(
        $"Invalid value for --numbers: '{numbers}' (expected: strict or none)."),
};

static void Inspect(string template)
{
    var occurrences = new[]
    {
        XLHFOccurrence.OddPages, XLHFOccurrence.EvenPages, XLHFOccurrence.FirstPage,
    };
    var usedOptions = XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats;

    using var workbook = new XLWorkbook(File.OpenRead(template));
    var found = new Dictionary<string, List<string>>(StringComparer.Ordinal);

    void Record(string field, string location)
    {
        if (!found.TryGetValue(field, out var locations))
            found[field] = locations = new List<string>();
        if (!locations.Contains(location))
            locations.Add(location);
    }

    foreach (var worksheet in workbook.Worksheets)
    {
        foreach (var row in worksheet.RowsUsed(usedOptions))
        {
            foreach (var cell in row.CellsUsed(usedOptions))
            {
                if (cell.HasFormula || !cell.Value.IsText)
                    continue;
                foreach (Match match in MergeFields.Pattern.Matches(cell.GetString()))
                    Record(match.Groups[1].Value.Trim(), $"{worksheet.Name}!{cell.Address}");
            }
        }

        foreach (var (headerFooter, kind) in new (IXLHeaderFooter, string)[]
                 {
                     (worksheet.PageSetup.Header, "header"),
                     (worksheet.PageSetup.Footer, "footer"),
                 })
        {
            foreach (var (item, part) in new (IXLHFItem, string)[]
                     {
                         (headerFooter.Left, "left"),
                         (headerFooter.Center, "center"),
                         (headerFooter.Right, "right"),
                     })
            {
                foreach (var occurrence in occurrences)
                {
                    var text = item.GetText(occurrence);
                    foreach (Match match in MergeFields.Pattern.Matches(text))
                        Record(match.Groups[1].Value.Trim(),
                            $"{worksheet.Name} ! {kind} {part} ({Describe(occurrence)})");
                }
            }
        }
    }

    if (found.Count == 0)
    {
        Console.WriteLine("No merge field found.");
        return;
    }

    Console.WriteLine($"{found.Count} merge field(s):");
    foreach (var field in found.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        Console.WriteLine($"  #{field.Key}#  ({field.Value.Count} location(s)) : {field.Value[0]}");
    Console.WriteLine();
    Console.WriteLine("Tip: the repeatable block is the row holding the most row fields.");
}

static string Describe(XLHFOccurrence occurrence) => occurrence switch
{
    XLHFOccurrence.OddPages => "odd",
    XLHFOccurrence.EvenPages => "even",
    XLHFOccurrence.FirstPage => "first page",
    _ => occurrence.ToString(),
};
