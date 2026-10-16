using System.Reflection;
using ClosedXML.Excel;
using Xunit;
using Xunit.Abstractions;

namespace XlsxMerge.Tests;

/// <summary>
/// Development helper (not a behaviour test): dumps the ClosedXML API surface used by the
/// library. Useful when upgrading ClosedXML to spot renamed/removed members.
/// </summary>
public class ClosedXmlApiDumpTests
{
    private readonly ITestOutputHelper _output;

    public ClosedXmlApiDumpTests(ITestOutputHelper output) => _output = output;

    private static void DumpMembers(ITestOutputHelper output, Type type)
    {
        output.WriteLine($"=== {type.FullName} ===");
        foreach (var m in type.GetMembers(
                     BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static
                     | BindingFlags.DeclaredOnly))
            output.WriteLine($"  {m.MemberType}: {m}");
    }

    [Fact]
    public void Dump_relevant_api()
    {
        DumpMembers(_output, typeof(IXLRichText));
        DumpMembers(_output, typeof(IXLFormattedText<IXLRichText>));
        DumpMembers(_output, typeof(IXLCell));
        _output.WriteLine("=== XLCellsUsedOptions ===");
        foreach (var value in Enum.GetValues<XLCellsUsedOptions>())
            _output.WriteLine($"  {value}");
        _output.WriteLine("=== IXLRange (row/column accessors) ===");
        foreach (var m in typeof(IXLRange).GetMethods(
                     BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            if (m.Name.Contains("Row") || m.Name.Contains("Column") || m.Name.Contains("Address")
                || m.Name.Contains("Merge"))
                _output.WriteLine($"  {m}");

        DumpMembers(_output, typeof(XLCellValue));
        DumpMembers(_output, typeof(IXLRow));
        DumpMembers(_output, typeof(IXLHFItem));
        DumpMembers(_output, typeof(IXLHeaderFooter));

        // Compile checks: these conversions/assignments must exist.
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("s");
        var cell = ws.Cell(1, 1);
        cell.Value = "text";
        cell.Value = 12.5;
        cell.Value = 42;
        cell.Value = true;
        cell.Value = new DateTime(2024, 1, 1);
        var rich = cell.GetRichText();
        _output.WriteLine($"rich.Count = {rich.Count}, rich.Text = {rich.Text}");
        wb.Dispose();
    }
}
