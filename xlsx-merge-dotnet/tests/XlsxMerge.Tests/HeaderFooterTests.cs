using ClosedXML.Excel;
using Xunit;
using Xunit.Abstractions;

namespace XlsxMerge.Tests;

public class HeaderFooterTests
{
    private readonly ITestOutputHelper _output;

    public HeaderFooterTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void Headers_and_footers_replaced_per_occurrence()
    {
        var template = new XLWorkbook();
        var sheet = template.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "#ID#";

        sheet.PageSetup.SetDifferentFirstPageOnHF(true);
        sheet.PageSetup.SetDifferentOddEvenPagesOnHF(true);
        sheet.PageSetup.Header.Left.AddText("#COURSE#", XLHFOccurrence.OddPages);
        sheet.PageSetup.Header.Right.AddText("#YEAR# #TERM#", XLHFOccurrence.OddPages);
        sheet.PageSetup.Header.Left.AddText("First #COURSE#", XLHFOccurrence.FirstPage);
        sheet.PageSetup.Footer.Center.AddText("Page #TERM#", XLHFOccurrence.EvenPages);
        sheet.PageSetup.Footer.Left.AddText("Unknown #ZORBLE#", XLHFOccurrence.OddPages);

        var data = new MergeData()
            .AddGlobal("COURSE", "Excel")
            .AddGlobal("YEAR", "2024 / 2025")
            .AddGlobal("TERM", "T1")
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var merged = workbook.Worksheet("Sheet1");

        Assert.Contains("Excel", merged.PageSetup.Header.Left.GetText(XLHFOccurrence.OddPages));
        Assert.DoesNotContain("#COURSE#", merged.PageSetup.Header.Left.GetText(XLHFOccurrence.OddPages));
        Assert.Contains("2024 / 2025 T1", merged.PageSetup.Header.Right.GetText(XLHFOccurrence.OddPages));
        Assert.Contains("First Excel", merged.PageSetup.Header.Left.GetText(XLHFOccurrence.FirstPage));
        Assert.Contains("Page T1", merged.PageSetup.Footer.Center.GetText(XLHFOccurrence.EvenPages));

        // Unknown field: kept as-is (as in the JS implementation).
        Assert.Contains("#ZORBLE#", merged.PageSetup.Footer.Left.GetText(XLHFOccurrence.OddPages));

        // The "different page" flags are preserved.
        Assert.True(merged.PageSetup.DifferentFirstPageOnHF);
        Assert.True(merged.PageSetup.DifferentOddEvenPagesOnHF);
    }

    [Fact]
    public void Formatted_header_keeps_at_least_the_text()
    {
        var template = new XLWorkbook();
        var sheet = template.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "#ID#";
        sheet.PageSetup.Header.Left.AddText("#COURSE#", XLHFOccurrence.OddPages).SetBold();
        sheet.PageSetup.Footer.Left.AddText("&BBold #COURSE#", XLHFOccurrence.OddPages);

        var data = new MergeData()
            .AddGlobal("COURSE", "Excel")
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var merged = workbook.Worksheet("Sheet1");

        var headerText = merged.PageSetup.Header.Left.GetText(XLHFOccurrence.OddPages);
        var footerText = merged.PageSetup.Footer.Left.GetText(XLHFOccurrence.OddPages);
        _output.WriteLine($"header: {headerText}");
        _output.WriteLine($"footer: {footerText}");

        Assert.Contains("Excel", headerText);
        Assert.Contains("Excel", footerText);
    }

    [Fact]
    public void Unused_occurrences_are_processed_too()
    {
        var template = new XLWorkbook();
        var sheet = template.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "#ID#";
        // differentOddEven disabled: the even header exists in the file but is not displayed.
        sheet.PageSetup.Header.Left.AddText("#COURSE#", XLHFOccurrence.EvenPages);

        var data = new MergeData()
            .AddGlobal("COURSE", "Excel")
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var merged = workbook.Worksheet("Sheet1");

        Assert.DoesNotContain("#COURSE#", merged.PageSetup.Header.Left.GetText(XLHFOccurrence.EvenPages));
    }
}
