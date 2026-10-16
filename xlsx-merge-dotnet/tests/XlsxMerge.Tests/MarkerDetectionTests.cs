using ClosedXML.Excel;
using Xunit;

namespace XlsxMerge.Tests;

public class MarkerDetectionTests
{
    private static XLWorkbook SheetWith(params Action<IXLWorksheet>[] rows)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        foreach (var build in rows)
            build(sheet);
        return workbook;
    }

    private static readonly IReadOnlyList<string> RowFields = new[] { "ID", "NAME" };

    [Fact]
    public void Row_with_the_most_fields_wins()
    {
        var template = SheetWith(
            sheet => sheet.Cell(3, 1).Value = "#ID# #NAME#",
            sheet => sheet.Cell(7, 1).Value = "#ID#");

        var marker = RowExpander.DetectMarkerRow(template.Worksheet("Sheet1"), RowFields, new MergeOptions());
        Assert.Equal(3, marker);
    }

    [Fact]
    public void Tie_between_rows_raises_an_explicit_error()
    {
        var template = SheetWith(
            sheet => sheet.Cell(3, 1).Value = "#ID#",
            sheet => sheet.Cell(7, 1).Value = "#NAME#");

        var exception = Assert.Throws<MergeException>(() =>
            RowExpander.DetectMarkerRow(template.Worksheet("Sheet1"), RowFields, new MergeOptions()));
        Assert.Contains("candidate rows", exception.Message);
        Assert.Contains("--marker", exception.Message);
    }

    [Fact]
    public void Forced_marker_field_absent()
    {
        var template = SheetWith(sheet => sheet.Cell(3, 1).Value = "#NAME#");
        var options = new MergeOptions { RowMarkerField = "ID" };

        var exception = Assert.Throws<MergeException>(() =>
            RowExpander.DetectMarkerRow(template.Worksheet("Sheet1"), RowFields, options));
        Assert.Contains("missing a row containing #ID#", exception.Message);
    }

    [Fact]
    public void Forced_marker_field_present_twice()
    {
        var template = SheetWith(
            sheet => sheet.Cell(3, 1).Value = "#ID#",
            sheet => sheet.Cell(9, 1).Value = "#ID# #NAME#");
        var options = new MergeOptions { RowMarkerField = "#ID#" };

        var exception = Assert.Throws<MergeException>(() =>
            RowExpander.DetectMarkerRow(template.Worksheet("Sheet1"), RowFields, options));
        Assert.Contains("Only one row", exception.Message);
    }

    [Fact]
    public void No_row_field_returns_a_clear_error()
    {
        var template = SheetWith(sheet => sheet.Cell(3, 1).Value = "Free text");

        var exception = Assert.Throws<MergeException>(() =>
            RowExpander.DetectMarkerRow(template.Worksheet("Sheet1"), Array.Empty<string>(), new MergeOptions()));
        Assert.Contains("--marker", exception.Message);
    }

    [Fact]
    public void Row_fields_derived_or_declared()
    {
        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" })
            .AddRow(new Dictionary<string, object?> { ["ID"] = "2" });
        Assert.Equal(new[] { "ID", "NAME" }, data.RowFieldNames);

        var declared = new MergeData
        {
            DeclaredRowFields = new List<string> { "#ID#", "NAME" },
        };
        Assert.Equal(new[] { "ID", "NAME" }, declared.RowFieldNames);
    }
}
