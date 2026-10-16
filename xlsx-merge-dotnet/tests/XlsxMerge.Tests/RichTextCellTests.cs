using ClosedXML.Excel;
using Xunit;

namespace XlsxMerge.Tests;

public class RichTextCellTests
{
    [Fact]
    public void Rich_text_cell_replaced_with_the_cell_font()
    {
        var template = new XLWorkbook();
        var sheet = template.Worksheets.Add("Sheet1");
        var cell = sheet.Cell(3, 1);
        var richText = cell.GetRichText();
        richText.AddText("Hello ", cell.Style.Font).SetBold();
        richText.AddText("#NAME#", cell.Style.Font);
        sheet.Cell(3, 2).Value = "#ID# #NAME#";

        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var merged = workbook.Worksheet("Sheet1");

        Assert.Equal("Hello Alice", merged.Cell(3, 1).GetString());
        Assert.Equal("1 Alice", merged.Cell(3, 2).GetString());
    }

    [Fact]
    public void Formula_cell_not_modified()
    {
        var template = new XLWorkbook();
        var sheet = template.Worksheets.Add("Sheet1");
        sheet.Cell(3, 1).Value = "#ID#";
        sheet.Cell(3, 2).FormulaA1 = "LEN(A3)";

        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "12345" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var merged = workbook.Worksheet("Sheet1");

        Assert.True(merged.Cell(3, 2).HasFormula);
        Assert.Equal("LEN(A3)", merged.Cell(3, 2).FormulaA1);
    }

    [Fact]
    public void Cell_without_field_intact()
    {
        var template = new XLWorkbook();
        var sheet = template.Worksheets.Add("Sheet1");
        sheet.Cell(3, 1).Value = "#ID#";
        sheet.Cell(3, 2).Value = 42.5;
        sheet.Cell(3, 3).Value = "Free text";

        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var merged = workbook.Worksheet("Sheet1");

        Assert.Equal(42.5, merged.Cell(3, 2).GetDouble());
        Assert.Equal("Free text", merged.Cell(3, 3).GetString());
    }
}
