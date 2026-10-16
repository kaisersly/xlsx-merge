using ClosedXML.Excel;
using Xunit;

namespace XlsxMerge.Tests;

public class RowExpanderTests
{
    /// <summary>Template: title (1), marker block (3) + style row (4), footer (5).</summary>
    private static XLWorkbook BuildTemplate(bool mergeInBlock = true, bool styleRow = true)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Title #COURSE#";

        sheet.Cell(3, 1).Value = "#ID#";
        sheet.Cell(3, 2).Value = "#NAME#";
        sheet.Cell(3, 1).Style.Fill.BackgroundColor = XLColor.Red;
        sheet.Row(3).Height = 20;
        if (mergeInBlock)
            sheet.Range("C3:D3").Merge();

        if (styleRow)
        {
            sheet.Cell(4, 1).Style.Fill.BackgroundColor = XLColor.Blue;
            sheet.Row(4).Height = 25;
        }

        sheet.Cell(5, 1).Value = "Footer";
        sheet.Range("A5:B5").Merge();
        return workbook;
    }

    [Fact]
    public void Expansion_alternating_styles_merges_and_heights()
    {
        var template = BuildTemplate();
        var data = new MergeData()
            .AddGlobal("COURSE", "Excel")
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" })
            .AddRow(new Dictionary<string, object?> { ["ID"] = "2", ["NAME"] = "Bob" })
            .AddRow(new Dictionary<string, object?> { ["ID"] = "3", ["NAME"] = "Carol" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var sheet = workbook.Worksheet("Sheet1");

        // Title, 3 generated rows (3..5), footer moved to 6.
        Assert.Equal("Title Excel", sheet.Cell(1, 1).GetString());
        Assert.Equal("1", sheet.Cell(3, 1).GetString());
        Assert.Equal("Alice", sheet.Cell(3, 2).GetString());
        Assert.Equal("2", sheet.Cell(4, 1).GetString());
        Assert.Equal("Bob", sheet.Cell(4, 2).GetString());
        Assert.Equal("3", sheet.Cell(5, 1).GetString());
        Assert.Equal("Carol", sheet.Cell(5, 2).GetString());
        Assert.Equal("Footer", sheet.Cell(6, 1).GetString());
        Assert.Equal(6, sheet.LastRowUsed(XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats)!.RowNumber());

        // Alternating styles: row 0 ⇒ marker row style, row 1 ⇒ second model row style.
        Assert.Equal(XLColor.Red, sheet.Cell(3, 1).Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.Blue, sheet.Cell(4, 1).Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.Red, sheet.Cell(5, 1).Style.Fill.BackgroundColor);
        Assert.Equal(20, sheet.Row(3).Height);
        Assert.Equal(25, sheet.Row(4).Height);
        Assert.Equal(20, sheet.Row(5).Height);

        // Block merges replicated, merge below the block moved.
        Assert.True(TestHelpers.IsMerged(sheet, 3, 3, 3, 4));
        Assert.True(TestHelpers.IsMerged(sheet, 4, 3, 4, 4));
        Assert.True(TestHelpers.IsMerged(sheet, 5, 3, 5, 4));
        Assert.True(TestHelpers.IsMerged(sheet, 6, 1, 6, 2));
        Assert.False(TestHelpers.IsMerged(sheet, 3, 1, 3, 2));
    }

    [Fact]
    public void Single_data_row_removes_the_style_row()
    {
        var template = BuildTemplate();
        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var sheet = workbook.Worksheet("Sheet1");

        Assert.Equal("1", sheet.Cell(3, 1).GetString());
        Assert.Equal("Alice", sheet.Cell(3, 2).GetString());
        Assert.Equal("Footer", sheet.Cell(4, 1).GetString());
        Assert.Equal(4, sheet.LastRowUsed(XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats)!.RowNumber());
        Assert.Equal(XLColor.Red, sheet.Cell(3, 1).Style.Fill.BackgroundColor);
    }

    [Fact]
    public void Zero_rows_removes_the_model_block()
    {
        var template = BuildTemplate();
        var data = new MergeData();
        var options = new MergeOptions { RowMarkerField = "ID" };

        using var workbook = TestHelpers.RenderAndLoad(template, data, options);
        var sheet = workbook.Worksheet("Sheet1");

        Assert.Equal("Footer", sheet.Cell(3, 1).GetString());
        Assert.Equal(3, sheet.LastRowUsed(XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats)!.RowNumber());
    }

    [Fact]
    public void Next_row_with_values_does_not_extend_the_block()
    {
        var template = BuildTemplate(styleRow: false);
        template.Worksheet("Sheet1").Cell(4, 1).Value = "Header";
        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" })
            .AddRow(new Dictionary<string, object?> { ["ID"] = "2", ["NAME"] = "Bob" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var sheet = workbook.Worksheet("Sheet1");

        Assert.Equal("1", sheet.Cell(3, 1).GetString());
        Assert.Equal("2", sheet.Cell(4, 1).GetString());
        Assert.Equal("Header", sheet.Cell(5, 1).GetString());
        // One-row block: every row gets the marker row style.
        Assert.Equal(XLColor.Red, sheet.Cell(3, 1).Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.Red, sheet.Cell(4, 1).Style.Fill.BackgroundColor);
    }

    [Fact]
    public void Forced_two_row_block_rejects_a_missing_style_row()
    {
        var template = BuildTemplate(styleRow: false);
        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" });
        var options = new MergeOptions { RepeatBlockRows = RepeatBlockRows.Two };

        var exception = Assert.Throws<MergeException>(() =>
            TestHelpers.Render(template, data, options));
        Assert.Contains("second style-model row", exception.Message);
    }

    [Fact]
    public void Forced_one_row_block_keeps_the_style_row_as_content()
    {
        var template = BuildTemplate();
        var data = new MergeData()
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" });
        var options = new MergeOptions { RepeatBlockRows = RepeatBlockRows.One };

        using var workbook = TestHelpers.RenderAndLoad(template, data, options);
        var sheet = workbook.Worksheet("Sheet1");

        Assert.Equal("1", sheet.Cell(3, 1).GetString());
        // The former style row (without values) stays as an empty row; Footer is at 5.
        Assert.True(sheet.Cell(4, 1).Value.IsBlank);
        Assert.Equal("Footer", sheet.Cell(5, 1).GetString());
    }

    [Fact]
    public void Sheet_without_marker_row_gets_only_global_fields()
    {
        var template = new XLWorkbook();
        var block = template.Worksheets.Add("Block");
        block.Cell(1, 1).Value = "Title #COURSE#";
        block.Cell(3, 1).Value = "#ID#";
        block.Cell(3, 2).Value = "#NAME#";

        // A sheet without any row field (cover page): no repeatable block expected.
        var cover = template.Worksheets.Add("Cover");
        cover.Cell(1, 1).Value = "#COURSE# — cover";
        cover.Cell(2, 1).Value = "Contact: #UNKNOWN#";
        cover.Cell(4, 1).Value = "Free text";

        var data = new MergeData()
            .AddGlobal("COURSE", "Excel")
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice" });

        using var workbook = TestHelpers.RenderAndLoad(template, data);

        var renderedBlock = workbook.Worksheet("Block");
        Assert.Equal("Title Excel", renderedBlock.Cell(1, 1).GetString());
        Assert.Equal("1", renderedBlock.Cell(3, 1).GetString());
        Assert.Equal("Alice", renderedBlock.Cell(3, 2).GetString());

        var renderedCover = workbook.Worksheet("Cover");
        Assert.Equal("Excel — cover", renderedCover.Cell(1, 1).GetString());
        // Unknown field left as-is, structure untouched (no block expansion).
        Assert.Equal("Contact: #UNKNOWN#", renderedCover.Cell(2, 1).GetString());
        Assert.Equal("Free text", renderedCover.Cell(4, 1).GetString());
        Assert.Equal(4, renderedCover.LastRowUsed(XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats)!.RowNumber());
    }

    [Fact]
    public void No_row_field_defined_keeps_the_template_block_untouched()
    {
        // Data with globals only (no row fields at all): the marker row cannot be
        // detected, so nothing is expanded nor removed — unknown fields stay visible.
        var template = BuildTemplate();
        var data = new MergeData().AddGlobal("COURSE", "Excel");

        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var sheet = workbook.Worksheet("Sheet1");

        Assert.Equal("Title Excel", sheet.Cell(1, 1).GetString());
        Assert.Equal("#ID#", sheet.Cell(3, 1).GetString());
        Assert.Equal("Footer", sheet.Cell(5, 1).GetString());
    }
}
