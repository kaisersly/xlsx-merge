using ClosedXML.Excel;
using Xunit;

namespace XlsxMerge.Tests;

public class TemplateGeneratorTests
{
    private static TemplateSpec BuildSpec()
    {
        var spec = new TemplateSpec();
        spec.GlobalHints.Add("COURSE");
        spec.GlobalHints.Add("DATE");

        var sheet = new TemplateSheetSpec { Name = "Group A", Title = "Group A — workshop" };
        sheet.Columns.Add(new TemplateColumnSpec
        {
            Field = "ID", Label = "Identifier", Element = "People", Identity = true,
        });
        sheet.Columns.Add(new TemplateColumnSpec
        {
            Field = "NAME", Label = "Name", Element = "People", Identity = true,
        });
        sheet.Columns.Add(new TemplateColumnSpec
        {
            Field = "SCORE", Label = "Score", Element = "Result", Group = "Session 1",
        });
        sheet.Columns.Add(new TemplateColumnSpec
        {
            Field = "GRADE", Label = "Grade", Element = "Result", Group = "Session 2",
        });
        spec.Sheets.Add(sheet);

        var second = new TemplateSheetSpec { Name = "All groups" };
        second.Columns.Add(new TemplateColumnSpec { Field = "ID", Label = "Identifier", Identity = true });
        second.Columns.Add(new TemplateColumnSpec { Field = "SCORE", Label = "Score", Group = "Session 1" });
        second.Columns.Add(new TemplateColumnSpec { Field = "GRADE", Label = "Grade", Group = "Session 2" });
        spec.Sheets.Add(second);
        return spec;
    }

    private static XLWorkbook Generate(TemplateSpec spec)
    {
        using var output = new MemoryStream();
        TemplateGenerator.Generate(spec, output);
        return new XLWorkbook(new MemoryStream(output.ToArray()));
    }

    [Fact]
    public void Generates_titles_labels_and_marker_row()
    {
        using var workbook = Generate(BuildSpec());
        var sheet = workbook.Worksheet("Group A");

        // Row 1: title (merged, text from spec).
        Assert.Equal("Group A — workshop", sheet.Cell(1, 1).GetString());
        // Row 2: global hints.
        Assert.Equal("#COURSE#  ·  #DATE#", sheet.Cell(2, 1).GetString());
        // Row 3: element runs (merged).
        Assert.Equal("People", sheet.Cell(3, 1).GetString());
        Assert.Equal("Result", sheet.Cell(3, 3).GetString());
        Assert.True(TestHelpers.IsMerged(sheet, 3, 1, 3, 2));
        Assert.True(TestHelpers.IsMerged(sheet, 3, 3, 3, 4));
        // Element titles rotated up except the identity run (« Identification »… well,
        // here the identity run is « People »), and the row grows to fit them.
        Assert.Equal(0, sheet.Cell(3, 1).Style.Alignment.TextRotation);
        Assert.Equal(90, sheet.Cell(3, 3).Style.Alignment.TextRotation);
        Assert.True(sheet.Row(3).Height > 20, $"element row height={sheet.Row(3).Height}");
        // The session/group row (row 4) stays flat.
        Assert.Equal(0, sheet.Cell(4, 3).Style.Alignment.TextRotation);
        // Row 4: group runs (merged).
        Assert.Equal("Session 1", sheet.Cell(4, 3).GetString());
        Assert.Equal("Session 2", sheet.Cell(4, 4).GetString());
        // Row 5: labels — non-identity headers rotated up, identity horizontal.
        Assert.Equal("Identifier", sheet.Cell(5, 1).GetString());
        Assert.Equal("Score", sheet.Cell(5, 3).GetString());
        Assert.Equal(0, sheet.Cell(5, 1).Style.Alignment.TextRotation);
        Assert.Equal(90, sheet.Cell(5, 3).Style.Alignment.TextRotation);
        Assert.Equal(90, sheet.Cell(5, 4).Style.Alignment.TextRotation);
        // The label row grows so the rotated labels are fully readable.
        Assert.True(sheet.Row(5).Height > 30, $"height={sheet.Row(5).Height}");
        // Row 6: marker row — every row field wrapped in #.
        Assert.Equal("#ID#", sheet.Cell(6, 1).GetString());
        Assert.Equal("#NAME#", sheet.Cell(6, 2).GetString());
        Assert.Equal("#SCORE#", sheet.Cell(6, 3).GetString());
        Assert.Equal("#GRADE#", sheet.Cell(6, 4).GetString());
        Assert.Equal(0, sheet.Cell(6, 3).Style.Alignment.TextRotation);
        Assert.Equal(6, sheet.LastRowUsed(XLCellsUsedOptions.Contents)!.RowNumber());

        // Identity column: content-driven width (label « Identifier » = 10 + 3).
        Assert.Equal(13d, sheet.Column(1).Width);
        // Non-identity columns: fixed narrow width.
        Assert.Equal(4d, sheet.Column(3).Width);
        Assert.Equal(4d, sheet.Column(4).Width);
    }

    [Fact]
    public void Groups_columns_with_an_outline_level_and_freezes_header()
    {
        using var workbook = Generate(BuildSpec());
        var sheet = workbook.Worksheet("Group A");

        // Grouped session columns are collapsible; identity columns are not.
        Assert.Equal(1, sheet.Column(3).OutlineLevel);
        Assert.Equal(1, sheet.Column(4).OutlineLevel);
        Assert.Equal(0, sheet.Column(1).OutlineLevel);
        Assert.Equal(0, sheet.Column(2).OutlineLevel);

        // Only the 5 header rows above the marker (row 6 = first data row) are frozen;
        // 2 identity columns frozen as well.
        Assert.Equal(5, sheet.SheetView.SplitRow);
        Assert.Equal(2, sheet.SheetView.SplitColumn);
    }

    [Fact]
    public void Second_sheet_without_elements_has_no_element_row()
    {
        using var workbook = Generate(BuildSpec());
        var sheet = workbook.Worksheet("All groups");

        Assert.Equal("All groups", sheet.Cell(1, 1).GetString()); // title defaults to name
        // No element row: row 2 = hints, row 3 = groups, row 4 = labels, row 5 = marker.
        Assert.Equal("Session 1", sheet.Cell(3, 2).GetString());
        Assert.Equal("#SCORE#", sheet.Cell(4 + 1, 2).GetString());
        Assert.Equal(4, sheet.SheetView.SplitRow);
        Assert.Equal(1, sheet.Column(2).OutlineLevel);
        Assert.Equal(0, sheet.Column(1).OutlineLevel);
    }

    [Fact]
    public void Sheet_names_are_sanitized()
    {
        var spec = new TemplateSpec();
        var sheet = new TemplateSheetSpec { Name = "Bad:Name/With?Chars*" };
        sheet.Columns.Add(new TemplateColumnSpec { Field = "ID" });
        spec.Sheets.Add(sheet);
        var duplicate = new TemplateSheetSpec { Name = "Bad:Name/With?Chars*" };
        duplicate.Columns.Add(new TemplateColumnSpec { Field = "ID" });
        spec.Sheets.Add(duplicate);

        using var workbook = Generate(spec);
        Assert.Equal("Bad Name With Chars", workbook.Worksheet(1).Name);
        Assert.Equal("Bad Name With Chars (2)", workbook.Worksheet(2).Name);
    }

    [Fact]
    public void Generated_template_is_renderable()
    {
        // Round trip: the generated marker row is detected and expanded by the merger,
        // global fields land in the title/hint rows.
        var spec = BuildSpec();
        var data = new MergeData()
            .AddGlobal("COURSE", "Excel")
            .AddGlobal("DATE", "2024-06-01")
            .AddRow(new Dictionary<string, object?> { ["ID"] = "1", ["NAME"] = "Alice", ["SCORE"] = "17.5", ["GRADE"] = "A" })
            .AddRow(new Dictionary<string, object?> { ["ID"] = "2", ["NAME"] = "Bob", ["SCORE"] = "8", ["GRADE"] = "C" });

        using var template = Generate(spec);
        using var rendered = TestHelpers.RenderAndLoad(template, data);
        var sheet = rendered.Worksheet("Group A");

        Assert.Equal("Group A — workshop", sheet.Cell(1, 1).GetString());
        Assert.Equal("#COURSE#  ·  #DATE#".Replace("#COURSE#", "Excel").Replace("#DATE#", "2024-06-01"),
            sheet.Cell(2, 1).GetString());
        // The marker row (6) is replaced by 2 data rows, elements/labels stay above.
        Assert.Equal("1", sheet.Cell(6, 1).GetString());
        Assert.Equal("Alice", sheet.Cell(6, 2).GetString());
        Assert.Equal(17.5, sheet.Cell(6, 3).GetDouble());
        Assert.Equal("2", sheet.Cell(7, 1).GetString());
        Assert.Equal("Bob", sheet.Cell(7, 2).GetString());
        Assert.Equal("C", sheet.Cell(7, 4).GetString());
    }

    [Fact]
    public void Rotated_labels_stretch_the_row_but_not_the_columns()
    {
        var spec = new TemplateSpec();
        var sheet = new TemplateSheetSpec { Name = "Rotate" };
        sheet.Columns.Add(new TemplateColumnSpec
        {
            Field = "ID", Label = "Identifier", Identity = true,
        });
        sheet.Columns.Add(new TemplateColumnSpec
        {
            Field = "SCORE", Label = "A very long column label for rotation",
        });
        spec.Sheets.Add(sheet);

        using var workbook = Generate(spec);
        var ws = workbook.Worksheet("Rotate");

        // Layout without hints/element/group rows: row 1 title, row 2 labels, row 3 marker.
        Assert.Equal(90, ws.Cell(2, 2).Style.Alignment.TextRotation);
        Assert.Equal(0, ws.Cell(2, 1).Style.Alignment.TextRotation);
        Assert.Equal(0, ws.Cell(3, 2).Style.Alignment.TextRotation);

        // Non-identity column: fixed width 4 (was previously driven by the field).
        Assert.Equal(4d, ws.Column(2).Width);
        // Identity column keeps a content-driven width.
        Assert.True(ws.Column(1).Width > 4, $"width={ws.Column(1).Width}");
        // …but the label row grows to read it.
        Assert.True(ws.Row(2).Height >= 200, $"height={ws.Row(2).Height}");
    }

    [Fact]
    public void Page_setup_narrow_landscape_fitted_and_header_rows_repeated()
    {
        using var workbook = Generate(BuildSpec());
        var sheet = workbook.Worksheet("Group A"); // rows 1..5 = header, marker = 6

        Assert.Equal(XLPageOrientation.Landscape, sheet.PageSetup.PageOrientation);
        // All columns on one page wide (height as needed).
        Assert.Equal(1, sheet.PageSetup.PagesWide);
        Assert.Equal(0, sheet.PageSetup.PagesTall);
        // Narrow margins (inches).
        Assert.Equal(0.25d, sheet.PageSetup.Margins.Left);
        Assert.Equal(0.25d, sheet.PageSetup.Margins.Right);
        Assert.Equal(0.25d, sheet.PageSetup.Margins.Top);
        Assert.Equal(0.25d, sheet.PageSetup.Margins.Bottom);
        Assert.Equal(0.2d, sheet.PageSetup.Margins.Header);
        Assert.Equal(0.2d, sheet.PageSetup.Margins.Footer);
        // Header rows (1..5, above the marker) repeat on every printed page.
        Assert.Equal(1, sheet.PageSetup.FirstRowToRepeatAtTop);
        Assert.Equal(5, sheet.PageSetup.LastRowToRepeatAtTop);

        // Second sheet: no element row ⇒ its header block ends at row 4.
        var second = workbook.Worksheet("All groups");
        Assert.Equal(XLPageOrientation.Landscape, second.PageSetup.PageOrientation);
        Assert.Equal(4, second.PageSetup.LastRowToRepeatAtTop);
    }

    [Fact]
    public void Validation_errors()
    {
        var empty = new TemplateSpec();
        Assert.Contains("at least one sheet",
            Assert.Throws<MergeException>(() => TemplateGenerator.Generate(empty, new MemoryStream())).Message);

        var noColumns = new TemplateSpec();
        noColumns.Sheets.Add(new TemplateSheetSpec { Name = "S" });
        Assert.Contains("no column",
            Assert.Throws<MergeException>(() => TemplateGenerator.Generate(noColumns, new MemoryStream())).Message);

        var badField = new TemplateSpec();
        var badSheet = new TemplateSheetSpec { Name = "S" };
        badSheet.Columns.Add(new TemplateColumnSpec { Field = "BAD#FIELD" });
        badField.Sheets.Add(badSheet);
        Assert.Contains("Invalid field",
            Assert.Throws<MergeException>(() => TemplateGenerator.Generate(badField, new MemoryStream())).Message);
    }

    [Fact]
    public void GenerateFile_writes_the_output()
    {
        var directory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "out", "template.xlsx");
            TemplateGenerator.GenerateFile(BuildSpec(), path);
            Assert.True(File.Exists(path));
            using var workbook = new XLWorkbook(path);
            Assert.Equal(2, workbook.Worksheets.Count);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
