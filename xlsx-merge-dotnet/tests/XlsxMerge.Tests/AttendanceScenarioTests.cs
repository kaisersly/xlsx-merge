using ClosedXML.Excel;
using XlsxMerge.IO;
using Xunit;

namespace XlsxMerge.Tests;

/// <summary>
/// End-to-end scenario, independent of any real-world domain: a training attendance sheet.
/// It exercises the whole pipeline the way an application would use it — a template built
/// with a spreadsheet editor (here produced in code) plus a data file in the generic
/// [globals]/[rows] format.
///
/// Template layout:
///   row 1  : title with global fields  #COURSE# #DATE#
///   row 2  : section label
///   row 4  : repeatable block marker   #EMPLOYEE_ID# #NAME# #SCORE# (merged cell E4:F4)
///   row 5  : alternating-style model row (styled, no value)
///   row 7  : footer with a global field and an unknown field
///   header : #COURSE# (odd) / first page uses a different header
///   footer : #DATE# (odd)
/// </summary>
public class AttendanceScenarioTests
{
    private const string DataCsv = """
        [globals]
        COURSE;Advanced Excel
        DATE;2024-06-01
        ROOM;B12

        [rows]
        EMPLOYEE_ID;NAME;SCORE
        42;Alice;17.5
        43;Bob;8
        44;Carol;13.25
        45;Dan;
        """;

    private static XLWorkbook BuildTemplate()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Attendance");

        // Title and section label (global fields).
        sheet.Cell(1, 1).Value = "#COURSE# — attendance";
        sheet.Cell(2, 1).Value = "Room: #ROOM#";
        sheet.Range("A1:D1").Merge();

        // Repeatable block: marker row + alternating-style model row.
        sheet.Cell(4, 1).Value = "#EMPLOYEE_ID#";
        sheet.Cell(4, 2).Value = "#NAME#";
        sheet.Cell(4, 3).Value = "#SCORE# / 20";
        sheet.Cell(4, 4).Value = "Signed";
        sheet.Range("E4:F4").Merge();
        sheet.Cell(4, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#DDEBF7");
        sheet.Row(4).Height = 18;
        sheet.Cell(5, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFFFF");

        // Footer line with a global and an unknown field.
        sheet.Cell(7, 1).Value = "Generated on #DATE# (#SOMEWHERE#)";

        // Headers / footers.
        sheet.PageSetup.SetDifferentFirstPageOnHF(true);
        sheet.PageSetup.Header.Left.AddText("#COURSE#", XLHFOccurrence.OddPages);
        sheet.PageSetup.Header.Left.AddText("#COURSE# — cover", XLHFOccurrence.FirstPage);
        sheet.PageSetup.Footer.Center.AddText("#DATE#", XLHFOccurrence.OddPages);

        return workbook;
    }

    [Fact]
    public void Full_attendance_pipeline()
    {
        var data = CsvMergeDataReader.Read(new StringReader(DataCsv));
        using var template = BuildTemplate();
        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var sheet = workbook.Worksheet("Attendance");

        // Global fields substituted, the block expanded with one row per data row (4..7).
        Assert.Equal("Advanced Excel — attendance", sheet.Cell(1, 1).GetString());
        Assert.Equal("Room: B12", sheet.Cell(2, 1).GetString());

        // Row fields that are strictly numeric become numbers.
        Assert.Equal(42d, sheet.Cell(4, 1).GetDouble());
        Assert.Equal("Alice", sheet.Cell(4, 2).GetString());
        Assert.Equal("17,5 / 20", sheet.Cell(4, 3).GetString());
        Assert.Equal(43d, sheet.Cell(5, 1).GetDouble());
        Assert.Equal(44d, sheet.Cell(6, 1).GetDouble());
        // A missing row value renders as an empty string inside composed text.
        Assert.Equal("Dan", sheet.Cell(7, 2).GetString());
        Assert.Equal(" / 20", sheet.Cell(7, 3).GetString());

        // The footer line is shifted by (4 rows - 2 model rows).
        Assert.Equal("Generated on 2024-06-01 (#SOMEWHERE#)", sheet.Cell(9, 1).GetString());

        // Alternating styles: row 0 and 2 = marker style, row 1 = model row style.
        Assert.Equal(XLColor.FromHtml("#DDEBF7"), sheet.Cell(4, 1).Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.FromHtml("#FFFFFF"), sheet.Cell(5, 1).Style.Fill.BackgroundColor);
        Assert.Equal(XLColor.FromHtml("#DDEBF7"), sheet.Cell(6, 1).Style.Fill.BackgroundColor);
        Assert.Equal(18, sheet.Row(4).Height);

        // Block merge replicated on every generated row.
        Assert.True(TestHelpers.IsMerged(sheet, 4, 5, 4, 6));
        Assert.True(TestHelpers.IsMerged(sheet, 5, 5, 5, 6));
        Assert.True(TestHelpers.IsMerged(sheet, 6, 5, 6, 6));
        // Title merge above the block untouched.
        Assert.True(TestHelpers.IsMerged(sheet, 1, 1, 1, 4));

        // Headers and footers.
        Assert.Contains("Advanced Excel", sheet.PageSetup.Header.Left.GetText(XLHFOccurrence.OddPages));
        Assert.Contains("Advanced Excel — cover", sheet.PageSetup.Header.Left.GetText(XLHFOccurrence.FirstPage));
        Assert.Contains("2024-06-01", sheet.PageSetup.Footer.Center.GetText(XLHFOccurrence.OddPages));
    }

    [Fact]
    public void Numeric_scores_become_numbers_in_single_field_cells()
    {
        // The same template with the score as the whole cell content.
        var template = BuildTemplate();
        template.Worksheet("Attendance").Cell(4, 3).Value = "#SCORE#";

        var data = CsvMergeDataReader.Read(new StringReader(DataCsv));
        using var workbook = TestHelpers.RenderAndLoad(template, data);
        var sheet = workbook.Worksheet("Attendance");

        Assert.Equal(17.5, sheet.Cell(4, 3).GetDouble());
        Assert.Equal(8d, sheet.Cell(5, 3).GetDouble());
        // A missing row value renders as an empty string.
        Assert.Equal(string.Empty, sheet.Cell(7, 3).GetString());
    }

    [Fact]
    public void Render_to_file_and_reload()
    {
        var data = CsvMergeDataReader.Read(new StringReader(DataCsv));
        using var template = BuildTemplate();

        var directory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
        try
        {
            var templatePath = Path.Combine(directory, "template.xlsx");
            var outputPath = Path.Combine(directory, "out", "attendance.xlsx");
            File.WriteAllBytes(templatePath, TestHelpers.ToBytes(template));

            XlsxMerger.RenderFile(templatePath, outputPath, data);

            Assert.True(File.Exists(outputPath));
            using var rendered = new XLWorkbook(outputPath);
            Assert.Equal("Alice", rendered.Worksheet("Attendance").Cell(4, 2).GetString());
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
