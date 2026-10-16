using XlsxMerge.IO;
using Xunit;

namespace XlsxMerge.Tests;

public class CsvMergeDataReaderTests
{
    private const string Sample = """
        [globals]
        COURSE;Advanced Excel
        DATE;2024-06-01

        [rows]
        EMPLOYEE_ID;NAME;SCORE
        42;Alice;17.5
        43;Bob;
        """;

    [Fact]
    public void Reads_both_sections()
    {
        var data = CsvMergeDataReader.Read(new StringReader(Sample));

        Assert.Equal("Advanced Excel", data.GlobalFields["COURSE"]);
        Assert.Equal("2024-06-01", data.GlobalFields["DATE"]);

        Assert.Equal(new[] { "EMPLOYEE_ID", "NAME", "SCORE" }, data.RowFieldNames);
        Assert.Equal(2, data.Rows.Count);
        Assert.Equal("42", data.Rows[0]["EMPLOYEE_ID"]);
        Assert.Equal("17.5", data.Rows[0]["SCORE"]);
        Assert.Equal("Bob", data.Rows[1]["NAME"]);
        Assert.Null(data.Rows[1]["SCORE"]);
    }

    [Fact]
    public void Quotes_and_semicolons_in_values()
    {
        var data = CsvMergeDataReader.Read(new StringReader("""
            [globals]
            COURSE;"Excel; advanced"

            [rows]
            NAME
            "Alice; aka Ali"
            """));

        Assert.Equal("Excel; advanced", data.GlobalFields["COURSE"]);
        Assert.Equal("Alice; aka Ali", data.Rows[0]["NAME"]);
    }

    [Fact]
    public void Content_before_a_section_is_an_error()
    {
        var exception = Assert.Throws<MergeException>(() =>
            CsvMergeDataReader.Read(new StringReader("COURSE;X\n[globals]\nA;B")));
        Assert.Contains("[globals] or [rows] section", exception.Message);
    }

    [Fact]
    public void Column_without_header_is_an_error()
    {
        var exception = Assert.Throws<MergeException>(() =>
            CsvMergeDataReader.Read(new StringReader("[rows]\nA;B\n1;2;3")));
        Assert.Contains("3 columns", exception.Message);
        Assert.Contains("2", exception.Message);
    }

    [Fact]
    public void Duplicate_header_is_an_error()
    {
        var exception = Assert.Throws<MergeException>(() =>
            CsvMergeDataReader.Read(new StringReader("[rows]\nA;#A#")));
        Assert.Contains("duplicate header", exception.Message);
    }

    [Fact]
    public void Empty_file_is_an_error()
    {
        var exception = Assert.Throws<MergeException>(() =>
            CsvMergeDataReader.Read(new StringReader("\n\n")));
        Assert.Contains("empty", exception.Message);
    }

    [Fact]
    public void Two_rows_sections_are_an_error()
    {
        var exception = Assert.Throws<MergeException>(() =>
            CsvMergeDataReader.Read(new StringReader("[rows]\nA\n1\n[rows]\nB")));
        Assert.Contains("already defined", exception.Message);
    }

    [Fact]
    public void Keys_normalized_with_hash()
    {
        var data = CsvMergeDataReader.Read(new StringReader("[globals]\n#COURSE#;Excel"));
        Assert.Equal("Excel", data.GlobalFields["COURSE"]);
    }

    [Fact]
    public void Custom_delimiter()
    {
        var data = CsvMergeDataReader.Read(new StringReader("[globals]\nCOURSE|Excel"), "|");
        Assert.Equal("Excel", data.GlobalFields["COURSE"]);
    }

    [Fact]
    public void Header_only_rows_declares_the_fields()
    {
        var data = CsvMergeDataReader.Read(new StringReader("[rows]\nID;NAME"));
        Assert.Empty(data.Rows);
        Assert.Equal(new[] { "ID", "NAME" }, data.RowFieldNames);
    }
}
