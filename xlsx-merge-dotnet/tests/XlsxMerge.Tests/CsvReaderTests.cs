using XlsxMerge.IO;
using Xunit;

namespace XlsxMerge.Tests;

public class CsvReaderTests
{
    [Fact]
    public void Reads_simple_fields()
    {
        var records = CsvReader.Parse("a;b;c\n1;2;3");
        Assert.Equal(2, records.Count);
        Assert.Equal(new[] { "a", "b", "c" }, records[0].Fields);
        Assert.Equal(new[] { "1", "2", "3" }, records[1].Fields);
        Assert.Equal(1, records[0].LineNumber);
        Assert.Equal(2, records[1].LineNumber);
    }

    [Fact]
    public void Handles_quotes_and_quoted_delimiters()
    {
        var records = CsvReader.Parse("\"Hello; world\";B\n\"with \"\"quotes\"\"\";C");
        Assert.Equal("Hello; world", records[0].Fields[0]);
        Assert.Equal("with \"quotes\"", records[1].Fields[0]);
    }

    [Fact]
    public void Handles_multi_line_fields()
    {
        var records = CsvReader.Parse("a;\"first\nline\";c");
        Assert.Equal("first\nline", records[0].Fields[1]);
        Assert.Equal(3, records[0].Fields.Count);
    }

    [Fact]
    public void Keeps_blank_lines_with_their_numbering()
    {
        var records = CsvReader.Parse("a\n\n\nb");
        Assert.Equal(4, records.Count);
        Assert.True(CsvReader.IsBlank(records[1]));
        Assert.True(CsvReader.IsBlank(records[2]));
        Assert.Equal(4, records[3].LineNumber);
    }

    [Fact]
    public void End_of_file_without_newline()
    {
        var records = CsvReader.Parse("a;b");
        Assert.Single(records);
        Assert.Equal(new[] { "a", "b" }, records[0].Fields);
    }

    [Fact]
    public void Unterminated_quote_is_an_error()
    {
        var exception = Assert.Throws<MergeException>(() => CsvReader.Parse("a;\"unterminated"));
        Assert.Contains("unterminated quote", exception.Message);
    }

    [Fact]
    public void Crlf_and_lone_cr()
    {
        Assert.Equal(2, CsvReader.Parse("a\r\nb").Count);
        Assert.Equal(2, CsvReader.Parse("a\rb").Count);
    }

    [Fact]
    public void Custom_delimiter()
    {
        var records = CsvReader.Parse("a|b", '|');
        Assert.Equal(new[] { "a", "b" }, records[0].Fields);
    }
}
