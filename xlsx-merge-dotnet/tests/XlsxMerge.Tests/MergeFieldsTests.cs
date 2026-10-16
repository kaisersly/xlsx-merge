using Xunit;

namespace XlsxMerge.Tests;

public class MergeFieldsTests
{
    private static readonly MergeOptions Options = new();

    [Fact]
    public void Single_field_covering_the_cell_becomes_a_typed_value()
    {
        var context = new Dictionary<string, object?> { ["SCORE"] = "17.5" };
        var value = MergeFields.ReplaceCell("#SCORE#", context, Options, out var changed);

        Assert.True(changed);
        Assert.True(value.IsNumber);
        Assert.Equal(17.5, value.GetNumber());
    }

    [Fact]
    public void Single_integer_field_becomes_a_number()
    {
        var context = new Dictionary<string, object?> { ["ID"] = "123456789" };
        var value = MergeFields.ReplaceCell("#ID#", context, Options, out var changed);

        Assert.True(changed);
        Assert.True(value.IsNumber);
        Assert.Equal(123456789d, value.GetNumber());
    }

    [Fact]
    public void Null_or_empty_value_produces_an_empty_cell()
    {
        var context = new Dictionary<string, object?> { ["X"] = null };
        var value = MergeFields.ReplaceCell("#X#", context, Options, out var changed);

        Assert.True(changed);
        Assert.True(value.IsText);
        Assert.Equal(string.Empty, value.GetText());
    }

    [Fact]
    public void Typed_value_is_preserved()
    {
        var context = new Dictionary<string, object?> { ["SCORE"] = 11.5 };
        var value = MergeFields.ReplaceCell("#SCORE#", context, Options, out var changed);

        Assert.True(changed);
        Assert.Equal(11.5, value.GetNumber());
    }

    [Fact]
    public void Unknown_field_is_left_untouched()
    {
        var context = new Dictionary<string, object?> { ["KNOWN"] = "yes" };
        var value = MergeFields.ReplaceCell("#UNKNOWN# #KNOWN#", context, Options, out var changed);

        Assert.True(changed);
        Assert.Equal("#UNKNOWN# yes", value.GetText());
    }

    [Fact]
    public void Composed_text_formats_numbers_with_the_culture()
    {
        var context = new Dictionary<string, object?> { ["SCORE"] = "17.5" };
        var value = MergeFields.ReplaceCell("Score : #SCORE# / 20", context, Options, out var changed);

        Assert.True(changed);
        Assert.Equal("Score : 17,5 / 20", value.GetText());
    }

    [Fact]
    public void Disabled_numeric_parsing_keeps_text()
    {
        var options = new MergeOptions { NumericTextParsing = NumericTextParsing.None };
        var context = new Dictionary<string, object?> { ["SCORE"] = "17.5" };
        var value = MergeFields.ReplaceCell("#SCORE#", context, options, out var changed);

        Assert.True(changed);
        Assert.True(value.IsText);
        Assert.Equal("17.5", value.GetText());
    }

    [Fact]
    public void Composed_text_uses_the_requested_culture()
    {
        var options = new MergeOptions { NumberCulture = new System.Globalization.CultureInfo("en-US") };
        var context = new Dictionary<string, object?> { ["SCORE"] = "17.5" };
        var value = MergeFields.ReplaceCell("Score: #SCORE#/20", context, options, out var changed);

        Assert.True(changed);
        Assert.Equal("Score: 17.5/20", value.GetText());
    }

    [Fact]
    public void Every_occurrence_of_the_same_field_is_replaced()
    {
        var context = new Dictionary<string, object?> { ["S"] = "S1" };
        var value = MergeFields.ReplaceCell("#S# and #S#", context, Options, out var changed);

        Assert.True(changed);
        Assert.Equal("S1 and S1", value.GetText());
    }

    [Fact]
    public void Cell_without_field_is_unchanged()
    {
        var context = new Dictionary<string, object?> { ["S"] = "S1" };
        _ = MergeFields.ReplaceCell("Free text", context, Options, out var changed);
        Assert.False(changed);
    }

    [Fact]
    public void Keys_normalized_with_or_without_hash()
    {
        var data = new MergeData().AddGlobal("#COURSE#", "Excel");
        var context = MergeFields.BuildGlobalContext(data);
        Assert.Equal("Excel", context["COURSE"]);
    }

    [Fact]
    public void Validation_rejects_invalid_keys()
    {
        var data = new MergeData();
        data.GlobalFields["BAD#KEY"] = "x";
        Assert.Throws<MergeException>(() => MergeFields.Validate(data));

        var empty = new MergeData();
        empty.GlobalFields["##"] = "x";
        Assert.Throws<MergeException>(() => MergeFields.Validate(empty));
    }

    [Theory]
    [InlineData("17.5", 17.5)]
    [InlineData("17,5", 17.5)]
    [InlineData("-3", -3)]
    [InlineData("+3.5", 3.5)]
    [InlineData(".5", 0.5)]
    [InlineData("123456789", 123456789)]
    public void Recognizes_strict_numbers(string text, double expected)
    {
        Assert.True(MergeFields.TryParseNumber(text, out var number));
        Assert.Equal(expected, number);
    }

    [Theory]
    [InlineData("")]
    [InlineData("1.2.3")]
    [InlineData("abc")]
    [InlineData("12a")]
    [InlineData(" 17.5")]
    public void Rejects_non_numbers(string text)
    {
        Assert.False(MergeFields.TryParseNumber(text, out _));
    }
}
