using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace XlsxMerge;

/// <summary>Detection, normalization and substitution of <c>#FIELD#</c> merge fields.</summary>
internal static class MergeFields
{
    /// <summary>All strings enclosed between two "#" (as in the JS implementation).</summary>
    public static readonly Regex Pattern = new(@"#([^#]+)#", RegexOptions.Compiled);

    private static readonly Regex StrictIntegerPattern =
        new(@"^[+-]?\d+$", RegexOptions.Compiled);

    private static readonly Regex StrictNumberPattern =
        new(@"^[+-]?(?:\d+[.,]\d+|\d+|[.,]\d+)$", RegexOptions.Compiled);

    /// <summary>Normalizes a field key: "#FORMATION#" or "FORMATION" ⇒ "FORMATION".</summary>
    public static string NormalizeKey(string key)
    {
        if (key is null)
            throw new MergeException("Null merge field key.");

        var normalized = key.Trim();
        if (normalized.Length >= 2 && normalized[0] == '#' && normalized[^1] == '#')
            normalized = normalized.Substring(1, normalized.Length - 2).Trim();

        if (normalized.Length == 0)
            throw new MergeException("A merge field cannot be empty.");
        if (normalized.IndexOf('#') >= 0)
            throw new MergeException(
                $"The merge field '{normalized}' cannot contain '#'. A field is delimited by two '#'.");

        return normalized;
    }

    /// <summary>Validates every key of a data set.</summary>
    public static void Validate(MergeData data)
    {
        foreach (var key in data.GlobalFields.Keys)
            NormalizeKey(key);
        foreach (var row in data.Rows)
        foreach (var key in row.Keys)
            NormalizeKey(key);
        if (data.DeclaredRowFields is not null)
        foreach (var field in data.DeclaredRowFields)
            NormalizeKey(field);
    }

    /// <summary>Normalized global context (keys without #).</summary>
    public static Dictionary<string, object?> BuildGlobalContext(MergeData data)
    {
        var context = new Dictionary<string, object?>(StringComparer.Ordinal);
        foreach (var pair in data.GlobalFields)
            context[NormalizeKey(pair.Key)] = pair.Value;
        return context;
    }

    /// <summary>
    /// Context of one row: global + every row field (missing value ⇒ empty string, as in the
    /// JS implementation) + the row-specific values (which win over the global ones).
    /// </summary>
    public static Dictionary<string, object?> BuildRowContext(
        IReadOnlyDictionary<string, object?> globalContext,
        IReadOnlyList<string> rowFields,
        IDictionary<string, object?> row)
    {
        var context = new Dictionary<string, object?>(globalContext, StringComparer.Ordinal);
        foreach (var field in rowFields)
            context[field] = null;
        foreach (var pair in row)
            context[NormalizeKey(pair.Key)] = pair.Value;
        return context;
    }

    /// <summary>Replaces the known fields of a context in some text; unknown fields are left untouched.</summary>
    public static string ReplaceText(
        string text,
        IReadOnlyDictionary<string, object?> context,
        MergeOptions options,
        out bool changed)
    {
        changed = false;
        if (text.IndexOf('#') < 0)
            return text;

        StringBuilder? builder = null;
        var last = 0;
        foreach (Match match in Pattern.Matches(text))
        {
            var key = match.Groups[1].Value.Trim();
            if (!context.TryGetValue(key, out var value))
                continue; // unknown field: kept as-is

            changed = true;
            builder ??= new StringBuilder(text.Length + 16);
            builder.Append(text, last, match.Index - last);
            builder.Append(FormatForText(value, options));
            last = match.Index + match.Length;
        }

        if (!changed)
            return text;

        builder!.Append(text, last, text.Length - last);
        return builder.ToString();
    }

    /// <summary>
    /// Replaces the fields of one cell. When the cell is exactly a single field, the value is
    /// returned typed (number, boolean, date); otherwise the result is text formatted with
    /// <see cref="MergeOptions.NumberCulture"/>.
    /// </summary>
    public static XLCellValue ReplaceCell(
        string text,
        IReadOnlyDictionary<string, object?> context,
        MergeOptions options,
        out bool changed)
    {
        changed = false;
        if (text.IndexOf('#') < 0)
            return (XLCellValue)text;

        // A single field covering the whole cell: typed value.
        var first = Pattern.Match(text);
        if (first.Success
            && first.Index == 0
            && first.Length == text.Length
            && context.TryGetValue(first.Groups[1].Value.Trim(), out var value))
        {
            changed = true;
            return ToCellValue(value, options);
        }

        var replaced = ReplaceText(text, context, options, out changed);
        return changed ? (XLCellValue)replaced : (XLCellValue)text;
    }

    /// <summary>Converts a data value into a cell value.</summary>
    public static XLCellValue ToCellValue(object? value, MergeOptions options)
    {
        if (value is null)
            return string.Empty;
        if (value is string s)
            return ToStringValue(s, options);
        if (value is bool b)
            return b;
        if (value is double d)
            return d;
        if (value is float f)
            return (double)f;
        if (value is decimal m)
            return (double)m;
        if (value is DateTime dateTime)
            return dateTime;
        if (value is TimeSpan timeSpan)
            return timeSpan;
        if (value is sbyte or byte or short or ushort or int or uint or long or ulong)
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);

        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    /// <summary>Formats a value that is inserted into a composed text.</summary>
    public static string FormatForText(object? value, MergeOptions options)
    {
        switch (value)
        {
            case null:
                return string.Empty;
            case string s:
                return ToStringValueAsText(s, options);
            case bool boolean:
                return boolean.ToString();
            case double number:
                return number.ToString(options.NumberCulture);
            case float single:
                return ((double)single).ToString(options.NumberCulture);
            case decimal decimalNumber:
                return decimalNumber.ToString(options.NumberCulture);
            case sbyte or byte or short or ushort or int or uint or long or ulong:
                return Convert.ToDouble(value, CultureInfo.InvariantCulture)
                    .ToString(options.NumberCulture);
            default:
                return Convert.ToString(value, options.NumberCulture) ?? string.Empty;
        }
    }

    private static XLCellValue ToStringValue(string value, MergeOptions options)
    {
        if (options.NumericTextParsing == NumericTextParsing.Strict
            && TryParseNumber(value, out var number))
            return number;
        return value;
    }

    private static string ToStringValueAsText(string value, MergeOptions options)
    {
        if (options.NumericTextParsing == NumericTextParsing.Strict
            && TryParseNumber(value, out var number))
            return number.ToString(options.NumberCulture);
        return value;
    }

    /// <summary>
    /// Recognizes strictly numeric strings ("123", "6.882", "6,882", "-3.5") and interprets
    /// them as invariant floating point numbers.
    /// </summary>
    public static bool TryParseNumber(string value, out double number)
    {
        number = 0;
        if (value.Length == 0 || !StrictNumberPattern.IsMatch(value))
            return false;

        return double.TryParse(
            value.Replace(',', '.'),
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out number);
    }

    /// <summary>True when the string is a strict integer.</summary>
    public static bool IsStrictInteger(string value) => StrictIntegerPattern.IsMatch(value);
}
