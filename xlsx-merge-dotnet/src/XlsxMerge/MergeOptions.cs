using System;
using System.Globalization;

namespace XlsxMerge;

/// <summary>Detection of the repeatable block (marker row plus optional style-model row).</summary>
public enum RepeatBlockRows
{
    /// <summary>
    /// Automatic detection: the row following the marker row is a second style-model row
    /// when it exists, contains cells (styling) but no value.
    /// </summary>
    Auto,

    /// <summary>The repeatable block consists of the marker row only.</summary>
    One,

    /// <summary>The repeatable block consists of the marker row and the next row (alternating styles).</summary>
    Two,
}

/// <summary>Conversion of numeric strings into numeric values.</summary>
public enum NumericTextParsing
{
    /// <summary>Text values stay text (even "6.882").</summary>
    None,

    /// <summary>Strictly numeric strings ("6.882", "-3"…) become numbers.</summary>
    Strict,
}

/// <summary>Options of the mail-merge.</summary>
public sealed class MergeOptions
{
    /// <summary>Repeatable block detection. Default: <see cref="RepeatBlockRows.Auto"/>.</summary>
    public RepeatBlockRows RepeatBlockRows { get; set; } = RepeatBlockRows.Auto;

    /// <summary>
    /// Forced marker field (e.g. "EMPLOYEE_ID"), with or without <c>#</c>.
    /// By default the marker row is detected automatically: the row containing the
    /// largest number of distinct row fields (ambiguity ⇒ explicit error).
    /// </summary>
    public string? RowMarkerField { get; set; }

    /// <summary>Culture used to format numbers inserted into text. Default: fr-FR.</summary>
    public CultureInfo NumberCulture { get; set; } = new("fr-FR");

    /// <summary>Conversion of numeric strings. Default: <see cref="NumericTextParsing.Strict"/>.</summary>
    public NumericTextParsing NumericTextParsing { get; set; } = NumericTextParsing.Strict;

    internal string? NormalizedMarkerField =>
        RowMarkerField is null ? null : MergeFields.NormalizeKey(RowMarkerField);
}
