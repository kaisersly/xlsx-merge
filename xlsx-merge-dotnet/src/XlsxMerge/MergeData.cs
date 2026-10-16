using System;
using System.Collections.Generic;

namespace XlsxMerge;

/// <summary>
/// Generic mail-merge data set.
/// <list type="bullet">
/// <item><see cref="GlobalFields"/>: fields available everywhere (worksheet body, headers, footers);</item>
/// <item><see cref="Rows"/>: rows to repeat, each carrying the row fields.</item>
/// </list>
/// Keys are given without the <c>#</c> delimiter ("FORMATION" for "#FORMATION#");
/// the form "#FORMATION#" is also accepted and normalized.
/// </summary>
public sealed class MergeData
{
    /// <summary>Global fields: replaced in every cell outside the repeatable block, and in headers/footers.</summary>
    public IDictionary<string, object?> GlobalFields { get; } =
        new Dictionary<string, object?>(StringComparer.Ordinal);

    /// <summary>Rows to repeat (one entry per generated row in the repeatable block).</summary>
    public IList<IDictionary<string, object?>> Rows { get; } =
        new List<IDictionary<string, object?>>();

    /// <summary>
    /// Explicitly declared row fields. Useful when no row carries data yet
    /// (a CSV with only a header line): the block stays detectable through the marker field.
    /// When absent, the fields are derived from the union of the <see cref="Rows"/> keys.
    /// </summary>
    public IList<string>? DeclaredRowFields { get; set; }

    /// <summary>Adds a global field.</summary>
    public MergeData AddGlobal(string key, object? value)
    {
        GlobalFields[MergeFields.NormalizeKey(key)] = value;
        return this;
    }

    /// <summary>Adds a row to repeat.</summary>
    public MergeData AddRow(IDictionary<string, object?> row)
    {
        Rows.Add(row ?? throw new ArgumentNullException(nameof(row)));
        return this;
    }

    /// <summary>Effective row fields (declared or derived from the rows), in appearance order.</summary>
    public IReadOnlyList<string> RowFieldNames
    {
        get
        {
            if (DeclaredRowFields is { Count: > 0 } declared)
            {
                var normalized = new List<string>(declared.Count);
                foreach (var field in declared)
                    normalized.Add(MergeFields.NormalizeKey(field));
                return normalized;
            }

            var names = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var row in Rows)
            {
                foreach (var key in row.Keys)
                {
                    var normalized = MergeFields.NormalizeKey(key);
                    if (seen.Add(normalized))
                        names.Add(normalized);
                }
            }

            return names;
        }
    }
}
