using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace XlsxMerge.IO;

/// <summary>
/// Reader of the generic two-section data format:
/// <code>
/// [globals]
/// COURSE;Advanced Excel
/// DATE;2024-06-01
///
/// [rows]
/// EMPLOYEE_ID;NAME;SCORE
/// 42;Alice;17.5
/// 43;Bob;14
/// </code>
/// The first line of <c>[rows]</c> holds the headers (the row field names).
/// </summary>
public static class CsvMergeDataReader
{
    private enum Section
    {
        None,
        Globals,
        Rows,
    }

    /// <summary>Reads a data file (UTF-8 by default, BOM detection).</summary>
    public static MergeData ReadFile(string path, string? delimiter = null, Encoding? encoding = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var reader = new StreamReader(
            path,
            encoding ?? Encoding.UTF8,
            detectEncodingFromByteOrderMarks: true);
        return Read(reader, delimiter);
    }

    /// <summary>Reads a data file from a text reader.</summary>
    public static MergeData Read(TextReader reader, string? delimiter = null)
    {
        ArgumentNullException.ThrowIfNull(reader);
        var separator = ParseDelimiter(delimiter);

        var data = new MergeData();
        var section = Section.None;
        IReadOnlyList<string>? headers = null;
        var sawAnyRecord = false;

        foreach (var record in CsvReader.ReadRecords(reader, separator))
        {
            if (CsvReader.IsBlank(record))
                continue;

            sawAnyRecord = true;
            var first = record.Fields[0].Trim();

            if (string.Equals(first, "[globals]", StringComparison.OrdinalIgnoreCase))
            {
                section = Section.Globals;
                continue;
            }

            if (string.Equals(first, "[rows]", StringComparison.OrdinalIgnoreCase))
            {
                if (headers is not null)
                    throw new MergeException(
                        $"Line {record.LineNumber}: the [rows] section is already defined.");
                section = Section.Rows;
                continue;
            }

            switch (section)
            {
                case Section.Globals:
                    ReadGlobal(data, record);
                    break;
                case Section.Rows:
                    if (headers is null)
                    {
                        // First non-blank line of [rows]: the headers.
                        headers = ReadHeaders(data, record);
                        continue;
                    }

                    ReadRow(data, record, headers);
                    break;
                default:
                    throw new MergeException(
                        $"Line {record.LineNumber}: a [globals] or [rows] section is expected before any content.");
            }
        }

        if (!sawAnyRecord)
            throw new MergeException("The data file is empty.");

        return data;
    }

    private static void ReadGlobal(MergeData data, CsvRecord record)
    {
        var fields = StripTrailingEmpty(record.Fields);
        if (fields.Count == 0)
            return;
        if (fields.Count > 2)
            throw new MergeException(
                $"Line {record.LineNumber}: too many columns in the [globals] section (key;value expected).");

        var key = fields[0].Trim();
        var value = fields.Count == 2 ? fields[1] : string.Empty;
        try
        {
            data.AddGlobal(key, value);
        }
        catch (MergeException ex)
        {
            throw new MergeException($"Line {record.LineNumber}: {ex.Message}", ex);
        }
    }

    private static IReadOnlyList<string>? ReadHeaders(MergeData data, CsvRecord record)
    {
        var fields = StripTrailingEmpty(record.Fields);
        if (fields.Count == 0)
            return null;

        var headers = new List<string>(fields.Count);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var field in fields)
        {
            string normalized;
            try
            {
                normalized = MergeFields.NormalizeKey(field.Trim());
            }
            catch (MergeException ex)
            {
                throw new MergeException($"Line {record.LineNumber}: {ex.Message}", ex);
            }

            if (!seen.Add(normalized))
                throw new MergeException(
                    $"Line {record.LineNumber}: duplicate header '{normalized}'.");
            headers.Add(normalized);
        }

        data.DeclaredRowFields = headers;
        return headers;
    }

    private static void ReadRow(MergeData data, CsvRecord record, IReadOnlyList<string> headers)
    {
        var fields = StripTrailingEmpty(record.Fields);
        if (fields.Count > headers.Count)
            throw new MergeException(
                $"Line {record.LineNumber}: {fields.Count} columns but the header row has {headers.Count}.");

        var row = new Dictionary<string, object?>(StringComparer.Ordinal);
        for (var i = 0; i < headers.Count; i++)
            row[headers[i]] = i < fields.Count && fields[i].Length > 0 ? fields[i] : null;
        data.Rows.Add(row);
    }

    private static IReadOnlyList<string> StripTrailingEmpty(IReadOnlyList<string> fields)
    {
        var end = fields.Count;
        while (end > 0 && fields[end - 1].Length == 0)
            end--;
        if (end == fields.Count)
            return fields;
        var result = new List<string>(end);
        for (var i = 0; i < end; i++)
            result.Add(fields[i]);
        return result;
    }

    private static char ParseDelimiter(string? delimiter)
    {
        if (delimiter is null)
            return ';';
        if (delimiter.Length != 1)
            throw new MergeException("The delimiter must be a single character (e.g. ';').");
        return delimiter[0];
    }
}
