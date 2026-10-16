using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace XlsxMerge.IO;

/// <summary>CSV record: fields plus the physical line number (1-based, start of the record).</summary>
public sealed class CsvRecord
{
    public CsvRecord(int lineNumber, IReadOnlyList<string> fields) =>
        (LineNumber, Fields) = (lineNumber, fields);

    public int LineNumber { get; }
    public IReadOnlyList<string> Fields { get; }
}

/// <summary>
/// Minimal CSV reader (RFC 4180): quotes, doubled quotes, multi-line fields,
/// defined by a single delimiter.
/// </summary>
public static class CsvReader
{
    /// <summary>Reads every record of a CSV text.</summary>
    public static IReadOnlyList<CsvRecord> ReadRecords(TextReader reader, char delimiter = ';')
    {
        ArgumentNullException.ThrowIfNull(reader);
        return Parse(reader.ReadToEnd(), delimiter);
    }

    /// <summary>Reads every record of a CSV string.</summary>
    public static IReadOnlyList<CsvRecord> Parse(string text, char delimiter = ';')
    {
        var records = new List<CsvRecord>();
        var fields = new List<string>();
        var field = new StringBuilder();
        var lineNumber = 1;
        var inQuotes = false;
        var quoteClosed = false;

        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        field.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                        quoteClosed = true;
                    }
                }
                else
                {
                    if (c == '\n')
                        lineNumber++;
                    field.Append(c);
                }

                continue;
            }

            if (quoteClosed && c != delimiter && c != '\r' && c != '\n')
                throw new MergeException(
                    $"Line {lineNumber}: unexpected character '{c}' after a closing quote.");

            if (c == '"' && field.Length == 0)
            {
                inQuotes = true;
            }
            else if (c == delimiter)
            {
                fields.Add(field.ToString());
                field.Clear();
                quoteClosed = false;
            }
            else if (c == '\r' || c == '\n')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i++;
                fields.Add(field.ToString());
                field.Clear();
                quoteClosed = false;
                records.Add(new CsvRecord(lineNumber, fields));
                fields = new List<string>();
                lineNumber++;
            }
            else
            {
                field.Append(c);
            }
        }

        if (inQuotes)
            throw new MergeException($"Line {lineNumber}: unterminated quote.");

        if (field.Length > 0 || fields.Count > 0)
        {
            fields.Add(field.ToString());
            records.Add(new CsvRecord(lineNumber, fields));
        }

        return records;
    }

    /// <summary>True when the record is empty (no non-blank field).</summary>
    public static bool IsBlank(CsvRecord record)
    {
        foreach (var field in record.Fields)
        if (!string.IsNullOrWhiteSpace(field))
            return false;
        return true;
    }
}
