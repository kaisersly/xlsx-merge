using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace XlsxMerge;

/// <summary>Repeatable block after expansion: rows <c>FirstRow .. FirstRow + RowCount - 1</c>.</summary>
internal sealed class ExpandedBlock
{
    public ExpandedBlock(int firstRow, int rowCount) => (FirstRow, RowCount) = (firstRow, rowCount);

    public int FirstRow { get; }
    public int RowCount { get; }

    public bool Contains(int rowNumber) =>
        RowCount > 0 && rowNumber >= FirstRow && rowNumber <= FirstRow + RowCount - 1;
}

/// <summary>
/// Expansion of the repeatable block: the marker row (and optionally a second style-model row)
/// is replaced by as many rows as there are data rows, with alternating styles.
/// </summary>
internal static class RowExpander
{
    private static readonly XLCellsUsedOptions UsedOptions =
        XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats;

    public static ExpandedBlock? Expand(
        XLWorkbook workbook,
        IXLWorksheet worksheet,
        MergeData data,
        MergeOptions options)
    {
        var rowFields = data.RowFieldNames;
        var markerRow = DetectMarkerRow(worksheet, rowFields, options);
        if (markerRow is null)
            return null; // no repeatable block on this sheet: global fields only
        var marker = markerRow.Value;

        var blockRowCount = ResolveBlockRowCount(worksheet, marker, options);
        var rowCount = data.Rows.Count;

        // Columns actually present in the model rows (like ExcelJS did: only existing
        // cells are duplicated).
        var columns = new SortedSet<int>();
        for (var i = 0; i < blockRowCount; i++)
        foreach (var cell in worksheet.Row(marker + i).CellsUsed(UsedOptions))
            columns.Add(cell.Address.ColumnNumber);

        if (columns.Count == 0)
            throw new MergeException(
                $"Model row {marker} of sheet '{worksheet.Name}' contains no cell.");

        // Merges entirely inside the model block: they are replicated on every generated
        // row (they do not survive the deletion of the model rows).
        bool IsInsideBlock(IXLRange range) =>
            range.RangeAddress.FirstAddress.RowNumber >= marker
            && range.RangeAddress.LastAddress.RowNumber <= marker + blockRowCount - 1;

        var blockMerges = worksheet.MergedRanges
            .Where(IsInsideBlock)
            .Select(range => (FirstRow: range.RangeAddress.FirstAddress.RowNumber,
                LastRow: range.RangeAddress.LastAddress.RowNumber,
                FirstColumn: range.RangeAddress.FirstAddress.ColumnNumber,
                LastColumn: range.RangeAddress.LastAddress.ColumnNumber))
            .ToList();
        foreach (var range in worksheet.MergedRanges.Where(IsInsideBlock).ToList())
            range.Unmerge();

        // Heights of the model rows.
        var heights = new double[blockRowCount];
        for (var i = 0; i < blockRowCount; i++)
            heights[i] = worksheet.Row(marker + i).Height;

        // Copy of the model rows into a temporary worksheet: ClosedXML styles are proxies
        // tied to cells, they cannot be "captured" any other way than by copying them.
        var temporaryName = CreateTemporaryWorksheet(workbook);
        var temporary = workbook.Worksheet(temporaryName);
        try
        {
            for (var i = 0; i < blockRowCount; i++)
            {
                foreach (var column in columns)
                {
                    // Style + value copy (CopyFrom would crash on conditional formats).
                    var source = worksheet.Cell(marker + i, column);
                    var target = temporary.Cell(i + 1, column);
                    target.Style = source.Style;
                    if (source.HasFormula)
                        target.FormulaA1 = source.FormulaA1;
                    else
                        target.Value = source.Value;
                }

                temporary.Row(i + 1).Height = heights[i];
            }

            // Insertion of the generated rows above the model rows.
            if (rowCount > 0)
                worksheet.Row(marker).InsertRowsAbove(rowCount);

            // Fill: values of the marker row, alternating styles (row i ⇒ style of model
            // i % b; row 0 takes the marker row style).
            for (var i = 0; i < rowCount; i++)
            {
                var targetRowNumber = marker + i;
                var template = i % blockRowCount;
                foreach (var column in columns)
                {
                    var styleSource = temporary.Cell(template + 1, column);
                    var valueSource = temporary.Cell(1, column);
                    var target = worksheet.Cell(targetRowNumber, column);
                    target.Style = styleSource.Style;
                    if (valueSource.HasFormula)
                        target.FormulaA1 = valueSource.FormulaA1;
                    else
                        target.Value = valueSource.Value;
                }

                worksheet.Row(targetRowNumber).Height = heights[template];
            }

            // Deletion of the model rows shifted below the generated block.
            for (var i = 0; i < blockRowCount; i++)
                worksheet.Row(marker + rowCount).Delete();
        }
        finally
        {
            temporary.Delete();
        }

        // Replication of the model block merges on the generated rows.
        foreach (var merge in blockMerges)
        {
            if (rowCount == 0)
                continue;

            if (merge.FirstRow == merge.LastRow)
            {
                var relativeRow = merge.FirstRow - marker;
                for (var i = 0; i < rowCount; i++)
                    worksheet
                        .Range(marker + i + relativeRow, merge.FirstColumn,
                            marker + i + relativeRow, merge.LastColumn)
                        .Merge();
            }
            else
            {
                // Merge spanning several model rows: it applies to the whole generated block.
                worksheet
                    .Range(marker, merge.FirstColumn, marker + rowCount - 1, merge.LastColumn)
                    .Merge();
            }
        }

        return new ExpandedBlock(marker, rowCount);
    }

    /// <summary>
    /// Detects the marker row: when a marker field is forced, the row containing it;
    /// otherwise the row containing the largest number of distinct row fields.
    /// Returns <c>null</c> when the sheet holds no row field at all — the sheet then
    /// only gets its global fields substituted (no repeatable block).
    /// </summary>
    internal static int? DetectMarkerRow(
        IXLWorksheet worksheet,
        IReadOnlyList<string> rowFields,
        MergeOptions options)
    {
        var forcedMarker = options.NormalizedMarkerField;
        if (forcedMarker is not null)
        {
            var rows = FindRowsWithField(worksheet, forcedMarker);
            if (rows.Count == 0)
                throw new MergeException(
                    $"Sheet '{worksheet.Name}' is missing a row containing #{forcedMarker}#.");
            if (rows.Count > 1)
                throw new MergeException(
                    $"Only one row may contain #{forcedMarker}# in sheet '{worksheet.Name}' (rows {string.Join(", ", rows)}).");
            return rows[0];
        }

        if (rowFields.Count == 0)
            return null; // no row field defined: nothing to repeat anywhere

        var fieldSet = new HashSet<string>(rowFields, StringComparer.Ordinal);
        var counts = new Dictionary<int, HashSet<string>>();
        foreach (var row in worksheet.RowsUsed(UsedOptions))
        {
            var found = new HashSet<string>(StringComparer.Ordinal);
            foreach (var cell in row.CellsUsed(UsedOptions))
            {
                if (cell.HasFormula || !cell.Value.IsText)
                    continue;
                var text = cell.GetString();
                if (text.IndexOf('#') < 0)
                    continue;
                foreach (Match match in MergeFields.Pattern.Matches(text))
                {
                    var key = match.Groups[1].Value.Trim();
                    if (fieldSet.Contains(key))
                        found.Add(key);
                }
            }

            if (found.Count > 0)
                counts[row.RowNumber()] = found;
        }

        if (counts.Count == 0)
            return null; // sheet without marker row: global fields only

        var max = counts.Values.Max(set => set.Count);
        var candidates = counts.Where(pair => pair.Value.Count == max).Select(pair => pair.Key).OrderBy(x => x).ToList();
        if (candidates.Count > 1)
            throw new MergeException(
                $"Multiple candidate rows for the repeatable block of sheet '{worksheet.Name}' " +
                $"(rows {string.Join(", ", candidates)}). Set the marker field (--marker).");

        return candidates[0];
    }

    /// <summary>Number of model rows in the block (1 or 2; the 2nd one is the alternating-style model).</summary>
    internal static int ResolveBlockRowCount(IXLWorksheet worksheet, int markerRow, MergeOptions options)
    {
        var usedRows = worksheet.RowsUsed(UsedOptions).ToDictionary(row => row.RowNumber());
        var nextRow = usedRows.TryGetValue(markerRow + 1, out var row) ? row : null;

        switch (options.RepeatBlockRows)
        {
            case RepeatBlockRows.One:
                return 1;
            case RepeatBlockRows.Two:
                if (nextRow is null)
                    throw new MergeException(
                        $"The second style-model row of the block (row {markerRow + 1}) was not found in sheet '{worksheet.Name}'.");
                return 2;
            default:
                // Auto: the next row is a style-model row when it exists, holds cells
                // (styling) but no value.
                if (nextRow is null)
                    return 1;
                var hasValues = nextRow.CellsUsed(UsedOptions).Any(cell => !cell.Value.IsBlank);
                return hasValues ? 1 : 2;
        }
    }

    private static List<int> FindRowsWithField(IXLWorksheet worksheet, string field)
    {
        var rows = new List<int>();
        foreach (var row in worksheet.RowsUsed(UsedOptions))
        foreach (var cell in row.CellsUsed(UsedOptions))
        {
            if (cell.HasFormula || !cell.Value.IsText)
                continue;
            var text = cell.GetString();
            if (text.IndexOf('#') < 0)
                continue;
            foreach (Match match in MergeFields.Pattern.Matches(text))
            {
                if (match.Groups[1].Value.Trim() == field)
                {
                    rows.Add(row.RowNumber());
                    break;
                }
            }
        }

        return rows;
    }

    private static string CreateTemporaryWorksheet(XLWorkbook workbook)
    {
        var baseName = "__xlsx_merge_tmp__";
        var name = baseName;
        var suffix = 1;
        while (workbook.Worksheets.Contains(name))
            name = baseName + suffix++;
        workbook.Worksheets.Add(name);
        return name;
    }
}
