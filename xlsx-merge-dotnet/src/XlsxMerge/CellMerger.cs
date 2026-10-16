using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsxMerge;

/// <summary>
/// Substitution of the merge fields in the cells of every worksheet.
/// Formula cells are not modified.
/// </summary>
internal static class CellMerger
{
    private static readonly XLCellsUsedOptions UsedOptions =
        XLCellsUsedOptions.Contents | XLCellsUsedOptions.NormalFormats;

    public static void Merge(
        IXLWorksheet worksheet,
        MergeData data,
        IReadOnlyDictionary<string, object?> globalContext,
        IReadOnlyList<string> rowFields,
        ExpandedBlock? block,
        MergeOptions options)
    {
        foreach (var row in worksheet.RowsUsed(UsedOptions))
        {
            var rowNumber = row.RowNumber();
            var context = block is not null && block.Contains(rowNumber)
                ? MergeFields.BuildRowContext(
                    globalContext, rowFields, data.Rows[rowNumber - block.FirstRow])
                : globalContext;

            foreach (var cell in row.CellsUsed(UsedOptions))
                MergeCell(cell, context, options);
        }
    }

    internal static void MergeCell(
        IXLCell cell,
        IReadOnlyDictionary<string, object?> context,
        MergeOptions options)
    {
        if (cell.HasFormula)
            return; // formulas are not modified
        if (cell.Value.IsBlank || !cell.Value.IsText)
            return; // numbers, dates and booleans never carry a field

        var text = cell.GetString();
        if (text.IndexOf('#') < 0)
            return;

        var value = MergeFields.ReplaceCell(text, context, options, out var changed);
        if (!changed)
            return;

        WriteValue(cell, value, text);
    }

    private static void WriteValue(IXLCell cell, XLCellValue value, string originalText)
    {
        if (!cell.HasRichText)
        {
            cell.Value = value;
            return;
        }

        var richText = cell.GetRichText();
        if (!value.IsText || richText.Count != 1)
        {
            // Multi-run rich text: simple replacement, the internal formatting is lost.
            cell.Value = value;
            return;
        }

        // Single-run rich text: rewrite the text keeping the cell font
        // (ClosedXML 0.105 does not expose the font of existing runs).
        var plain = value.GetText();
        if (plain == originalText)
            return;
        richText.ClearText();
        richText.AddText(plain, cell.Style.Font);
    }
}
