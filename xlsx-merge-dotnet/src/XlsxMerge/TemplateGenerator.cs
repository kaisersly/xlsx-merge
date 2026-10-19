using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace XlsxMerge;

/// <summary>
/// Generates ready-to-use .xlsx templates from a <see cref="TemplateSpec"/>: title row,
/// optional global-field hint row, merged element headers, outline groups per column
/// group (collapsible), field labels and a marker row containing every row field — the
/// layout the <see cref="XlsxMerger"/> expects to expand.
/// </summary>
public static class TemplateGenerator
{
    /// <summary>OOXML <c>textRotation</c> 90° counterclockwise = « Rotate Text Up ».</summary>
    private const int TextRotationUp = 90;

    private static readonly XLColor TitleFill = XLColor.FromHtml("#1F4E79");
    private static readonly XLColor ElementFill = XLColor.FromHtml("#DDEBF7");
    private static readonly XLColor GroupFill = XLColor.FromHtml("#BDD7EE");
    private static readonly XLColor LabelFill = XLColor.FromHtml("#F2F2F2");
    private static readonly XLColor MarkerFill = XLColor.FromHtml("#FFF2CC");
    private static readonly XLColor GridColor = XLColor.FromHtml("#9CA3AF");

    /// <summary>Generates a template into an output stream (left open).</summary>
    public static void Generate(TemplateSpec spec, Stream output)
    {
        ArgumentNullException.ThrowIfNull(spec);
        ArgumentNullException.ThrowIfNull(output);
        Validate(spec);

        using var workbook = new XLWorkbook();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheetSpec in spec.Sheets)
            BuildSheet(workbook, sheetSpec, spec, usedNames);
        workbook.SaveAs(output);
    }

    /// <summary>Generates a template file (written only when generation succeeds).</summary>
    public static void GenerateFile(TemplateSpec spec, string outputPath)
    {
        ArgumentException.ThrowIfNullOrEmpty(outputPath);
        byte[] rendered;
        using (var output = new MemoryStream())
        {
            Generate(spec, output);
            rendered = output.ToArray();
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);
        File.WriteAllBytes(outputPath, rendered);
    }

    private static void Validate(TemplateSpec spec)
    {
        if (spec.Sheets.Count == 0)
            throw new MergeException("A template needs at least one sheet.");

        foreach (var sheet in spec.Sheets)
        {
            if (string.IsNullOrWhiteSpace(sheet.Name))
                throw new MergeException("Template sheet names cannot be empty.");
            if (sheet.Columns.Count == 0)
                throw new MergeException(
                    $"Template sheet '{sheet.Name}' has no column.");

            foreach (var column in sheet.Columns)
            {
                try
                {
                    MergeFields.NormalizeKey(column.Field);
                }
                catch (MergeException ex)
                {
                    throw new MergeException(
                        $"Invalid field in template sheet '{sheet.Name}': {ex.Message}", ex);
                }
            }
        }
    }

    private static void BuildSheet(
        XLWorkbook workbook, TemplateSheetSpec sheetSpec, TemplateSpec spec, HashSet<string> usedNames)
    {
        var ws = workbook.Worksheets.Add(SanitizeSheetName(sheetSpec.Name, usedNames));
        var columns = sheetSpec.Columns;
        var count = columns.Count;

        var row = 1;

        // Row 1 — title.
        var title = string.IsNullOrWhiteSpace(sheetSpec.Title) ? sheetSpec.Name : sheetSpec.Title!;
        var titleRange = Put(ws, row, 1, count, title);
        titleRange.Style.Font.Bold = true;
        titleRange.Style.Font.FontSize = 14;
        titleRange.Style.Font.FontColor = XLColor.White;
        titleRange.Style.Fill.BackgroundColor = TitleFill;
        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Row(row).Height = 28;
        row++;

        // Row 2 — global fields hint.
        if (spec.GlobalHints.Count > 0)
        {
            var hints = string.Join("  ·  ",
                spec.GlobalHints.Select(field => $"#{MergeFields.NormalizeKey(field)}#"));
            var hintRange = Put(ws, row, 1, count, hints);
            hintRange.Style.Font.FontColor = XLColor.FromHtml("#6B7280");
            hintRange.Style.Font.Italic = true;
            hintRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            ws.Row(row).Height = 18;
            row++;
        }

        // Element header row (merged runs), when at least one column has an element.
        // Same treatment as the field labels: every run that is not purely identity is
        // rotated up (« Rotate Text Up ») and the row grows to fit the longest title;
        // the identity run (« Identification ») stays flat.
        if (columns.Any(c => !string.IsNullOrEmpty(c.Element)))
        {
            var maxRotatedElement = WriteRuns(
                ws, row, columns, c => c.Element, ElementFill, bold: true,
                rotateNonIdentityRuns: true);
            ws.Row(row).Height = maxRotatedElement == 0
                ? 20
                : Math.Max(20, Math.Clamp(maxRotatedElement * 7.0 + 6, 40, 300));
            row++;
        }

        // Column group header row (merged runs) + Excel outline groups.
        if (columns.Any(c => !string.IsNullOrEmpty(c.Group)))
        {
            WriteRuns(ws, row, columns, c => c.Group, GroupFill, bold: true);
            foreach (var run in ContiguousRuns(columns, c => c.Group))
            {
                if (string.IsNullOrEmpty(run.Label))
                    continue;
                ws.Columns(run.First, run.Last).Group();
            }

            ws.Row(row).Height = 20;
            row++;
        }

        // Field labels. Headers of non-identity columns are rotated up (« Rotation du
        // texte vers le haut », OOXML textRotation 90° counterclockwise) so a long label
        // does not dictate the column width; the label row grows to fit the longest one.
        var rotatedLabelLength = 0;
        for (var i = 0; i < count; i++)
        {
            var cell = ws.Cell(row, i + 1);
            var column = columns[i];
            var label = string.IsNullOrWhiteSpace(column.Label) ? column.Field : column.Label!;
            cell.Value = label;
            cell.Style.Font.Italic = true;
            cell.Style.Font.FontColor = XLColor.FromHtml("#374151");
            cell.Style.Fill.BackgroundColor = LabelFill;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            if (!column.Identity)
            {
                cell.Style.Alignment.TextRotation = TextRotationUp;
                rotatedLabelLength = Math.Max(rotatedLabelLength, label.Length);
            }
        }

        // A rotated label is read along the row height: grow the row to fit it.
        ws.Row(row).Height = rotatedLabelLength == 0
            ? 30
            : Math.Max(30, Math.Clamp(rotatedLabelLength * 7.0 + 6, 40, 300));
        row++;

        // Marker row — every row field, which is what the merger detects.
        for (var i = 0; i < count; i++)
        {
            var cell = ws.Cell(row, i + 1);
            cell.Value = $"#{MergeFields.NormalizeKey(columns[i].Field)}#";
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = MarkerFill;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        ws.Row(row).Height = 22;
        var markerRow = row;

        // Grid over the whole header block.
        var block = ws.Range(1, 1, markerRow, count);
        block.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
        block.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
        block.Style.Border.SetOutsideBorderColor(GridColor);
        block.Style.Border.SetInsideBorderColor(GridColor);

        // Freezing: only the rows ABOVE the marker row stay visible — the marker row
        // itself becomes the first data row after a merge and must scroll normally.
        ws.SheetView.FreezeRows(markerRow - 1);
        var identityColumns = 0;
        while (identityColumns < count && columns[identityColumns].Identity)
            identityColumns++;
        if (identityColumns > 0)
            ws.SheetView.FreezeColumns(identityColumns);

        // Deterministic widths (AdjustToContents depends on fonts/DPI).
        for (var i = 0; i < count; i++)
        {
            var column = columns[i];
            if (!column.Identity)
            {
                // Rotated headers: fixed narrow width.
                ws.Column(i + 1).Width = 4;
                continue;
            }
            var labelLength = (string.IsNullOrWhiteSpace(column.Label) ? column.Field : column.Label!).Length;
            var fieldLength = $"#{MergeFields.NormalizeKey(column.Field)}#".Length;
            ws.Column(i + 1).Width = Math.Clamp(Math.Max(labelLength, fieldLength) + 3, 10, 34);
        }
    }

    /// <summary>
    /// Sets one merged (or single) run of headers and its value/style. When
    /// <paramref name="rotateNonIdentityRuns"/> is set, runs that are not purely
    /// identity are rotated up (« Rotate Text Up »); returns the length of the longest
    /// rotated label (0 when nothing was rotated) so the caller can size the row.
    /// </summary>
    private static int WriteRuns(
        IXLWorksheet ws,
        int row,
        IList<TemplateColumnSpec> columns,
        Func<TemplateColumnSpec, string?> selector,
        XLColor fill,
        bool bold,
        bool rotateNonIdentityRuns = false)
    {
        var maxRotatedLabelLength = 0;
        foreach (var run in ContiguousRuns(columns, selector))
        {
            if (string.IsNullOrEmpty(run.Label))
                continue;
            var range = run.First == run.Last
                ? ws.Range(row, run.First, row, run.First)
                : ws.Range(row, run.First, row, run.Last);
            if (run.First != run.Last)
                range.Merge();
            range.Cell(1, 1).Value = run.Label;
            range.Style.Fill.BackgroundColor = fill;
            range.Style.Font.Bold = bold;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            if (rotateNonIdentityRuns && !RunIsIdentity(columns, run))
            {
                range.Style.Alignment.TextRotation = TextRotationUp;
                maxRotatedLabelLength = Math.Max(maxRotatedLabelLength, run.Label.Length);
            }
        }

        return maxRotatedLabelLength;
    }

    /// <summary>True when every column of the run is an identity column.</summary>
    private static bool RunIsIdentity(IList<TemplateColumnSpec> columns, Run run)
    {
        for (var column = run.First; column <= run.Last; column++)
        if (!columns[column - 1].Identity)
            return false;
        return true;
    }

    private readonly record struct Run(int First, int Last, string? Label);

    private static IEnumerable<Run> ContiguousRuns(
        IList<TemplateColumnSpec> columns, Func<TemplateColumnSpec, string?> selector)
    {
        var first = 0;
        for (var i = 1; i <= columns.Count; i++)
        {
            var continues = i < columns.Count
                            && string.Equals(selector(columns[i]), selector(columns[first]),
                                StringComparison.Ordinal);
            if (continues)
                continue;
            yield return new Run(first + 1, i, selector(columns[first]));
            first = i;
        }
    }

    private static IXLRange Put(IXLWorksheet ws, int row, int firstColumn, int count, string value)
    {
        var range = count == 1
            ? ws.Range(row, firstColumn, row, firstColumn)
            : ws.Range(row, firstColumn, row, firstColumn + count - 1);
        if (count > 1)
            range.Merge();
        range.Cell(1, 1).Value = value;
        return range;
    }

    private static string SanitizeSheetName(string name, HashSet<string> used)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        var builder = new StringBuilder(name.Length);
        foreach (var character in name)
            builder.Append(Array.IndexOf(invalid, character) >= 0 ? ' ' : character);

        var clean = builder.ToString().Trim();
        if (clean.Length == 0)
            clean = "Sheet";
        if (clean.Length > 31)
            clean = clean.Substring(0, 31);

        var candidate = clean;
        var suffix = 2;
        while (!used.Add(candidate))
        {
            var suffixText = $" ({suffix++})";
            candidate = clean.Substring(0, Math.Min(clean.Length, 31 - suffixText.Length))
                        + suffixText;
        }

        return candidate;
    }
}
