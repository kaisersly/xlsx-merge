using System;
using System.IO;
using ClosedXML.Excel;

namespace XlsxMerge;

/// <summary>
/// Mail-merge: applies a data set (<see cref="MergeData"/>) to an .xlsx template.
/// Every worksheet is processed: repeatable block expansion, field substitution in the
/// cells then in the headers and footers.
/// </summary>
public static class XlsxMerger
{
    /// <summary>Mail-merges a template into an output (both streams are left open).</summary>
    public static void Render(Stream template, Stream output, MergeData data, MergeOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(data);

        options ??= new MergeOptions();
        MergeFields.Validate(data);

        using var workbook = new XLWorkbook(template);
        RenderWorkbook(workbook, data, options);
        workbook.SaveAs(output);
    }

    /// <summary>Mail-merges a template into an in-memory output.</summary>
    public static byte[] RenderToArray(Stream template, MergeData data, MergeOptions? options = null)
    {
        using var output = new MemoryStream();
        Render(template, output, data, options);
        return output.ToArray();
    }

    /// <summary>
    /// Mail-merges a template file into an output file. The output file is only written
    /// when the render succeeds.
    /// </summary>
    public static void RenderFile(
        string templatePath,
        string outputPath,
        MergeData data,
        MergeOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(templatePath);
        ArgumentException.ThrowIfNullOrEmpty(outputPath);
        ArgumentNullException.ThrowIfNull(data);

        options ??= new MergeOptions();
        MergeFields.Validate(data);

        byte[] rendered;
        using (var template = File.OpenRead(templatePath))
        using (var workbook = new XLWorkbook(template))
        {
            RenderWorkbook(workbook, data, options);
            using var output = new MemoryStream();
            workbook.SaveAs(output);
            rendered = output.ToArray();
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);
        File.WriteAllBytes(outputPath, rendered);
    }

    private static void RenderWorkbook(XLWorkbook workbook, MergeData data, MergeOptions options)
    {
        var globalContext = MergeFields.BuildGlobalContext(data);
        var rowFields = data.RowFieldNames;

        // Snapshot: temporary worksheets are created/deleted while processing.
        foreach (var worksheet in workbook.Worksheets.ToArray())
        {
            var block = RowExpander.Expand(workbook, worksheet, data, options);
            CellMerger.Merge(worksheet, data, globalContext, rowFields, block, options);
            HeaderFooterMerger.Merge(worksheet, globalContext, options);
        }
    }
}
