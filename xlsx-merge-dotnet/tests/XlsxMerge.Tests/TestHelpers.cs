using ClosedXML.Excel;

namespace XlsxMerge.Tests;

internal static class TestHelpers
{
    public static byte[] ToBytes(XLWorkbook workbook)
    {
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public static byte[] Render(XLWorkbook template, MergeData data, MergeOptions? options = null)
    {
        using var input = new MemoryStream(ToBytes(template));
        using var output = new MemoryStream();
        XlsxMerger.Render(input, output, data, options);
        return output.ToArray();
    }

    public static XLWorkbook Load(byte[] bytes) => new(new MemoryStream(bytes));

    public static XLWorkbook RenderAndLoad(
        XLWorkbook template,
        MergeData data,
        MergeOptions? options = null) =>
        Load(Render(template, data, options));

    public static bool IsMerged(
        IXLWorksheet sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        foreach (var range in sheet.MergedRanges)
        {
            var address = range.RangeAddress;
            if (address.FirstAddress.RowNumber == firstRow
                && address.FirstAddress.ColumnNumber == firstColumn
                && address.LastAddress.RowNumber == lastRow
                && address.LastAddress.ColumnNumber == lastColumn)
                return true;
        }

        return false;
    }
}
