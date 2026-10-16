using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsxMerge;

/// <summary>
/// Substitution of the merge fields in the headers and footers of each worksheet, for the
/// odd, even and first-page occurrences (the <c>DifferentOddEvenPagesOnHF</c> /
/// <c>DifferentFirstPageOnHF</c> flags are kept as-is).
/// </summary>
internal static class HeaderFooterMerger
{
    private static readonly XLHFOccurrence[] Occurrences =
    {
        XLHFOccurrence.OddPages,
        XLHFOccurrence.EvenPages,
        XLHFOccurrence.FirstPage,
    };

    public static void Merge(
        IXLWorksheet worksheet,
        IReadOnlyDictionary<string, object?> globalContext,
        MergeOptions options)
    {
        var headerFooters = new IXLHeaderFooter[] { worksheet.PageSetup.Header, worksheet.PageSetup.Footer };
        foreach (var headerFooter in headerFooters)
        {
            var items = new IXLHFItem[] { headerFooter.Left, headerFooter.Center, headerFooter.Right };
            foreach (var item in items)
            {
                foreach (var occurrence in Occurrences)
                {
                    var text = item.GetText(occurrence);
                    if (text.Length == 0)
                        continue;

                    var replaced = MergeFields.ReplaceText(text, globalContext, options, out var changed);
                    if (!changed)
                        continue;

                    item.Clear(occurrence);
                    item.AddText(replaced, occurrence);
                }
            }
        }
    }
}
