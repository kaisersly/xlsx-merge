using System;
using System.Collections.Generic;

namespace XlsxMerge;

/// <summary>One column of a generated template (header block + marker row).</summary>
public sealed class TemplateColumnSpec
{
    /// <summary>Row field name, with or without <c>#</c> (written as <c>#FIELD#</c> in the marker row).</summary>
    public string Field { get; set; } = string.Empty;

    /// <summary>Human-readable label (defaults to <see cref="Field"/>).</summary>
    public string? Label { get; set; }

    /// <summary>Optional element label: contiguous columns sharing it get a merged header cell.</summary>
    public string? Element { get; set; }

    /// <summary>
    /// Optional group label (e.g. a session): contiguous columns sharing it get a merged
    /// header cell <em>and</em> an Excel outline group (collapsible columns).
    /// </summary>
    public string? Group { get; set; }

    /// <summary>Identity column (leading columns, kept frozen when scrolling right).</summary>
    public bool Identity { get; set; }
}

/// <summary>One worksheet of a generated template.</summary>
public sealed class TemplateSheetSpec
{
    /// <summary>Worksheet name (sanitized: invalid characters, 31 chars max).</summary>
    public string Name { get; set; } = "Sheet";

    /// <summary>Title row text (defaults to <see cref="Name"/>).</summary>
    public string? Title { get; set; }

    public IList<TemplateColumnSpec> Columns { get; } = new List<TemplateColumnSpec>();
}

/// <summary>
/// Description of a template to generate: one worksheet per <see cref="TemplateSheetSpec"/>,
/// each with a title row, optional global-field hints, merged element/group headers,
/// field labels and a marker row holding every row field.
/// </summary>
public sealed class TemplateSpec
{
    public IList<TemplateSheetSpec> Sheets { get; } = new List<TemplateSheetSpec>();

    /// <summary>Global fields shown as a hint row under the title (e.g. FORMATION, DATE).</summary>
    public IList<string> GlobalHints { get; } = new List<string>();
}
