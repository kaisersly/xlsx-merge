# XlsxMerge (.NET)

Mail-merge for `.xlsx` templates using [ClosedXML](https://github.com/ClosedXML/ClosedXML).
This library is **domain-agnostic**: it knows nothing about any specific use case, it only
substitutes `#FIELD#` merge fields and expands a repeatable block.

It processes **every worksheet** and also rewrites **headers and footers** (odd / even /
first-page occurrences).

## Layout

```
src/XlsxMerge/        The library (.NET 8, ClosedXML 0.105.1)
src/XlsxMerge.Cli/    Command line tool (commands: render, inspect)
tests/XlsxMerge.Tests/  Unit tests + an end-to-end scenario (training attendance sheet)
samples/              Example data file
```

## Build and test

Requires the [.NET SDK 8](https://dotnet.microsoft.com/download/dotnet/8.0).

```bash
dotnet build
dotnet test
```

## Template model

In the `.xlsx` template:

- **fields** are written `#FIELD#` in cells, headers and footers. A field unknown to the
  data set is **left untouched**;
- the **repeatable block** is a row containing a row field (e.g. `#EMPLOYEE_ID#`). The row
  directly below is treated as a **second style-model row** (banding) if it exists, holds
  cells (styling) but **no value**; otherwise the block is a single row. Force it with
  `RepeatBlockRows` (`One` / `Two`);
- generated rows take the **values of the marker row** and alternate the **styles** of the
  model rows (row 0 ⇒ marker style, row 1 ⇒ second-model style, …);
- **merges** of the block are replicated on every generated row; those below the block are
  moved automatically.

### Marker row detection

By default it is the row containing the **largest number of distinct row fields**
(ambiguity ⇒ explicit error asking for `--marker`). The marker field can be forced with
`MergeOptions.RowMarkerField` / `--marker EMPLOYEE_ID`.

A sheet **may not contain any marker row** (cover page, legend…): in that case only the
global fields are substituted on that sheet — no block expansion, the structure is left
untouched. This is also what happens when the data set defines no row field at all.
A forced marker (`RowMarkerField`) is still an error when missing, since it was asked for
explicitly.

## Data format (sections `[globals]` / `[rows]`)

```csv
[globals]
COURSE;Advanced Excel
DATE;2024-06-01

[rows]
EMPLOYEE_ID;NAME;SCORE
42;Alice;17.5
43;Bob;8
```

- UTF-8 (BOM accepted), default delimiter `;` (configurable), RFC 4180 quotes;
- the first non-blank line of `[rows]` holds the **row field names**;
- `[globals]` lines are `KEY;VALUE`;
- keys may be written with or without `#` (`#COURSE#` ≡ `COURSE`).

### Values and typing

- a cell that is **exactly** one field receives the **typed** value (number, bool, date);
- strictly numeric strings (`17.5`, `-3.5`) become numbers (disable with
  `NumericTextParsing.None` / `--numbers none`);
- inside composed text (`Score : #SCORE# / 20`), numbers are formatted with the requested
  culture (`--culture fr-FR` ⇒ decimal comma);
- a missing/`null` value renders an empty cell; **formulas** are not modified (but copied
  as-is into the repeatable block).

## Library usage

```csharp
using XlsxMerge;
using XlsxMerge.IO;

var data = CsvMergeDataReader.ReadFile("data.csv");
data.AddGlobal("LOCATION", "Room B12");        // fields added from code
data.AddRow(new Dictionary<string, object?>   // rows added from code
{
    ["EMPLOYEE_ID"] = "42",
    ["NAME"] = "Alice",
    ["SCORE"] = 17.5,                          // typed value
});

XlsxMerger.RenderFile("template.xlsx", "output.xlsx", data, new MergeOptions
{
    RowMarkerField = "EMPLOYEE_ID",            // optional
    NumberCulture = new CultureInfo("fr-FR"),
});
```

`XlsxMerger.Render(Stream, Stream, …)` and `RenderToArray(…)` are also available.

## Template generation

`TemplateGenerator` builds ready-to-use templates from a `TemplateSpec` — useful when the
data columns are only known at runtime (e.g. generated from a data file):

```csharp
var spec = new TemplateSpec();
spec.GlobalHints.Add("COURSE");                 // hint row under the title

var sheet = new TemplateSheetSpec { Name = "Session 1" };
sheet.Columns.Add(new TemplateColumnSpec       // frozen identity columns
    { Field = "EMPLOYEE_ID", Label = "ID", Element = "People", Identity = true });
sheet.Columns.Add(new TemplateColumnSpec
    { Field = "SCORE", Label = "Score", Element = "Result", Group = "Session 1" });
spec.Sheets.Add(sheet);

TemplateGenerator.GenerateFile(spec, "template.xlsx");
```

Each sheet gets: a title row, an optional global-fields hint row, merged element headers,
merged **column-group headers with a collapsible Excel outline** (one group per `Group`
label), field labels, and a marker row holding every row field — ready to be expanded by
`XlsxMerger`.

Field labels **and element titles of non-identity columns are rotated up** (« Rotate
Text Up », `textRotation` 90°): long headers no longer dictate the column width, and
those header rows automatically grow to keep them readable. Identity columns
(« Identification ») stay horizontal. Non-identity columns use a **fixed width of 4**;
identity columns keep a content-driven width. The layout is self-contained (no dependency on any existing template).

## CLI

```bash
dotnet run --project src/XlsxMerge.Cli -- render \
  --template template.xlsx --data samples/attendance.csv --out output.xlsx \
  [--marker EMPLOYEE_ID] [--block auto|1|2] [--numbers strict|none] \
  [--culture fr-FR] [--delimiter ';']

# List the merge fields of a template (helps building it):
dotnet run --project src/XlsxMerge.Cli -- inspect --template template.xlsx
```

## Testing notes

The test suite is fully synthetic — templates are built in code, data sets come from
inline CSV. The end-to-end scenario (`AttendanceScenarioTests`) documents the expected
behaviour of the whole pipeline on a neutral use case (a training attendance sheet).

## Known limitations

- ClosedXML rewrites the file: template-specific binary metadata (print settings, …) may
  be regenerated;
- the internal formatting of **multi-run rich text** is not preserved when replacing the
  text (replaced as plain text).
