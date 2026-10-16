"""Python wrapper for the .NET ``XlsxMerge`` mail-merge library (via pythonnet).

It loads the .NET runtime (CoreCLR) and calls the library directly:

    >>> from xlsx_merge import render
    >>> render("template.xlsx", "data.csv", "output.xlsx")

The wrapper is domain-agnostic: it only substitutes ``#FIELD#`` merge fields and
expands a repeatable block, exactly like the .NET library.

Loading the .NET DLL
--------------------
Point the wrapper at a directory holding ``XlsxMerge.dll`` (and its dependency
``ClosedXML.dll``) using the ``XLSX_MERGE_DLL`` environment variable:

    export XLSX_MERGE_DLL=/path/to/xlsx-merge-dotnet/src/XlsxMerge/bin/Release/net8.0
    dotnet build -c Release          # if not built yet

Without the variable, the wrapper looks inside the package directory, then in a
default layout (``<repo>/src/XlsxMerge/bin/{Release,Debug}/net8.0``) relative to
the package.
"""

from __future__ import annotations

import os
import pathlib
from typing import Any, Iterable, Mapping, Optional

__all__ = [
    "CsvParseError",
    "MergeError",
    "load_dotnet",
    "parse_csv",
    "parse_file",
    "render",
    "render_data",
]
__version__ = "0.1.0"

from .parser import (  # noqa: E402  (pure Python, no .NET needed)
    CsvParseError,
    parse_csv,
    parse_file,
)

_DLL_NAME = "XlsxMerge.dll"
_DEPENDENCIES = ("ClosedXML.dll", "DocumentFormat.OpenXml.dll")
_loaded: Optional[tuple] = None
_resolver = None  # held to avoid the CLR handler being garbage-collected


class MergeError(RuntimeError):
    """Mail-merge error (invalid data/template, or missing environment)."""


def _candidate_dirs() -> Iterable[pathlib.Path]:
    env = os.environ.get("XLSX_MERGE_DLL")
    if env:
        # resolve() so a relative path works with clr.AddReference (needs absolute).
        yield pathlib.Path(env).expanduser().resolve()
    package_dir = pathlib.Path(__file__).resolve().parent
    yield package_dir
    # Two layouts, same root (python/xlsx_merge → root):
    #   - monorepo: root/xlsx-merge-dotnet/src/XlsxMerge/bin/<config>/net8.0
    #   - standalone checkout: root/src/XlsxMerge/bin/<config>/net8.0
    root = package_dir.parent.parent
    for configuration in ("Release", "Debug"):
        yield root / "xlsx-merge-dotnet" / "src" / "XlsxMerge" / "bin" / configuration / "net8.0"
        yield root / "src" / "XlsxMerge" / "bin" / configuration / "net8.0"


def _install_assembly_resolver(dll_dir: pathlib.Path) -> None:
    """The CLR must resolve ClosedXML / DocumentFormat.OpenXml, which sit next to
    XlsxMerge.dll (outside pythonnet's default probing paths)."""
    global _resolver

    from System import AppDomain  # type: ignore
    from System.Reflection import Assembly, AssemblyName  # type: ignore

    def _resolve(_sender, args):
        name = AssemblyName(args.Name).Name
        candidate = dll_dir / f"{name}.dll"
        if candidate.is_file():
            return Assembly.LoadFrom(str(candidate))
        return None

    _resolver = _resolve
    AppDomain.CurrentDomain.AssemblyResolve += _resolve


def load_dotnet() -> tuple:
    """Loads the .NET runtime and assemblies; returns the useful types."""
    global _loaded
    if _loaded is not None:
        return _loaded

    try:
        from pythonnet import load
    except ImportError as exc:  # pragma: no cover
        raise MergeError("pythonnet is required: pip install pythonnet") from exc

    try:
        load("coreclr")
    except Exception as exc:  # pragma: no cover - depends on the machine
        raise MergeError(
            "Could not load the .NET runtime (coreclr). "
            "Install the .NET runtime 8 (e.g. apt install dotnet-runtime-8.0)."
        ) from exc

    import clr  # noqa: E402  (must be imported after load)

    dll_dir = next((d for d in _candidate_dirs() if (d / _DLL_NAME).is_file()), None)
    if dll_dir is None:
        searched = "\n  ".join(str(d) for d in _candidate_dirs())
        raise MergeError(
            f"{_DLL_NAME} not found. Build the library first:\n"
            f"  dotnet build -c Release\n"
            f"Searched paths:\n  {searched}\n"
            "Or set XLSX_MERGE_DLL to the directory containing the DLL."
        )

    _install_assembly_resolver(dll_dir)
    for dependency in _DEPENDENCIES:
        dependency_path = dll_dir / dependency
        if dependency_path.is_file():
            clr.AddReference(str(dependency_path))
    clr.AddReference(str(dll_dir / _DLL_NAME))

    from XlsxMerge import (  # type: ignore  # noqa: E402
        MergeData,
        MergeOptions,
        NumericTextParsing,
        RepeatBlockRows,
        XlsxMerger,
    )
    from XlsxMerge.IO import CsvMergeDataReader  # type: ignore  # noqa: E402

    _loaded = (
        MergeData,
        MergeOptions,
        NumericTextParsing,
        RepeatBlockRows,
        XlsxMerger,
        CsvMergeDataReader,
    )
    return _loaded


def _build_options(
    MergeOptions: Any,
    RepeatBlockRows: Any,
    NumericTextParsing: Any,
    culture: str,
    marker: Optional[str],
    block: str,
    numbers: str,
) -> Any:
    from System.Globalization import CultureInfo  # type: ignore

    options = MergeOptions()
    options.NumberCulture = CultureInfo(culture)
    options.RowMarkerField = marker or None

    blocks = {
        "auto": RepeatBlockRows.Auto,
        "1": RepeatBlockRows.One,
        "2": RepeatBlockRows.Two,
    }
    try:
        options.RepeatBlockRows = blocks[str(block).lower()]
    except KeyError as exc:
        raise MergeError(f"Invalid value for block: {block!r} (expected: auto, 1 or 2).") from exc

    numbers_values = {
        # "None" is a Python keyword: access it via getattr.
        "strict": NumericTextParsing.Strict,
        "none": getattr(NumericTextParsing, "None"),
    }
    try:
        options.NumericTextParsing = numbers_values[str(numbers).lower()]
    except KeyError as exc:
        raise MergeError(
            f"Invalid value for numbers: {numbers!r} (expected: strict or none)."
        ) from exc

    return options


def _read_merge_data(
    CsvMergeDataReader: Any, data: Any, csv_text: Optional[str], delimiter: str
) -> Any:
    """Reads the data set from a CSV file or from in-memory CSV text."""
    if (data is None) == (csv_text is None):
        raise MergeError("Provide either data (a CSV file path) or csv_text (CSV content).")

    if csv_text is not None:
        from System.IO import StringReader  # type: ignore

        reader = StringReader(str(csv_text))
        return CsvMergeDataReader.Read(reader, str(delimiter))

    return CsvMergeDataReader.ReadFile(str(data), str(delimiter), None)


def _open_template(template: Any) -> Any:
    """Opens the template: a file path (System.IO.File) or xlsx bytes (MemoryStream)."""
    if isinstance(template, (bytes, bytearray)):
        from System.IO import MemoryStream  # type: ignore

        return MemoryStream(bytes(template))

    from System.IO import File  # type: ignore

    path = str(template)
    if not pathlib.Path(path).is_file():
        raise MergeError(f"Template not found: {path}")
    return File.OpenRead(path)


def _render_template(
    template: Any,
    output: Optional["str | os.PathLike"],
    merge_data: Any,
    culture: str,
    marker: Optional[str],
    block: str,
    numbers: str,
) -> "str | bytes":
    """Renders in memory; writes the output file only when asked."""
    (
        _,
        MergeOptions,
        NumericTextParsing,
        RepeatBlockRows,
        XlsxMerger,
        _,
    ) = load_dotnet()

    from System.IO import MemoryStream  # type: ignore

    options = _build_options(
        MergeOptions, RepeatBlockRows, NumericTextParsing, culture, marker, block, numbers
    )

    template_stream = _open_template(template)
    output_stream = MemoryStream()
    try:
        XlsxMerger.Render(template_stream, output_stream, merge_data, options)
        rendered = bytes(output_stream.ToArray())
    finally:
        output_stream.Dispose()
        template_stream.Dispose()

    if output is None:
        return rendered

    output_path = pathlib.Path(output)
    if output_path.parent != pathlib.Path(""):
        output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(rendered)
    return str(output_path)


def render(
    template: "str | os.PathLike | bytes",
    data: "str | os.PathLike | None" = None,
    output: "str | os.PathLike | None" = None,
    *,
    csv_text: Optional[str] = None,
    culture: str = "en-US",
    marker: Optional[str] = None,
    block: str = "auto",
    numbers: str = "strict",
    delimiter: str = ";",
) -> "str | bytes":
    """Mail-merges the ``template`` with a CSV data set.

    - ``data``: path to a data file (``[globals]`` / ``[rows]`` sections);
    - ``csv_text``: or the CSV content directly, as a string;
    - ``output``: path of the .xlsx file to write; **if ``None``, nothing is written to
      disk and the .xlsx bytes are returned**;
    - ``template``: path of the .xlsx template, or its **bytes**.

    Returns the written file path, or the workbook ``bytes`` when ``output`` is ``None``.
    """
    (
        _,
        _,
        _,
        _,
        _,
        CsvMergeDataReader,
    ) = load_dotnet()

    try:
        merge_data = _read_merge_data(CsvMergeDataReader, data, csv_text, delimiter)
        return _render_template(template, output, merge_data, culture, marker, block, numbers)
    except MergeError:
        raise
    except Exception as exc:
        raise MergeError(str(exc)) from exc


def _to_dotnet(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        return value
    from System import Boolean, DateTime, Double, Int64  # type: ignore

    if isinstance(value, bool):
        return Boolean(value)
    if isinstance(value, int):
        return Int64(value)
    if isinstance(value, float):
        return Double(value)
    try:
        from datetime import date, datetime

        if isinstance(value, datetime):
            return DateTime(
                value.year, value.month, value.day, value.hour, value.minute, value.second
            )
        if isinstance(value, date):
            return DateTime(value.year, value.month, value.day)
    except ImportError:  # pragma: no cover
        pass
    return str(value)


def render_data(
    template: "str | os.PathLike | bytes",
    output: "str | os.PathLike | None" = None,
    *,
    globals: Optional[Mapping[str, Any]] = None,
    rows: Iterable[Mapping[str, Any]] = (),
    fields: Optional[Iterable[str]] = None,
    culture: str = "en-US",
    marker: Optional[str] = None,
    block: str = "auto",
    numbers: str = "strict",
) -> "str | bytes":
    """Mail-merges the ``template`` with in-memory data (Python dicts).

    ``output=None`` (default): nothing is written to disk, the .xlsx bytes are
    returned. ``template`` may be a path or the template bytes.
    """
    (
        MergeData,
        _,
        _,
        _,
        _,
        _,
    ) = load_dotnet()

    from System import Object, String  # type: ignore
    from System.Collections.Generic import Dictionary, List  # type: ignore

    try:
        merge_data = MergeData()
        for key, value in (globals or {}).items():
            merge_data.AddGlobal(str(key), _to_dotnet(value))

        for row in rows:
            dotnet_row: Any = Dictionary[String, Object]()
            for key, value in row.items():
                dotnet_row[str(key)] = _to_dotnet(value)
            merge_data.Rows.Add(dotnet_row)

        if fields:
            declared = List[String]()
            for field in fields:
                declared.Add(str(field))
            merge_data.DeclaredRowFields = declared

        return _render_template(template, output, merge_data, culture, marker, block, numbers)
    except MergeError:
        raise
    except Exception as exc:
        raise MergeError(str(exc)) from exc
