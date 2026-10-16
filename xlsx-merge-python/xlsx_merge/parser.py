"""Reader of the generic two-section data format (pure Python, no .NET needed).

Expected format (UTF-8, default delimiter ``;``, RFC 4180 quotes):

    [globals]
    COURSE;Advanced Excel
    DATE;2024-06-01

    [rows]
    EMPLOYEE_ID;NAME;SCORE
    42;Alice;17.5
    43;Bob;8

Result: ``(globals, rows)`` — a dict of the global fields and the list of rows to
repeat (one dict per row). Values are raw strings; typing (numbers…) happens at
render time.
"""

from __future__ import annotations

import pathlib
from typing import Dict, List, Tuple

__all__ = ["CsvParseError", "parse_csv", "parse_file", "read_records"]


class CsvParseError(ValueError):
    """Invalid data file (message aimed at the user)."""


def parse_file(path: "str | pathlib.Path", delimiter: str = ";") -> Tuple[
    Dict[str, str], List[Dict[str, str]]
]:
    """Reads a data file (UTF-8, BOM accepted) and returns (globals, rows)."""
    return parse_csv(pathlib.Path(path).read_text(encoding="utf-8-sig"), delimiter)


def parse_csv(text: str, delimiter: str = ";") -> Tuple[Dict[str, str], List[Dict[str, str]]]:
    """Parses a CSV content and returns ``(globals, rows)``.

    - ``globals``: ``key → value`` dict of the ``[globals]`` section;
    - ``rows``: list of ``key → value`` dicts, one entry per line of the ``[rows]``
      section (missing columns → ``""``).

    Keys are normalized (blank and border ``#`` removed: ``#COURSE#`` ≡ ``COURSE``).
    Raises :class:`CsvParseError` on error, with the line number.
    """
    if len(delimiter) != 1:
        raise CsvParseError("The delimiter must be a single character (e.g. ';').")

    records = read_records(text, delimiter)

    globals_: Dict[str, str] = {}
    rows: List[Dict[str, str]] = []
    headers: List[str] | None = None
    section: str | None = None
    saw_any_record = False

    for line_number, fields in records:
        if all(not field.strip() for field in fields):
            continue  # blank line

        saw_any_record = True
        first = fields[0].strip()

        if first.lower() == "[globals]":
            section = "globals"
            continue
        if first.lower() == "[rows]":
            if headers is not None:
                raise CsvParseError(f"Line {line_number}: the [rows] section is already defined.")
            section = "rows"
            continue

        # Trailing empty fields carry no information.
        while fields and fields[-1] == "":
            fields = fields[:-1]

        if section == "globals":
            if len(fields) > 2:
                raise CsvParseError(
                    f"Line {line_number}: too many columns in the [globals] section "
                    "(key;value expected)."
                )
            key = _normalize_key(fields[0], line_number)
            globals_[key] = fields[1] if len(fields) == 2 else ""

        elif section == "rows":
            if headers is None:
                # First non-blank line of [rows]: the headers.
                headers = [_normalize_key(field, line_number) for field in fields]
                duplicates = sorted({h for h in headers if headers.count(h) > 1})
                if duplicates:
                    raise CsvParseError(
                        f"Line {line_number}: duplicate header '{duplicates[0]}'."
                    )
                continue

            if len(fields) > len(headers):
                raise CsvParseError(
                    f"Line {line_number}: {len(fields)} columns but the header row has "
                    f"{len(headers)}."
                )
            rows.append(
                {
                    header: (fields[index] if index < len(fields) else "")
                    for index, header in enumerate(headers)
                }
            )

        else:
            raise CsvParseError(
                f"Line {line_number}: a [globals] or [rows] section is expected before any content."
            )

    if not saw_any_record:
        raise CsvParseError("The data file is empty.")

    return globals_, rows


def _normalize_key(key: str, line_number: int) -> str:
    normalized = key.strip()
    if len(normalized) >= 2 and normalized.startswith("#") and normalized.endswith("#"):
        normalized = normalized[1:-1].strip()
    if not normalized:
        raise CsvParseError(f"Line {line_number}: a merge field cannot be empty.")
    if "#" in normalized:
        raise CsvParseError(
            f"Line {line_number}: the merge field '{normalized}' cannot contain '#' "
            "(a field is delimited by two '#')."
        )
    return normalized


def read_records(text: str, delimiter: str = ";") -> List[Tuple[int, List[str]]]:
    """RFC 4180 CSV parsing: quotes, doubled quotes, multi-line fields.

    Public utility (reused by applications building their own data readers).
    Returns a list of ``(start line number, fields)``.
    """
    records: List[Tuple[int, List[str]]] = []
    fields: List[str] = []
    field: List[str] = []
    line_number = 1
    in_quotes = False
    quote_closed = False
    index = 0
    length = len(text)

    while index < length:
        char = text[index]

        if in_quotes:
            if char == '"':
                if index + 1 < length and text[index + 1] == '"':
                    field.append('"')
                    index += 2
                    continue
                in_quotes = False
                quote_closed = True
            else:
                if char == "\n":
                    line_number += 1
                field.append(char)
            index += 1
            continue

        if quote_closed and char not in (delimiter, "\r", "\n"):
            raise CsvParseError(
                f"Line {line_number}: unexpected character '{char}' after a closing quote."
            )

        if char == '"' and not field:
            in_quotes = True
        elif char == delimiter:
            fields.append("".join(field))
            field = []
            quote_closed = False
        elif char in ("\r", "\n"):
            if char == "\r" and index + 1 < length and text[index + 1] == "\n":
                index += 1
            fields.append("".join(field))
            field = []
            quote_closed = False
            records.append((line_number, fields))
            fields = []
            line_number += 1
        else:
            field.append(char)
        index += 1

    if in_quotes:
        raise CsvParseError(f"Line {line_number}: unterminated quote.")

    if field or fields:
        fields.append("".join(field))
        records.append((line_number, fields))

    return records
