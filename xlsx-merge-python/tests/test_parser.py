"""Tests of the generic CSV parser (pure Python, no .NET dependency)."""

from __future__ import annotations

import pathlib
import sys

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from xlsx_merge import CsvParseError, parse_csv, parse_file  # noqa: E402

SAMPLE = """
[globals]
COURSE;Advanced Excel
DATE;2024-06-01
#ROOM#;B12

[rows]
EMPLOYEE_ID;NAME;SCORE
42;Alice;17.5
43;Bob;
"""


def test_basic_parse():
    globals_, rows = parse_csv(SAMPLE)

    assert globals_ == {
        "COURSE": "Advanced Excel",
        "DATE": "2024-06-01",
        "ROOM": "B12",  # key normalized ("#ROOM#")
    }
    assert rows == [
        {"EMPLOYEE_ID": "42", "NAME": "Alice", "SCORE": "17.5"},
        {"EMPLOYEE_ID": "43", "NAME": "Bob", "SCORE": ""},
    ]


def test_quotes_semicolons_and_multiline():
    globals_, rows = parse_csv(
        '[globals]\nCOURSE;"Excel; advanced"\n\n'
        '[rows]\nNAME;NOTE\n"Alice; aka Ali";"1.5"\n"With ""quotes""";\n"multi\nline";x\n'
    )

    assert globals_["COURSE"] == "Excel; advanced"
    assert rows[0] == {"NAME": "Alice; aka Ali", "NOTE": "1.5"}
    assert rows[1] == {"NAME": 'With "quotes"', "NOTE": ""}
    assert rows[2] == {"NAME": "multi\nline", "NOTE": "x"}


def test_crlf_and_blank_lines():
    globals_, rows = parse_csv("[globals]\r\nA;1\r\n\r\n[rows]\r\nX;Y\r\n1;2\r\n")
    assert globals_ == {"A": "1"}
    assert rows == [{"X": "1", "Y": "2"}]


def test_missing_columns_filled_with_empty():
    _, rows = parse_csv("[rows]\nA;B;C;D\n1;2\n")
    assert rows == [{"A": "1", "B": "2", "C": "", "D": ""}]


def test_trailing_empty_columns_ignored():
    _, rows = parse_csv("[rows]\nA;B;C\n1;2;\n")
    assert rows == [{"A": "1", "B": "2", "C": ""}]


@pytest.mark.parametrize(
    "texte,message",
    [
        ("COURSE;X\n[globals]\nA;B", "[globals] or [rows] section"),
        ("[rows]\nA\n1\n[rows]\nB", "already defined"),
        ("[rows]\nA;B\n1;2;3", "3 columns"),
        ("[rows]\nA;#A#", "duplicate header"),
        ("\n\n", "empty"),
        ('[globals]\nA;"unterminated', "unterminated quote"),
        ("[globals]\n##;X", "cannot be empty"),
        ("[globals]\nA#B;X", "cannot contain"),
        ('[globals]\nA;"X"Y', "unexpected character"),
    ],
)
def test_errors(texte, message):
    with pytest.raises(CsvParseError) as info:
        parse_csv(texte)
    assert message in str(info.value)


def test_invalid_delimiter():
    with pytest.raises(CsvParseError):
        parse_csv("A;;B", ";;")


def test_parse_file_with_bom(tmp_path):
    path = tmp_path / "data.csv"
    path.write_bytes(b"\xef\xbb\xbf[globals]\nA;1\n")
    globals_, rows = parse_file(path)
    assert globals_ == {"A": "1"}
    assert rows == []


def test_header_only_rows_declares_the_fields():
    globals_, rows = parse_csv("[rows]\nID;NAME")
    assert globals_ == {}
    assert rows == []
