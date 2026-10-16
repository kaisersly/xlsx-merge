"""Tests of the render functions (need the .NET DLL built — see conftest)."""

from __future__ import annotations

import pathlib
import sys

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from xlsx_merge import MergeError  # noqa: E402
from conftest import DATA_CSV, workbook_text  # noqa: E402


def _data_file(tmp_path) -> pathlib.Path:
    data = tmp_path / "data.csv"
    data.write_text(DATA_CSV, encoding="utf-8")
    return data


def test_render_from_csv_file(xlsx_merge, template_path, tmp_path):
    output = tmp_path / "out.xlsx"

    result = xlsx_merge.render(template_path, _data_file(tmp_path), output)

    assert result == str(output)
    content = workbook_text(output.read_bytes())
    assert "#EMPLOYEE_ID#" not in content
    assert "Alice" in content
    assert "Advanced Excel" in content


def test_render_with_csv_text_and_bytes_output(xlsx_merge, template_path):
    """CSV in a str + xlsx returned as bytes: nothing written to disk."""
    template_bytes = template_path.read_bytes()

    result = xlsx_merge.render(template_bytes, csv_text=DATA_CSV)

    assert isinstance(result, bytes)
    content = workbook_text(result)
    assert "#EMPLOYEE_ID#" not in content
    assert "Alice" in content
    assert "Generated on 2024-06-01 (#SOMEWHERE#)" in content


def test_render_data_in_memory(xlsx_merge, template_path):
    result = xlsx_merge.render_data(
        template_path,
        globals={"COURSE": "Excel", "DATE": "2024-06-01", "ROOM": "B12"},
        rows=[
            {"EMPLOYEE_ID": 1, "NAME": "Alice", "SCORE": 17.5},
            {"EMPLOYEE_ID": 2, "NAME": "Bob"},
        ],
    )

    assert isinstance(result, bytes)
    content = workbook_text(result)
    assert "Alice" in content
    assert "Bob" in content


def test_marker_and_options(xlsx_merge, template_path, tmp_path):
    output = tmp_path / "out.xlsx"
    xlsx_merge.render(
        template_path,
        _data_file(tmp_path),
        output,
        marker="EMPLOYEE_ID",
        block="2",
        numbers="strict",
        culture="fr-FR",
        delimiter=";",
    )
    assert output.is_file()


def test_data_and_csv_text_are_exclusive(xlsx_merge, template_path, tmp_path):
    with pytest.raises(MergeError):
        xlsx_merge.render(template_path, _data_file(tmp_path), csv_text=DATA_CSV)
    with pytest.raises(MergeError):
        xlsx_merge.render(template_path)


def test_invalid_data_is_a_merge_error(xlsx_merge, template_path, tmp_path):
    bad = tmp_path / "bad.csv"
    bad.write_text("Content outside a section\n", encoding="utf-8")
    with pytest.raises(MergeError):
        xlsx_merge.render(template_path, bad, tmp_path / "out.xlsx")
