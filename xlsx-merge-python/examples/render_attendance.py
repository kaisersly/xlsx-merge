#!/usr/bin/env python3
"""Example: mail-merge of a training attendance sheet.

Usage (build the .NET DLL first, see README):

    python examples/render_attendance.py [template.xlsx] [data.csv] [output.xlsx]
"""

from __future__ import annotations

import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from xlsx_merge import MergeError, render  # noqa: E402

SAMPLES = (
    pathlib.Path(__file__).resolve().parent.parent.parent
    / "xlsx-merge-dotnet"
    / "samples"
)


def main() -> int:
    template = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else SAMPLES / "attendance-template.xlsx"
    data = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else SAMPLES / "attendance.csv"
    output = pathlib.Path(sys.argv[3]) if len(sys.argv) > 3 else pathlib.Path("attendance-output.xlsx")

    try:
        result = render(template, data, output)
    except MergeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Output file: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
