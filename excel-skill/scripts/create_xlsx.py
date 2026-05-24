#!/usr/bin/env python3
"""Create an Excel .xlsx workbook from structured content using XlsxWriter."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", value.strip().lower()).strip("_")
    return slug or "workbook"


def clean_sheet_name(value: str, fallback: str) -> str:
    name = re.sub(r"[\[\]:*?/\\]", " ", value or fallback).strip()
    return (name or fallback)[:31]


def unwrap_sheet(sheet: dict[str, Any]) -> dict[str, Any]:
    if any(key in sheet for key in ("name", "headers", "rows", "formulas", "chart")):
        return sheet

    for key in ("sheet", "worksheet", "content", "data", "example_key"):
        nested = sheet.get(key)
        if isinstance(nested, dict):
            return unwrap_sheet(nested)

    if len(sheet) == 1:
        nested = next(iter(sheet.values()))
        if isinstance(nested, dict):
            return unwrap_sheet(nested)

    return sheet


def default_sheets(title: str) -> list[dict[str, Any]]:
    return [
        {
            "name": "Summary",
            "title": title,
            "headers": ["Metric", "Value"],
            "rows": [
                ["Total Items", 3],
                ["Completed Items", 2],
                ["Pending Items", 1],
            ],
            "formulas": [
                {"cell": "B6", "label": "Completion Rate", "formula": "=B3/B2", "format": "percent"},
            ],
            "chart": {
                "type": "column",
                "title": "Summary Metrics",
                "categories_column": 0,
                "values_column": 1,
            },
        }
    ]


def normalize_sheets(title: str, sheets: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    if not sheets:
        return default_sheets(title)

    normalized: list[dict[str, Any]] = []
    for index, sheet in enumerate(sheets, start=1):
        sheet = unwrap_sheet(sheet)
        headers = [str(item).strip() for item in sheet.get("headers", []) if str(item).strip()]
        rows = sheet.get("rows", [])
        safe_rows = rows if isinstance(rows, list) else []
        normalized.append(
            {
                "name": clean_sheet_name(str(sheet.get("name") or sheet.get("title") or ""), f"Sheet {index}"),
                "title": str(sheet.get("title") or sheet.get("name") or "").strip(),
                "headers": headers,
                "rows": safe_rows,
                "formulas": sheet.get("formulas", []) if isinstance(sheet.get("formulas", []), list) else [],
                "chart": sheet.get("chart") if isinstance(sheet.get("chart"), dict) else None,
            }
        )

    return normalized


def build_formats(workbook: xlsxwriter.Workbook) -> dict[str, Any]:
    return {
        "title": workbook.add_format(
            {"bold": True, "font_size": 16, "font_color": "#1C2833", "bottom": 1}
        ),
        "header": workbook.add_format(
            {"bold": True, "font_color": "white", "bg_color": "#2E4053", "border": 1}
        ),
        "text": workbook.add_format({"border": 1}),
        "number": workbook.add_format({"border": 1, "num_format": "#,##0.00"}),
        "integer": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "percent": workbook.add_format({"border": 1, "num_format": "0.0%"}),
        "formula_label": workbook.add_format({"bold": True, "bg_color": "#D5D8DC", "border": 1}),
        "formula": workbook.add_format({"bold": True, "border": 1, "num_format": "#,##0.00"}),
        "formula_percent": workbook.add_format({"bold": True, "border": 1, "num_format": "0.0%"}),
    }


def write_value(worksheet, row: int, col: int, value: Any, formats: dict[str, Any]) -> None:
    if isinstance(value, bool):
        worksheet.write_boolean(row, col, value, formats["text"])
    elif isinstance(value, int):
        worksheet.write_number(row, col, value, formats["integer"])
    elif isinstance(value, float):
        worksheet.write_number(row, col, value, formats["number"])
    elif isinstance(value, str) and value.startswith("="):
        worksheet.write_formula(row, col, value, formats["number"])
    else:
        worksheet.write(row, col, value, formats["text"])


def add_chart(workbook, worksheet, sheet: dict[str, Any], first_data_row: int, last_data_row: int) -> None:
    chart_spec = sheet.get("chart")
    headers = sheet.get("headers", [])
    if not chart_spec or not headers or last_data_row < first_data_row:
        return

    values_column = int(chart_spec.get("values_column", 1))
    categories_column = int(chart_spec.get("categories_column", 0))
    if values_column >= len(headers) or categories_column >= len(headers):
        return

    chart_type = chart_spec.get("type", "column")
    if chart_type not in {"column", "bar", "line", "pie"}:
        chart_type = "column"

    chart = workbook.add_chart({"type": chart_type})
    sheet_name = sheet["name"]
    chart.add_series(
        {
            "name": str(chart_spec.get("series_name") or headers[values_column]),
            "categories": [sheet_name, first_data_row, categories_column, last_data_row, categories_column],
            "values": [sheet_name, first_data_row, values_column, last_data_row, values_column],
        }
    )
    chart.set_title({"name": str(chart_spec.get("title") or "Chart")})
    chart.set_legend({"position": "bottom"})
    worksheet.insert_chart(first_data_row, len(headers) + 2, chart, {"x_scale": 1.25, "y_scale": 1.15})


def create_workbook(
    title: str,
    sheets: list[dict[str, Any]] | None = None,
    output_path: str | None = None,
) -> str:
    """Create a .xlsx file and return its absolute path."""
    normalized_sheets = normalize_sheets(title, sheets)
    output = Path(output_path) if output_path else OUTPUT_DIR / f"{slugify(title)}.xlsx"
    if not output.is_absolute():
        output = OUTPUT_DIR / output

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(output))
    formats = build_formats(workbook)

    for sheet in normalized_sheets:
        worksheet = workbook.add_worksheet(sheet["name"])
        row = 0

        if sheet.get("title"):
            width = max(len(sheet.get("headers", [])) - 1, 1)
            worksheet.merge_range(row, 0, row, width, sheet["title"], formats["title"])
            row += 2

        headers = sheet.get("headers", [])
        rows = sheet.get("rows", [])
        if headers:
            for col, header in enumerate(headers):
                worksheet.write(row, col, header, formats["header"])
                worksheet.set_column(col, col, max(14, min(32, len(str(header)) + 4)))

            first_data_row = row + 1
            for row_offset, values in enumerate(rows):
                values = values if isinstance(values, list) else [values.get(header, "") for header in headers]
                for col, value in enumerate(values[: len(headers)]):
                    write_value(worksheet, first_data_row + row_offset, col, value, formats)

            last_data_row = first_data_row + len(rows) - 1
            worksheet.freeze_panes(first_data_row, 0)
            worksheet.autofilter(row, 0, max(row, last_data_row), max(0, len(headers) - 1))
            row = max(first_data_row + len(rows) + 1, row + 2)
            add_chart(workbook, worksheet, sheet, first_data_row, last_data_row)

        for formula in sheet.get("formulas", []):
            cell = str(formula.get("cell", "")).strip()
            expression = str(formula.get("formula", "")).strip()
            if not cell or not expression.startswith("="):
                continue
            label = str(formula.get("label") or "Formula")
            formula_format = formats["formula_percent"] if formula.get("format") == "percent" else formats["formula"]
            worksheet.write(row, 0, label, formats["formula_label"])
            worksheet.write_formula(cell, expression, formula_format)
            row += 1

    workbook.close()
    return str(output.resolve())


def load_sheets(value: str | None) -> list[dict[str, Any]] | None:
    if not value:
        return None
    path = Path(value)
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return json.loads(value)


def main() -> None:
    parser = argparse.ArgumentParser(description="Create an Excel workbook.")
    parser.add_argument("title", help="Workbook title")
    parser.add_argument("--sheets", help="JSON sheet list or path to JSON")
    parser.add_argument("--output", help="Output .xlsx path")
    args = parser.parse_args()

    result = create_workbook(args.title, load_sheets(args.sheets), args.output)
    print(result)


if __name__ == "__main__":
    main()
