#!/usr/bin/env python3
"""Read basic text and values from an Excel .xlsx workbook."""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree


MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
OFFICE_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def read_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    try:
        root = ElementTree.fromstring(workbook.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    strings: list[str] = []
    for item in root.findall(f"{MAIN_NS}si"):
        parts = [node.text or "" for node in item.iter(f"{MAIN_NS}t")]
        strings.append("".join(parts))
    return strings


def read_sheet_names(workbook: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook_root = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
    rel_root = ElementTree.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rel_root.findall(f"{REL_NS}Relationship")
    }

    sheets: list[tuple[str, str]] = []
    for sheet in workbook_root.findall(f".//{MAIN_NS}sheet"):
        name = sheet.attrib.get("name", "Sheet")
        relation_id = sheet.attrib.get(f"{OFFICE_REL_NS}id")
        target = rels.get(relation_id or "", "")
        if target:
            path = "xl/" + target.lstrip("/")
            sheets.append((name, path))
    return sheets


def cell_column(cell_ref: str) -> int:
    letters = re.sub(r"[^A-Z]", "", cell_ref.upper())
    result = 0
    for letter in letters:
        result = result * 26 + ord(letter) - ord("A") + 1
    return result - 1


def cell_value(cell, shared_strings: list[str]) -> str:
    value_node = cell.find(f"{MAIN_NS}v")
    if value_node is None or value_node.text is None:
        inline = cell.find(f".//{MAIN_NS}t")
        return inline.text if inline is not None and inline.text else ""

    value = value_node.text
    if cell.attrib.get("t") == "s":
        try:
            return shared_strings[int(value)]
        except (IndexError, ValueError):
            return value
    return value


def read_workbook(path: str, max_rows_per_sheet: int = 50) -> str:
    xlsx_path = Path(path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Excel workbook not found: {xlsx_path}")

    output: list[str] = []
    with zipfile.ZipFile(xlsx_path) as workbook:
        shared_strings = read_shared_strings(workbook)
        for sheet_name, sheet_path in read_sheet_names(workbook):
            output.append(f"# {sheet_name}")
            sheet_root = ElementTree.fromstring(workbook.read(sheet_path))
            row_count = 0
            for row in sheet_root.findall(f".//{MAIN_NS}row"):
                values: list[str] = []
                current_col = 0
                for cell in row.findall(f"{MAIN_NS}c"):
                    col = cell_column(cell.attrib.get("r", "A1"))
                    while current_col < col:
                        values.append("")
                        current_col += 1
                    values.append(cell_value(cell, shared_strings))
                    current_col += 1
                if any(values):
                    output.append(" | ".join(values))
                    row_count += 1
                if row_count >= max_rows_per_sheet:
                    break
    return "\n".join(output)


def main() -> None:
    parser = argparse.ArgumentParser(description="Read values from an Excel workbook.")
    parser.add_argument("path", help="Path to .xlsx file")
    parser.add_argument("--max-rows-per-sheet", type=int, default=50)
    args = parser.parse_args()
    print(read_workbook(args.path, args.max_rows_per_sheet))


if __name__ == "__main__":
    main()
