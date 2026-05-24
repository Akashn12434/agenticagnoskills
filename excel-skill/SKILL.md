---
name: excel-skill
description: Creates, reads, analyzes, formats, and structures Microsoft Excel (.xlsx) workbooks.
metadata:
  version: "1.0.0"
  tags: ["xlsx", "excel", "spreadsheet", "data-analysis", "office-documents"]
---

# Excel Spreadsheet Skill

Use this skill whenever a user wants to create, read, analyze, summarize, or structure a Microsoft Excel workbook (.xlsx).

## When to Use

Use this skill if the user asks to:

- Create a new Excel workbook (.xlsx)
- Generate a spreadsheet from a topic, prompt, dataset, table, plan, or report
- Create budgets, trackers, schedules, inventories, dashboards, reports, or comparison tables
- Convert structured content into formatted Excel sheets
- Add formulas, totals, percentages, averages, or calculated fields
- Add charts or visual summaries
- Read or extract worksheet values from an existing .xlsx file
- Create professional, business, educational, financial, operational, or analytical spreadsheets

## Workbook Creation Rules

When generating a new Excel workbook:

- Use the registered `create_excel_workbook` tool or `scripts/create_xlsx.py`.
- Always generate and save a real `.xlsx` file.
- Use Python `XlsxWriter` for workbook creation.
- Do not stop at a text table unless the user explicitly asks for text-only content.
- Use clear worksheet names.
- Use header rows for every table.
- Use professional formatting for titles, headers, numbers, dates, and totals.
- Freeze header rows when tables are present.
- Set useful column widths.
- Add formulas when calculations are needed.
- Add charts when they make trends, comparisons, or proportions clearer.
- Avoid empty worksheets.

## Reading and Extraction Rules

When reading an existing Excel workbook:

- Use the registered `read_excel_workbook` tool or `scripts/read_xlsx.py`.
- Extract workbook sheet names, rows, and cell values.
- Preserve sheet and row order as much as possible.
- Summarize or analyze the extracted content only after reading the file.

## Design Requirements

Consult `references/guide.md` and follow its recommendations for:

- Workbook Structure
- Table Formatting
- Number Formatting
- Formula Guidelines
- Chart Guidelines
- Dashboard and Summary Layouts
- Spreadsheet Guardrails

## Validation Checklist

Before finalizing any Excel workbook:

- Workbook has at least one useful worksheet.
- Each table has clear headers.
- Column widths are readable.
- Header rows are frozen when appropriate.
- Numbers use suitable formats.
- Formulas are accurate and placed in clear cells.
- Charts, if used, reference the correct data range.
- Worksheets are logically named.
- No empty placeholder sheets remain.
- The saved `.xlsx` file path is returned to the user.

## Script Usage

When the user asks to create a new Excel workbook, call `create_excel_workbook` or `scripts/create_xlsx.py` to generate and save a real `.xlsx` file using Python `XlsxWriter`.

The workbook generator accepts:

- `title`: the workbook title.
- `sheets`: an optional list of worksheet objects.
- `output_path`: an optional `.xlsx` destination.

Each worksheet object should use this shape:

```json
{
  "name": "Worksheet Name",
  "title": "Optional Sheet Title",
  "headers": ["Column A", "Column B", "Column C"],
  "rows": [
    ["Value A1", "Value B1", 100],
    ["Value A2", "Value B2", 200]
  ],
  "formulas": [
    {"cell": "C4", "formula": "=SUM(C2:C3)", "label": "Total"}
  ],
  "chart": {
    "type": "column",
    "title": "Chart Title",
    "categories_column": 0,
    "values_column": 2
  }
}
```

When the user asks to read or extract content from an Excel workbook, call `read_excel_workbook` or `scripts/read_xlsx.py`.

Use `scripts/verify.py` to validate the excel-skill structure.

## Best Practices

- Build the workbook structure first, then generate the Excel file.
- Prefer one worksheet per clear topic or dataset.
- Use a summary sheet when the workbook has multiple sheets.
- Keep table headers short and descriptive.
- Use formulas for totals, averages, percentages, and derived values.
- Use charts for comparisons, trends, categories, or summary insights.
- If the user provides data, preserve the important values and organize them professionally.
- Return the saved `.xlsx` path after generation.
