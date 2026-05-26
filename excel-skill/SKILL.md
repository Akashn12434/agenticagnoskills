---
name: excel-skill
description: Creates, reads, analyzes, formats, and structures Microsoft Excel (.xlsx) workbooks.
metadata:
  version: "1.0.0"
  tags: ["xlsx", "excel", "spreadsheet", "data-analysis", "office-documents"]
---

# Excel Spreadsheet Skill

Use for Excel workbook creation, reading, analysis, summarization, formatting, or restructuring.

## When to Use

- Create `.xlsx` budgets, trackers, schedules, inventories, dashboards, reports, tables, or comparisons.
- Convert topics, datasets, prompts, plans, or reports into formatted worksheets.
- Add formulas, totals, percentages, averages, calculated fields, charts, or visual summaries.
- Read, extract, summarize, or analyze worksheet values from existing workbooks.

## Workbook Creation Rules

- Use `create_excel_workbook` or `scripts/create_xlsx.py`; save a real `.xlsx` unless the user asks for text-only content.
- Use clear sheet names, table headers, readable column widths, and professional number formats.
- Freeze header rows and add filters when tables are present.
- Add formulas for useful calculations and charts for clear numeric comparisons or trends.
- Avoid empty worksheets and unrelated tables on one sheet.

## Reading and Extraction Rules

- Use `read_excel_workbook` or `scripts/read_xlsx.py`.
- Preserve workbook sheet order and row order as much as possible.
- Analyze or summarize only after reading the file.

## Design Requirements

Follow `references/guide.md` for workbook structure, tables, number formats, formulas, charts, dashboards, templates, and spreadsheet guardrails.

## Validation Checklist

- Workbook has at least one useful worksheet.
- Tables have clear headers and readable columns.
- Numbers, formulas, and charts use correct ranges and formats.
- Worksheets are logically named.
- No empty placeholder sheets remain.
- The saved `.xlsx` path is returned.

## Script Usage

Generator inputs:

- `title`: workbook title.
- `sheets`: optional worksheet objects with `name`, optional `title`, `headers`, `rows`, optional `formulas`, and optional `chart`.
- `output_path`: optional `.xlsx` destination.

Reader inputs:

- `path`: `.xlsx` file path.
- `max_rows_per_sheet`: optional row limit.

Use `scripts/verify.py` to validate the skill structure.

## Best Practices

- Build workbook structure before calling the generator.
- Prefer one worksheet per clear topic or dataset.
- Use a summary sheet when the workbook has multiple sheets.
- Preserve user-provided data and organize it professionally.
