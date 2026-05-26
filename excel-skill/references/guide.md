# Excel Workbook Design Guide

## Standard Workbook Structure

- Summary sheet: title, purpose, key metrics, highlights, navigation notes.
- Data sheets: clear name, optional title, headers, rows, formulas or totals, optional chart.
- Reference sheet: assumptions, definitions, notes, or data sources.

## Table Formatting

- Use bold, high-contrast headers.
- Freeze header rows and apply filters when useful.
- Set readable widths and consistent alignment.
- Use alternating row colors for larger tables.
- Avoid empty columns, very long sheet names, merged cells inside data tables, and multiple unrelated tables on one worksheet.

## Number Formatting

| Data Type | Recommended Format |
|:---|:---|
| Currency | `$#,##0.00` |
| Whole Numbers | `#,##0` |
| Percentages | `0.0%` |
| Decimals | `0.00` |
| Dates | `yyyy-mm-dd` |

Format numeric values as numbers; keep labels as text.

## Formula Guidelines

- Use formulas for totals, averages, percentages, growth, counts, and simple derived metrics.
- Common formulas: `SUM`, `AVERAGE`, `COUNT`, `MAX`, `MIN`.
- Place formulas near related tables, label results clearly, and verify referenced ranges.
- Avoid complex formulas unless requested.

## Chart Guidelines

- Column: category comparisons.
- Bar: long category labels.
- Line: trends over time.
- Pie: simple proportions with few categories.
- Use clean numeric ranges, clear titles, few series, and no charts for tiny or non-numeric datasets.

## Spreadsheet Templates

Budget:
1. Summary
2. Income
3. Expenses
4. Monthly Totals
5. Charts

Sales Report:
1. Summary
2. Sales Data
3. Regional Breakdown
4. Product Breakdown
5. Trend Chart

Project Tracker:
1. Summary
2. Task List
3. Timeline
4. Risks
5. Status Dashboard

Educational Data:
1. Summary
2. Topic Data
3. Comparison Table
4. Calculations
5. Chart

## Spreadsheet Guardrails

- Every worksheet should have a clear purpose.
- Tables should start near the top-left.
- Headers must be understandable without extra explanation.
- Use formulas and charts only when they add value.
- Keep the workbook readable on standard laptop screens.
