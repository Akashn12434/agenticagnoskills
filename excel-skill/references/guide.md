# Excel Workbook Design Guide

## Standard Workbook Structure

### Summary Sheet

Use a summary sheet when the workbook contains multiple sheets or a dashboard-style overview.

Contains:

- Workbook title
- Short description or purpose
- Key metrics, totals, or highlights
- Navigation notes when useful

### Data Sheets

Each data sheet should contain:

- Clear worksheet name
- Optional title above the table
- Header row
- Data rows
- Formulas or totals where useful
- Chart when visual comparison or trends help

### Final / Reference Sheet

Use a reference sheet when the workbook needs:

- Assumptions
- Definitions
- Notes
- Data sources

---

## Table Formatting

Recommended:

- Bold header row
- High-contrast header fill
- Frozen header row
- Filterable table ranges when possible
- Readable column widths
- Consistent alignment
- Alternating row colors for larger tables

Avoid:

- Empty columns between related data
- Very long worksheet names
- Unformatted number columns
- Merged cells inside data tables
- Multiple unrelated tables on one worksheet

---

## Number Formatting

| Data Type | Recommended Format |
|:---|:---|
| Currency | `$#,##0.00` |
| Whole Numbers | `#,##0` |
| Percentages | `0.0%` |
| Decimals | `0.00` |
| Dates | `yyyy-mm-dd` |

Use number formats only when the values are numeric. Keep labels as text.

---

## Formula Guidelines

Use formulas for:

- Totals
- Averages
- Percentages
- Growth or change
- Counts
- Simple derived metrics

Common formulas:

- `=SUM(A2:A10)`
- `=AVERAGE(B2:B10)`
- `=COUNT(A2:A10)`
- `=MAX(C2:C10)`
- `=MIN(C2:C10)`

Formula guardrails:

- Place formulas near the related table.
- Label formula results clearly.
- Reference the correct row and column ranges.
- Avoid overly complex formulas unless the user asks for them.

---

## Chart Guidelines

Use charts when they make insight clearer:

- Column chart: compare categories
- Bar chart: compare categories with long labels
- Line chart: show trends over time
- Pie chart: show simple proportions with few categories

Chart guardrails:

- Use a clear chart title.
- Reference clean table ranges.
- Avoid too many series.
- Place charts to the right of or below the table.
- Do not create charts for tiny or non-numeric datasets.

---

## Spreadsheet Templates

## Budget Template

1. Summary
2. Income
3. Expenses
4. Monthly Totals
5. Charts

## Sales Report Template

1. Summary
2. Sales Data
3. Regional Breakdown
4. Product Breakdown
5. Trend Chart

## Project Tracker Template

1. Summary
2. Task List
3. Timeline
4. Risks
5. Status Dashboard

## Educational Data Template

1. Summary
2. Topic Data
3. Comparison Table
4. Calculations
5. Chart

---

## Spreadsheet Guardrails

- Every worksheet should have a clear purpose.
- Tables should start near the top-left of the sheet.
- Headers must be understandable without extra explanation.
- Use formulas only when they add value.
- Use charts only when there is numeric data to visualize.
- Keep the workbook readable on standard laptop screens.
