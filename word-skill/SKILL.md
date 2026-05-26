---
name: word-skill
description: Creates, reads, and improves Microsoft Word (.docx) documents.
metadata:
  version: "1.0.0"
  tags: ["docx", "word", "document", "office-documents"]
---

# Word Document Skill

Use for Word creation, reading, analysis, editing, summarization, or document restructuring.

## When to Use

- Create `.docx` reports, essays, articles, proposals, letters, memos, training docs, or project docs.
- Convert topics, notes, articles, reports, or structured content into formatted documents.
- Extract, summarize, analyze, rewrite, or reorganize DOCX paragraphs, headings, lists, and tables.

## Document Creation Rules

- Use `create_word_document` or `scripts/create_docx.py`; save a real `.docx` unless the user asks for outline-only content.
- Use a clear title, logical headings, concise paragraphs, and consistent formatting.
- Add bullets for grouped ideas and tables only for structured comparisons or data.
- Avoid empty sections and long walls of text.
- Include a conclusion, summary, recommendations, or next steps when appropriate.

## Reading and Extraction Rules

- Use `read_word_document` or `scripts/read_docx.py`.
- Preserve paragraph and table order as much as possible.
- Analyze or summarize only after reading the file.

## Design Requirements

Follow `references/guide.md` for structure, density, headings, formatting, templates, lists, tables, and readability guardrails.

## Validation Checklist

- Title and section order are clear.
- Each main section has useful paragraphs, bullets, or a table.
- Heading hierarchy and formatting are consistent.
- Paragraphs are readable and not overly dense.
- Tables have headers and meaningful rows.
- The saved `.docx` path is returned.

## Script Usage

Generator inputs:

- `title`: document title.
- `subtitle`: optional subtitle.
- `sections`: optional section objects with `heading`, `paragraphs`, `bullets`, and optional `table`.
- `output_path`: optional `.docx` destination.

Reader input:

- `path`: `.docx` file path.

Use `scripts/verify.py` to validate the skill structure.

## Best Practices

- Build the outline before calling the generator.
- Prefer 3 to 7 sections for general topics.
- Preserve user-provided source details while organizing them professionally.
- Use tables for comparisons, timelines, facts, feature matrices, or schedules.
