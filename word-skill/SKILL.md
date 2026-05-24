---
name: word-skill
description: Creates, reads, analyzes, edits, and structures Microsoft Word (.docx) documents.
metadata:
  version: "1.0.0"
  tags: ["docx", "word", "document", "office-documents"]
---

# Word Document Skill

Use this skill whenever a user wants to create, read, edit, analyze, summarize, or restructure a Microsoft Word document (.docx).

## When to Use

Use this skill if the user asks to:

- Create a new Word document (.docx)
- Generate a document from a topic, prompt, notes, article, or report
- Create a report, essay, article, proposal, letter, memo, project document, or training document
- Convert structured content into a formatted Word document
- Extract text, headings, paragraphs, bullet points, or tables from a DOCX file
- Analyze document structure, readability, or organization
- Summarize document content
- Improve document flow, tone, formatting, and readability
- Add, remove, rewrite, or reorganize sections
- Create professional, educational, technical, academic, business, or marketing documents

## Document Creation Rules

When generating a new Word document:

- Use the registered `create_word_document` tool or `scripts/create_docx.py`.
- Always generate and save a real `.docx` file.
- Do not stop at a text outline unless the user explicitly asks for outline-only content.
- Use a clear document title.
- Use logical section headings.
- Include at least one explanatory paragraph in each main section when possible.
- Use concise paragraphs instead of long text blocks.
- Use bullet points for grouped ideas.
- Use tables only when structured comparison or tabular data improves clarity.
- Keep formatting readable, professional, and consistent.
- Avoid empty sections.
- Include a conclusion, summary, recommendations, or next steps when appropriate.

## Reading and Extraction Rules

When reading an existing Word document:

- Use the registered `read_word_document` tool or `scripts/read_docx.py`.
- Extract readable text from paragraphs and tables.
- Preserve the order of the document content as much as possible.
- Summarize or analyze the extracted content only after reading the file.

## Design Requirements

Consult `references/guide.md` and follow its recommendations for:

- Document Structure
- Content Density
- Heading Guidelines
- Professional Formatting
- Document Templates
- Readability Guardrails
- Lists and Tables

## Validation Checklist

Before finalizing any Word document:

- The document has a clear title.
- Section order is logical.
- Every main section has useful content.
- Every main section has at least one paragraph or a meaningful bullet list.
- Headings use a consistent hierarchy.
- Paragraphs are readable and not overly dense.
- Bullet lists are concise and relevant.
- Tables, if used, have clear headers and rows.
- Formatting is consistent.
- No empty sections remain.
- The document is suitable for reading, printing, and sharing.
- The saved `.docx` file path is returned to the user.

## Script Usage

When the user asks to create a new Word document, call `create_word_document` or `scripts/create_docx.py` to generate and save a real `.docx` file using Python `python-docx`.

The document generator accepts:

- `title`: the document title.
- `subtitle`: an optional subtitle.
- `sections`: an optional list of section objects.
- `output_path`: an optional `.docx` destination.

Each section object should use this shape:

```json
{
  "heading": "Section Heading",
  "paragraphs": ["Short paragraph one.", "Short paragraph two."],
  "bullets": ["Bullet one", "Bullet two"],
  "table": {
    "headers": ["Column A", "Column B"],
    "rows": [["Value A", "Value B"]]
  }
}
```

When the user asks to read or extract content from a Word document, call `read_word_document` or `scripts/read_docx.py`.

Use `scripts/verify.py` to validate the word-skill structure.

## Best Practices

- Build the document outline first, then generate the Word file.
- Prefer 3 to 7 sections for a general topic.
- Use headings to make the document easy to scan.
- Keep each section focused on one main idea.
- Use bullets for lists, benefits, causes, effects, steps, or examples.
- Use tables for comparisons, timelines, facts, feature matrices, or schedules.
- If the user provides source content, preserve the important details and organize them professionally.
- Return the saved `.docx` path after generation.
