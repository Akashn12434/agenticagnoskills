---
name: ppt-skill
description: Creates and improves Microsoft PowerPoint (.pptx) presentations.
metadata:
  version: "1.0.0"
  tags: ["pptx", "powerpoint", "presentation", "office-documents"]
---

# PowerPoint Presentation Skill

Use for PowerPoint creation, reading, analysis, editing, summarization, or slide restructuring.

## When to Use

- Create `.pptx` decks from topics, notes, articles, reports, or structured outlines.
- Extract, summarize, analyze, update, reorder, or improve slide content.
- Build business, educational, technical, marketing, training, or project presentations.

## Presentation Creation Rules

- Use `create_powerpoint_presentation` or `scripts/create_ppt.py`; save a real `.pptx` unless the user asks for outline-only content.
- Use 16:9 widescreen, clear slide titles, concise bullets, readable spacing, and consistent formatting.
- Avoid empty slides, dense paragraphs, tiny text, and placeholder content.
- Include a summary, conclusion, thank-you, or questions slide when appropriate.

## Design Requirements

Follow `references/guide.md` for density, typography, colors, layouts, flow, templates, and design guardrails.

## Validation Checklist

- Every slide has a title and useful content.
- Slide order is logical and complete.
- Text is readable and not overcrowded.
- Formatting is consistent.
- The saved `.pptx` path is returned.

## Script Usage

Generator inputs:

- `topic`: presentation topic.
- `slides`: optional slide objects with `title`, optional `subtitle`, and `bullets`.
- `output_path`: optional `.pptx` destination.

Use `scripts/verify.py` to validate the skill structure.

## Best Practices

- Build slide content before calling the generator.
- Keep each slide focused on one idea.
- Prefer 7 to 10 slides for general educational topics.
- Pass custom slide content when available instead of using fallback slides.
