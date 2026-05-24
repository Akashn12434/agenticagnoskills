---
name: ppt-skill
description: Creates, reads, analyzes, edits, and rearranges Microsoft PowerPoint (.pptx) presentations.
metadata:
  version: "1.0.0"
  tags: ["pptx", "powerpoint", "presentation", "office-documents"]
---

# PowerPoint Presentation Skill

Use this skill whenever a user wants to create, read, edit, analyze, summarize, or restructure a PowerPoint presentation (.pptx).

## When to Use

Use this skill if the user asks to:

- Create a new PowerPoint presentation (.pptx)
- Generate slides from a topic, prompt, notes, article, or report
- Convert structured content into presentation slides
- Extract text, titles, bullet points, or notes from a PPTX file
- Analyze presentation structure and slide organization
- Summarize slide content
- Improve presentation flow and readability
- Add, remove, duplicate, or reorder slides
- Update slide titles, bullet points, or slide content
- Create business, educational, technical, marketing, training, or project presentations


## Presentation Creation Rules

When generating a new presentation:

- use 16:9 widescreen format.
- Every slide must have a title.
- Use concise bullet points instead of long paragraphs.
- Keep content readable and uncluttered.
- Ensure text fits within slide boundaries.
- Use consistent formatting across all slides.
- Avoid empty placeholders.
- Use professional layouts.
- Include a conclusion or thank-you slide when appropriate.

## Design Requirements

Consult references/guide.md and follow its recommendations for:

- Content Density
- Font Guidelines
- Color Guidelines
- Layout Rules and Recommendations
- Presentation Flow
- Presentation Templates
- Design Guardrails


## Validation Checklist

Before finalizing any presentation:

- Every slide has a title.
- Content is readable.
- No excessive text blocks exist.
- Formatting is consistent.
- Slide order is logical.
- Presentation structure is complete.
- No empty slides remain.
- Presentation is suitable for projection and screen viewing.

## Script Usage

When the user asks to create a new presentation, call `scripts/create_ppt.py` or the
registered `create_powerpoint_presentation` tool to generate and save a real `.pptx`
file. Do not stop at a text outline unless the user explicitly asks for outline-only
content.

The presentation generator accepts:

- `topic`: the presentation topic.
- `slides`: an optional list of slide objects with `title`, `subtitle`, and `bullets`.
- `output_path`: an optional `.pptx` destination.

Use scripts/verify.py to validate the ppt-skill structure.

## Best Practices

- Build slide content first, then generate the PowerPoint file.
- Keep each slide focused on one idea.
- Prefer 7 to 10 slides for a general educational topic.
- Return the saved `.pptx` path to the user after generation.
- If custom slide content is available, pass it to the generator instead of using the fallback outline.



