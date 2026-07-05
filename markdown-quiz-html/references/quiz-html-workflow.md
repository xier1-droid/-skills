# Markdown Quiz HTML Workflow

## Recommended input structure

Use this skill when the source markdown has a predictable pattern such as:

- one group heading per section
- one word box per group
- numbered fill-in-the-blank questions under each group
- optional OCR notes or uncertainty markers that have already been manually reviewed

## Output expectations

Generate two files:

- a self-contained `.html` quiz page with inline CSS and JS
- a companion answer file that preserves the original order

## Implementation notes

- Prefer offline-friendly HTML with no external dependencies.
- Keep group order and question order unchanged.
- Use the reviewed source as the answer authority.
- Accept only clearly supported answer variants.
- Include group-level grading, total scoring, and reset behavior when useful.
- Keep the page readable on desktop and mobile.

## Validation checklist

- Count groups and questions before generating output.
- Confirm every question appears exactly once in both outputs.
- Verify UTF-8 rendering for mixed Chinese and English text.
- Check for broken punctuation, duplicated prompts, or missing answers.
- Test that the generated HTML opens directly in a browser without a build step.
