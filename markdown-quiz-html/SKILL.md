---
name: markdown-quiz-html
description: Legacy entry point for building offline quiz HTML from reviewed markdown question banks. Use the newer offline-quiz-html skill for new work, but keep this alias for existing prompts and automations.
---

# Markdown Quiz HTML

## Legacy Alias

This skill is kept for compatibility.

For new work, use `$offline-quiz-html`.

When you do use this legacy entry point, follow the same markdown quiz workflow:

1. Inspect the reviewed markdown source first.
2. Preserve group order and question order.
3. Extract answers from the reviewed text, not OCR guesses.
4. Generate a self-contained HTML page.
5. Generate the companion answer file in the same order as the source.

## Notes

- Keep the implementation simple unless the user asks for richer quiz features.
- Prefer one clean HTML file plus the companion answer file when the markdown source requires it.
