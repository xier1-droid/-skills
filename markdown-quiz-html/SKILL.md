---
name: markdown-quiz-html
description: Generate self-contained quiz HTML pages and companion answer files from reviewed markdown question banks. Use when Codex needs to turn a markdown-based exercise set, fill-in-the-blank quiz, or exam bank into an offline HTML practice page plus a matching answer file.
---

# Markdown Quiz HTML

## Purpose

Turn a cleaned markdown question bank into two deliverables:

- a single-file HTML practice page that works offline
- a companion answer file that preserves the original questions and answers

Use this skill for quiz banks that follow a repeatable structure such as `group heading -> word box -> numbered questions`.

## Workflow

1. Inspect the source markdown first.
2. Identify the repeating group structure and confirm the question count per group.
3. Extract the authoritative answers from the reviewed text, not from OCR guesses.
4. Generate a self-contained HTML page with inline CSS and JS.
5. Add per-group controls, scoring, and clear answer reveal behavior.
6. Generate the answer file in the same order as the source.
7. Validate that every group and question is present exactly once.

## HTML Requirements

- Keep the page self-contained so it opens by double-clicking.
- Preserve group order and question order.
- Render each group separately with its word box.
- Use clear input fields and show grading feedback after submission.
- Support group-level grading, overall grading, and reset actions when appropriate.
- Make the layout readable on desktop and mobile.
- Use UTF-8 and avoid introducing encoding issues in Chinese and English mixed text.

## Answer Extraction Rules

- Prefer the user-reviewed source text over OCR-derived guesses.
- If the source has an explicit answer in parentheses, use it as the primary answer.
- Accept close inflections only when the reviewed content or sentence context clearly supports them.
- If a line remains ambiguous, keep the original wording intact and add the most defensible answer without rewriting the question.
- Preserve the source order and numbering exactly.

## Common Checks

- Count groups and questions before writing output.
- Compare the HTML answer key against the answer file.
- Scan for broken Unicode, missing punctuation, or duplicated questions.
- Confirm the generated HTML contains the expected grading logic before finishing.

## Notes

- Keep the implementation simple unless the user asks for richer quiz features.
- Prefer one clean HTML file over a multi-file frontend unless the task clearly needs more.
- If the markdown structure is inconsistent, normalize only the minimum needed to keep the quiz usable.
