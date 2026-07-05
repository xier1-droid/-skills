---
name: offline-quiz-html
description: Build self-contained offline quiz HTML from reviewed markdown or XLSX question banks, including companion answer files for markdown inputs. Use when Codex needs to turn a markdown exercise set, fill-in-the-blank quiz, exam bank, or spreadsheet question bank into a single offline practice page.
---

# Offline Quiz HTML

## Overview

Turn reviewed quiz sources into a single offline practice page that opens by double-clicking.

Use this skill for:

- markdown question banks with grouped headings, numbered questions, and explicit answers
- XLSX question banks that follow the reviewed workbook layout
- generating a markdown companion answer file when the source is markdown

## Workflow

1. Inspect the source file first and confirm whether it is markdown or XLSX.
2. Preserve the original group order and question order.
3. Extract answers from the reviewed text or the workbook cells, not from OCR guesses.
4. Generate a self-contained HTML page with inline CSS and JS.
5. For markdown sources, also generate a companion `.answers.md` file unless an explicit answer path is provided.
6. Validate that every group and question is present exactly once.

## Input Rules

- Treat markdown headings as group boundaries when present.
- Treat numbered markdown lines as questions.
- Use explicit answer markers when present, and fall back to the final parenthetical answer only when it is clearly part of the reviewed source.
- For XLSX sources, use the reviewed worksheet headers and question rows rather than visible sheet names alone.
- Keep Chinese and English text UTF-8 safe.

## Output Rules

- Keep the HTML self-contained: no CDN assets, no network dependencies, no external fonts.
- Render both choice questions and text-answer questions.
- Support submit, reset, group-level grading, and all-up grading.
- Keep the layout readable on desktop and mobile.
- Prefer one HTML file, plus one answer file only when markdown sources are involved.

## Validation

- Confirm the total question count and group count.
- Confirm the HTML opens offline.
- Confirm at least one choice question and one text-answer question grade correctly when both appear.
- Confirm the markdown answer file matches the question order.
- Use `scripts/build_offline_quiz_html.py` as the canonical generator.
