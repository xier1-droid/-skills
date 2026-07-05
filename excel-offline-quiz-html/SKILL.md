---
name: excel-offline-quiz-html
description: Legacy entry point for building offline quiz HTML from XLSX question banks. Use the newer offline-quiz-html skill for new work, but keep this alias for existing prompts and automations.
---

# Excel Offline Quiz HTML

## Legacy Alias

This skill is kept for compatibility.

For new work, use `$offline-quiz-html`.

When you do use this legacy entry point, keep following the reviewed workbook workflow:

1. Inspect the workbook structure first.
2. Identify the true question sheets and ignore summary or blank sheets.
3. Determine the grouping key from the actual columns.
4. Keep the page self-contained and offline.
5. Validate total question count and group distribution before finishing.

## Notes

- Prefer the newer unified skill for any new task.
- Keep the page single-file unless the user explicitly asks otherwise.
