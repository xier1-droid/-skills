---
name: document-editing
description: Edit DOCX/Word documents safely when preserving layout matters, especially for templates, reports, and multi-section files. Use this skill when locating real files, filling specific sections, editing Chinese-named documents, or verifying that only intended parts changed.
---

# Document Editing

## Overview

Use this skill for Word/DOCX tasks where structure matters more than freeform prose. It favors minimal edits, stable anchors, and post-edit validation over broad rewrites.

## Workflow

1. Inspect the real files first.
- Enumerate the directory and confirm the exact filename(s) before editing.
- Do not rely on memory for Chinese names, copied paths, or generated outputs.

2. Find stable anchors.
- Prefer headings, section titles, table labels, and unique phrases.
- Do not use paragraph indexes unless the document is known to be static.
- After any insert/delete above a section, re-find anchors before editing later sections.

3. Make the smallest safe change.
- Preserve untouched sections, headers, footers, tables, and signatures unless the user asked otherwise.
- Prefer local insertions or replacements over rewriting whole chapters.
- Keep formulas, code blocks, and lists consistent with the source style.

4. Handle paths and filenames safely.
- Treat non-ASCII paths as data, not shell text.
- Prefer directory enumeration and resolved paths over hardcoded strings.
- If a path fails once because of encoding, switch to a Unicode-safe lookup strategy instead of retrying the same command.

5. Validate after editing.
- Re-open the result and confirm the edited sections are in the right place.
- Check that later chapters were not shifted or overwritten.
- If visual rendering is available, inspect page images once. If the renderer is missing, state that clearly and stop retrying it.

## Guardrails

- Do not bulk-rewrite a report unless the user explicitly wants a full rewrite.
- Do not assume section numbers remain valid after earlier edits.
- Do not keep guessing when source content is missing or ambiguous; ask for the missing piece.
- Do not treat a rendering failure as a document failure if the environment lacks the renderer.
- Do not delete files in bulk; remove only explicitly named files when the user asks.

## Practical Checks

- Before editing: confirm file name, target section, and source document.
- During editing: anchor by heading text, then insert or replace only within that section.
- After editing: verify the exact edited paragraphs and the first untouched section that follows.
