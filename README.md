# Skills

这个仓库收集可复用的 Codex skills。

## Skills

- [中文 DOCX 编辑](./chinese-docx-editing/SKILL.md)
- [Offline Quiz HTML](./offline-quiz-html/SKILL.md)
- [Markdown Quiz HTML (Legacy)](./markdown-quiz-html/SKILL.md)
- [Excel Offline Quiz HTML (Legacy)](./excel-offline-quiz-html/SKILL.md)

## 使用方式

- 在 Codex 里调用 `$chinese-docx-editing` 处理中文 Word 文档。
- 在 Codex 里调用 `$offline-quiz-html` 把审核过的 markdown 题库或 XLSX 题库生成离线刷题 HTML。
- 旧入口 `$markdown-quiz-html` 和 `$excel-offline-quiz-html` 保留兼容，建议新任务优先使用 `$offline-quiz-html`。
- 适合需要保留排版、题目顺序和离线刷题体验的题库整理任务。
