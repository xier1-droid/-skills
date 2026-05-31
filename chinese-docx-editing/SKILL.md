---
name: chinese-docx-editing
description: Safely edit Chinese DOCX/Word templates and reports while preserving layout, anchors, tables, headers, and untouched sections. Use this skill when filling Chinese-named documents, updating specific report sections, or verifying that only intended paragraphs changed.
---

# 中文 DOCX 编辑

## Overview

Use this skill for Chinese Word/DOCX tasks where structure matters more than freeform prose. It favors minimal edits, stable anchors, and post-edit validation over broad rewrites.

## Workflow

1. 先确认真实文件。
- 先枚举目录，确认实际文件名再编辑。
- 中文文件名、复制路径和生成输出都不要靠记忆。

2. 找稳定锚点。
- 优先用标题、章节名、表格标签和唯一短语定位。
- 不要依赖段落序号，除非文档是静态模板且前面不会插入内容。
- 上方有增删后，编辑后续章节前要重新找锚点。

3. 做最小安全修改。
- 默认保留未触碰章节、页眉页脚、表格和签名区。
- 优先局部插入或替换，不要整章重写。
- 公式、代码块和列表要跟原文风格一致。

4. 安全处理路径和文件名。
- 把非 ASCII 路径当成数据，不要直接硬拼命令文本。
- 优先用目录枚举和已解析路径，不要手写中文路径字符串。
- 如果路径因为编码失败过一次，马上换 Unicode 安全的查找方式，不要原样重试。

5. 编辑后验证。
- 重新打开结果，确认修改落在正确位置。
- 检查后续章节有没有被错位或覆盖。
- 如果能渲染页面图，就只做一次视觉检查；如果渲染器缺失，就明确说明，不要反复重试。

## Guardrails

- 不要在用户没要求时整份重写报告。
- 前面插入过内容后，不要再假设章节编号还有效。
- 来源内容缺失或含糊时，不要继续猜，直接补齐信息。
- 环境没有渲染器时，不要把渲染失败当成文档失败。
- 不要批量删除文件；用户要删时只处理明确指定的单个文件。

## Practical Checks

- 编辑前：确认文件名、目标章节和来源文档。
- 编辑中：先按标题文字定位，再只在该章节内插入或替换。
- 编辑后：核对修改段落和后面第一个未触碰章节。
