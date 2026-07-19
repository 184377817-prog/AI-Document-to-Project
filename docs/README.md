# 文档中心

这里集中放置项目案例、架构、评测方法和可交付文档。建议招聘方按下列顺序阅读。

## 核心材料

| 顺序 | 文档 | 适合了解 |
| ---: | --- | --- |
| 1 | [项目案例复盘](01-case-study.md) | 业务问题、产品方案、关键决策、版本数据与复盘 |
| 2 | [Agent 架构设计](02-architecture.md) | Code 与 LLM 分工、中间节点、校验和状态机 |
| 3 | [评测与上线门槛](03-evaluation.md) | 自动评测、人工复核、复杂度分层和发布判断 |
| 4 | [公开版 PRD](04-product-requirements.md) | 用户流程、功能需求、异常处理、埋点与验收标准 |
| 5 | [公开版 AI Specification](05-ai-specification.md) | 输入输出契约、字段规则、Code/AI 边界与降级策略 |
| 6 | [匿名评测报告](06-evaluation-report.md) | 逐样例证据、可复算指标、自动/人工分歧与版本决策 |

## 可验证材料

- [交互原型](../prototype/index.html)：演示上传、后台生成、预览与确认创建。
- [示例 Excel](../demo/sample_project_plan.xlsx)：真实 `.xlsx` 输入。
- [端到端转换脚本](../src/excel_to_project_demo.py)：从 Excel 读取到项目 JSON。
- [匿名评测数据](../evaluation/anonymized_results.csv)：20 个样例的逐条结果。
- [评测汇总脚本](../evaluation/summarize_results.py)：复算均分、通过分布和耗时。

## 可在线预览的成品

| 文档 | PDF 在线阅读 | Word 源文件 |
| --- | --- | --- |
| AI 产品项目推进 SOP | [打开 PDF](guides/AI产品项目推进SOP.pdf) | [下载 Word](source/AI产品项目推进SOP.docx) |
| AI 产品项目复盘指南 | [打开 PDF](guides/AI产品项目复盘指南.pdf) | [下载 Word](source/AI产品项目复盘指南.docx) |

## 为什么 Word 页面看起来是空的？

`.docx` 是压缩封装的二进制 Office 文件，GitHub 的代码浏览器不能像 Markdown、图片或 PDF 一样直接渲染它。因此 Word 文件页面通常只显示 `View raw`，并不代表文件为空。

- 想直接阅读：打开上表中的 PDF。
- 想编辑或查看 Word 样式：点击 `Download raw file` 下载源文件。
- 想看生成方式：查看仓库 `src/` 下的文档生成脚本。

两份 PDF 均由对应 Word 文件转换，并经过逐页渲染检查。
