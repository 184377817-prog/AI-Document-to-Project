# AI Document to Project

一套将零散的 AI 产品经验整理为可执行项目文档的方法与示例，覆盖项目立项、方案设计、Agent 设计、开发测试、上线复盘等关键环节。

本仓库展示的不只是最终 Word 文档，也保留了可复现的 Python 生成源码和流程图资产，便于查看文档工程化的实现方式。

## 项目内容

| 内容 | 说明 |
| --- | --- |
| [AI 产品项目推进 SOP](docs/AI产品项目推进SOP.docx) | 从需求分析到上线复盘的七阶段推进框架，包含流程图、阶段目标、关键产出物及评审要点 |
| [AI 产品项目复盘指南](docs/AI产品项目复盘指南.docx) | 面向 AI 产品全生命周期的复盘框架，覆盖业务价值、数据、模型、Agent、评测、安全、成本与运营指标 |
| [项目推进流程图](assets/ai_stage_flow.png) | 七阶段项目推进路径 |
| [生命周期闭环图](assets/ai_lifecycle_flywheel.png) | 从业务目标到持续优化的反馈闭环 |

## 图示预览

![AI 产品项目七阶段推进流程](assets/ai_stage_flow.png)

![AI 产品全生命周期闭环](assets/ai_lifecycle_flywheel.png)

## 核心思路

1. **从文档到行动**：每个阶段都明确目标、输入、输出、协作角色与验收标准。
2. **把 AI 特性纳入项目管理**：除传统功能测试外，增加准确性、稳定性、幻觉、权限、成本与人工介入等评估维度。
3. **沉淀可复用资产**：将 PRD、AI Spec、Prompt/Context、评测集、监控指标和复盘结论纳入版本化管理。
4. **文档工程化**：使用 Python 生成结构化 Word 文档及图示，减少重复排版并保证输出一致性。

## 仓库结构

```text
.
├── assets/                      # 流程图与文档插图
├── docs/                        # 可直接阅读的 Word 成品
├── src/
│   ├── build_ai_product_sop.py  # 生成基础 SOP
│   ├── add_stage_diagram.py     # 生成流程图并写入 SOP
│   └── build_ai_project_retro.py# 生成复盘指南及生命周期图
├── requirements.txt
└── README.md
```

## 本地生成

需要 Python 3.10 或更高版本。

```bash
python -m venv .venv
```

Windows：

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python src/build_ai_product_sop.py
python src/add_stage_diagram.py
python src/build_ai_project_retro.py
```

macOS / Linux：

```bash
source .venv/bin/activate
pip install -r requirements.txt
python src/build_ai_product_sop.py
python src/add_stage_diagram.py
python src/build_ai_project_retro.py
```

生成结果会写入 `docs/`，图示会写入 `assets/`。

## 可重点查看

- 项目生命周期的阶段拆解与交付物设计
- AI Spec、Agent 工具调用和安全边界的定义方式
- 评测指标、上线门槛和持续监控体系
- 使用 Python 与 `python-docx` 实现结构化文档生成

## 说明

仓库内容为通用方法论与脱敏示例，不包含原公司、客户或业务系统的内部代码与数据。文档与代码在 AI 辅助下完成，由作者负责需求定义、内容判断、结构设计、校验与最终交付。
