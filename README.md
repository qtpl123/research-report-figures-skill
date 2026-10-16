# research-report-figures

一个用于生成**「专业调研 + 图文结合」Word 报告**的 Agent Skill。

把「主题 → 调研 → 配图 → Word 报告 → 版面检查」固化成可重复执行的流程，让 AI 助手按同一套标准产出带目录、图号、图注、来源与参考文献的 `.docx`，并渲染逐页检查排版。

## 它解决什么问题

让 AI 写报告不难，难的是**配图靠谱**。常见的失败方式是：随手画个流程图冒充分子结构、编一个看起来合理的文献引用、图表数据来源不明。

本 Skill 用三条硬规则把这些问题挡在生成之前：

1. **分子结构式、化学结构、抗体等生物大分子、真实实验数据 —— 必须引用真实来源原图**，禁止自绘或用流程图冒充
2. **禁止编造文献出处、数据和图片来源**，查不到就写「未找到可核实来源」
3. **真实图片必须有原始链接和许可信息**，缺任意一项，构建脚本直接拒绝生成文档

## 三层配图策略

| 层级 | 适用内容 | 工具 |
| --- | --- | --- |
| 第 1 层 · 技术示意图 | 工艺流程、技术路线、层级结构、数据对比 | Mermaid / Graphviz / matplotlib |
| 第 2 层 · 真实图片 | 实物照片、分子结构、文献原图、真实谱图 | 联网检索 + 来源登记 |
| 第 3 层 · AI 生成图 | 仅封面与装饰性插图，默认关闭 | 需在图注标注「AI 生成」 |

## 安装

把本仓库放入你的 Agent 技能目录，例如：

```
~/.workbuddy/skills/research-report-figures/     # WorkBuddy / Claude Code 风格的用户级技能目录
```

目录结构：

```
research-report-figures/
├── SKILL.md                  # 技能主文件（流程与红线规则）
├── references/
│   └── content-syntax.md     # content.md 语法参考
├── scripts/                  # 8 个工具脚本
└── examples/                 # 示例报告
```

## 依赖

| 依赖 | 用途 | 说明 |
| --- | --- | --- |
| Python 3 | 运行构建与图表脚本 | 需安装 `python-docx`、`Pillow`、`matplotlib`、`pypdfium2` |
| Node.js | Mermaid 渲染 | 需 `playwright`、`mermaid` |
| PowerShell 7+ | Graphviz 渲染、Word 导 PDF | Windows 环境 |
| Graphviz | 结构图渲染 | 需 `dot` 在 PATH 中 |
| Microsoft Word | Word 转 PDF（质检环节） | 仅 Windows + 本机装有 Word 时需要 |

安装 Python 依赖：

```bash
pip install python-docx Pillow matplotlib pypdfium2
```

## 路径配置（重要）

`SKILL.md` 中使用 `<TOOLS>`、`<PYTHON>` 作为占位符。首次使用前，请替换为你本机的实际路径，或让助手在使用时替换：

- `<TOOLS>` → 本仓库 `scripts/` 目录的绝对路径
- `<PYTHON>` → 装有上述依赖的 Python 解释器路径
- `node` / `pwsh` → 本机对应可执行文件

`scripts/mermaid_render.mjs` 支持用环境变量指定依赖位置：

| 环境变量 | 含义 | 默认值 |
| --- | --- | --- |
| `CODEX_NODE_MODULES` | 包含 playwright 的 `node_modules` 目录 | `<TOOLS>/../node/node_modules` |
| `MERMAID_JS` | `mermaid.min.js` 的完整路径 | `<CODEX_NODE_MODULES>/mermaid/dist/mermaid.min.js` |

## 使用

在支持 Skill 的 AI 助手中直接提出需求即可触发，例如：

> 帮我写一份《钠离子电池技术路线调研报告》，要图文结合，配工艺流程图和数据图表

助手会按 `SKILL.md` 的标准流程执行：联网调研 → 配图规划 → 生成图与登记来源 → 装配 Word → 逐页版面检查 → 交付。

产出目录结构：

```
<工作目录>/reports/<主题>/
├── content.md            # 报告正文
├── assets/               # 图片、图形源码、manifest.json
├── assets/qa/            # 版面检查产物（内部使用）
└── <报告标题>.docx        # 最终交付
```

## 示例

`examples/` 下为完整示例报告，可直接查看产出效果。

## 说明

- 本技能为**环境无关**设计：所有路径通过占位符与环境变量配置，不绑定任何特定机器。
- 联网检索环节需要你的助手具备网络访问能力，并在受限 sandbox 中经你批准。
- 真实图片的许可信息由使用者自行确认，对外发布前请核对原始来源的授权条款。

## License

[MIT](LICENSE)
