---
name: research-report-figures
description: 生成“专业调研 + 图文结合”的 Word 报告。按主题自动判断哪些知识点需要配图，优先用 Mermaid/Graphviz/matplotlib 生成准确的技术示意图与数据图，必要时联网检索有明确许可的真实图片，最后用 python-docx 装配成带目录、图号、图注、来源和参考文献的 .docx，并渲染逐页检查排版。当用户要求写调研报告、技术报告、综述、行业分析并希望配图，或要求“图文结合的报告/文档”时使用。
---

# 专业调研 + 图文结合 Word 报告

本 Skill 把“主题 → 调研 → 配图 → Word 报告 → 版面检查”固化为可重复执行的流程。

## 工具位置

所有工具脚本位于本 Skill 的 `scripts/` 目录（下称 `<TOOLS>`）。**使用前请先把 `<TOOLS>` 替换为你在本机的实际绝对路径**，或在环境变量中配置。

| 用途 | 命令 |
| --- | --- |
| 运行环境（Python） | `<PYTHON>`（需已安装：python-docx、Pillow、matplotlib、pypdfium2） |
| Mermaid 渲染 | `node <TOOLS>/mermaid_render.mjs` |
| Graphviz 渲染 | `pwsh -File <TOOLS>/graphviz_render.ps1` |
| 数据图表渲染 | `<PYTHON> <TOOLS>/chart_render.py` |
| 下载真实图片并登记来源 | `<PYTHON> <TOOLS>/fetch_image.py` |
| 装配 Word | `<PYTHON> <TOOLS>/report_build.py` |
| 版面质检 | `<PYTHON> <TOOLS>/report_check.py` |
| Word 导出 PDF | `pwsh -File <TOOLS>/word_to_pdf.ps1` |
| PDF 转页面图片 | `<PYTHON> <TOOLS>/pdf_to_png.py` |

其中：

- `<PYTHON>`：指向你本机装有上述依赖的 Python 解释器（如虚拟环境的 `python.exe` / `python`）
- `node`：Node.js 可执行文件
- `pwsh`：PowerShell 可执行文件（Windows 需 PowerShell 7+ 以支持脚本参数）

`mermaid_render.mjs` 需要的依赖（playwright、mermaid）可通过环境变量指定：

- `CODEX_NODE_MODULES`：包含 playwright 的 `node_modules` 目录
- `MERMAID_JS`：`mermaid.min.js` 的完整路径

默认按 `<TOOLS>/../node/node_modules` 查找。

联网检索可使用任意可用的检索工具或 Skill（沙箱环境默认禁止联网时，需经用户批准后执行）。

## 一、三层配图策略（默认行为）

按下面顺序为每个知识点选择配图方式，能不用图片就写成文字或表格。

**第 1 层：技术示意图优先（默认选择，准确可控）**

| 内容类型 | 使用工具 |
| --- | --- |
| 工艺流程、技术路线、上下游关系、决策分支、分子机制框架、实验原理 | Mermaid（`flowchart` / `sequenceDiagram` / `stateDiagram`） |
| 层级结构、正交关系、多分支归属、网络关系 | Graphviz（`dot`） |
| 关键数据对比、趋势、分布、占比、相关性 | matplotlib（`chart_render.py`） |
| 规格对比、多对象同维度比较 | Word 表格，不必配图 |

本层只负责"流程、层级、数据"类知识。**分子结构式、化学结构、抗体等生物大分子结构不属于本层**——Mermaid/Graphviz 画不出分子和 Y 形抗体，这类内容一律走第 2 层真实文献图。若个别生物结构确需自绘示意，必须形态准确、有辨识度（如抗体必须画出 Y 形并标注 Fab/Fab'/Fc/铰链区），禁止"两根线加一个框"式敷衍。

**第 2 层：真实图片检索（实物照片 + 文献原图）**

以下两类内容，AI 和自绘示意图都不能替代，必须检索真实图片：

（a）**实物照片**：设备外观、实验仪器、产品实物、生产现场、真实实验照片。

（b）**文献原图**：分子结构式、化学结构、抗体等生物大分子结构、作用机制图、真实谱图与实验数据图。读者理解化学与生物学概念完全依赖这类图，**禁止自绘、禁止用流程图冒充**。优先从开放获取文献（CC BY / CC BY-SA / 公有领域，如 PMC、MDPI、Frontiers 等期刊论文）引用原图；小分子化学结构也可取自 PubChem、DrugBank 等权威数据库的结构图，同属真实来源。

做法：

1. 用 anysearch 检索可信来源：官方网站、开放获取论文（优先 CC BY 许可）、维基共享资源等有明确许可的来源。
2. 用 `fetch_image.py` 下载并登记：原始链接、作者/机构、出处（含期刊与 Figure 编号）、许可、获取日期、SHA256。
3. 检索不到可核实来源时，不要用别的图顶替，也不要退回第 1 层自绘；直接在正文写明"未找到可核实的图片来源"，或按标准流程第 1 步向用户求助获取文献。

**第 3 层：AI 生成图（默认关闭）**

仅在用户明确要求时用于封面、概念示意、装饰性插图，且**必须**在图注中标注“AI 生成示意图，非真实结构/数据”。当前环境未接入 AI 生图能力。

## 二、红线规则（构建脚本会强制校验，违反则拒绝生成文档）

1. **分子结构式与化学结构必须引用真实来源的原图**（开放获取文献、PubChem 等权威数据库），逐图注明出处与许可；**禁止自绘、禁止用流程图或简笔示意冒充**。实验数据、设备照片、真实实验结果同理，必须用真实图片。抗体等生物大分子结构优先用文献原图；确需自绘示意时必须形态准确（如抗体的 Y 形与 Fab/Fc/铰链区分区标注），并在图注标明"结构示意图，非真实结构"。
2. **禁止编造文献出处、数据和图片来源。** 查不到来源就写“未找到可核实来源”，不要凑一个看起来合理的引用。
3. **真实图片必须有原始链接和许可信息**，缺任意一项都会校验失败。许可不明时写“未标注/需确认”，并在正文提示对外使用前需再确认。
4. **示意图与数据图必须写明依据来源**。数据曲线若为示意形态（非实测值），必须在图注中写明“示意形态，非实测数据”。
5. **区分三类图并保持可追溯**：真实照片（photo）、技术示意图（schematic）、数据图（chart）。报告结论性描述不得把示意图表述为实测结果。
6. 引用外部图表数据时，同步记录出处与许可；不得把他人图表改头换面作为原创。

## 三、标准流程

### 0 准备目录

```
<工作目录>/reports/<主题>/
  content.md          报告正文（Markdown 语法）
  assets/             图片、图形源码、manifest.json
  <报告标题>.docx     最终交付（用标题命名，规则见第 7 步交付）
  assets/qa/          版面检查产物（PDF 与页面 PNG，仅内部使用）
```

### 1 联网调研并整理来源

用 anysearch 检索（中文 + 英文各一轮，法规与厂商资料单列），把可核查的来源整理成参考文献列表。来源不足时继续补充检索，不要凭印象写结论。

**关键文献拿不到全文时，向用户求助，不要硬凑：**

- 检索到的关键文献若被付费墙挡住、只能看摘要、或 PDF 下载失败（403/404/需机构权限），不要跳过，也不要只凭摘要或二手转述写内容。
- 整理一份求助清单发给用户，每条包含：文献标题、作者、期刊与年份、DOI/链接、报告哪一节需要它、需要其中的哪张图或哪部分内容。
- 明确说明："这些文献我无法获取全文，如果您方便，请下载 PDF 后发给我。"
- 用户提供的 PDF 只用于本次报告的阅读与引用，引用格式仍按原文献题录标注。
- 拿到 PDF 后再读原文、提取图表与参数；最终仍拿不到的，改用其他开放获取来源替代，或在正文明确标注"该结论基于摘要/二手综述，未核对原文"。

### 2 做配图规划

写完大纲后，列出“章节 → 是否需要图 → 图类型 → 理由 → 来源策略”的清单，只给确实需要图的知识点配图。判断标准：

- 描述顺序、分支、层级、循环 → Mermaid / Graphviz
- 描述数量关系、趋势、对比 → matplotlib
- 描述分子结构、化学结构、抗体等大分子形态 → 真实文献原图（见第 2 层）
- 描述外观、实物、现场 → 真实图片检索
- 只是并列罗列参数 → 用表格

### 3 生成图与登记

技术示意图（Mermaid）：

```powershell
cd <报告目录>
node <TOOLS>/mermaid_render.mjs assets/fig-01-xxx.mmd assets/fig-01-xxx.png --width 1750 --scale 2 --svg assets/fig-01-xxx.svg
```

渲染器会在图形高宽比失衡时给出提示；出现“图形过高/过宽”提示时，改成 `flowchart LR`/`TB` 或拆成两张图，不要直接塞进 Word。

结构图（Graphviz）：

```powershell
pwsh -File <TOOLS>/graphviz_render.ps1 -InFile assets\fig-02-xxx.dot -OutFile assets\fig-02-xxx.png -Dpi 200
```

数据图（matplotlib）：先写 JSON 规格，再渲染。

```powershell
<PYTHON> <TOOLS>/chart_render.py assets\fig-03-xxx.json
```

支持 `bar` / `grouped_bar` / `stacked_bar` / `line` / `pie` / `scatter` / `box` / `chromatogram`（层析洗脱曲线示意）。

真实图片（需要用户批准联网）：

```powershell
<PYTHON> <TOOLS>/fetch_image.py --id fig-04 --url "<图片直链>" `
  --out assets/fig-04-xxx.jpg --title "设备实物" --source "来源机构" `
  --license "CC BY 2.0，作者 XXX" --author "XXX" --base "<报告目录>"
```

### 4 维护图片清单 manifest.json

生成的图要手工补进 `assets/manifest.json`（真实图片由脚本自动写入）。字段：

```json
{
  "id": "fig-01",
  "file": "assets/fig-01-xxx.png",
  "type": "schematic",
  "title": "图题（也会作为默认图注）",
  "generator": "mermaid",
  "source": "依据 XXX 公开资料整理",
  "note": "本图为技术示意图，非真实照片",
  "url": "真实图片必填：原始链接",
  "license": "真实图片必填：许可",
  "author": "作者或机构"
}
```

`type` 取值 `photo` / `schematic` / `chart` / `ai`，校验规则见红线部分。

### 5 写 content.md 并装配

语法见 `references/content-syntax.md`。装配：

```powershell
cd <报告目录>
<PYTHON> <TOOLS>/report_build.py content.md --out "<报告标题>.docx"
```

脚本会自动：校验图片清单与红线、按出现顺序编号、图片居中并限制在版心内、生成图注与来源行、表题在上表在下、插入目录域并设置“打开时更新域”。

### 6 质检与逐页检查（不可跳过）

```powershell
<PYTHON> <TOOLS>/report_check.py "<报告标题>.docx"
pwsh -File <TOOLS>/word_to_pdf.ps1 -Docx "<报告标题>.docx" -Pdf assets\qa\report.pdf
<PYTHON> <TOOLS>/pdf_to_png.py assets\qa\report.pdf assets\qa\pages --scale 1.6
```

必须逐页打开 PNG 检查：图片是否遮挡文字、是否超出页边距、图注是否与图分离、表格是否跨页断开或丢表头、是否有大面积无意义留白。发现问题的处理办法：

- 图注与图分离 → 减小 `MAX_IMAGE_HEIGHT_CM` 或把图放到章节开头
- 表格跨页且表头丢失 → 检查表头重复属性，或缩短表头文字
- 图过高导致整页独占 → 改图形布局或拆图
- 页面出现大块空白 → 调整正文长度分布或图尺寸，必要时在合适位置分页

修改后必须重新装配、重新导出、重新检查，直到没有阻塞问题。

### 7 交付

交付文件**必须用报告标题命名，不要叫 report.docx**。命名规则：取报告标题或其浓缩版（建议不超过 25 个字），去掉文件名非法字符（\ / : * ? " < > |）。示例：标题《ADC 抗体药物偶联物纯化工艺调研报告——从偶联化学到 DAR 分离》→ 文件名 `ADC纯化工艺调研报告.docx`。装配命令的 `--out` 参数即用此文件名；重命名应在装配前确定，避免装配后再改名导致 manifest 路径不一致。

同时说明：图片清单 `assets/manifest.json`、参考文献、以及内部质检产物 `assets/qa/` 的位置。质检产物不主动交付，除非用户需要。

## 四、写作与排版约定

- 正文中文用宋体五号、1.5 倍行距、首行缩进 2 字符；标题用微软雅黑、纯黑。
- 图题在图下方，表题在表上方；图号按出现顺序自动生成，正文用 `[[图:fig-01]]` 引用。
- 图注下方必须有一行来源/许可说明。
- 报告用词：先给结论，再给依据；不确定的内容明确标注不确定性。
- 不为凑图而配图：一段纯文字说明清楚的内容，不必生成图。

## 五、常见问题

- **Mermaid 渲染失败**：多为语法错误，先把 `.mmd` 内容精简到最小可复现版本定位问题。
- **中文变成方框**：检查系统中文字体是否存在（`C:\Windows\Fonts\msyh.ttc`），三个渲染器都依赖它。
- **Word 导出 PDF 失败**：确认本机 Word 可正常启动；脚本会释放 COM 对象，失败时先关闭残留的 WINWORD.EXE 进程。
- **联网检索被拒**：沙箱默认禁止网络，需要用户批准后重试，不要绕过。
- **图片下载 429/403**：更换来源或稍后重试；不要伪造图片内容。
- **关键文献 PDF 拿不到**：整理题录清单（标题、作者、DOI、报告哪节需要、要哪张图）向用户求助，用户可协助下载；不要凭摘要硬写。
