# SSR-PPT-SKILL

PPT 制作与美化 Skill，适用于 **Codex** 与 **Claude Code**。它从风格配方出发确定视觉方向，配合设计字体与中文排版、逐页策划、方向比较和实际渲染验收，制作有记忆点的演示文稿；支持离线 HTML 与原生可编辑 PPTX。

显示名称为 **SSR-PPT-SKILL**，调用名为 **`ssr-ppt-skill`**（Codex 中写作 `$ssr-ppt-skill`，Claude Code 中写作 `/ssr-ppt-skill`，也可以在对话中直接提出需求自动触发）。

## 设计原则：推荐优先，不硬套

Skill 由两部分组成：

- **方法**：找焦点、定构图、排节奏、做验收。
- **推荐资源**：风格配方、版式原型、字体搭配、图表做法。每个配方和原型都先写明“适用”与“不适用”。

使用推荐资源时：

- 内容合适就优先参考。
- 不合适就不用，按方法从内容出发自行设计。
- 配方和原型都可以修改、组合，清单也不封闭，效果好的自定方案可以在全套中复用。
- 不为套用资源而凑内容、删内容或扭曲内容。

## 功能

- **视觉方向**：13 套风格配方，涵盖瑞士编辑网格、暗场发布会、数据新闻、日式杂志、东方当代、工程蓝图、粗野海报、静奢品牌、温暖人文、咨询商务、光感科技、学术研究、纪实影像。选用时结合主题改编色彩、视觉母题与尺度；都不贴切时自定方向。
- **版式原型**：22 个常见页面任务的构图起点，包括封面、章节页、金句、大数字、主图加结论、指标组、前后对比、方案比较、2×2 矩阵、时间线、流程、架构图、议题树、出血图、图文分屏、照片组、人物、编号要点、长文页、表格和结束页。每个原型都写明适用与不适用，并给出构图、尺度、变体、常见错误和 PPTX 实现要点。
- **字体与中文排版**：免费可商用中文字体清单与搭配、字号阶梯、中文断行与中西混排规则。HTML 通过子集化把设计字体内联进单文件（中文字体从十几 MB 压到约 100 KB）；PPTX 按字体策略设置西文、东亚与主题字体。
- **反“AI 味”清单**：列出等宽卡片墙、左侧色条、紫蓝渐变、emoji 图标、空洞标题等常见廉价做法，以及对应的改法。
- **逐页策划**：每页记录能量（高冲击 / 标准 / 密集 / 呼吸）、核心信息、依据、焦点、构图与可编辑要求，并提前规划高潮页。
- **方向比较与出彩检查**：较长演示先为封面和关键页做 2–3 个不同方向并排比较；验收不只检查有没有错，还做去品牌、截图、模糊与灰度测试。
- **图片与图表**：生成图提示词规范、调色与遮罩手法、没有素材时的图形方案、统一图标；图表选型与美化默认值，HTML 用手写 SVG，PPTX 用原生图表并覆盖默认样式。
- **HTML 播放器**：
  - 悬浮控制条自动隐藏，配进度条与淡入、滑动或缩放过渡。
  - 支持入场动画（含线条描绘）、逐步揭示、概览（O）、演讲备注（N）、黑屏（B）、全屏（F）。
  - 网址记住页码，触屏可滑动翻页，按画布尺寸打印。
- **PPTX**：正文、表格、数据图表和指定的示意图使用原生对象，并在副本中做编辑测试。

用户未指定格式时，默认交付 HTML、PPTX 和 `planning.md`。用户只要求一种格式或只要求大纲时，按该范围执行。

## 安装

把本仓库的 `ssr-ppt-skill` 文件夹放进对应工具的技能目录：

| 工具 | 技能目录 |
|---|---|
| Codex | 设置了 `CODEX_HOME` 时为 `$CODEX_HOME/skills`，否则为 `~/.codex/skills` |
| Claude Code（个人） | `~/.claude/skills` |
| Claude Code（单个项目） | `<项目>/.claude/skills` |

### Windows / PowerShell

```powershell
git clone https://github.com/EuronSSR-01/SSR-PPT-SKILL-2.git
# Codex：
$ssrSkillsDir = if ($env:CODEX_HOME) { Join-Path $env:CODEX_HOME 'skills' } else { Join-Path $HOME '.codex/skills' }
# Claude Code 改用：$ssrSkillsDir = Join-Path $HOME '.claude/skills'
$ssrSkillTarget = Join-Path $ssrSkillsDir 'ssr-ppt-skill'
if (Test-Path -LiteralPath $ssrSkillTarget) { throw '已存在同名 Skill，请先决定是否更新。' }
New-Item -ItemType Directory -Path $ssrSkillsDir -Force | Out-Null
Copy-Item -LiteralPath './SSR-PPT-SKILL-2/ssr-ppt-skill' -Destination $ssrSkillTarget -Recurse
```

### macOS / Linux

```bash
git clone https://github.com/EuronSSR-01/SSR-PPT-SKILL-2.git
ssr_skills_dir="${CODEX_HOME:-$HOME/.codex}/skills"   # Claude Code 改用：ssr_skills_dir="$HOME/.claude/skills"
mkdir -p "$ssr_skills_dir"
if [ -e "$ssr_skills_dir/ssr-ppt-skill" ]; then
  printf '%s\n' '已存在同名 Skill，请先决定是否更新。'
else
  cp -R ./SSR-PPT-SKILL-2/ssr-ppt-skill "$ssr_skills_dir/"
fi
```

### 可选依赖

脚本只在需要时使用，缺少时 Skill 会说明受影响的验收项：

```bash
python -m pip install fonttools brotli playwright pillow pymupdf python-pptx
```

- 浏览器：Playwright 会自动使用本机已安装的 Chrome 或 Edge；都没有时运行 `python -m playwright install chromium`。
- PPTX 渲染需要 LibreOffice。
- Claude Code 中新建 PPTX 使用 PptxGenJS，Skill 会把它安装到私有构建目录，不改动你项目的依赖。

## 使用

```text
使用 ssr-ppt-skill，基于这些资料制作面向管理层的 10 页汇报，
交付 HTML 演示和可编辑 PPTX，保留策划稿。资料和需求已完整，请连续完成并实际预览验收。
```

```text
使用 ssr-ppt-skill 美化这份演示，保留事实、数据、总页数和原生可编辑对象。
根据主题自主确定视觉方向，重点改善构图、中文排版、数据图表和整套节奏，完成后实际预览。
```

```text
使用 ssr-ppt-skill，把这篇报告做成 8 页离线 HTML 演示，只交付 HTML 和策划稿。
```

```text
使用 ssr-ppt-skill，先为这份产品介绍整理大纲和逐页策划，并给我两个封面方向，暂不制作成品。
```

提供受众、目的、总页数、模板或品牌要求、素材与输出格式，可以减少需要澄清的信息。精确的总页数包含封面和结束页。

## 工作方式

需求与证据整理 → 大纲与逐页策划（含能量与高潮页）→ 视觉定调（参考或自定配方、字体、色彩）→ 逐页构图（参考或自定原型）→ 方向比较 → 代表页校准 → 整套制作 → 实际预览、反模式与出彩检查 → 修复 → 交付。

根据任务规模缩短流程，局部修改沿用现有视觉体系。两种格式使用相同的内容和数据，分别制作、分别验收。用户模板与品牌优先；没有参考时由 Skill 自主确定视觉概念，不要求用户先挑风格。

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/embed_fonts.py` | 按 HTML 实际用到的字符子集化字体，以 base64 内联；报告缺字；可变字体可固定为单一字重 |
| `scripts/preview.py html` | 逐页截图（完全揭示状态）；检查离线、控制台、键盘、网址页码、打印页数与纸张；生成总览、灰度与模糊版本 |
| `scripts/preview.py pptx` | 经 LibreOffice 渲染逐页 PNG；统计原生对象与实际字体；`--edit-test` 在副本中改文字、表格和图表后重新载入并渲染 |
| `scripts/preview.py compare` | 把多个方向并排放在一张图中比较 |
| `scripts/set_theme_fonts.py` | 设置 PPTX 主题的标题与正文字体（含东亚字体），让图表中文和新增文本使用设计字体 |

生成的图片必须实际查看（Claude Code 用 Read 读取 PNG），脚本输出的数字只用于辅助定位问题。

## 环境与编辑范围

这是工作指令、HTML 播放器与辅助脚本，不包含 Office 或 PPTX 转换引擎。

- Codex 桌面环境优先使用官方 `presentations` Skill、bundled runtime 与 `@oai/artifact-tool`。
- Claude Code 用 PptxGenJS 新建 PPTX，用 python-pptx 修改已有文件或模板，用 LibreOffice 渲染。
- 照片、标志、图标和插画是图片；文字、表格和数据图表的原生编辑能力单独验证。
- HTML 动画和复杂网页效果不保证在 PPTX 中原样保留；整页截图或整页 SVG 不算原生可编辑 PPTX。
- PPTX 不随文件携带字体，使用非系统字体时需要在演示电脑上安装，Skill 会说明所选的字体策略。
- 默认由当前智能体执行，不要求其他模型或外部 CLI。Web Brain / Inner Web Brain 仅在用户为当前任务明确要求时使用。

缺少必要能力时，Skill 会说明已完成的内容与未完成的验收项，不把未检查的文件称为完成。

## 已验证范围

本次更新在 Windows 11 上用 Claude Code 按 Skill 流程制作了一份 6 页中文示例演示（“R3 数据新闻”配方改编，数据全部虚构），包括封面、大数字页、原生条形图、逐步揭示页、表格页和结束页。验证内容：

- **HTML**（Chrome 155）：
  - `embed_fonts.py` 将 Noto Sans SC（可变字体）与 Noto Serif SC（固定为 700）子集化，分别约 124 KB 与 90 KB，单文件 HTML 约 323 KB，没有缺字。
  - `preview.py html` 检查无外部请求、无控制台错误；键盘翻页、首尾边界和网址页码正常；打印为 6 页，纸张 960×540pt。
  - 另用交互测试覆盖了逐步揭示、交叉淡入、概览（方向键、Tab、点击、Esc）、备注、黑屏、空闲隐藏、按钮与空格不重复翻页、触屏滑动、网址直达和减少动态效果模式，共 28 项全部通过。
- **PPTX**（PptxGenJS 4.0.1，LibreOffice 26.8 渲染）：
  - 文字、图表和表格都是原生对象，没有整页图片。
  - `--edit-test` 在副本中修改文字、表格单元格和图表值，保存、重新载入并渲染，三处修改均生效。
  - `set_theme_fonts.py` 写入了主题东亚字体。
- **看图验收**：逐页查看了 HTML 与 PPTX 的渲染图，按检查结果修复了封面图形与标题重叠、图表标注重叠、页面下部空洞、概览高亮双重描边、PPTX 柱内标签对比度不足等问题。
- **文档核对**：`charts.md` 中引用的 PptxGenJS 选项名已在 4.0.1 源码中核对；python-pptx 的 API 名称已在 1.0.2 中核对。

版式原型库另用一份 11 页 HTML 示例验证（“R1 瑞士编辑网格”配方，数据虚构），覆盖 L02、L03、L07–L14 和 L20：

- `preview.py html` 检查全部通过：11 页打印，无外部请求，无报错。
- 看图发现并修复了以下问题，相关数值和避免方法已写回 `layouts.md`：
  - 指标组的单位换行；
  - 矩阵结论句末行只剩一个字，标签与圆点对应不清；
  - 2×2 象限按栏宽设计时放不下；
  - 时间线末端标签出界、连线穿过刻度标签；
  - 架构图多个箭头挤在同一点。
- 第一份示例已覆盖 L01、L05、L06、L19、L21、L22。

尚未覆盖：

- PowerPoint 实机测试。LibreOffice 看不出主题东亚字体缺失的影响，这一项依据 OOXML 规则处理，未经 PowerPoint 实测。
- 本轮没有在 Codex 中重新运行。
- 摄影或生成图片的设计效果，包括依赖照片的原型 L15–L18。
- 除 R1、R3 以外的 11 套配方的实际效果；原型 L04 也未单独渲染。

审美效果因主题和受众而异，示例只用于检验流程与交付能力。

## 文件结构

```text
ssr-ppt-skill/
├── SKILL.md
├── agents/openai.yaml          # Codex 界面元数据
├── assets/
│   ├── player.css
│   └── player.js
├── references/
│   ├── planning.md             # 逐页策划与能量节奏
│   ├── art-direction.md        # 视觉定调、设计数值、节奏、构图
│   ├── style-recipes.md        # 13 套风格配方（含适用 / 不适用）
│   ├── layouts.md              # 22 个版式原型（含适用 / 不适用）
│   ├── typography.md           # 字体、字号阶梯、中文排版、字体落地
│   ├── anti-patterns.md        # 平庸与“AI 味”反模式
│   ├── imagery.md              # 图片、生成图、图标、无素材方案
│   ├── charts.md               # 图表与表格
│   ├── visual-review.md        # 方向比较、代表页校准、出彩检查
│   ├── html.md                 # HTML 结构、播放器、检查
│   └── pptx.md                 # 原生 PPTX 制作与验收
└── scripts/
    ├── embed_fonts.py
    ├── preview.py
    └── set_theme_fonts.py
```

播放器资源应内联进最终 HTML，不引用技能安装目录。
