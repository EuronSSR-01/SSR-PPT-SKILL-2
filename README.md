# SSR-PPT-SKILL

演示制作技能，采用通用的 Agent Skills（`SKILL.md`）格式，可用于 Codex、Claude Code、Cursor、Grok Build、Antigravity、DeepSeek Harness 等 agent：借助风格配方、版式原型、字体与中文排版、图标与配图、逐页策划和实际渲染验收，做出好看、有记忆点、字少图多的演示；支持公司模板、离线 HTML 和原生可编辑 PPTX。仅供能直接查看图片的模型使用。

名称保持 **SSR-PPT-SKILL**，技能名为 `ssr-ppt-skill`。调用方式见下文“使用”。

## 这次升级

- **设计规则回到第一版**：风格配方、版式原型、视觉定调、反模式、字体、图表和策划的写法回到 SKILL 2 的第一版（8de60d2）。后来加的“内容先行”和“图示优先就是多画关系图”的规则撤下，它们让成品文字越来越多、信息图越来越密。脚本和工具的改进保留。
- **图示优先＝图片、图标、示意插画**：凡是能用图来表示的文字，尽量用图来表示。这里的图首先是好看的图：表头、卡片、流程节点和关键数字配图标；讲主体、场景、产品时上网搜授权合适的图片，或用 SVG 画出产品、物件和场景的示意。图里基本不放字，不追求塞满文字的信息图。页面只留结论、关键数字和短标签，解释放备注，同一句话不在多页重复。
- **自带图标库**：`assets/icons/` 收录 127 个 Lucide 线性图标（ISC 许可，索引含中文含义）。`pptx_kit.py` 的 `icon()` 把图标做成 PowerPoint／WPS 里能改颜色、缩放不糊的原生形状；`svg_picture()` 把 SVG 插画按 2 倍分辨率转成图片插入。
- **开场两问**：继承已有偏好；尚未明确的，开始前用同一条消息问：① 是否允许先询问几个问题；② 表达方式选“图示优先”还是“图文均可”。“需求明确时连续完成”只表示不必逐步确认，不等于不许提问；只有“不要问我”“全部你自己判断”这类针对整个任务的说法才算不问，“配色你自己判断”只把配色交给 agent。不让问又没说表达方式时，按图示优先并在策划中写明。
- **只供能看图的模型使用**：制作和验收都要实际查看渲染后的页面；模型不能看图时直接停止。OCR、自动检查或另一个模型的描述都不能代替看图。
- **把好模型会做的事写成步骤和脚本**：
  - `preview.py pptx --export wps`（或 `powerpoint`）：在 Windows 上让 WPS／PowerPoint 把最终 PPTX 导出为 PDF 再检查，看到的就是观众看到的换行和字体；
  - `--zoom`：每页另出四个放大的象限，整页看一遍后再细看一遍；
  - 版面检查新增警告：一行只剩一个字（标点不计，含强制换行前）、文字跑出卡片或底框、至少 3 页的内容区偏向同一侧；
  - 编辑测试按对象类型汇总实际测了什么，避免验收记录写成“测过图表和表格”。
- **门槛不变**：事实与口径准确（不补造数字、关系和可比性）、版面完成、对象符合编辑要求、模板保护；记忆点是目标，用“出彩检查”自查。保留 13 套配方、22 个原型，推荐优先、不硬套。

未指定格式默认 HTML＋PPTX＋`planning.md`；仅要求一种格式、大纲、样稿或局部修改，就只完成约定范围。精确页数包含封面与结束页。

## 工作方式

建立需求与证据 → 大纲与策划稿 → 定义视觉方向 → 制作交付格式 → 实际预览验收 → 交付。

方向不明确时先做方向比较，再用代表页校准，最后推广到全套。视觉部分包括字体与中文排版、风格配方、构图原型、图标、配图与插画手法。HTML 播放器保留缩放、键盘／触屏翻页、逐步揭示、概览、备注、全屏、黑屏、网址页码与打印。


## 安装

把本仓库的 `ssr-ppt-skill` 文件夹放进 agent 读取技能的目录。下表是各 agent 文档列出的用户级目录（2026 年 10 月查阅，以各 agent 最新文档为准）：

| Agent | 用户级技能目录 | 说明 |
|---|---|---|
| Claude Code | `~/.claude/skills` | 项目级为 `<项目>/.claude/skills` |
| Codex | `$CODEX_HOME/skills`，未设置时为 `~/.codex/skills` | |
| Cursor | `~/.cursor/skills` 或 `~/.agents/skills` | 项目级为 `.cursor/skills` 或 `.agents/skills`；文档还写明兼容读取 `.claude/skills`、`.codex/skills` |
| Grok Build | `~/.grok/skills` | 也会读取 `~/.claude/skills` 和 `~/.agents/skills` |
| Antigravity（agy CLI） | `~/.gemini/antigravity-cli/skills` | 桌面应用与 IDE 为 `~/.gemini/config/skills`；项目级为 `<项目>/.agents/skills` |
| DeepSeek Harness（dsh） | `~/.dsh/skills` | 来自社区插件 dsh-skills-manager 的说明，未见官方文档；可能需要该插件才能加载 |

**同一个 agent 会读取多个目录时，只装一份**，否则可能出现重名技能。例如已装到 `~/.claude/skills` 时，Grok Build 已能读到，不必再装到 `~/.grok/skills`。

用仓库自带的安装脚本（只依赖 Python 标准库），`--list` 会显示各目标路径与安装状态：

```bash
git clone https://github.com/EuronSSR-01/SSR-PPT-SKILL-2.git
cd SSR-PPT-SKILL-2
python install.py --list
python install.py --agent claude --agent agy           # 按需选择，可重复
python install.py --agent cursor --project /path/to/repo   # 装到某个项目的 .agents/skills
```

可选的 `--agent` 有 `claude`、`codex`、`cursor`、`grok`、`agy`、`agy-app`、`agents`（通用的 `~/.agents/skills`）和 `dsh`。目标已存在时脚本不会覆盖；加 `--force` 会先把旧版本移到技能目录旁边的 `skill-backups/` 中备份，再安装新版本。备份放在技能目录之外，避免 agent 把它当成第二个同名技能加载。

### 不支持技能的 agent

把下面这段放进项目的 `AGENTS.md`（或该 agent 读取的规则文件），并把路径改成实际位置：

```markdown
## 制作 PPT / 演示文稿
做 PPT、幻灯片、演示文稿或美化已有 PPT 时，先完整阅读
`<技能目录>/ssr-ppt-skill/SKILL.md` 并按其流程执行；需要时再读其中 references/ 下对应的文件，脚本在 scripts/ 下。
```

### 依赖

先运行环境自检，它只用 Python 标准库，会列出已有能力并给出建议路线：

```bash
python ssr-ppt-skill/scripts/doctor.py
```

按需安装，缺少时 Skill 会说明受影响的验收项：

```bash
python -m pip install fonttools brotli playwright pillow pymupdf python-pptx
```

- 浏览器：Playwright 会自动使用本机已安装的 Chrome 或 Edge；都没有时运行 `python -m playwright install chromium`。
- PPTX 渲染：Windows 上装了 WPS 或 PowerPoint 时用 `preview.py pptx --export wps`（或 `powerpoint`）；否则用 LibreOffice，或用 `--pdf` 传入由最终 PPTX 导出的 PDF。
- 新建 PPTX 默认用 python-pptx；已有 PptxGenJS 也可以用。

## 使用

调用方式因 agent 而异：Codex 写 `$ssr-ppt-skill`；Claude Code、Cursor、Grok Build 可以用 `/ssr-ppt-skill`；Antigravity 可以用 `/skills` 查看已加载的技能。各 agent 也会根据描述自动匹配，直接说“用 ssr-ppt-skill 做……”即可。

```text
用 ssr-ppt-skill，基于这些资料做 10 页管理层汇报，
交付离线 HTML、可编辑 PPTX 和策划稿。
需求已明确，后续自行判断，不再追问，完成后实际预览与编辑验证。
```

```text
用 ssr-ppt-skill，在这份公司 PPT 模板中制作 6 页方案。
仅交付 PPTX，保留母版、字体与品牌规范；遇到关键缺口可以逐个问我。
```

```text
用 ssr-ppt-skill，先整理大纲和逐页策划，暂不制作成品。
```

## 样例与验证

[双格式样例](ssr-ppt-skill/examples/dual-format/README.md) 共享一份虚构数据生成 4 页 HTML 和原生 PPTX，含长中文、分组流程、形状矩阵、统计图、原生表格和独立图片。它是可复现的有限路径，不是通用 HTML 转 PPTX 引擎。

本次在 Codex／Windows 中验证：
- 9 项回归测试通过，含损坏内联图片、减少动态效果、分组文字、表格原有勾号、图表未改值、未支持图型和多主题／母版保留。
- 4 页 HTML：无外部请求、无控制台或页面错误，键盘与网址页码通过，打印 4 页且尺寸正确。
- 4 页 PPTX：重新读取与 LibreOffice 渲染，11 个指定编辑目标保存重开后符合预期，包括组内文字、矩阵、图表工作表、图片裁剪和表格。
- 逐页查看 HTML、PPTX 和编辑副本；修正大数字间距、无依据的方案高亮及默认阴影。
- 补充三组新材料的双格式检查，8 个指定编辑目标通过；修正教学样例里无依据的周报发送方，并重新验证受影响页面。观察与限制见 [新材料检查记录](tests/design-review.md)。

随后在 Claude Code／Windows 中复核（Chrome 155、LibreOffice 26.8）：
- 回归测试增至 10 项，全部通过。新增一项：内嵌工作表里的空 `<v/>` 单元格不能导致检查崩溃。修复前，`preview.py pptx --edit-test` 遇到 PptxGenJS 生成的图表会直接报错退出，连报告都不写。
- PptxGenJS 单系列逐点着色实测可行：数据仍是一个真实系列，标签可放在柱外。它生成的图表在自动编辑测试中如实标为 `unsupported`，处理方式见 [pptx.md](ssr-ppt-skill/references/pptx.md)。

为多 agent 使用补充（同一环境）：
- **环境自检 `doctor.py`**：在完整环境与隔离环境（缺少部分包）下都给出了对应路线。
- **版面检查**：
  - HTML：在早先版本的两份示例上，抓到了此前靠看图发现的 4 个问题（标注压字、单位换行、标题末行孤字、时间线标签出界），还抓到 1 处看图时漏掉的文字裁切。
  - PPTX：在专门构造的缺陷文件上，5 类问题全部检出。
  - 几份正常示例上没有误报。
  - 抓不到文字压住图形、箭头互相挤压这类问题。
- **`pptx_kit.py`**：重建了示例中的 4 页：中西文字体齐全，图表为单一真实系列，7 个编辑目标全部通过，其中包括图表。
- **安装脚本 `install.py`**：在临时目录中测试了列表、安装、跳过、强制替换（备份放在技能目录之外）与项目级安装。
- **回归测试**：增至 13 项，全部通过。

为“图示优先”补充（同一环境）：
- 起因：DSH（DeepSeek v4.1 Flash）用本技能做的一份 12 页汇报字仍偏多，多数页是“一排数字卡片＋几张要点卡片”。材料为用户本地的真实资料，不在仓库中。
- `--visual-first` 在这份汇报上提示了 5 页（共 11 处），包括四段百字左右的整段说明和几张 60–85 字的要点卡片；另有两页每张卡片约 40–50 字，没有提示。仓库自带的双格式样例 HTML 与 PPTX 都没有提示；两份本地示例中，偏图示的一份只在 HTML 版提示 1 处（三条行动要点合计约 61 字），演示原型的一份提示 4 处，都在文字较多的原型页。
- 修复误报：LibreOffice 渲染后，表格一行被当成一个文本块，最后一格的单个数字被报成“末行孤字”；同一张表，旧版报 2 处，新版 0 处。
- 修复：没有 LibreOffice 且不加 `--edit-test` 时，`preview.py pptx` 因输出目录未创建而崩溃。
- 回归测试增至 19 项，全部通过。
- 尚未验证：同一材料、同一模型下新旧版本的对照，需要在 DSH 等 agent 中实际运行。

为看图前提与设计底线补充（同一环境）：
- 起因：DSH 用新版再做同类汇报，字数下降，出现了真正的结构图和流程图；过程记录写着逐页看过图且“通过”，成品仍有虚线压字、多页下方大片空白、计划 12pt 的正文实际约 9.75pt、短数量条被最小宽度放大约 2.5 倍、回路标签补出材料没有的责任，可以做成图表的分类数值画成了普通形状。
- 改动：只供能看图的模型使用，删除“不能看图时”的降级路线；基础版面完成度进入门槛；逐页看图改为具体问题，记录发现的问题和修改；保留文字按信息单元判断；补充图形闭环、数量比例、图表内嵌数据、现状与建议分开等规则；`pptx_kit.py` 写清字号的 px 与 pt 换算。
- 回归测试 19 项（含 4 项浏览器测试）全部通过；双格式样例重新生成后严格检查通过，版面 0 错误 0 警告，11 个编辑目标通过。
- 尚未验证：新规则对 DSH 等 agent 的实际效果，需要用同一材料再跑一次对照。

为线条压字和外部 PDF 补充（同一环境）：
- **线条穿字提示**：`preview.py pptx` 的版面检查新增警告级的“线条穿过文字”。逐字判断，线要穿过字形中部才算；后画的填充形状挡住的线、贴合整段文字的删除线、文字下方的下划线都不算。只作提示，需看图判断，有意的设计可以保留。
  - 在 DSH 第三次的成品上报出第 6 页压在虚线上的标签，这是此前只能靠看图发现的问题。
  - 在另外 9 份 PPTX（共 55 页，含模板、原稿、DSH 第二次的成品和本仓库的示例）上没有误报。
  - HTML 的版面检查仍抓不到文字压住图形。
- **`--pdf`**：`preview.py pptx --pdf <PDF>` 用 WPS、PowerPoint 等由最终 PPTX 导出的 PDF 生成逐页图并做版面检查，适合没有 LibreOffice 的环境。报告记录 PDF 与 PPTX 的哈希；PDF 比 PPTX 旧时给出警告；页数不符时渲染检查失败。对象清单、字体和编辑测试仍读 PPTX。
  - 已用 LibreOffice 导出的 PDF 验证；本机 WPS 的 COM 接口无法调用，WPS 导出的 PDF 尚未实测。
- 回归测试增至 21 项（含 4 项浏览器测试），全部通过。

回到第一版设计规则并加入图标与配图（同一环境）：
- 起因：用户比较 DSH 的几次结果后认为，第二次（技能第一版）做的模板版基本可接受，只是字多；之后几版越改越偏，第四次正文字数约为第三次的两倍，页面写满口径说明和重复内容，还出现了文字跑出卡片、孤字、全套内容区偏右等问题。用户想要的“图”是好看的图片、图标和示意插画，不是更多的信息图。
- 改动：设计规则回到第一版；“图示优先”改为用图片、图标、示意插画代替文字；新增图标库、`icon()`、`svg_picture()`、`--export`、`--zoom` 和三类版面警告；开场两问的措辞修正。
- 图标：127 个图标全部转成原生形状，在 LibreOffice 渲染中逐个看过，形状、线宽和颜色正确。
- 新的版面警告在 11 份 PPTX 上的结果：
  - 第四次模板版报出 4 处孤字（含“间；”“性。”和强制换行前单独一行的“在”）、2 处文字跑出卡片，以及全套内容区左 79pt、右 38pt 的偏移，都与看图一致；
  - 第二次和第三次分别报出 4 处、1 处此前没人发现的孤字，看图确认属实；
  - 原稿、模板、本仓库的双格式样例和另外四份演示文件都没有误报。
- 修复：LibreOffice 会忽略空的效果列表，照样画出形状引用的主题阴影；`pptx_kit` 的矩形、圆形和图标改为去掉主题样式引用，LibreOffice 渲染中不再出现阴影（WPS／PowerPoint 未实测）。
- 回归测试增至 29 项（含 5 项浏览器测试），全部通过；双格式样例重新生成后严格检查通过，版面 0 错误 0 警告，11 个编辑目标通过。
- `--export` 尚未在能正常调用 WPS 的电脑上实测：本机 WPS 的 COM 接口无法调用，脚本如实报告导出失败、渲染记为失败。

按复核收紧措辞（同一环境）：
- 开场两问：只针对某件事的“你自己判断”（如“配色你自己判断”）只交出那件事，不改变整个任务的询问方式。
- 少字：计算过程放备注；影响观众判断的口径、估算状态和前提，在相关页面用尽可能简短的说明保留（通常一句，必要时分别标注）。此前策划说明把“口径如何推算”整体移出页面，与来源一节的要求矛盾。
- 版心：模板或已确定的视觉方案有侧栏、图片带或不对称网格时，按对应网格执行，并在策划稿写明依据；公司模板仍以“左右相等，或分别对齐页眉的标题与标志”为具体参照。“内容区偏一侧”警告的说明改为与实现一致：至少 3 页偏向同一侧就报，不要求多数页。
- 字数提示：写明是近似阅读量（汉字计 1，英文单词和数字计 2），不是合格线。
- `doctor.py`：写明只检测 COM 注册、没有实际调用；能否导出以 `--export` 的结果为准。本机就是注册了但调不通。
- 反模式新增：原文一个字没减，只在旁边加了图标。
- 回归测试 29 项全部通过。

[维护测试说明](tests/README.md) 记录命令、版本与证据范围；[跨 agent 对比测试](tests/cross-agent/README.md) 提供固定任务与评分表。

尚未覆盖：Cursor、Grok Build、Antigravity、DeepSeek Harness 中的实际运行（本机没有安装）；PowerPoint／WPS 实机编辑；多模型效果对照；所有风格的验证。

## 脚本与边界

| 脚本 | 用途 |
|---|---|
| `install.py`（仓库根目录） | 把技能安装到各 agent 的技能目录 |
| `scripts/doctor.py` | 环境自检：已有能力与建议路线（只用标准库） |
| `scripts/pptx_kit.py` | python-pptx 基础操作：画布坐标、中西文字体、无阴影形状、图标（原生形状）、SVG 插画、单系列高亮图表、无样式表格、裁切图片 |
| `assets/icons/` | 127 个 Lucide 线性图标（ISC 许可），`INDEX.md` 列出名称与中文含义 |
| `scripts/embed_fonts.py` | HTML 字体子集化、内联及缺字报告；嵌入前核许可 |
| `scripts/preview.py html` | 离线浏览、图片解码、字体加载、键盘、逐页截图、打印与版面检查；`--zoom` 放大局部；每页正文字数，`--visual-first` 时提示长文字和字多的页 |
| `scripts/preview.py pptx` | `--export wps／powerpoint` 由目标软件导出 PDF，或 LibreOffice 渲染，或 `--pdf` 传入 PDF；版面检查（孤字、文字跑出卡片、线条穿字、内容区偏一侧等）、对象与所有主题字体清单；指定对象编辑与重开，按类型汇总；`--zoom` 放大局部；每页正文字数，`--visual-first` 时提示长文字（不需要渲染） |
| `scripts/preview.py sheet / compare` | 总览、灰度／模糊图与方向对照 |
| `scripts/set_theme_fonts.py` | 指定主题的字体修复；多主题要求明确选择范围 |

`--strict` 只要求脚本覆盖的必需自动检查通过，不能代替看图与目标应用测试。PPTX 编辑用 `--edit-test`，关键目标用 `--edit-targets` JSON 指定，详见 [对象合同与验证](ssr-ppt-skill/references/pptx.md)。

这是技能指令、播放器与辅助脚本，不包含 Office。图标是原生形状，可改颜色和尺寸；照片和 SVG 插画是独立图片，插画内部要改时编辑 SVG 源文件后重新生成；文字与数据图的编辑能力另测。复杂网页效果无需逐像素复刻，整页截图不算可编辑 PPTX。字体可用性、嵌入和分发依工具与授权处理。

复用已有工具链，默认当前智能体执行。Web Brain／Inner Web Brain 仅在用户为当前任务明确要求时使用。缺少必需能力时说明阶段成果和未完成项。

## 文件导航

- [SKILL.md](ssr-ppt-skill/SKILL.md)：六个步骤的入口，开场两问与表达方式。
- [策划](ssr-ppt-skill/references/planning.md)、[视觉定调](ssr-ppt-skill/references/art-direction.md)、[13 套配方](ssr-ppt-skill/references/style-recipes.md)、[22 个原型](ssr-ppt-skill/references/layouts.md)。
- [字体](ssr-ppt-skill/references/typography.md)、[图标与配图](ssr-ppt-skill/references/imagery.md)、[图表和表格](ssr-ppt-skill/references/charts.md)、[反模式](ssr-ppt-skill/references/anti-patterns.md)、[图标索引](ssr-ppt-skill/assets/icons/INDEX.md)。
- [HTML](ssr-ppt-skill/references/html.md)、[PPTX](ssr-ppt-skill/references/pptx.md)、[实际验收](ssr-ppt-skill/references/visual-review.md)。

播放器资源内联进成品，不引用技能安装目录；构建结果放在已忽略的 `.build/`。
