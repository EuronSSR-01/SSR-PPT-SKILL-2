# SSR-PPT-SKILL

演示制作技能，采用通用的 Agent Skills（`SKILL.md`）格式，可用于 Codex、Claude Code、Cursor、Grok Build、Antigravity 等 agent：先组织内容、提炼信息、选择表达，再做视觉设计；保留策划稿与实际预览验收，支持公司模板、离线 HTML 和原生可编辑 PPTX。

名称保持 **SSR-PPT-SKILL**，技能名为 `ssr-ppt-skill`。调用方式见下文“使用”。

## 这次升级

- **内容先行**：观点与证据、文字精简、数据口径、机制关系、表格重表达，避免为了好看遗漏条件或混淆事实。
- **按用户偏好沟通**：继承已有偏好；未知时先问一次“允许逐个询问，还是自行判断连续完成”，再按所选模式进行。允许询问则逐个问必要问题，自主模式连续判断，缺事实不编造。策划、方向比较不自动成为审批关卡。
- **门槛与目标分开**：事实、可读性和编辑合同是必须通过的门槛；记忆点是目标，有设计自由时用“出彩检查”（去品牌、截图、模糊、灰度测试与精修）自查，公司模板任务检查模板内的精致度。字体比例、色彩面积、配方、高潮页与版式数量是可调整方法。保留 13 套配方、22 个原型，不要求硬套。
- **图片按作用选**：真实证据用真实来源，概念氛围可用专用生图模型或搜索，流程与架构用有语义的 SVG／原生图解。
- **表格按操作选**：可精简、聚合、重组为图表或矩阵；需要增删行列用原生表格，只需改字、移动的矩阵可用原生色块和文字。统计图默认保留底层数据编辑。
- **制作路线分流**：新建双格式先验证 HTML／SVG 代表页与原生 PPTX 代表页；公司模板和局部修改原生优先，保护主题、母版、占位符、备注和链接。
- **检查真正的变化**：递归定位分组内对象，核对修改前／预期／实际值；明确失败、未执行、工具不支持与不适用，避免把“对象还存在”当作成功。

未指定格式默认 HTML＋PPTX＋`planning.md`；仅要求一种格式、大纲、样稿或局部修改，就只完成约定范围。精确页数包含封面与结束页。

## 工作方式

明确任务 → 形成策划 → 设计与适配 → 实际预览验收 → 完成交付。

内容问题回到策划，布局或实现问题回到制作。方向不明确时才做方向比较，代表页挑最能暴露风险的内容；短任务不机械运行多套流程。

视觉部分保留字体与中文排版、风格配方、构图原型和图片手法。HTML 播放器保留缩放、键盘／触屏翻页、逐步揭示、概览、备注、全屏、黑屏、网址页码与打印。


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
- PPTX 渲染需要 LibreOffice。
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

[教学案例](ssr-ppt-skill/references/worked-examples.md) 展示科研资源、零售补货、产品比较和培训须知的材料提炼、表达选择、页面与检查点，包含两张可直接查看的 SVG 页面，以及“保留文字更清楚”的完整案例。

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

[维护测试说明](tests/README.md) 记录命令、版本与证据范围；[跨 agent 对比测试](tests/cross-agent/README.md) 提供固定任务与评分表。

尚未覆盖：Cursor、Grok Build、Antigravity、DeepSeek Harness 中的实际运行（本机没有安装）；PowerPoint／WPS 实机编辑；多模型效果对照；所有风格的验证。

## 脚本与边界

| 脚本 | 用途 |
|---|---|
| `install.py`（仓库根目录） | 把技能安装到各 agent 的技能目录 |
| `scripts/doctor.py` | 环境自检：已有能力与建议路线（只用标准库） |
| `scripts/pptx_kit.py` | python-pptx 基础操作：画布坐标、中西文字体、无阴影形状、单系列高亮图表、无样式表格、裁切图片 |
| `scripts/embed_fonts.py` | HTML 字体子集化、内联及缺字报告；嵌入前核许可 |
| `scripts/preview.py html` | 离线浏览、图片解码、字体加载、键盘、逐页截图、打印与版面检查 |
| `scripts/preview.py pptx` | LibreOffice 渲染、版面检查、对象与所有主题字体清单；指定对象编辑与重开 |
| `scripts/preview.py sheet / compare` | 总览、灰度／模糊图与方向对照 |
| `scripts/set_theme_fonts.py` | 指定主题的字体修复；多主题要求明确选择范围 |

`--strict` 只要求脚本覆盖的必需自动检查通过，不能代替看图与目标应用测试。PPTX 编辑用 `--edit-test`，关键目标用 `--edit-targets` JSON 指定，详见 [对象合同与验证](ssr-ppt-skill/references/pptx.md)。

这是技能指令、播放器与辅助脚本，不包含 Office。照片、图标和插画保留独立图片；文字与数据图的编辑能力另测。复杂网页效果无需逐像素复刻，整页截图不算可编辑 PPTX。字体可用性、嵌入和分发依工具与授权处理。

复用已有工具链，默认当前智能体执行。Web Brain／Inner Web Brain 仅在用户为当前任务明确要求时使用。缺少必需能力时说明阶段成果和未完成项。

## 文件导航

- [SKILL.md](ssr-ppt-skill/SKILL.md)：五阶段入口与按需读取路由。
- [策划](ssr-ppt-skill/references/planning.md)、[内容表达](ssr-ppt-skill/references/content-expression.md)、[教学案例](ssr-ppt-skill/references/worked-examples.md)。
- [设计参考](ssr-ppt-skill/references/design-reference.md)、[视觉定调](ssr-ppt-skill/references/art-direction.md)、[13 套配方](ssr-ppt-skill/references/style-recipes.md)、[22 个原型](ssr-ppt-skill/references/layouts.md)。
- [字体](ssr-ppt-skill/references/typography.md)、[图片](ssr-ppt-skill/references/imagery.md)、[图表和表格](ssr-ppt-skill/references/charts.md)、[视觉诊断](ssr-ppt-skill/references/anti-patterns.md)。
- [HTML](ssr-ppt-skill/references/html.md)、[PPTX](ssr-ppt-skill/references/pptx.md)、[实际验收](ssr-ppt-skill/references/visual-review.md)。

播放器资源内联进成品，不引用技能安装目录；构建结果放在已忽略的 `.build/`。
