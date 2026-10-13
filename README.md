# SSR-PPT-SKILL

面向 **Codex** 与 **Claude Code** 的演示制作技能：先组织内容、提炼信息、选择表达，再做视觉设计；保留策划稿与实际预览验收，支持公司模板、离线 HTML 和原生可编辑 PPTX。

名称保持 **SSR-PPT-SKILL**；调用名为 `ssr-ppt-skill`，Codex 使用 `$ssr-ppt-skill`，Claude Code 使用 `/ssr-ppt-skill`。

## 这次升级

- **内容先行**：观点与证据、文字精简、数据口径、机制关系、表格重表达，避免为了好看遗漏条件或混淆事实。
- **按用户偏好沟通**：继承已有偏好；未知时问一次。允许询问则逐个问必要问题，自主模式连续判断，缺事实不编造。策划、方向比较不自动成为审批关卡。
- **提高下限，保留自由**：事实、可读性和编辑合同是底线；字体比例、色彩面积、配方、高潮页与版式数量是可调整方法。保留 13 套配方、22 个原型，不要求硬套。
- **图片按作用选**：真实证据用真实来源，概念氛围可用专用生图模型或搜索，流程与架构用有语义的 SVG／原生图解。
- **表格按操作选**：可精简、聚合、重组为图表或矩阵；需要增删行列用原生表格，只需改字、移动的矩阵可用原生色块和文字。统计图默认保留底层数据编辑。
- **制作路线分流**：新建双格式先验证 HTML／SVG 代表页与原生 PPTX 代表页；公司模板和局部修改原生优先，保护主题、母版、占位符、备注和链接。
- **检查真正的变化**：递归定位分组内对象，核对修改前／预期／实际值；明确失败、未执行、工具不支持与不适用，避免把“对象还存在”当作成功。

未指定格式默认 HTML＋PPTX＋`planning.md`；仅要求一种格式、大纲、样稿或局部修改，就只完成约定范围。精确页数包含封面与结束页。

## 工作方式

明确任务 → 形成策划 → 设计与适配 → 实际预览验收 → 完成交付。

内容问题回到策划，布局或实现问题回到制作。方向探索由不确定性触发，代表页验证由风险触发；短任务不机械运行多套流程。

视觉部分保留字体与中文排版、风格配方、构图原型和图片手法。HTML 播放器保留缩放、键盘／触屏翻页、逐步揭示、概览、备注、全屏、黑屏、网址页码与打印。


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
- 新建 PPTX 可用已有的 PptxGenJS 或其他原生工具；样例使用 python-pptx，不强制新增项目依赖。

## 使用

```text
使用 $ssr-ppt-skill，基于这些资料做 10 页管理层汇报，
交付离线 HTML、可编辑 PPTX 和策划稿。
需求已明确，后续自行判断，不再追问，完成后实际预览与编辑验证。
```

```text
使用 $ssr-ppt-skill，在这份公司 PPT 模板中制作 6 页方案。
仅交付 PPTX，保留母版、字体与品牌规范；遇到关键缺口可以逐个问我。
```

```text
使用 $ssr-ppt-skill，先整理大纲和逐页策划，暂不制作成品。
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

[维护测试说明](tests/README.md) 记录命令、版本与证据范围。历史版本的其他样稿和模型评价不作为本次验收证据。当前没有 PowerPoint／WPS 实机编辑、多模型效果对照或所有风格的验证；不承诺模型间能力差距被消除。

## 脚本与边界

| 脚本 | 用途 |
|---|---|
| `scripts/embed_fonts.py` | HTML 字体子集化、内联及缺字报告；嵌入前核许可 |
| `scripts/preview.py html` | 离线浏览、图片解码、字体加载、键盘、逐页截图与打印检查 |
| `scripts/preview.py pptx` | LibreOffice 渲染、对象与所有主题字体清单；指定对象编辑与重开 |
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
