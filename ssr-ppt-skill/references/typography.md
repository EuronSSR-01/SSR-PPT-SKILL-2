# 字体与中文排版

字体是演示气质的重要来源，也直接决定可读性。公司指定字体或系统字体同样能做好设计，前提是用字号、字重和间距建立清楚的层级。没有模板约束却全套只用系统默认字体的单一字重，是“看起来普通”最常见的原因之一。本文件给出可选字体、搭配、字号阶梯、中文排版细则，以及两种格式的字体落地方法。

## 可选字体

以下字体常见于免费商用场景。使用前以官方页面的最新许可为准，并在 `sources.md` 或构建记录中写明字体来源。获取顺序：用户提供的字体文件 → 本机已安装（Windows `C:\Windows\Fonts`、`%LOCALAPPDATA%\Microsoft\Windows\Fonts`，macOS `/Library/Fonts`、`~/Library/Fonts`，Linux `/usr/share/fonts`、`~/.local/share/fonts`）→ 环境允许联网时从官方来源下载。

| 类别 | 字体 | 许可与来源 | 性格 |
|---|---|---|---|
| 中文黑体 | 思源黑体 Source Han Sans / Noto Sans SC | SIL OFL；Adobe、Google Fonts | 中性、字重全，通用首选 |
| 中文黑体 | HarmonyOS Sans SC | 华为官方许可 | 现代、几何、科技 |
| 中文黑体 | MiSans | 小米官方许可 | 现代、清爽 |
| 中文黑体 | 阿里巴巴普惠体 | 阿里官方许可 | 圆润、亲和 |
| 中文黑体 | OPPO Sans | OPPO 官方许可 | 干净、略圆 |
| 中文宋体 | 思源宋体 Source Han Serif / Noto Serif SC | SIL OFL | 编辑感、文化感、权威 |
| 中文特色 | 得意黑 Smiley Sans | SIL OFL | 窄斜、有态度，只用于大标题 |
| 中文特色 | 霞鹜文楷 LXGW WenKai | SIL OFL | 楷书感、人文，用于引文与点缀 |
| 西文无衬线 | Inter、Manrope、Space Grotesk、Barlow | SIL OFL；Google Fonts / GitHub | 数字清晰，适合数据 |
| 西文展示 | Bebas Neue、Oswald、Barlow Condensed | SIL OFL | 窄体大数字、海报感 |
| 西文衬线 | Playfair Display、Bodoni Moda、Cormorant Garamond、Source Serif | SIL OFL | 优雅、杂志、奢华 |
| 等宽 | JetBrains Mono、IBM Plex Mono | SIL OFL | 技术标注、编号、代码 |

**系统字体**：Windows 有微软雅黑、等线、宋体；macOS 有苹方 PingFang SC、宋体-简 Songti SC。可按授权用于已确认的目标环境，不打包嵌入 HTML 或随文件分发。商业字体使用已确认的授权；授权不明确时选可用且许可清楚的方案，沿用当前沟通模式。

## 搭配原则

- 通常从少量字体家族开始，用字号、字重、间距和位置建立层级；数量和字重差不是硬性门槛，品牌与阅读效果优先。
- 宋体标题配黑体正文，能立刻获得编辑感；黑体标题配宋体正文适合长阅读。
- 数字和西文单独指定字体：CSS 中把西文字体放在中文字体之前，例如 `font-family: "Inter", "Noto Sans SC", sans-serif`，数字与字母会使用 Inter，汉字回落到中文字体。排好后检查两种字体的基线与字高是否协调。
- 粗体必须使用真实的字重文件或可变字体，不能依赖浏览器或软件的合成粗体。中文不使用斜体表示强调。

| 搭配 | 标题 | 正文 | 数字 / 西文 | 适合 |
|---|---|---|---|---|
| 现代无衬线 | 思源黑体 Heavy | 思源黑体 Regular | Inter | 通用、商务、科技 |
| 编辑对比 | 思源宋体 Bold | 思源黑体 Regular | Inter / Source Serif | 数据新闻、研究、文化 |
| 东方雅致 | 思源宋体 SemiBold | 思源宋体 Regular | Cormorant Garamond | 品牌故事、文化 |
| 年轻有力 | 得意黑 | 思源黑体 Medium | Bebas Neue / Oswald | 创意、活动 |
| 人文温度 | 阿里巴巴普惠体 Bold | 阿里巴巴普惠体 Regular | 同家族 | 教育、公益、健康 |
| 产品科技 | HarmonyOS Sans / MiSans Bold | 同家族 Regular | JetBrains Mono | 产品发布、技术 |
| 手写点缀 | 霞鹜文楷（仅引文） | 思源黑体 Regular | Inter | 故事、引言页 |

## 字号阶梯

以下是 1280×720 对应 13⅓×7.5 英寸的起点，此时 pt = px × 0.75。其他尺寸按 [pptx.md](pptx.md) 的实际画布比例换算，不能只看长宽比。字号以显示场景和实际预览为准。

| 层级 | HTML（px） | PPTX（pt） | 说明 |
|---|---|---|---|
| 展示字（封面、大数字、金句） | 120–240 | 90–180 | 高冲击页使用；行高 0.95–1.1 |
| 页标题 | 40–64 | 30–48 | 结论句不超过两行 |
| 小标题 / 引导语 | 24–32 | 18–24 | |
| 正文 | 20–26 | 15–20 | 按投影距离、屏幕与阅读模式放大，避免堆字 |
| 标注 / 来源 / 页码 | 13–16 | 10–12 | 辅助文字也需实际读清，关键条件不能藏在小字中 |

尺度反差可制造重点，但不用固定字号比约束每一页。层级仍不清时先重组内容、拆分或转移次要细节，保留事实与关键条件，再调整排版。

## 中文排版细则

- **行高**：正文 1.5–1.7；标题 1.15–1.3；展示字 0.95–1.1。
- **字距**：正文 0 至 0.02em；中文大标题 0 至 -0.02em，不要加大字距去装“高级感”；小号全大写西文标签可加 0.08–0.16em。
- **中西混排**：汉字与西文或数字之间留一个空格（如“增长 23%”“使用 Python 处理”）；全角标点前后不加空格。
- **断行**：
  - 标题在语义边界手动断行，避免末行只剩一两个字。
  - 不拆开数字与单位、人名、专有名词。
  - 行首不出现句读标点，行尾不留开引号或开括号。
  - HTML 标题可加 `text-wrap: balance`，正文可用 `line-break: strict`；这些属性只能辅助，最终仍要看实际渲染结果。
- **行长**：幻灯片正文每行约 18–32 个汉字；更长时分栏或缩短。
- **对齐**：中文正文默认左对齐。两端对齐只用于较宽的纯中文段落，夹杂长西文单词时不用。
- **数字**：表格与对比数字用等宽数字（CSS `font-variant-numeric: tabular-nums`）。大数字配小单位：单位约为数字字号的 0.35–0.5 倍，与数字基线对齐。
- **强调**：优先用字重、颜色或尺度突出必要重点，避免过多强调相互竞争。

## HTML：子集化并内联字体

单文件离线 HTML 想用设计字体，就要把字体按实际用到的字子集化，再以 base64 内联。中文字体原文件通常 10–25MB，子集后一般只有几十到几百 KB。

使用本 Skill 的 `scripts/embed_fonts.py`（依赖 Python `fonttools`，输出 WOFF2 还需 `brotli`；缺少 brotli 时自动改用 WOFF）：

```bash
python scripts/embed_fonts.py build/index.html \
  --font "Noto Sans SC=fonts/NotoSansSC-VF.ttf" \
  --font "Noto Serif SC=fonts/NotoSerifSC-VF.ttf@700" \
  --font "Inter=fonts/Inter-Variable.ttf" \
  -o build/index.html
```

- 参数格式为 `族名=文件路径[#TTC 序号][@字重或字重范围][:italic]`。可变字体不写字重时保留全部字重范围；写单个字重（如 `@700`）时先固定为该字重的静态字体，体积更小。
- 脚本收集 HTML 中出现的全部字符（包括播放器脚本里的提示文字）以及基础 ASCII，生成 `<style id="ssr-embedded-fonts">` 插入 `</head>` 前。重复运行会替换旧的字体块，所以每次改完文字都要重新运行。
- 脚本会列出字体缺失的字符。出现缺字时换字体，或给对应文字指定能覆盖的回退字体。
- 页面 CSS 里照常写 `font-family`，并保留系统回退字体：`"Noto Sans SC", "PingFang SC", "Microsoft YaHei", sans-serif`。
- 每个嵌入字体都要确认许可允许网页嵌入与分发。系统字体和未授权的商业字体不得嵌入。

## PPTX：字体设置与可用性

PPTX 能否嵌入字体取决于工具和许可；这里的通用生成路线不自动嵌入字体。观看或编辑的电脑没有安装该字体时，会替换成默认字体，排版随之走样。制作前在策划稿里选定以下一种策略，并在交付说明中写明：

1. **稳妥**：正文用目标电脑大概率已有的字体（Windows 微软雅黑或等线，macOS 苹方），标题与数字用设计字体并随交付附上字体文件和安装说明。
2. **设计优先**：全套使用设计字体，交付字体文件（确认许可允许分发），说明需先安装。
3. **用户环境已知**：用户确认演示电脑已安装的字体。

技术要点：

- 中文需通过运行属性、占位符或主题获得正确的东亚字体。OOXML 的文字运行属性里，`<a:latin>` 管西文，`<a:ea>` 管中日韩文字；仅有 latin 不证明中文会使用预期字体，需结合主题与目标环境核对。
  - python-pptx 的 `font.name` 只写 `a:latin`，需要另外写入 `a:ea`。
  - PptxGenJS（4.0）文本的 `fontFace` 会同时写 latin、ea、cs；但 `pptx.theme = { headFontFace, bodyFontFace }` 只写主题的 latin。图表的坐标轴与数据标签字体也只写 latin，所以图表里的中文类别名依赖主题的东亚字体。
- 新建文件且确有主题字体缺项时，可用 `scripts/set_theme_fonts.py` 设置 latin／ea／Hans。公司模板默认保留主题、占位符和运行级字体；只有任务要求修复时才改。多主题文件必须用 `--theme ppt/theme/theme1.xml`（可重复）指定范围；明确要全改才用 `--all-themes`，用 `-o` 保留原件。
- 用 `scripts/preview.py pptx` 输出的字体清单核对幻灯片和所有主题；多母版需另核对每页实际关联的主题／版式。清单不证明字形覆盖和目标应用未回退。LibreOffice 渲染不一定能暴露东亚字体缺失的问题，不能只靠看图判断。
- 图表的坐标轴、数据标签、图例字体要单独设置，不会自动继承页面字体。
- 渲染验收所用的电脑必须装有目标字体，否则看到的是替换后的效果，不能据此判断设计。渲染图与预期不符时，先查字体是否被替换。
