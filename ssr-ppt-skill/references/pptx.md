# 原生可编辑 PPTX

## 实现选择

### Codex 桌面

调用 `mcp__codex_app__load_workspace_dependencies` 获取 bundled Node.js、Python 与模块目录；不要硬编码用户路径或 runtime 版本。

如果环境提供 `presentations:Presentations` Skill：

- 读取其 `references/implementation.md` 与 `artifact_tool_docs/API_QUICK_START.md`；制作原生表格与图表时再读取 `references/native_evidence.md`。
- 使用 JavaScript ES modules 与 `@oai/artifact-tool`，遵循它当前的本地 authoring 和 finalization 要求；不复制过时 API，也不修改 bundled runtime。
- 新建或美化版式时，同时读取它的 `style_guidelines.md`；新建或重做封面时再读取 `references/cover_art_direction.md`。
- 两套视觉指南冲突时，以用户要求和本 Skill 策划稿中的视觉定调为准。

### Claude Code

- **新建演示**：用 PptxGenJS（Node.js）。先检查是否已安装；没有时在私有构建目录执行 `npm init -y` 与 `npm install pptxgenjs`，不改用户项目的依赖。它支持原生文本、形状、表格、图表、图片、备注和母版。
  - 设 `pptx.layout = 'LAYOUT_WIDE'` 得到 13.333×7.5 英寸画布；默认的 `LAYOUT_16x9` 只有 10×5.625 英寸。
  - 导出后运行 `python scripts/set_theme_fonts.py deck.pptx --heading <标题字体> --body <正文字体>`，补上主题的东亚字体（原因见 [typography.md](typography.md)）。
- **修改已有 PPTX 或基于用户模板制作**：用 python-pptx，它可以打开现有文件、使用模板版式与占位符，并支持逐点设置图表颜色。东亚字体与部分样式需要直接写 OOXML（lxml）。
- 会话中如果有 Anthropic 的 `pptx` Skill，可以参考其中的 OOXML 编辑与渲染方法，但视觉规则以本 Skill 为准。
- **渲染**：`python scripts/preview.py pptx deck.pptx --out <目录>` 会调用 LibreOffice 转成 PDF，再用 PyMuPDF 逐页渲染为 PNG，统计每页的原生对象，并列出幻灯片与主题实际使用的西文 / 东亚字体。加 `--edit-test` 还会在副本上做编辑测试。脚本会在 PATH 与常见安装路径中查找 `soffice`。
- **LibreOffice 与 PowerPoint 的已知差异**，看图时不要误判：
  - LibreOffice 会自动在汉字与西文、数字之间加间距，叠加手动空格后显得偏宽；PowerPoint 不会这样显示。
  - LibreOffice 在西文字体含有汉字时直接用它显示中文，因此主题东亚字体缺失的问题在 LibreOffice 渲染中看不出来，要以 `report.json` 中的字体清单为准。

### 其他环境

使用项目已安装的原生 PPTX 工具，保留现有工具链。不要假定浏览器打印、HTML 转图片或插入 SVG 就能完成原生可编辑交付。缺少必要的 authoring 或渲染能力时，报告具体缺项和已完成的部分，不降低输出要求。

### 单位

画布尺寸以模板或策划稿为准。各工具的单位不同：

- Artifact Tool 的几何通常使用 96 DPI 像素。
- PptxGenJS 使用英寸；`LAYOUT_WIDE` 为 13.333×7.5 英寸，像素换英寸除以 96。
- python-pptx 使用 EMU（914400 EMU = 1 英寸）。

HTML 的 px 换算为 pt 时乘以 0.75（1280px 宽对应 960pt）。字号单位不能直接混用。

## 对象合同

| 内容 | PPTX 中的形式 |
|---|---|
| 标题、正文、数据标注 | 原生文字及段落 |
| 表格 | 原生表格，保留必要的行列与单元格文字 |
| 数据图表 | 原生图表，保留类别、系列、数值、单位及正确的数据引用 |
| 要求编辑的流程 / 结构图 | 原生形状、连接线与文字 |
| 照片、标志、插画、图标 | 图片，可移动裁剪；不宣称图片内容可以逐元素编辑 |
| 用户明确要求保留的原图 | 保留原图，并说明其编辑范围 |

整页截图或整页 SVG 不能满足文字、表格和图表的编辑合同。SVG 可以作为局部图标或艺术素材；插入或转换为形状后，其中的文字和图表不一定能恢复为原生对象。

从 HTML 制作 PPTX 时，优先复用策划稿和数据重建原生对象，不承诺任意 CSS、WebGL、滤镜和动态状态都能转换。某个视觉效果难以保留时，保留核心信息和所需的编辑能力，并记录具体差异。用户要求严格复刻而现有工具做不到时，说明限制。

## 原生页面的美化要点

可编辑不是放弃设计的理由。把策划稿中的视觉定调落实到原生页面：

- **字体**：每段文字同时设置西文与东亚字体（latin / ea）；用 `scripts/set_theme_fonts.py` 设置主题的标题与正文字体（含东亚字体）；按 [typography.md](typography.md) 的 PPTX 字体策略处理字体可用性。字号用 pt，按 px × 0.75 换算。
- **母版与版式**：把背景、页边距参考线、页码、来源区等固定元素放进母版（PptxGenJS 用 `defineSlideMaster`，python-pptx 用模板版式），用户新增页面时能保持一致。
- **文本框**：
  - 内边距设为 0 或固定小值，让文字真正对齐网格。
  - 关闭自动缩放，避免字号被悄悄缩小；按设计固定行距（中文正文 1.3–1.5 倍）。
  - 标题文本框宽度按设计断行预留，用实际渲染核对换行。
- **形状**：明确设置填充与线条，不使用默认主题样式、阴影和渐变；需要遮罩时用带透明度的纯色或渐变矩形，并作为独立形状。
- **图片**：按版位比例裁切（PptxGenJS 用 `sizing: { type: 'cover' }`，或预先用 Pillow 裁好），不拉伸；需要统一调色、颗粒或双色调时预先处理图片（见 [imagery.md](imagery.md)）。
- **图表与表格**：按 [charts.md](charts.md) 去除默认样式，设置字体、颜色、标签和间距。
- **备注**：讲稿、来源和口径放入演讲者备注。

先在代表页验证两种格式中的大标题断行、主视觉裁剪和数据图重点，按 PPTX 实际的文本度量微调位置与间距，保持阅读顺序和构图比例。照片或局部艺术素材可以与原生文字、图表组合；不为复刻网页效果而把必需对象栅格化。

## 最终文件验收

- **官方 presentations 环境**：读取其 `references/finalization.md`，用现有 finalizer 检查任务要求、对象、字体与文件结构，导出到新文件。按实际情况声明总页数和必须含原生表格或图表的页码，不虚构模板覆盖要求。
- **其他环境**：用 `scripts/preview.py pptx` 或现有的检查与渲染工具完成同样范围的验证。也可以用 Python 标准库的 `zipfile` 与 `xml.etree.ElementTree` 检查 OOXML 中的 `p:sp/a:t`、`a:tbl`、`c:chart`、`a:ea` 和数据引用；这些只是结构证据，不能代替预览和编辑测试。

实际验收必须包含：

1. 重新读取导出的最终 PPTX，检查总页数、关键文本和必需的原生对象。
2. 渲染最终导出文件的每一页，按 [visual-review.md](visual-review.md) 查看全套节奏和完整尺寸的页面。不能只看导出前的内存对象或 HTML 预览。检查视觉焦点、中文排版、字体替换、溢出、遮挡、图表数据、单位、标签、素材裁剪和模板匹配。
3. 在测试副本中修改至少一处代表性正文；有原生表格则改一个单元格，有数据图表则改一个值。保存、重新载入并渲染，核对实际变化，并检查图表的类别和数值顺序与源数据一致。
4. 如果 PowerPoint、LibreOffice 等目标应用可用且无需新增授权，优先在目标应用中打开测试副本，完成编辑、保存和重开。只用库完成的测试，报告为“原生对象及重新导入验证”，不能称为 PowerPoint 实机验收；用 LibreOffice 渲染的，如实写明渲染器。
5. 修复后重新导出并查看相关页面；两种格式共用的数据有变化时，两者都要重新检查。

默认不把构建源码、测试副本或验收截图当成最终交付。保留用户原文件；最终 PPTX 和策划稿使用真实的绝对路径。
