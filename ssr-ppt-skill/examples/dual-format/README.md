# 双格式最小样例

一份虚构资料，四页明确布局，同时生成 HTML／SVG 和原生 PPTX。用于观察从内容到对象的映射，不是任意 HTML、SVG 或 CSS 的通用转换器；也不代表真实项目事实。

## 运行

在仓库根目录，使用已有 Python 环境（python-pptx、Pillow；预览另需 Playwright、PyMuPDF、浏览器与 LibreOffice）：

```bash
python ssr-ppt-skill/examples/dual-format/generate.py --out .build/dual-format
python ssr-ppt-skill/scripts/preview.py html .build/dual-format/index.html --out .build/review/html --browser chrome --strict
python ssr-ppt-skill/scripts/preview.py pptx .build/dual-format/deck.pptx --out .build/review/pptx --edit-test --edit-targets .build/dual-format/edit-targets.json --strict
```

浏览器可选 chromium／msedge。环境必须允许启动本机浏览器。临时目录不可写时，将当前进程 TEMP／TMP（POSIX 用 TMPDIR）指向可写的构建临时目录，再运行，不更改系统配置。

输入 `data.json`，输出 `index.html`、`deck.pptx`、`planning.md`、`edit-targets.json` 和一张抽象插画。插画由代码生成，仅演示独立图片与裁剪，不代表专用生图模型的照片质量。所有内容可离线运行，不访问素材网站。

## 四页覆盖

| 页 | 表达 | 原生对象与目标 |
|---|---|---|
| 1 | 长中文结论、大数字、状态与概念插画 | 标题编辑、独立图片裁剪 |
| 2 | 门店提交—规则判断—仓库备货，退回补充及周报 | 分组文字、带绑定的主流程连接线、整组移动 |
| 3 | 三方案比较矩阵，保留成本、周期和条件 | 形状与文字分组、具体价格格编辑、移动 |
| 4 | 上线趋势、预测收益、待核事项 | 原生柱图与工作表、原生表格格子 |

`generate.py` 为这些页面写了明确布局，只共享简单文字／形状辅助函数和业务数据，不建设通用场景树。HTML 与 PPTX 的图表分别实现，字体度量、基线和表格线可略有差异，数据、条件与关系保持一致。

1280×720 px 映射到 13⅓×7.5 英寸：px ÷ 96 为英寸，px × 0.75 为 pt；其他画布必须重新计算。样例使用 Microsoft YaHei，运行和观看机器需有该字体；没有则修改 FONT 并重新渲染。系统字体未随样例分发。新建样例明确设置包括备注主题在内的全部字体；公司模板不能照搬全局修改。

## 验证范围

已用 Python 3.14.3、python-pptx 1.0.2、Pillow 12.1.0、Playwright 1.58.0、Chrome 156.0.8078.12、LibreOffice 26.2.3.2、PyMuPDF 1.27.2.2 运行。版本为测试记录，不是唯一支持版本。

严格自动检查通过，4 页双格式和编辑副本实际看图；11 个编辑目标符合预期。修复了大数字与说明间距、矩阵无依据的单列高亮、PPTX 默认阴影和重复图表标题。字体、组内文字、价格、图表缓存与工作表均按当前文件复核。

编辑测试为库层保存／重开＋LibreOffice 渲染。自动连线跟随、图片替换、模板中的特殊媒体和 PowerPoint／WPS 实机编辑未在该样例验证。LibreOffice 的临时 PDF 在本机出现 MuPDF 结构树诊断，PNG 正常生成并查看；该路线不作为 PDF 结构或无障碍验收。HTML 打印另行通过页数与尺寸检查。

不要把副本的 `[E…]` 编辑标记或测试数值交付为真实内容；用户成品取原始导出文件。报告的 `manual_review` 默认为未执行，需实际看图后另记结果，不由生成器预填“通过”。
