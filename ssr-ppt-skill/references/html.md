# HTML 演示制作与验收

## 交付约定

默认生成单个 `index.html`，CSS、播放器脚本、可分发的设计字体和必要媒体内联，直接打开即可演示；使用系统字体时明确目标环境与回退策略。不依赖 CDN、远端字体或构建服务器。素材很大时可按用户要求交付完整目录，确保相对路径正确、资源随包提供。

默认 16:9，画布 1280×720；用户模板的比例优先。用固定画布整体缩放，窄视口下页面完整可见，而不是把正文重新排成网页长文。按 [art-direction.md](art-direction.md) 把策划中的视觉焦点、字体层级、素材处理与构图做成页面。两种格式共享内容与视觉意图；网页交互和动画无需搬进 PPTX，但静态核心内容必须完整。

## 页面结构与设计变量

把视觉定调中的色彩、字体与间距写成 CSS 变量，页面只引用变量：

```css
:root {
  --bg: #FBF9F4; --surface: #F3EFE6; --ink: #1A1A1A; --ink-2: #666158; --line: #E2DED5;
  --accent: #D7263D; --data-1: #0F4C81; --data-2: #9DB4C0; --data-3: #B5A88F;
  --font-title: "Noto Serif SC", "Songti SC", "SimSun", serif;
  --font-body: "Inter", "Noto Sans SC", "PingFang SC", "Microsoft YaHei", sans-serif;
  --margin: 80px; --gap: 24px;
}
.ssr-slide { background: var(--bg); color: var(--ink); font-family: var(--font-body); }
```

- 每一页都要有不透明的背景，因为翻页时新页会叠在旧页上淡入。
- 用绝对定位或 CSS Grid 在 1280×720 画布内精确排版，尺寸用 px。播放器按比例整体缩放画布，不要用 vw 或 vh。
- 使用可分发的设计字体时，用 `scripts/embed_fonts.py` 子集化并内联；新增文字后重建子集并检查缺字。若按目标环境使用系统字体，明确所需字体与回退策略，不打包未经许可的文件。

## 播放器

`assets/player.css` 和 `assets/player.js` 负责缩放、翻页、逐步揭示、入场动画、概览、备注、全屏和打印。把文件内容分别内联到 `<style>` 与页面末尾的 `<script>` 中，不要在成品中引用 Skill 的安装目录。播放器不访问任何外部服务，也不提供页面视觉模板。

页面只需提供以下结构，控制条、进度条和备注面板由脚本生成：

```html
<div class="ssr-player">
  <main class="ssr-stage" aria-label="演示画布">
    <div id="ssr-deck" data-width="1280" data-height="720" data-transition="fade">
      <section class="ssr-slide" aria-label="第 1 页：页面标题">
        <!-- 该页实际的正文、图表与素材 -->
        <aside class="ssr-notes">讲稿要点、来源与口径（不在页面上显示）</aside>
      </section>
      <!-- 按策划稿添加其余 .ssr-slide -->
    </div>
  </main>
</div>
```

| 功能 | 写法 |
|---|---|
| 翻页过渡 | `#ssr-deck` 上设置 `data-transition="fade"`（默认），或 `slide`、`zoom`、`none` |
| 入场动画 | 元素加 `data-anim="fade / up / down / left / right / scale / mask"`，用 `style="--d: 200ms"` 设置延迟；父元素加 `data-stagger="80"` 可让子元素依次出现 |
| 线条描绘 | SVG 线条或其所在的组加 `data-anim="draw"`，脚本会设置 `pathLength`。描绘只作用于描边，适合折线、连线和路线；不要用在本身就有虚线样式的线条上 |
| 逐步揭示 | 元素加 `data-step="1"`、`data-step="2"`……，每按一次“下一步”出现一组。动画与揭示不要加在同一个元素上 |
| 演讲备注 | 页内 `<aside class="ssr-notes">`；按 N 打开备注面板，供排练和独立阅读时使用 |
| 主题色 | 在 `.ssr-player` 上覆盖 `--ssr-backdrop`（画布外的底色）和 `--ssr-accent`（进度条与高亮） |

快捷键：

- 翻页：→、↓、PageDown、空格前进；←、↑、PageUp、Shift+空格后退；Home、End 跳到首尾。
- 功能：O 概览，N 备注，F 全屏，B 或 . 黑屏。
- 其他：网址 `#5` 直接打开第 5 页，刷新后停在当前页；触屏左右滑动翻页。

控制条在鼠标静止约 2.5 秒后自动隐藏。脚本还暴露 `window.SSRPlayer`（`index`、`count`、`totalSteps`、`go(i, { reveal: 'all' })` 等），供预览脚本截图使用。

动效原则：

- 先完成静态构图，再按讲述需要增加揭示或入场动画，主要用于高冲击页和逐步论证。
- 每页的入场动画控制在约 1.5 秒内结束。
- 动画结束后的状态就是元素原本的样式。打印和概览显示全部内容；“减少动态效果”只关闭过渡与入场动效，手动逐步揭示仍按讲述推进。
- 不能依靠悬停、点击或动画才能读到必要的证据。

图表可以用 HTML/CSS、内联 SVG 或已有的本地图表工具实现（见 [charts.md](charts.md)）。两种格式都要时，图表数据同时用于原生 PPTX 图表。普通文本保留为真实文字。

## 实际打开检查

用 `scripts/preview.py html` 或可用的浏览器工具打开**最终 HTML**。agent 自带的文件预览面板（如 Codex 的 `open_in_codex`）只展示文件，检查仍需浏览器实际操作或自动化工具。

```bash
python scripts/preview.py html build/index.html --out build/review/html --strict
```

脚本会：

- 拦截外部请求以验证能否离线使用，并记录控制台错误、页面异常和失败的请求。
- 等待字体与图片；记录内联图片解码失败，逐页截取完全揭示状态并生成总览、灰度和模糊版本。加载成功不能代替实际字形与裁切检查。
- 测试键盘翻页、首尾页边界和网址页码，截取宽屏与小视口下带控制条的画面。
- 打印成 PDF，核对页数与纸张尺寸。
- 对每页完全揭示后的状态做**版面检查**，有问题的页面另存一张标注图到 `layout/`（红框为错误，橙框为警告）：
  - 错误：文字超出页面、被溢出隐藏的容器裁掉、与其他文字重叠。
  - 警告：末行只剩约一个字、字号小于 12px、文字与背景对比度不足、疑似占位文字。

版面检查只看几何关系，抓不到构图、含义、裁切质量，以及文字压住图形这类问题。故意出血的元素加 `data-bleed`；确认无误、要跳过检查的局部加 `data-lint-ignore`，并在策划稿里写明原因。

然后：

1. 先处理版面检查的错误，再实际查看生成的图片：先看总览，再逐张看完整尺寸的页面，标注图指出了要重点看的位置。按 [visual-review.md](visual-review.md) 检查焦点、构图、素材、页面边缘、换行、标签和重叠。版面检查不能代替看图；当前模型不能看图时，按 visual-review 的“不能看图时”处理。
2. 动效页面至少在真实播放状态下看一次，确认过渡、入场和逐步揭示的顺序符合讲述（可截取动画中间帧，或在会话可用的浏览器中操作）。
3. 环境允许时测试全屏。
4. 修复本次引入的问题后，重新运行并查看受影响的页面。

报告区分 passed、failed、not_run、unsupported 与是否必需；严格模式只汇总脚本覆盖的自动项目，跳过打印也不是通过。实际看图、动效、SVG／CSS 素材和所需交互仍单独检查并记录，不能以进程返回 0 代替完整验收。验收记录保留在构建目录。
