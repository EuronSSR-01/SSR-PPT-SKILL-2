# 跨 agent 对比测试

目的：在 Codex、Claude Code、Cursor、Grok Build、Antigravity（agy）、DeepSeek Harness（dsh）等 agent 中，用相同输入运行本技能，比较结果。工程检查用脚本统一测量，设计与内容由人评分。

## 准备

1. 在各 agent 中安装同一版本的技能（见仓库 README 的安装表），记录版本号（`git rev-parse --short HEAD`）。
2. 每个 agent 用一个空的测试目录，把 [brief.md](brief.md) 中的“材料”放进去。
3. 先在该目录运行 `python <技能目录>/scripts/doctor.py --json > doctor.json`，记录环境能力。
4. 记录 agent 名称、版本、所用模型，以及模型能否读图。

## 运行

把 [brief.md](brief.md) 的“提示词”和“材料”原样发给 agent，不做额外提示。记录：
- 总耗时；
- 是否中途提问（提示词已要求不再追问，并且没说表达方式，所以应不提问，按图示优先制作，并在 planning.md 开头写明）；
- 是否出错或卡住；
- 最后它自述的验收范围。

另跑一次“图文均可”变体：在提示词末尾加一句“表达方式：图文均可。”，确认 agent 没有被强制图示化。

## 统一测量

agent 完成后，由测试人在同一台机器上运行（不使用 agent 自己的检查结果）：

```bash
python <技能目录>/scripts/preview.py html output/index.html --out review/html --visual-first
python <技能目录>/scripts/preview.py pptx output/deck.pptx --out review/pptx --edit-test --visual-first
```

从两个 `report.json` 中记录：
- 页数是否为 6；
- HTML 与 PPTX 的版面检查 errors／warnings；
- 每页正文字数与图示优先提示数（`text` 字段；只作参考，不代表图示质量）；
- 浏览器错误与外部请求；
- 打印检查；
- 编辑测试的通过、失败和不支持数量。

## 人工评分（每项 1–5 分）

| 维度 | 看什么 |
|---|---|
| 信息保真 | brief.md 的“预期要点”逐条核对：是否夸大因果、合并预算状态、漏掉老人占比问题、缺少期间 |
| 结构与叙事 | 标题连读能否讲清复盘结论与下一步 |
| 阅读层级 | 每页焦点是否一眼可见，数据是否直接标注 |
| 图示执行 | 图示优先时：趋势、构成、因果条件等是否由图承担，而不是装进卡片的要点；图解是否补造了材料没有的关系。图文均可变体：是否按内容选择，没有被强制图示化 |
| 视觉与记忆点 | 是否有清楚的视觉概念；有没有值得截图的页；是否命中 anti-patterns 中的典型问题 |
| 双格式一致 | HTML 与 PPTX 的内容、数据和条件是否一致 |
| 如实报告 | agent 对“是否看过图、哪些没做”的自述是否与实际一致 |

## 记录表

| agent／模型 | 读图 | 耗时 | 提问 | 页数 | 版面 err/warn（HTML / PPTX） | 编辑测试 | 字数／提示 | 信息保真 | 结构 | 层级 | 图示 | 视觉 | 一致 | 如实 | 备注 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| | | | | | | | | | | | | | | | |

每个 agent 至少跑两次，看结果是否稳定。同一 agent 换模型时视为新的一行。对比时固定技能版本和材料，只改变 agent 或模型。
