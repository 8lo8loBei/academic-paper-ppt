# academic-paper-ppt

把研究论文（单篇、多篇或已有学术幻灯片）做成组会、journal club、读书报告用的 **可编辑 PPTX**，并附逐页演讲者备注。

这是一个 Agent Skill：仓库本身不生成 PPT，而是给 Codex / Claude Code 等编码 Agent 一套可以照做、可以量化验收的制作规范，外加一个自动审计脚本。规范的核心约束是「版式不靠感觉」——坐标、字号、图幅、留白都有数值，做完逐项量。

## 它解决什么问题

让 Agent 直接「把这篇论文做成 PPT」，常见的失败是：标题写成「结果分析」这类空话、图被缩成缩略图、页面下半截留白、正文字号过小、公式是截图、备注空着或只有速记。本 skill 把这些问题逐条写成硬规则，并提供脚本自动检查其中可机读的部分。

## 目录结构

```
.agents/skills/academic-paper-ppt/
├── SKILL.md                        # 入口：制作流程、门禁、十条不可协商规则
├── agents/openai.yaml              # Codex 界面显示名与默认提示词
├── references/
│   ├── layout-specification.md     # 版面规格：全稿唯一的数值来源（三带、字号阶梯、图幅下限）
│   ├── page-recipes.md             # 页面配方：各页型的槽位坐标与达标线
│   ├── cover-and-visuals.md        # 封面构图、原图处理、原生公式
│   ├── content-and-evidence.md     # 论证链、选材与证据摘取
│   ├── design-system.md            # 配色、字体、目录与参考文献
│   ├── speaker-notes.md            # 演讲者备注的结构与长度
│   └── review-requirements.md      # 分步验收清单
└── scripts/audit_deck.py           # 按版面规则自动审计 PPTX
```

## 安装

把 skill 目录放到 Agent 会扫描的 skills 目录下即可。

**Codex**（个人级，所有项目可用）：

```bash
git clone https://github.com/8lo8loBei/academic-paper-ppt.git
mkdir -p ~/.agents/skills
cp -r academic-paper-ppt/.agents/skills/academic-paper-ppt ~/.agents/skills/
```

也可以直接在本仓库目录里启动 Codex：仓库自带 `.agents/skills/`，会被识别为项目级 skill。

**Claude Code**：

```bash
mkdir -p ~/.claude/skills
cp -r academic-paper-ppt/.agents/skills/academic-paper-ppt ~/.claude/skills/
```

更新时重新 `git pull` 后再复制一次（或用 `ln -s` 软链接到仓库目录，免去手动同步）。

## 使用

在 Agent 里点名 skill 并给出论文与需求，例如：

```
用 $academic-paper-ppt 把 ./papers/LigandMPNN.pdf 做成 10 分钟组会 PPT，润色封面并写逐页讲稿。
```

Claude Code 中可以直接描述任务（会按 description 自动触发），或写 `/academic-paper-ppt ...`。

建议在提示里写清以下信息，缺省时 skill 按「8–12 分钟、单篇论文」起稿并在交付说明中写明假设：

| 信息 | 示例 |
|---|---|
| 时长与听众 | 15 分钟，课题组内部，听众熟悉蛋白设计 |
| 论文篇数与主次 | 主讲 ProteinMPNN，对比 LigandMPNN、SolubleMPNN |
| 模板 / 配色 | 用学校模板；或指定深蓝白方案 |
| 不可改动的部分 | 第 3–5 页已定稿，只改结果页 |
| 汇报人信息 | 姓名、日期、单位（不提供则不编造） |

### Agent 会按什么流程做

1. **定范围**：时长、听众、篇数、模板。
2. **写论证链**：已有方法能做什么 → 什么问题卡着 → 论文改了哪一步 → 什么证据支持；据此列页面清单。
3. **选证据**：每页先定一句判断，再挑原图 / 原表 / 公式，PDF 抽图按墨迹紧裁白边。
4. **做三张样张**：封面、方法页、最密的结果页，量达标后再铺全稿。
5. **铺全稿**、**写备注**（每页 40–200 字，四段结构）。
6. **渲染复核**：PPTX → PDF → 逐页图片，按验收清单逐项检查。

### 常见任务的提示写法

- **局部改稿**：`按批注只改第 6、9 页：图放大、标题改成结论句，其他页不动。`——skill 只动指定范围，并给同页改前/改后对照。
- **多论文分享**：主线工作篇幅最大，目录双列，末页为按首次引用编号的参考文献。
- **只改备注**：`给现有 deck.pptx 每页补演讲者备注，复杂页加备答和出处。`

## 交付物

- 带原生演讲者备注的可编辑 **PPTX**（公式为可编辑 Office Math，不是图片）。
- 交付说明：页数、预计时长、主要证据页、仍存在的缺项（如无法核验的公式或外链）。
- 渲染出的逐页图片只用于复核，不替代 PPTX。

## 自动审计脚本

`scripts/audit_deck.py` 检查 PPTX 中可机读的规则：16:9 画布、备注缺失、页码、元标题（「结果分析」「Thank you」等）、字号下限、浅灰文字、分点前置标签、正文段数、图片白边（墨迹占比）、图占页面积、内容上下沿与横向跨度、封面上下带、末页类型，并统计 OMML 公式个数。

依赖：Python 3.8+、`python-pptx`；装了 `Pillow` 才会做图片白边检查。

```bash
pip install python-pptx Pillow

python .agents/skills/academic-paper-ppt/scripts/audit_deck.py deck.pptx          # 人读报告
python .agents/skills/academic-paper-ppt/scripts/audit_deck.py deck.pptx --json   # JSON，便于 Agent 解析
```

输出示例：

```
deck.pptx: 7 页，OMML 公式 2 个
  [WARN] p 7 font-size: 13 pt；若是正文分点须 ≥17：「J. Devlin et a…」
FAIL 0 / WARN 1
```

- `FAIL`：违反硬指标，必须修；存在任一 FAIL 时退出码为 1，可直接接入 CI 或 Agent 的自检循环。
- `WARN`：疑似问题，需人工判断。例如参考文献页的 13 pt 条目会报字号 WARN，属预期。

脚本不能替代目视复核：对齐、图文对应、内容是否准确仍要看渲染页。渲染可用 LibreOffice：

```bash
soffice --headless --convert-to pdf --outdir render deck.pptx
```

## 规则速查（十条不可协商）

完整数值见 [`layout-specification.md`](.agents/skills/academic-paper-ppt/references/layout-specification.md)，此处只列要点：

1. 16:9 画布；内容带 1.95–6.40 英寸，页脚线 6.62，页底 0.3 英寸内无对象。
2. 页面要满：含图页内容下沿 ≥5.90 英寸，纯文字页 ≥5.20。
3. 图要够大：单图页图墨迹 ≥32% 页面积、高 ≥3.6 英寸；PDF 图先紧裁白边。
4. 字号阶梯：标题 32 pt，分点 17 pt，图注 13 pt，脚注 12 pt。
5. 北大红 `#8B0012` 页眉之外每页只出现一处。
6. 标题写判断，不写「XX 的读法」「结果分析」「小结」；分点不加前置标签。
7. 公式先写 LaTeX 再转可编辑 Office Math，每式三层解释，一页 ≤3 式。
8. 目录在第二页，3–8 项，条目点出科学内容。
9. 封面顶带有信息锚点、底带有汇报人与日期，主图 ≥15% 页面积。
10. 不设单独 Thank you 页；末页为总结或参考文献（多篇分享必须是参考文献）。

## 定制

- **换配色 / 模板**：在提示中指定即可，skill 会整套替换色值与母版，但保留三带、字号阶梯、图幅与留白指标。`design-system.md` 内置红黑白、深绿白、深蓝白三套方案。
- **改标准**：只改 `layout-specification.md`（唯一数值来源）；如涉及可机读阈值，同步修改 `audit_deck.py` 顶部的常量（如 `MIN_BODY_PT`、`FIG_AREA_MIN`）。
