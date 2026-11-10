# 银行螺丝钉直播 Keynote 生成 —— 交付包（给 Codex）

这是一套把「飞书直播大纲 → 成品 Keynote」半自动化的工具。目标：走势图、表格、柱状图、设计页**全部原生可编辑或可复用模板页编辑**，样式统一符合《银行螺丝钉 PPT 制作规范》。默认最终产物是 `.key`，不要生成 PPTX，也不要把 PPTX 当中间产物。

> ⚠️ 本包里的脚本含**绝对路径**（原作者机器 `/Users/fanlili/...`）。Codex 运行前需把路径改成自己环境的实际路径（见下文「需要改的路径」）。

---

## 0. 环境依赖

```bash
pip install python-pptx openpyxl pillow matplotlib
```
- 字体：**Source Han Sans CN Regular**（思源黑体），需安装到系统，否则中文可能回退。
- Wind 取数依赖 Node：`node`（运行 `wind-mcp-skill/scripts/cli.mjs`），且需配置 **Wind API Key**（见第 4 节）。没有 Key 时，指数行情数据无法自动取，可改用人工导出的 Excel 内联（见第 4 节 c）。
- Keynote 依赖 macOS Keynote 自动化；本 Skill 使用 `scripts/keynote_direct.py` 直接复制和修改 `.key`。

---

## 1. 目录结构

```
交付包/
├─ README_给Codex.md          ← 本文件
├─ code/
│   ├─ ppt_engine.py          ★引擎：所有页型/图型函数 + build_deck()
│   ├─ chart_gen.py           ★Wind 取数 fetch_kline(code,start,end,freq)
│   ├─ build_456.py           ★最完整的示例（456期16页，强烈建议照着改）
│   ├─ build_449.py           示例：从「数据需求xlsx」读内联数据建页
│   ├─ p4_render.py           matplotlib 双轴图（备份，一般不用）
│   └─ make_deck.py           spec.json 入口（可选）
├─ template/
│   └─ 母版（ppt格式）.pptx     ★必需，所有版式从这里来
├─ spec/
│   └─ 银行螺丝钉PPT规范_完整合并版V3.docx   完整制作规范
├─ rules/                     ★样式与流程规则（原作者的"记忆"，务必读）
│   ├─ feedback_ppt453_chart_rules.md   全套样式规则（字号/配色/走势图/表格/排版）
│   ├─ reference_feishu_cli.md          用 lark-cli 读飞书大纲
│   ├─ reference_combo_name_map.md      投顾组合简称→全称映射 + 全分页取数
│   ├─ feedback_wind_constituent_cap.md Wind 100行上限 + 代码坑 + 长区间重试
│   ├─ wind_mcp_capability_manual.md    Wind 8大数据域工具清单 + 参数字典
│   └─ feedback_wind_api_workflow.md    Wind 取数工作流与前置沟通习惯
├─ wind-mcp-skill/            Wind 取数底层（Node CLI + 说明）
│   ├─ SKILL.md / README.md / references/ / scripts/cli.mjs
├─ scripts/
│   └─ keynote_direct.py      ★Keynote 直连：复制模板、填文字、加图片、导出PDF验稿
├─ examples/                  示例数据（可直接 openpyxl 读，照格式替换）
│   ├─ 投顾组合净值数据_*.xlsx
│   └─ 【第456期直播】PPT数据需求-快照2.xlsx
└─ SKILL.md                   技能说明（Keynote 直出流程总览）
```

## Keynote 模板

默认母版：技能内置的 `assets/指数模版.key`

新任务先复制该 `.key`，再在复制件里增删改页面。不要从 PPTX 转换。

---

## 2. 标准流程（务必按序）

1. **读大纲**：用 lark-cli 读飞书 docx，取 **H3 标题原文**（用 blocks 的 children 树，不要扁平 raw_content 猜）和每页文字/来源；只取 V2「以此为准」，忽略 V1/运营内容。见 `rules/reference_feishu_cli.md`。
2. **逐页看参考图**：下载每页 image.png，确认**真实图表类型**（单/双/多线、双Y轴叠加、表格、图+表、信息图、设计页）和截图内嵌文字（复用历史图的文字也要转写进 PPT）。
3. **出「AI 解析稿」发用户确认**（关键步，不可跳过）：逐页列①页面类型②标题=H3原文③正文④图表样式⑤是否要数据⑥数据需求(指数代码+时段+周/日频+字段+红点)；并标数据获取方式（✅Wind / ⚠️人工导出）。**用户确认后再取数、再生成。**
4. **取数**（见第 4 节）。
5. **生成 Keynote**：用 `scripts/keynote_direct.py create` 复制模板 `.key`；用 `duplicate-slide`、`add-text`、`add-image` 修改本期 `.key`。
6. **校验**：从 `.key` 导出 PDF 做视觉核对；无断线/溢出/叠字/压图，标题=H3，字号/配色/黄框/数据来源齐全，Y轴整数刻度。

---

## 3. 引擎页型（ppt_engine.py，build_deck 或直接调函数）

- `cover_page(prs,layout,title)` 封面
- `section_page` 过渡/标题页（只有标题无正文图）
- `arrows_page(prs,layout,title,points)` 箭头要点页
- `box_page(prs,layout,title,boxes,bubble)` N框+气泡
- `summary_table_page` 6轮牛熊表
- `chart_page` 单线+红点高低+高亮区间
- `multi_chart_page` / `add_multi_line_chart` 双线/多线（按共同交易日对齐，不断线）
- `multi_chart_table_page` / `chart_below_table_page` 图+表
- `add_overlay_dual_axis` **双Y轴叠加图**（两个原生图叠放，详见规则）
- `img_page(...,big=True)` 图片页（纯图页 big=True 自动放大铺满）
- 通用表格、柱状图：见 `build_456.py` 里的 `_table_page` / `_bar_chart`

**关键样式（详见 `rules/feedback_ppt453_chart_rules.md`）**：
主色橙(品牌橙) `#F17640`；多线配色板①橙②浅蓝`#7CA7E0`③灰`#929292`④中蓝⑤红⑥玫红⑦金黄；黄框 `#FFBF00` 无阴影；**Y轴整数刻度**；**图例靠左手动画**；高低点红圆点用原生 marker；**多线按共同交易日对齐绝不断线**；**所有图必须有"数据来源"**；纯图页放到最大；标题=H3原文（多页同标题正常，不造副标题）。

---

## 4. 取数（三个来源）

### a. 指数行情 / K线 → Wind（chart_gen.fetch_kline）
```python
from chart_gen import fetch_kline
data = fetch_kline('000906.CSI', '2007-01-15', '2026-06-12', freq='W')  # W=周 D=日
# 返回 [{'date':'YYYY-MM-DD','close':float}, ...] 升序
```
- 底层调 `wind-mcp-skill/scripts/cli.mjs`，需先配置 Wind API Key：`node cli.mjs setup-key <KEY>`（详见 `wind-mcp-skill/SKILL.md`）。
- **长区间（十几年）会随机某年瞬时报错** → 已支持按年分段重试；建议像 `build_456.py` 的 `_kl()` 再包一层"每年多轮重试+跳过死段"。
- **指数代码坑**：中证800=`000906.CSI`（`.SH` 会被解析成个股！）；中证全指=`000985.CSI`，全收益加 `H` 前缀（如 `H00906.CSI`/`H00985.CSI`）。
- **Wind 有 100 行/100 计数硬上限**：全市场逐只成分股分布、家数、市值中位数取不全 → 改用 Wind 终端导出 Excel（见 c）。详见 `rules/feedback_wind_constituent_cap.md` 和 `rules/wind_mcp_capability_manual.md`（含 K线 period 必须用数字：10=日 11=周 12=月）。

### b. 投顾组合净值（内部数据）→ 导出 Excel 后 openpyxl 读
内部投顾组合净值（指数增强C、全市场指数C1+业绩基准、主动优选C 等）不在 Wind，来自「螺丝钉数据工具箱」。原作者用浏览器进「投顾组合净值」tab→设日期区间→导出 Excel→openpyxl 读。**Codex 没有该工具时，直接用 `examples/投顾组合净值数据_*.xlsx`**（列：交易日期/组合名称/组合单位净值/基准单位净值）。
- **简称→全称映射**（用户只说简称时自动对应，见 `rules/reference_combo_name_map.md`）：主动优选→螺丝钉金钉宝主动优选C；全市场→全市场指数**C1**；指数增强→指数增强C；全球→全球指数C；365天/90天/月薪宝→银钉宝对应C。
- 读法见 `build_456.py` 的 `combo()` 函数。

### c. 数据需求 xlsx（外汇/利率/成分股涨幅/上市公司数等）→ openpyxl 读为内联 data
用户每期会发《PPT数据需求》xlsx，每个 sheet 对应一页的权威数据。直接 openpyxl 读成 `data:[{date,close},...]` 内联进图，不走 Wind。示例见 `examples/【第456期直播】PPT数据需求-快照2.xlsx` 和 `build_456.py`（`need_sheet3/need_listed/need_csi_returns`）、`build_449.py`（`series()`）。

### d. 截图类（估值表/星级信号板/止盈图）→ 用户发图
这类图浏览器抓不到（微信公众号需登录），固定走用户发图，直接 `add_picture` 放进页面（见 `build_456.py` 的 `_p15`）。

---

## 5. 最快上手

用 `scripts/keynote_direct.py create --out ~/Downloads/第XXX期直播.key` 复制模板，然后按页面需要调用 `duplicate-slide`、`add-text`、`add-image`。数据图可先用 Python/Wind 生成 PNG，再直接放入 `.key`。

---

## 6. 更新（2026-06，重要）

**新增能力与文件：**
- `scripts/keynote_direct.py` —— Keynote 直连工具：复制 `指数模版.key`、复用模板页、填文字、加图片、导出 PDF 验稿。默认路线，不经过 PPTX。
- `code/replace_valuation.py` —— 估值表换图工具：基于 `templates/456人工成品_模板.pptx` 第16页，一条命令替换右侧估值表图，版式不变。用法 `python3 replace_valuation.py <估值表图.png> -o 输出.pptx`。
- `templates/456人工成品_模板.pptx` —— **人工成品模板**（Keynote导出）。**复用工作流**：设计/概念页（封面/过渡/四圆/宽基树状/对比框/星级&估值表/资产配置/总结等）直接复用模板里做好的那一页，只改文字/数字、替换估值表图；数据图页（走势图/柱状/分布）用引擎原生重画。这样设计页100%还原人工审美、数据图可编辑自动更新。
- `rules/reference_ppt_design_assets.md` —— 设计素材库（色板/形状/版式组件规格，详见 `素材库/设计素材库.md`）。
- `rules/reference_ppt_human_template.md` —— 复用人工模板的经验。

**两条强约束（已并入 rules/feedback_ppt453_chart_rules.md 的 A–M）：**
- **生成后必自渲染预览自检（不要盲改）**：从 `.key` 导出 PDF，再 Read PDF 对应页，检查无遮挡/溢出/贴边、标题无过程词、配色=素材库。
- **设计页必先Read原版参考图、1:1照搬**（规则M）：层级/形状/颜色/连接线/文字配色都照原版，禁止自由发挥。

**品牌色已统一**：主橙 `#F17640`（封面/表头/箭头/线），黄框 `#FFBF00` 不变。
