# PPT数据提需（工具箱和 Wind 都拿不到的数据，请人工导出）

- 文档：飞书 docx「PPT数据提需」 https://epndqwwg0a.feishu.cn/docx/VST4d5u30oeY5NxwhZtcAKCUnId （token `VST4d5u30oeY5NxwhZtcAKCUnId`）。顶部 callout 是数据工具箱链接。

## 什么时候写
取数时按顺序试：大纲/用户给的来源 → 数据工具箱（`数据工具箱dbms.md`）→ Wind。**三者都拿不到**，或者能拿到但口径/区间不全（例：工具箱沪深300 PE/PB 只有 2024-07 以来）、Wind 无权限或返回 0 行、需要盘中分时、份额/成交额/规模等 Wind MCP 有 100 行上限的数据、内部截图类，才写提需。能自己算出来的（回撤、涨幅、年化、比值）不提需。

## 格式（照文档里已有的写法，新的一条放最上面）
插在顶部 callout 和第一个已有 `<h1>` 之间（先 `--scope outline --detail with-ids` 拿到第一个已有日期标题的 block id，用它前面那个空 `<h1>` 做锚点 `block_insert_after`；写前确认当天这一期还没提过，避免重复）：

```xml
<h1>YYYYMMDD数据需求（期望交付MMDD）</h1>
<h2>【第NNN期直播】标题</h2>
<table><colgroup><col/><col/><col/><col/></colgroup><tbody>
<tr><td vertical-align="middle"><p>序号</p></td><td vertical-align="middle"><p>数据需求</p></td><td vertical-align="middle"><p>时间</p></td><td><p>样式参考</p></td></tr>
<tr><td vertical-align="middle"><p>1</p></td>
    <td vertical-align="middle"><p><b>指标/标的名称</b>（代码，如有）</p><p>频率、字段、单位；用在 PNN 哪张图</p><p>为什么要人工：工具箱/Wind 试过的结果</p></td>
    <td vertical-align="middle"><p>YYYY-MM-DD 至 YYYY-MM-DD（或「更新至最新」）</p></td>
    <td vertical-align="middle"><img path="@./参考图.jpg"/><p>参考哪一期哪一页</p></td></tr>
</tbody></table>
```
- 「数据需求」写到人工拿着就能导：名称、代码、频率（日/周/月）、字段（收盘/份额/规模/成交额…）、单位、区间；写明用在哪一页。
- 「时间」写数据区间；文档标题写期望交付日期（一般是做 PPT 的当天或第二天，按直播日期倒推）。
- 「样式参考」放历史 PPT 那一页的截图（从底稿导出），说明照哪页做。
- 写完 `--scope section` 读回核对，然后告诉用户：提了几条、各是什么、期望何时交付。

## 拿到数据以后
- 人工一般回 Excel：按 `examples/【第456期直播】PPT数据需求-快照2.xlsx` / `code/build_449.py` 的方式读，内联进图表（`keynote_tool.py set`）。
- 和历史 PPT 比对（`compare_history.py`），在页面对照表里写来源「人工导出（PPT数据提需 YYYYMMDD 第N条）」。
