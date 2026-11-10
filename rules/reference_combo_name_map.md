---
name: reference-combo-name-map
description: 投顾组合简称→全称映射表 + 默认导出组合清单 + 导出必须全分页取数规则。从dbms.yhlsd.com取组合净值时用
metadata: 
  node_type: memory
  type: reference
  originSessionId: 0916e9f5-463b-48e5-94e4-b00430082a93
---

从 https://dbms.yhlsd.com/ 取投顾组合净值时的两条长期规则（用户2026-06确认，2026-08已从旧的「螺丝钉数据工具箱」aiforce入口切换到dbms.yhlsd.com，配合 [[feedback_ppt453_chart_rules]]）：

## 1. 简称 → 全称 自动对应（用户只说简称，我自动映射）
默认要导出的就是下面这几个（用户截图红框）：

| 简称 | 全称（选这个组合） |
|---|---|
| 主动优选 | 螺丝钉金钉宝主动优选C |
| 中证全指全市场 / 全市场 | 螺丝钉金钉宝全市场指数**C1** |
| 全球组合 / 全球 | 螺丝钉金钉宝全球指数C |
| 指数增强 | 螺丝钉金钉宝指数增强C |
| 365天 | 螺丝钉银钉宝365天C |
| 90天 | 螺丝钉银钉宝90天C |
| 月薪宝 | 螺丝钉银钉宝月薪宝C |

例：用户说"画主动优选组合的走势图" → 自动取「螺丝钉金钉宝主动优选C」再导出/作图。
⚠️注意区分：全市场默认是 **C1**，下拉里还有"全市场指数C"(不带1)是另一个组合，别选错。（456期P13之前误用了不带1的C，需向用户确认改成C1）

## 2. 必须按时间区间取全部分页，不能只看第一页
REST API 返回是**分页**的（`page`/`pageSize` 参数）。**严禁只取第一页**。要：设好规定的起止日期区间（`beginDate`/`endDate`）→ `pageSize` 设够大（如2000）一次性覆盖区间内全部行，或翻完所有页拿到区间内全部交易日数据，不得只取第一页就当全量用。

取法/读法详见 [[feedback_ppt453_chart_rules]]（dbms.yhlsd.com 段）：真实登录浏览器（飞书/豆包扫码）进 https://dbms.yhlsd.com/ → 选组合/设日期 → 浏览器 JS `fetch('/api/portfolio-nav?portfolioName=...&beginDate=...&endDate=...&page=1&pageSize=2000', {credentials:'include'})` 复用已登录会话直接取数据，比等"历史导出"异步任务更快。
