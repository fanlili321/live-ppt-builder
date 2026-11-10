---
name: feedback-wind-constituent-cap
description: Wind MCP 有100行/100计数硬上限，全市场逐只成分股分布、家数、市值中位数等取不全，需Wind终端导出
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 0916e9f5-463b-48e5-94e4-b00430082a93
---

Wind MCP（wind-mcp-skill）做**全市场/全成分股逐只统计**时有硬上限，实测：
- `stock_data search_stocks`、`analytics_data get_financial_data` 都把"聚合/计数"意图塌缩成**最多返回 100 行的股票列表**（excelTotalCount 封顶 100）；只有当某筛选真实命中 <100 只时 excelTotalCount 才反映真实家数（如"年初至今涨幅>300%"=25 只准确）。
- 因此**全市场涨幅分布直方图、跑赢/跑输指数家数、市值中位数（需逐只成分股）这类无法用 MCP 取全**——故事级分桶（>10% / 0~10% / 跌等）都 >100 被截断，分桶反推也失效。

**可靠能取的**：单指数行情/涨跌幅（get_index_price_indicators，如中证全指年初至今涨跌幅）、指数总市值1（→ 平均市值=总市值÷成分股数，仅固定成分数的指数可算）、Top100 涨幅榜。

**结论**：涉及全市场逐只分布/家数/市值中位数的页（如456期P5个股涨幅分布、P10市值中位数），数据需从 **Wind 终端一次性导出 Excel**（成分股+年初至今涨跌幅/市值），再 openpyxl 读为内联 data 出原生图。配合 [[feedback-ppt453-chart-rules]]。

补：中证800 指数 = `000906.CSI`（.SH 会被解析成股票"浙商中拓"）；中证全指=000985.CSI（+全收益H00985），中证800全收益=H00906.CSI。中证800 行情数据从 2007-01-04 起（2004 基日无K线）。Wind 长区间(十几年)多年取数会随机某年瞬时'content'报错，需**按年分段+每年多轮重试**容错。
