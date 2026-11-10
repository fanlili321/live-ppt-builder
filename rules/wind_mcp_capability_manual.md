# Wind MCP Skill 完整能力手册

## ⚠️ 强制规则（每次调用Wind API前必须执行）

**在构造任何 Wind API 调用参数之前，必须先查阅本手册对应章节，确认：**
1. `server_type` 和 `tool_name` 是否在工具清单中
2. 所有参数名和参数值是否与手册完全一致，禁止凭经验或通用知识推断
3. 枚举类参数（`period`、`indexes`、`freq` 等）必须逐字对照手册取值，不得自行发明

**典型错误示范（严禁重复）：**
- ❌ `"period":"W"` → ✅ `"period":"11"`（周K）
- ❌ `"period":"M"` → ✅ `"period":"12"`（月K）
- ❌ `"period":"D"` → ✅ `"period":"10"`（日K）
- ❌ `"lang":"zh"` → ✅ `"lang":"中文"`

## 一、8大数据域 & 工具清单

### 1. A股数据 (stock_data)
- `search_stocks` — 全市场A股筛选，参数：question
- `get_stock_price_indicators` — 行情快照，参数：windcode + indexes
- `get_stock_kline` — K线历史，参数：windcode + begin_date + end_date
- `get_stock_quote` — 分钟行情，参数：windcode
- `get_stock_basicinfo` — 公司档案，参数：question
- `get_stock_fundamentals` — 财务数据，参数：question
- `get_stock_equity_holders` — 股东股本，参数：question
- `get_stock_events` — 公司事件，参数：question
- `get_stock_technicals` — 技术指标，参数：question
- `get_risk_metrics` — 风险指标，参数：question

### 2. 港美股 (global_stock_data)
- `search_global_stocks` — 港美股筛选，参数：question
- `get_global_stock_price_indicators` — 行情快照，参数：windcode + indexes
- `get_global_stock_kline` — K线历史，参数：windcode + begin_date + end_date
- `get_global_stock_quote` — 分钟行情，参数：windcode
- `get_global_stock_basicinfo/fundamentals/equity_holders/events/technicals/risk_metrics` — 各类基本面，参数：question

### 3. 基金数据 (fund_data)
- `search_funds` — 基金筛选，参数：question
- `get_fund_price_indicators` — 行情快照，参数：windcode + indexes
- `get_fund_kline` — K线历史，参数：windcode + begin_date + end_date
- `get_fund_quote` — 分钟行情，参数：windcode
- `get_fund_info/financials/holdings/performance/holders/company_info` — 各类基金信息，参数：question

### 4. 指数/板块 (index_data)
- `get_index_price_indicators` — 行情快照，参数：windcode + indexes
- `get_index_kline` — K线历史，参数：windcode + begin_date + end_date（**注：曾出现服务不可用**）
- `get_index_quote` — 分钟行情，参数：windcode
- `get_index_basicinfo` — 指数档案，参数：question
- `get_index_fundamentals` — 基本面/成分变动，参数：question
- `get_index_technicals` — 技术指标，参数：question

### 5. 债券数据 (bond_data)
- `get_bond_basicinfo/issuer_info/market_data/financial_data` — 参数：question

### 6. 公告/新闻 (financial_docs)
- `get_company_announcements` — 公司公告，参数：query
- `get_financial_news` — 财经新闻，参数：query

### 7. 宏观经济 (economic_data)
- `get_economic_data` — EDB指标，参数：metricIdsStr + 可选日期/频率
- 支持自然语言：中国CPI同比、美国失业率、中国社融规模、美联储利率等

### 8. 通用取数 (analytics_data)
- `get_financial_data` — 自然语言取数，参数：question
- ⚠️ **不得用于指数成分股精确查询**（容易串数据，见300成长事故）

---

## 二、K线参数 (period) — 必须用数字，不能用字母
- 1=1分钟, 3=5分钟, 4=10分钟, 5=15分钟, 6=30分钟, 7=60分钟, 8=120分钟, 9=240分钟
- 10=日K, **11=周K**, **12=月K**, 13=年K
- ⚠️ 错误示例：`"period":"W"` 或 `"period":"M"` → 会导致TEMPORARILY_UNAVAILABLE
- ✅ 正确示例：`"period":"11"`（周）、`"period":"12"`（月）、`"period":"10"`（日）
- aftime: 0=前复权, 1=后复权
- count: 正数=从begin向后取N条, 负数=从end向前取N条

---

## 三、行情快照关键字段 (indexes参数，逐字抄写)

### 基础行情
`中文简称` `最新成交价` `前收盘价` `今日开盘价` `今日最高价` `今日最低价`
`成交量` `成交额` `涨跌` `涨跌幅` `5分钟涨跌幅` `换手率` `量比` `总市值1` `流通市值`

### 估值
`市净率` `市净率(LF)` `市盈率(TTM)` `市盈率(LYR)` `市盈率(预测)` `股息率`

### 多周期涨跌幅
`5日涨跌幅` `10日涨跌幅` `20日涨跌幅` `60日涨跌幅` `120日涨跌幅` `250日涨跌幅` `年初至今涨跌幅`

### 技术指标
`指数平滑异同移动平均` `DIF快线` `随机指标K值` `随机指标D值` `随机指标J值`
`6周期相对强弱指标` `12周期相对强弱指标`
`布林中轨` `布林上轨` `布林下轨`
`5周期移动平均` `10周期移动平均` `20周期移动平均` `60周期移动平均` `120周期移动平均` `250日均线`

### 资金流向
`当日主力净流入额` `当日主力净流入占比`
`近5日主力净流入额` `近10日主力净流入额` `近20日主力净流入额`

### 基金专属
`最新净值` `累计净值` `近一周净值增长率` `近一月净值增长率` `近一年净值增长率`
`近三年净值增长率` `IOPV` `贴水率` `基金规模` `七日年化收益率` `万份基金收益`

### 债券专属
`最新YTM` `麦氏久期` `修正久期` `涨跌BP`
`最新价(净)` `最新价(全)` `转股溢价率` `纯债溢价率`
