# 螺丝钉数据工具箱（dbms.yhlsd.com）

团队内部金融数据工作台（2026-08 起替代旧 aiforce 入口）。按权限可查：指数行情（后端调 Wind）、基金净值、两融规模、各资产代表指数、**螺丝钉星级**（A股/港股/红利/全球/黄金）、投顾组合净值、估值、投研数据（指数盈利增速）、**黄金开采成本**。系统记录登录、查询、导出操作。**只读，绝不调用 import / sync / adjustment / generate / approve 等写接口。**

## 两种取法
### A. 内置浏览器登录状态（退路：凭证失效时，或要查 v1 没有的接口）
1. 内置浏览器打开 https://dbms.yhlsd.com/ ，登录状态会保留（2026-11-09 登录，11-10 仍有效），取数时不用找用户。只有跳到 accounts.feishu.cn（登录过期）时，才请用户自己点飞书登录（不代输账号），登好继续。
   - 每期记一下是否需要重新登录，若频繁过期，再建议用户走 B。
2. 在页面里用 JS 复用登录会话（`javascript_tool`）：`await fetch(url,{credentials:'include'}).then(r=>r.json())`。

已实测可用的只读接口（2026-11-09）：

| 数据 | 接口 | 参数 | 返回要点 |
|---|---|---|---|
| A股螺丝钉星级（中证全指） | `GET /api/screw-star` | `begin`,`end`（YYYY-MM-DD） | ratingDate, starRating, chinaIndex, chinaIndexTotalReturn |
| 分市场星级 | `GET /api/screw-star/markets/{hk\|dividend\|…}` | `begin`,`end` | starRating, indexCode / totalReturnIndexCode（港股 HSI.HI/HSIRH.HI，红利 000922.CSI/H00922.CSI）, 估值指标 |
| 全球 / 黄金星级 | `GET /api/screw-star/special/{global\|gold}` | `begin`,`end` | 黄金：londonGold, shanghaiGold, **miningCost**, starRating, ruleVersion |
| 黄金开采成本 | `GET /api/screw-star/special/gold/mining-costs` | 无 | items（按生效日的成本序列）、current、rates（WGC 最新季度成本、同比、月增速） |
| 投顾组合净值 | `GET /api/portfolio-nav` | `portfolioName`,`beginDate`,`endDate`,`page`,`pageSize` | **必须取全部分页** |
| 组合名称列表 | `GET /api/portfolio-nav/names` | 无 | 全称列表（全市场默认 C1，见 `reference_combo_name_map.md`） |
| 两融余额 | `GET /api/margin-balance/series` | 无 | tradeDate, balance… |
| 指数盈利增速 | `GET /api/research-data/index-earnings-growth/indices`、`…/periods` | 无 | 指数列表、报告期 |
| 指数行情 / 基金净值 | 页面「指数查询」「基金净值」操作（后端 POST `/api/query-history/wind-query`、`/api/fund-nav/query`） | 指数代码、全收益开关、日期、频次、字段（收盘/最高/最低）、指标（累计收益、最大回撤、年化、区间收盘最大涨幅） | POST 会写查询历史，优先用 Wind 或 API v1 |

### B. 开放接口 API v1（**默认用这个**，2026-11-10 已配置凭证，完全免登录）
- 脚本：`scripts/dbms_api.py`（`check` 测凭证；`get <路径> k=v… --out x.json` 自动翻页；`index-query` 查 Wind 行情）。只允许 GET 和 POST /index-queries。
- 凭证存放：用户本人执行 `mkdir -p ~/.config/dbms && pbpaste > ~/.config/dbms/token && chmod 600 ~/.config/dbms/token`（先复制凭证）。AI 不经手、不回显凭证。
- 凭证在工具箱「API Client → 管理凭证」生成（需要 api_client:manage 权限，没有就找管理员），完整密钥只显示一次，**默认 90 天有效**。脚本报 401/403 = 过期或无权限，提醒用户重新生成替换。
- 2026-11-10 管理员已发凭证（`st_` 开头，存 `~/.config/dbms/token`，约 2027-02 到期，到期前提醒用户找管理员更新）。`ak_` 开头的是 Wind 密钥，不是 v1 凭证。
- 实测接口行为（2026-11-10）：
  - 返回统一 `{"data": {"items": [...], "total", "truncated"}, "meta": {requestId, rowCount, nextCursor}}`；基金净值是 `data.records`。
  - **星级 `/star-ratings/{screw|hk|dividend|global}`**：参数 `begin`、`end`（YYYY-MM-DD）、`limit`（默认 100，单次最多 1000），按日期倒序，超出时 `truncated=true`。→ 一律用 `dbms_api.py star <kind> 起 止`（自动分批取全）。
    - 2026-11-10 教训：曾以为「begin+end 同传会 500」，实为测试时 zsh 不拆分变量，多个参数被当成一个字符串塞进了 begin；管理员查日志确认后已加 422 参数校验。**参数一律分开传，不要手拼 query 字符串。**`gold` 一次返回全部周度记录，日期字段是 `reportDate`（另有 `marketCutoffDate` 行情截止日，用它对齐行情）。
  - `/portfolio-nav`、`/valuations/dingda`、`/funds/{code}/nav`：`beginDate`/`endDate` 有效。
- 不需要工具箱的数据（指数行情）直接走 Wind；工具箱只用于星级、估值、投顾组合净值、开采成本等独有数据。
- 基址 `https://dbms.yhlsd.com/api/v1`，请求头 `Authorization: Bearer $API_TOKEN`。权限范围：fund_nav:query、star:read、dingda_valuation:read、screw_valuation:read、research_data:read、portfolio_nav:read、index:query。
- 接口：`GET /funds/{code}/nav`（单位净值与复权净值）、`GET /star-ratings/{type}`（screw/hk/dividend/global/gold）、`GET /valuations/dingda`（钉大估值历史，例 `?indexCode=000688&beginDate=…&endDate=…&limit=100`）、`GET /valuations/daily`（每日估值历史）、`GET /portfolio-nav`、`POST /index-queries`（Wind 指数行情）。
- Token 由用户在工具里自己申请/生成，存到本机（如 `~/.config/dbms/token`，权限 600），脚本从文件读取；**不要把 Token 写进 skill、聊天或代码仓库**。
- 页面上的「配置 API Key」是 **Wind MCP 的 Key**（存在浏览器本地），不是这里的 API Token。

## 口径注意
- **黄金开采成本有两套口径（2026-11-09 发现）**：
  - 工具箱（黄金星级用的）：月度序列，按 WGC 最新季度成本的同比算月增速逐月递推。2026-11-01 生效值 **1798.05**（rates：2026Q2 WGC 1929.54、一年前 1585.18、月增速 1.81%）；10-02 1766.09、09-07 1753.81、08-01 1741.62。
  - WGC 原始季度值：2026Q2 1930（485 期 P14/P16 定稿按用户要求用此口径，6 月起统一 1930）。
  - PPT 里成本、比值和黄金星级同框出现时，要么统一口径，要么注明差异，**动手前问用户用哪套**。
- 星级以工具箱为准（和小程序一致），如 2026-11-06 黄金星级 3.12、A股 4.87、港股 4.45、红利 4.13、全球 3.64。
- 交付时来源写「螺丝钉数据工具箱」+ 接口/页面 + 查询条件。
