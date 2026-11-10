# Keynote 环境与 AppleScript 能力边界（实测）
## 环境
- 本机是**正版 Apple Keynote 15.3**，只是 .app 被改名 `/Applications/Keynote Creator Studio.app`（bundle id `com.apple.Keynote`、含Keynote.framework、© Apple）。别被名字迷惑。
- osascript 报 `-600 应用程序没有运行` = 没先启动。**先 `open -a "/Applications/Keynote Creator Studio.app"` 再 tell**。
- 用 `sdef` 判App真伪不可靠（本机无Xcode，sdef报错返空=假阴性）。

## 能做 / 不能做
| 能做 | 不能做 |
|---|---|
| 文字 font/size/color（可读可设，用于核验修复字号） | **图表线颜色/图例/坐标轴**（图表对象只有 position/height/width/locked） |
| `add chart` 建原生图+填数据 | 用 AppleScript 改**已存在**图表的数据（-10006）→ 改用 `scripts/keynote_iwa/keynote_tool.py set` |
| 表格 cell 值、部分列宽 | 表格 cell 填色/表头（走 python-pptx→key 更稳） |
| 满页贴图、`base slide` 切母版、跨页删/移、导PDF、pptx↔key转换 | 跨文档 duplicate slide（实测不生效）→ 改用界面拷贝粘贴 `scripts/copy_slides.py`（需辅助功能权限，见 `rules/多期拼页.md`） |

**结论**：图表配色/样式 → 一律 **python-pptx 出 styled 原生图 → Keynote 转 .key**（转后仍原生可编辑）；纯 Keynote 直连（add chart）出的图**改不了色**，别用。

## 2026-11 更新：已有图表换数据
- `scripts/keynote_iwa/keynote_tool.py` 直接改 `.key` 里 `Index/Slide-*.iwa` 的图表数据网格（行名、各系列数值、行 ID），**样式完全保留**，行数可增减，空格=空单元格。已在 485 期实测（P2、P14、P16、P24 等）。
- 换数据后：Y 轴若为自动范围会跟着变；图下方起止日期、红点数值、气泡标注是独立文本框，要另外改。
- 所以「历史页 + 数据待更新」的走势图，**首选复制历史 .key 换数据**；python-pptx→key 只用于没有历史出处的新图。
