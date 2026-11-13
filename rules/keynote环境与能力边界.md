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

## 486 期补充（2026-11-12）
| 能做 | 不能做 / 坑 |
|---|---|
| `duplicate slide` 文档内整页复制（可借带气泡的页改造新页） | AppleScript **跨页/同页复制形状**都不行（「Shapes can not be copied」）→ 用 `scripts/gui_paste_shape.js` 界面拷贝粘贴 |
| `make new image`：要 `tell 幻灯片 to make new image with properties {file: (POSIX file …) as alias, …}` | 写成 `make new image at end of iWork items …` 报 -10000 |
| `make new text item`，可设字体/字号/颜色 | `add chart` 建的原生图用主题默认配色（红黄蓝绿）、坐标轴格式也不对，且改不了色 → 新图一律 python-pptx 画好再转 |
| 图片换内容：zip 替换 `Data/` 里的图片文件，保留原图片对象（边框、位置） | 换成**比例不同**的图会被拉伸（AppleScript 改 width/height 改不过来）→ 比例差大时删掉重插 |
| 图片/形状的边框、阴影：改 `DocumentStylesheet.iwa` 里样式的描边数值，或把对象的样式引用换成标准样式（同长度编号原地替换 + 同步 `Metadata.iwa` 该页的外部引用）→ `set_image_borders.py` / `set_frame_strokes.py` | AppleScript 读写不了边框/阴影/填充 |
| 读用户批注：批注文字在各页 `Slide-*.iwa` 里，和旧版本比对新增的中文字符串即可找到 | Keynote AppleScript 读不到批注 |
- **坑**：
  - Keynote **不能打开 `/tmp` 里的文件**（弹「操作不被允许」，脚本一直卡住）→ 临时文件放 `output/NNN_claude输出/tmp/`。
  - **锁屏/显示器睡眠时导入 pptx 会卡死**（open 超时、无窗口）；解锁后若仍不响应，在没有打开文档的前提下退出重开 Keynote。
  - 粘贴 pptx 转来的页后会弹「警告」窗口，之后 System Events 的 `window 1` 是警告窗；用窗口名取主窗口，先关掉警告窗。
  - **删除对象一次只删一个**：`repeat with itm in …` 里 `delete itm` 会索引失效（-1719）；`delete (every shape of s whose …)` 一次删多个、或先存引用列表再删，486 期都删错了对象（标题被隐藏、留下一个三角）。正确做法：`repeat while (count of (every shape of s whose 唯一条件)) > 0` → `delete (first shape of s whose 唯一条件)`，删完重新列出核对。
  - **出错后不要指望 `close saving no` 回到上次保存**：Keynote 自动保存会把半改状态写进文件（486 期两次「关掉不存，再开内容还是改过的」）。改前留备份；出错就把这份移到废纸篓，从备份/源文件重新复制一份再做。
  - **按文字开头匹配会误中标题**（标题与正文同开头，且页面上常有重叠的占位标题）→ 匹配时加字号/位置条件，改前打印匹配到几个。
  - 文本框垂直居中：改字、加宽后整段会上下移，三角等对齐物要按新渲染位置重算。
  - pptx 里字体名写**字体族**「Source Han Sans CN」+ bold 属性；写「SourceHanSansCN-Bold」会回退成别的字体。图例要一行放下需给 legend 设 manualLayout（x/y/w/h）。
  - 新拷入的页换母版用 `set base slide of 新页 to base slide of 底稿页`；页码字段脚本建不了，手动加一个文本框（照其他页页码的字体/颜色/位置）。

## 486 期王磊反馈修改补充（2026-11-13）
### AppleScript
- **列表引用会指错对象**：`every text item of s` / `iWork item i` 循环里拿到的引用，删改其他对象后会指向别的东西（曾把正文写进标题、删掉标题）。一律用**精确 `whose` 条件**直接定位：正文 `first text item of s whose width is 1740`，▶ `first shape of s whose width is 44 and height is 39`，标题用 `default title item of s`（被隐藏时 `set title showing of s to true`）。改前打印匹配数，必须 = 1。
- `whose` 里不能做类型转换（`whose (width as integer) is 717` 报错）→ 用范围 `whose width > 700 and width < 730`；`whose opacity is 21 or opacity is 11` 会匹配全部 → 用 `opacity > 5 and opacity < 25`。
- 新建文本框默认**居中对齐**：宽度按字数设（约 字数×字号+12），或建好后再设位置。软换行（段内换行）= `character id 8232`。
- **只改一个字用 `set character k of object text of b to "x"`**：保留原段落样式；整段重设 `object text` 可能换掉段落样式，行距跟着变。
- 新插入的图片在最上层，会盖住已有文字 → 插完图把被盖住的文字重建，或先插图后建文字。
- 个别形状要 `set locked to false` 后才能删。`note` 是保留字，不能做变量名；text item delimiters 要在 `tell application "Keynote"` 外面设。
- **导出页图跳过隐藏页**：导出的第 N 张 = 可见第 N 页，不是 slide N。王磊、汤爱学说的页码都是他们看到的 PDF/在线页码 = 可见页码，动手前先把「可见页码 → slide 序号」对照表列出来。
### IWA（直接改 .key 内部）
- **图例位置**：图表对象字段 `10000.3` = 图例框，origin x/y 是 float32。图例相对**绘图区**定位：Y 轴刻度文字更宽的图，绘图区右移，同样的数值图例会偏右 → 加 x 偏移补回（486 期 Slide-5445404 用 -1249.525，其他页 -1258.525）。改完导出量像素核对。
- **行距**：段落样式 `12.13` = 行距；空的 `6a00` 表示显式 1.0，字段不存在表示继承父样式（父样式可能是 0.9）。几页行距差几像素时，比的是这个字段，不是文字本身。
- **文本框内边距**在形状样式里（486 期 P12 用的样式 inset 5.625，基础样式 4）；同一组页面要用同一个形状样式。
- **⚠️ 样式引用只能换成该页 `Metadata.iwa` 外部引用里已有的样式**：换成未登记的样式编号，Keynote 打开即崩溃（CFHash / TSWPFontHeightInfoForFont）。要么换成已登记的，要么直接改 `DocumentStylesheet.iwa` 里那个样式的数值。改前留备份，崩溃了用备份还原。
- 等长替换最安全；字段长度变了要同步改外层 message 的长度字节（varint），否则文件损坏。
