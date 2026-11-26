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

## 487 期补充（2026-11-16）
### 界面自动化（GUI）
- **动手前先确认用户没在用电脑**：`ioreg -c IOHIDSystem | awk '/HIDIdleTime/ {print int($NF/1000000000); exit}'` 大于 20 秒才做点击/粘贴类操作；用户正在操作就停下等，不抢鼠标。
- **Keynote 窗口可能在外接显示器上**：坐标会是负数（如外接屏在左：画布左上 L=-1526, T=452）。每次 GUI 前重新截屏量 L/T/S；`scripts/gui_axis.js` 已按 L/T 相对 (273,328) 的偏移换算面板位置。点到别的窗口（如 Claude）= 坐标过期。
- **批注便利贴会挡住点击**：点图表时避开便利贴所在位置（选图表空白处或另一端的数据点）。
- 颜色面板里按 Esc 会取消当前系列选中，要重新选系列再开描边面板。
- 纵轴范围：`scripts/gui_axis.js`；新建文字框左对齐：`scripts/gui_align_left.js`（均在 skill scripts/）。
### AppleScript / 页面
- `set base slide` 换版式后，隐藏的正文占位符会重新出现：紧接着 `set body showing of slide N to false`。
- `delete shape` 在本机会删错对象（删到相邻形状或标题）：不用的形状移出画布（y≥1300），不删。表格、图片、文字删除正常。
- **画布外的参考图会伸进页面**：用户常把参考截图放在页面上方外面，下边缘一旦 >0 就会出现在导出图顶部。每次导出后检查每页 `position` 的 y<0 且 y+height>0 的对象。
- 协作在线文档（包格式，含 ActivityStream/OperationStorage）可用 `keynote_tool.py charts/get/verify` **只读**读取图表数据；AppleScript 读文字表格、导出页图也不改文件。终审核对以在线文档为准（用户上传后常在线上继续改）。

## 487 期（5星级）补充（2026-11-17）
- 图表系列数可增减：`scripts/set_cols.py`（改列数，复用原列 ID）、`scripts/set_drop_col.py`（删某一列）；新增系列沿用图表已有样式槽。
- 折线线宽：系列样式在 DocumentStylesheet（type 5028），10000 下第 48 号字段是折线宽度：`scripts/set_line_width.py in out 6 48 Slide-xxx.iwa …`。系列样式多页共用，改一页会连带别页；只改一页时请用户在界面改。
- 值轴范围：`scripts/set_axis_range.py in out Slide-x.iwa grid 轴序号 min max`（值轴 nonstyle type 5027，min=18.1、max=17.1；图表信息 field 14 列出 Y1/Y2 轴）。Keynote 刻度条数自动（实测 1.5 个百分点只画 2 段），0 线不一定能出现。
- 图表类型 field 1（10=混合、11=双轴）；面积改柱形不能只改 field 1（坐标轴会跑到左边、格式错乱），请用户在界面「序列」里改。
- 双轴叠加图首选：复制本期已有 Keynote 原生双轴页换数据，不用 python-pptx（用户 2026-11-17 明确：不要出现 PPT 格式，要直接编辑 Keynote）。
- 双图并排页：原生图表各放一张跳过页，可见页贴导出图片（用户 2026-11-17 定）。
- 协作在线文档的本地副本：`cp -R` 后用 Keynote `save … in` 另存为普通 .key 再改。
- 界面代点前必须截屏确认前台是目标窗口（487 期两次前台是微信/在线文档/飞书）；用户空闲 <15 秒不点。

## 488 期补充（2026-11-19）
### 换截图（同位置、同尺寸替换图片）
- 用户发来新截图替换页面上的旧图（如 P23 雪球业绩曲线）：
  1. 手机截图常是横屏拍的竖图，先转正（PIL `rotate(90, expand=True)`，按文字方向判断）；
  2. 先从包里取出旧图（`Data/xxx-NNNN.png`）看裁切范围和像素尺寸，新图按**同一范围、同一像素尺寸**裁好（缩略图 `xxx-small-NNNN.png` 同步按原尺寸缩放）；
  3. 在 Keynote 里关闭文件 → 留改前备份 → 在一个临时目录里建 `Data/` 放好新图，用 `zip -0` 只更新这几个条目：`(cd 临时目录 && zip -0 -q 绝对路径/成品.key Data/xxx.png Data/xxx-small.png)`；
  4. `keynote_tool.py verify 备份 成品`：「修改」只应有这几个图片，「缺失」= 无。位置、边框、叠在上面的图例文字都不变。
- **再次强调安全铁律第 8 条**：488 期这一步先用了 Python zipfile 整包重写，中文文件名的 12 个图片全部变乱码丢失（verify 报缺失），靠备份恢复。改包内文件只用 `zip` 命令。

### 删形状
- **不要按文字批量删形状**。488 期按便签文字删时，脚本误删了 P14 全部 8 个形状；而且便签文字其实留在一个红色虚线框里（框是从便签复制出来的），按文字删就会把框一起删。
- 要删某个形状：先列出该页所有形状（序号 + 文字 + 位置尺寸），确认**只命中 1 个**再删；删完核对形状数只少 1 个，并导出页图和删前像素比对。
- 用户的便签/批注本身一律不删不清空（见 `复核记录与交付.md`）。

### AppleScript 保留字（不能做变量名）
- 已踩过：`note`、`by`、`key`（488 期 `set key to …` 报 -10006）。变量名用 `kw`、`nn` 这类缩写。
- `repeat with n in {5, 15, 24}` 里的 `n` 是引用，用作 `slide n` 前先 `set n to nn as integer`。
- 导出页图跳过隐藏页（488 期 `skipped slides` 设 true/false 实测都是 25 张 = 可见页）：第 N 张图 = 可见第 N 页，不是 slide N。

## 488 期补充（2026-11-20，王磊反馈后）
### 给已有图片加底框（如「左边也用框框起来」）
- Keynote 脚本**不能复制图片**（`duplicate image` 报 "Images can not be copied"），也**不能调图层顺序**；新插入的底图（框、卡片）一定在最上层，会盖住原来的二维码/截图和文字。
- 做法：从包里取出原图（`unzip -o -q 成品.key "Data/xxx-NNNN.png" -d 临时目录`，先 `sips` 看像素尺寸确认没裁切），用 `kit.py` 的 `Layer.img(路径, x, y, w, h)` 把图片连同 #FEAD00 6pt 边框一起画进底图；删掉该页旧底图、旧图片和文字，按「底图 → 文字」顺序重建（`SlideScript.text`）。底图 PNG 用**新文件名**（如 `p18_frames_v4`），避免以后粘到在线文档时复用同名旧图。
- 代价：图片不能再单独拖动/换图，以后换截图要重新生成底图；交付时告诉用户。
- 两栏并排的框：一行放两个步骤（Bold 36）需要约 1010px 宽，左框至少约 1100px；右框里最长一行 38 号字约 610px，右框约 700px。

### 文件被替换前先关掉
- 脚本打开过的文档如果没关，又用 `mv` 把磁盘上的文件换成新包（如 `keynote_tool.py set` 的输出），Keynote 里那份仍是旧内容，再执行 AppleScript 会报 -10000「AppleEvent 处理程序失败」。换包前先 `close (every document whose name contains "…") saving no`（只关自己的文件，不碰在线文档），换完再开。

### 更新到新日期时连带检查气泡里的派生数字
- 只补最新一天的数据，也要重算气泡/正文里由它算出来的数（488 期两融补到 11-19 后，「从巅峰撤去了7000多亿」变成 8107 亿，改为「8000多亿」）。改了要告诉用户。

## 490 期补充（2026-11-26）
- **`iWork item N` / `text item N` / `shape N` 的序号各算各的，删过对象后还会整体前移**：按序号改文字前，先 dump 一遍（`class`、`position`、`object text`）确认是哪个对象；删了对象之后要重新 dump，不能沿用旧序号。490 期先删图片再按旧序号改标题，结果把页码占位符改成了标题文字。
- **页码占位符改坏了，界面开关「幻灯片编号」、重套版式都恢复不了**（Keynote 会缓存占位符里的旧文字），AppleScript 也不能插入页码字段，「插入」菜单里没有页码项，占位符也不能拷贝。**修法**：`duplicate` 一张页码正常、同版式的页 → 删掉它的内容对象（保留标题和页码）→ 设标题文字/位置 → 重新放图片、文字（kit 脚本重跑）→ 删掉坏页。删对象时如果把标题也删了，`set title showing of slide to true` 后再设 `default title item`。改完逐页读页码对象，`object text` 必须是「￼」。
- **改页码以外的文字一律用 `default title item` 改标题**，不要用 `iWork item 2` 之类的序号猜。
- kit 透明图层有时会盖在后建的文本框上面（点击选中的是图层）：选中图层 → 排列 ▸ 移到最后面；仍点不中就把文本框临时挪到空白处操作完再挪回。
- 纵轴：`set_axis_range.py` 只改最小/最大值，Keynote 会自己重排刻度数（490 期 -2.2~4.4 变成只有 3 个刻度、没有 0 线）。要固定刻度数用界面：图表 ▸ 坐标轴 ▸ 值(Y) 的最大、最小、步数；**百分比坐标轴在界面里填小数**（500% 填 5，填 500 会变成 50000%）。本机窗口布局下三个输入框在屏幕点坐标约 (1010,312) 最大、(1010,344) 最小、(1115,312) 步数（每次先列出 AXTextField 位置再填）。
- 表格单元格写数字会显示成 0.1446：百分比一律写成文本「14.46%」（先 `set format of cell to text`），保留两位小数。
- 表格合并单元格后 `set width of column` 可能不生效：用 `set width of table` 整体缩放。
- 标题两行（90 号不缩字号）：标题框 y≈70~85、正文框和 ▶ 往下挪到 y≈338，走势框从 y≈450 起；导出页图核对第二行不压正文。
- **AppleScript 删除不能边遍历边删**（490 期本地文件也出事）：`repeat with x in (every shape …)` 里 delete，列表序号会失效——要么报 -1728/-1719，要么删到别的对象（490 期 P2 把页内的两个框删了）。`whose` 过滤结果也一样，改掉一个后再取 `item 2` 会越界。做法：每次只删/改一个，删完重新查询再删下一个；删前后核对对象数。按序号 `shape i` 删也不可靠（序号和 `every shape` 列出的顺序可能不同）。**页面外的历史残留（放映不可见）一律不删**，删错的代价远大于留着。
- 出错后用改前备份**整份还原**再重做，不要逐项手工补。
- `set_axis_range.py` 改上下限会改变刻度格数（490 期 P9 把上限设成 0.6000000000000001 想消掉「-0%」，6 格变 3 格）：改完必须导出看格数；百分比轴 0 刻度显示「-0%」是浮点问题，改不掉就保留，交付时写明。
- **AppleScript 不能复制表格**（`duplicate table` 报 Tables can not be copied）。拆表：界面选中表格（先点页面空白处取消选择，再点表格里某格，避开批注便签）→ 编辑 ▸ 拷贝 → 点空白处 → 粘贴，得到第二张表；两张表各删另一半的列。表头合并单元格所在的列删一列会连同整组一起删，删前删后都数 `column count`。宽度用 `set width of table`，位置用 `set position`。
- **只换数据、尺寸不变的 kit 图层**（估值表放大、柱状图等）：重渲染后直接 `zip_update` 替换包里同名 PNG（`Data/xxx-NNNNN.png` 和 `-small-` 预览图），不用删图重插、不用 GUI 移到最后面。新图像素尺寸必须与旧图一致（差几像素就在右/下补透明边，不要拉伸）；替换后 verify + 导出看。只适用于本地单文件 .key，协作在线文档不行。
- **协作在线文档里改表格单元格底色，可能把同一张表其他行的底色冲掉**（490 期 P3：只给第 9、11 行设浅黄，线上第 1～7 行的浅黄全没了，本地试改正常）。线上改表格底色时**整张表每一格都显式设一遍**，改完导出线上页图看全表。
- 本地副本和在线文档同源时（复制上传的），在线文档开着会导致本地副本 `open` 后取不到 document：先关在线文档（确认 `modified` 为 false），试改完再开。
