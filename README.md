# live-ppt-builder 同事上手说明（2026-11-26）

银行螺丝钉每周直播 PPT（Keynote）从大纲到定稿的完整流程。范丽丽请假时，同事按本说明装好环境，在 Claude Code 里说「NNN期直播大纲已确认，开始做ppt吧 <飞书大纲链接>」即可按同一套规则做。

流程和规则以 `SKILL.md` 为准（Step 0–7，四个环节：AI 初稿 → 用户修改 → 汤爱学复核 → 王磊审），本文件只讲**怎么装、要哪些账号、哪些东西只在原作者电脑上有**。旧版 `README_给Codex.md`（2026-06，PPTX 路线）已过时，仅作历史参考。

## 1. 一起用到的几个 skill（都在 github.com/fanlili321 下）

| skill | 仓库 | 用途 |
|---|---|---|
| live-ppt-builder | `live-ppt-builder` | 做 PPT（本仓库） |
| ppt-review | `ppt-review-skill` | 成品逐页复核（交 v1 前、上传前必跑） |
| live-ppt-publisher | `live-ppt-publisher`（skill 在 `skill/live-ppt-publisher/`） | 定稿后发布、归档 |
| outline-builder | `outline-builder-skill` | 大纲（PPT 之前的环节） |
| ppt-visual-guide | `ppt-visual-guide` | 新画信息图时参考 |
| image-material-library | `image-material-library` | 历史 PPT 配图检索 |

另需：`wind-mcp-skill`（Wind 取数，可选）、飞书 `lark-*` 系列 skill（随 lark-cli 安装）。

## 2. 安装

```bash
mkdir -p ~/GitHub && cd ~/GitHub
git clone https://github.com/fanlili321/live-ppt-builder.git
git clone https://github.com/fanlili321/ppt-review-skill.git
git clone https://github.com/fanlili321/live-ppt-publisher.git
mkdir -p ~/.claude/skills
ln -s ~/GitHub/live-ppt-builder ~/.claude/skills/live-ppt-builder
ln -s ~/GitHub/ppt-review-skill ~/.claude/skills/ppt-review
ln -s ~/GitHub/live-ppt-publisher/skill/live-ppt-publisher ~/.claude/skills/live-ppt-publisher
```
每次开工前 `git pull` 三个仓库，拿到最新规则。脚本路径都按脚本所在目录查找，装在哪里都能跑。

## 3. 本机环境

- **macOS + Keynote**（正版 Keynote；如果 .app 改过名，按 `rules/keynote环境与能力边界.md` 开头的说明先 `open -a` 启动）。
- **系统设置 → 隐私与安全性 → 辅助功能 / 自动化**：给终端（或 Claude 桌面 App）打开，界面操作（复制表格、坐标轴步数、移到最后面等）要用。
- Python 3：`pip3 install pillow python-pptx openpyxl matplotlib`。
- 字体：思源黑体 Source Han Sans CN（全部字重）。
- Node（Wind 取数用）。
- 命令行工具：`pdfinfo`/`pdffonts`/`pdftoppm`（`brew install poppler`），PDF 自查用。

## 4. 账号和凭证（各人自己申请，**不要提交进仓库**）

| 用途 | 怎么配 |
|---|---|
| 飞书（读大纲、写过程记录/数据提需/数据复核） | `npm i -g @larksuite/cli`，`lark-cli auth login` 用自己的飞书账号；需能访问「每周直播大纲」文件夹和下面列的文档 |
| 直播小助理机器人发群消息 | lark-cli 档案名 `live-poster-bot`（App ID `cli_a94b5144b1381cb3`），找范丽丽或管理员配置；没有时先只在聊天里给出话术，由本人手动发 |
| 螺丝钉数据工具箱 https://dbms.yhlsd.com/ | 在工具箱「API Client → 管理凭证」生成自己的 Token，存 `~/.config/dbms/token`（`chmod 600`），90 天有效；用法见 `rules/数据工具箱dbms.md` |
| Wind（可选） | `wind-mcp-skill` 里 `node scripts/cli.mjs setup-key <KEY>`；积分不足时按 `rules/复核记录与交付.md`「Wind 不可用时」用工具箱复核 |
| iCloud Keynote 在线文档 | 本人 Apple ID 登录 iCloud，上传位置 `~/Library/Mobile Documents/com~apple~Keynote/Documents/` |

## 5. 用到的飞书文档和群（规则里都写了，这里汇总）

- 过程记录：「PPT制作过程记录」 `YMHnda1X8oVqfyxdRTRcoiSWnuh`
- 人工数据需求：「PPT数据提需」 `VST4d5u30oeY5NxwhZtcAKCUnId`
- 给老板看的：「PPT数据复核」 wiki `NN4uwNq6riHMxvkNDGxcz86HnNg`
- 内容|PPT制作群（发汤爱学）`oc_918c9be8ab6950e746bc308c8c32a334`；226 课程群（发王磊）`oc_ad278b7a15d31ab7a5ced19569769db8`
- 历史成品库、大纲文件夹：见 `rules/大纲与历史成品库.md`

## 6. 只在原作者电脑上有的东西（同事电脑上怎么办）

- 图片素材库 `~/Desktop/范丽丽./图片素材库/library.db`：没有就跳过，直接下载来源期 PDF 核对页码（`rules/多期拼页.md` 1.5）。
- `history_keys/` 历史期 .key 缓存：没有就按 `rules/大纲与历史成品库.md` 从飞书下载。
- 每期产出放 `<项目目录>/output/NNN_claude输出/`，定稿后整夹移废纸篓（`rules/复核记录与交付.md`）。

## 7. 改了规则怎么同步

本机改完 → 在仓库开分支提交 → 提 PR 合并进 main，让大家拿到同一份规则：
```bash
cd ~/GitHub/live-ppt-builder && git checkout -b 规则-YYYYMMDD && git add -A && git commit -m "…" && git push -u origin HEAD
gh pr create --fill
```
项目里如果另有一份 skill 副本（如 `.agents/skills/live-ppt-builder`），合并后和仓库逐文件比对一致再收工。
