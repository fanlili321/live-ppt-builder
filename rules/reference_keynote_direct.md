---
name: reference-keynote-direct
description: 直播演示稿直接生成 Keynote 的主流程，不经过 PPTX
---

# Keynote 直出流程

默认只生成 `.key`。不要生成 PPTX，不要把 PPTX 当中间产物。

## 模板

技能内置的 `assets/指数模版.key`

该文件是 zip 格式 Keynote 包，包含 `Index/*.iwa`、`Data/*`、`Metadata/*` 和预览图，适合作为真实 Keynote 模板源。

## 推荐步骤

1. 复制模板为本期 `.key`：
   `python3 scripts/keynote_direct.py create --out ~/Downloads/第XXX期直播.key`
2. 打开本期 `.key`：
   `python3 scripts/keynote_direct.py open ~/Downloads/第XXX期直播.key`
3. 用模板页复用版式：
   `python3 scripts/keynote_direct.py duplicate-slide ~/Downloads/第XXX期直播.key --source 1`
4. 填标题、正文、来源：
   `python3 scripts/keynote_direct.py add-text ~/Downloads/第XXX期直播.key --slide 1 --text "标题" --x 80 --y 80 --w 1100 --h 80`
5. 添加图片或渲染后的走势图：
   `python3 scripts/keynote_direct.py add-image ~/Downloads/第XXX期直播.key --slide 1 --image chart.png --x 80 --y 190 --w 1100 --h 520`
6. 导出 PDF 验稿：
   `python3 scripts/keynote_direct.py export-pdf ~/Downloads/第XXX期直播.key --out /tmp/第XXX期直播.pdf`

## 约束

- `.key` 是最终交付文件。
- PDF 只用于视觉验稿。
- PPTX 只在用户明确要求兼容 PowerPoint 时才生成。
- 图表如果 Keynote 原生图表自动化不足，先用 matplotlib/数据脚本生成高分辨率 PNG 放入 Keynote；不要因此回退到 PPTX。
