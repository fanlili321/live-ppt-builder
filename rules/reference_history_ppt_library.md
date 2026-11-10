---
name: reference-history-ppt-library
description: 历史直播成品PPT库的默认检索范围（多个飞书Drive文件夹token）。用户说"历史直播ppt"时默认只在这些文件夹里找
metadata:
  node_type: memory
  type: reference
---

用户确认（2026-08-06）：以后说"历史直播PPT" / "历史成品库"，默认检索范围 = 以下几个飞书 Drive 文件夹（不含子文件夹以外的其它猜测位置）。**不包含**"直播ppt生成（AI）"文件夹（token=`QXmuf05g0l31abdim1LcUYpenRh`）——那个是AI生成产物的存放处，不是历史成品参考库。

## 默认检索范围

| 文件夹名 | token | 归属 | 说明 |
|---|---|---|---|
| 已直播ppt | `fldcn82dxPE6sznG1xWP2OvcLvc` | 范丽丽（本人） | **主库**，437个子文件夹，每期一个，覆盖第9期～第466期+，最常用、最权威 |
| 02-私募直播ppt | `fldcnFMmJZrDXZnmo8jVcekNSec` | 汤爱学 | 私募向直播的历史成品 |
| 02-03公募-直播PPT | `fldcnXKPrVanO164gH0GKEbY2Qg` | 李菁（杜珊珊编辑） | 公募向直播的历史成品 |
| 02-03公募-钉大对外演讲PPT | `DU88fYBMblNlNmdah5lcHpl8nAc` | 杜珊珊（范丽丽编辑） | 钉大对外演讲类PPT历史成品 |

## 用法
- 大纲/用户提到"某期第X页"、"复用XX期的设计"时，先在**已直播ppt**里按期数关键词（如"第456期"）找对应子文件夹，找不到再查其余两个。
- 找到对应期数文件夹后，进去找该期的原始 PPT/Keynote/PDF，下载抽取指定页，禁止用大纲里的缩略截图代替（见 [[reference_ppt_human_template]] 和 SKILL.md 的 `HISTORICAL_FULL_PAGE` 规则）。
- 若用户明确说要找私募或公募专场的历史稿，直接对应到第2、3个文件夹，不用逐个翻。
