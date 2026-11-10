---
name: reference-feishu-cli
description: 读取飞书文档/表格的稳定通道——用 lark-cli 命令行，不要用浏览器扩展
metadata: 
  node_type: memory
  type: reference
  originSessionId: 0916e9f5-463b-48e5-94e4-b00430082a93
---

读取用户的飞书文档/表格，**用本地 `lark-cli`**（路径 `/Users/fanlili/.npm-global/bin/lark-cli`，配置在 `~/.lark-cli/config.json`），这是稳定通道，不依赖 Claude for Chrome 浏览器扩展（扩展会频繁掉线，user 用的浏览器也不一定装了扩展）。

调用前设 `export PATH="/Users/fanlili/.npm-global/bin:$PATH" LARK_CLI_NO_PROXY=1`（避免代理警告/串改）。

常用：
- 文档纯文本：`lark-cli api GET /open-apis/docx/v1/documents/{doc_token}/raw_content --jq '.data.content'`
- 文档全部 block（含图片 token，block_type=27 为图片，字段 image.token）：`lark-cli api GET /open-apis/docx/v1/documents/{doc_token}/blocks --params '{"page_size":500}' --page-all`
- 下载图片/附件：`lark-cli api GET /open-apis/drive/v1/medias/{file_token}/download -o 输出路径`

doc_token 从飞书链接 `/docx/XXXX` 里取 XXXX。
以后 user 发飞书链接，直接用这个读，不要再去试浏览器扩展。
