# -*- coding: utf-8 -*-
"""445期：从飞书大纲图片直接生成PPT。

用于落实“只有图就直接放图”的规则。大纲中的删除页不制作，常见疑问解答做过渡页。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E
import live_helpers as hl

OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/445期直播PPT.pptx"
FETCH = Path("/Users/fanlili/Downloads/live-ppt-tool/current/445_fetch.json")
IMG_DIR = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_445_outline")
LARK = "/Users/fanlili/.npm-global/bin/lark-cli"


def clean_title(text):
    text = re.sub(r"<text[^>]*>(.*?)</text>", r"\1", text, flags=re.S)
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = text.replace("~~", "")
    text = re.sub(r"\s+", " ", text).strip()
    for prefix in ("（新增）",):
        if text.startswith(prefix):
            text = text[len(prefix):].strip()
    return text


def sections():
    md = json.loads(FETCH.read_text(encoding="utf-8"))["data"]["markdown"]
    md = re.split(r"<text[^>]*>【标题及直播海报宣传文案】", md)[0]
    parts = re.split(r"^##\s+", md, flags=re.M)
    out = []
    for part in parts[1:]:
        lines = part.splitlines()
        raw_title = lines[0].strip()
        body = "\n".join(lines[1:])
        title = clean_title(raw_title)
        if "本页删除" in raw_title:
            continue
        token_match = re.search(r'<image[^>]*token="([^"]+)"', body)
        token = token_match.group(1) if token_match else None
        note = ""
        if "2024年改为2025年" in body or "2024改为2025" in body:
            note = "大纲要求：2024改为2025。"
        if "2024年改为去年" in body:
            note = "大纲要求：2024年改为去年。"
        if "图片待更新为最新" in body:
            note = "大纲要求：图片待更新为公众号最新图。"
        out.append({"title": title, "token": token, "note": note})
    return out


def download(token):
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    path = IMG_DIR / f"{token}.png"
    if path.exists() and path.stat().st_size > 0:
        return path
    subprocess.run(
        [LARK, "docs", "+media-download", "--token", token, "--output", f"./{token}.png", "--overwrite"],
        check=True,
        text=True,
        capture_output=True,
        cwd=str(IMG_DIR),
    )
    return path


def build():
    prs = E.Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")
    hl.cover(prs, cover, "必学：个税汇算清缴开始了，\n如何操作享优惠？")

    for item in sections():
        title = item["title"]
        if title.startswith("常见疑问解答"):
            hl.section(prs, body, "常见疑问解答")
            continue
        if title == "总结":
            hl.summary(prs, body, "总结", [
                "2025年度的个税汇算清缴开始了，如果有房租/房贷、子女教育等支出，要记得申报，说不定能省一笔钱。",
                "近些年比较热门的个人养老金，也属于税前扣除项目之一，2025年交过个人养老金的朋友，在本次汇算清缴中也可以享受税收优惠。",
                "如果对个人养老金感兴趣，可以读一读螺丝钉的新书《个人养老金投资指南》。",
            ])
            continue
        if item["token"]:
            img = download(item["token"])
            points = [item["note"]] if item["note"] else None
            pending = "待更新" in item["note"]
            E.img_page(prs, body, title, str(img), points=points, pending=pending)
        else:
            hl.ph_note(prs, body, title, "大纲未提供图片，待补充素材或原生重做。")

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"✅ 445完成！{len(prs.slides)} 张 → {OUT}")


if __name__ == "__main__":
    build()
