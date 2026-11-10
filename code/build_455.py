# -*- coding: utf-8 -*-
"""455期：黄金大幅波动，当前投资机会如何？"""
import sys
from pathlib import Path

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E
from live_helpers import section, summary

from pptx import Presentation


IMG = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_455")
OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/455期直播PPT_根据解析稿生成.pptx"


def p(n):
    return str(IMG / f"v{n:02d}.png")


def img_page(prs, body, title, n, points=None, pending=False):
    path = Path(p(n))
    if path.exists():
        return E.img_page(prs, body, title, str(path), points=points, pending=pending)
    return E.img_page(prs, body, title, p(4), points=["参考图缺失，待补充。"], pending=True)


def build():
    prs = Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")

    E.cover_page(prs, cover, "黄金大幅波动，\n当前投资机会如何？")

    img_page(prs, body, "黄金最近波动较大", 3,
             points=["展示今年以来伦敦金叠加上海金走势，并分别标注最大回撤及时间。"],
             pending=True)
    img_page(prs, body, "过去200多年里，黄金略微跑赢通货膨胀", 4)
    img_page(prs, body, "1971年之后，黄金收益率大幅提升", 5,
             points=["参考415期第3页，数据待更新。"], pending=True)
    img_page(prs, body, "黄金收益提升的原因：从金本位到纸币本位", 6)
    img_page(prs, body, "黄金收益提升的原因：从金本位到纸币本位", 7)
    img_page(prs, body, "美联储废除金本位（1971年）后，黄金经历了三轮牛熊市", 8,
             points=["参考415期第5页，数据待更新。"], pending=True)
    img_page(prs, body, "第一轮牛熊市：1971年-2000年", 9)
    img_page(prs, body, "第二轮牛熊市：2001年-2016年", 10)
    img_page(prs, body, "第三轮牛熊市：2017年至今", 11,
             points=["参考415期第9页，数据待更新。"], pending=True)
    img_page(prs, body, "黄金的波动风险", 12,
             points=["参考415期第10页，数据待更新。"], pending=True)
    img_page(prs, body, "影响黄金走势的最大因素：美元", 13,
             points=["参考415期第11页，数据待更新。"], pending=True)
    img_page(prs, body, "影响黄金走势的最大因素：美元", 14,
             points=["参考415期第12页，数据待更新。"], pending=True)
    img_page(prs, body, "影响黄金走势的其他因素：开采成本", 15,
             points=["参考5月30号高净值线下活动第22页，数据待更新。",
                     "开采成本需同时显示“美元/盎司”和“元/克”；当前美元/盎司后新增元/克。",
                     "气泡：通常情况下，黄金价格在开采成本的1倍以内，属于相对低估。"],
             pending=True)
    img_page(prs, body, "影响黄金走势的其他因素：地区冲突、金融危机等风险事件", 16,
             points=["参考415期第14页，数据待更新。"], pending=True)
    img_page(prs, body, "黄金估值如何判断？", 17,
             points=["参考415期第15页，数据待更新。"], pending=True)
    img_page(prs, body, "黄金估值如何判断？", 18,
             points=["参考415期第16页，数据待更新。"], pending=True)
    img_page(prs, body, "螺丝钉黄金牛熊信号板", 19,
             points=["参考415期第17页，图片替换为最新。"], pending=True)
    img_page(prs, body, "黄金基金", 20,
             points=["参考415期第20页，数据待更新。",
                     "右侧图里的ETF换为富国黄金ETF 518680，管理费率0.15%。"],
             pending=True)
    section(prs, body, "常见疑问解答")
    img_page(prs, body, "Q1：近期波动后，黄金当前是否适合投资？黄金历史上什么时候出现过4-5星级？", 21,
             points=["参考415期第24页，数据待更新。",
                     "将当前文字“在1星级附近”删除。",
                     "走势图上用绿色圈圈出黄金几次4-5星位置：2015-2016年、2020年初、2022年中。"],
             pending=True)
    img_page(prs, body, "Q2：黄金为何波动变大？", 22,
             points=["参考441期第10页，数据待更新。"], pending=True)
    summary(prs, body, "总结", [
        "过去几十年黄金一共经历三轮牛熊市。",
        "最近一轮牛市从2017年至今，最高涨幅达到X（数据待提供）。",
        "虽然近期波动，当前估值并不便宜。",
        "可参考螺丝钉黄金星级及牛熊信号板判断估值。",
        "买入习惯低估买入，之后出现低估会介绍。",
    ])

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"完成：{len(prs.slides)}页 -> {OUT}")


if __name__ == "__main__":
    build()
