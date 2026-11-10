# -*- coding: utf-8 -*-
"""453期：复盘A股过去6轮牛熊市。

离线生成版：Wind 接口不可用时，使用本地已生成/下载的453期图表素材，
并严格保留原大纲里的标题和正文要点。
"""
import sys
from pathlib import Path

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E
from live_helpers import summary

from pptx import Presentation


IMG = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_453")
OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/453期直播PPT_根据解析稿生成.pptx"


def img(name):
    return str(IMG / name)


def img_page(prs, body, title, image_name, points=None, pending=False):
    path = IMG / image_name
    if path.exists():
        return E.img_page(prs, body, title, str(path), points=points, pending=pending)
    return E.img_page(prs, body, title, img("img01.png"), points=["参考图缺失，待补充。"], pending=True)


def build():
    prs = Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")

    E.cover_page(prs, cover, "复盘A股过去6轮牛熊市：\n牛熊市出现，是什么原因？")

    img_page(prs, body, "过去20年，A股经历的6轮牛熊市", "img02.png",
             points=["449期第8页，最后一行的日期、天数、涨跌幅数据待更新。",
                     "这里的涨跌幅口径是最大涨跌幅，不是累计涨跌幅。"],
             pending=True)

    img_page(prs, body, "第一轮牛市：2005-2007年，A股历史最大牛市", "gen_r1bull.png",
             points=["2005-2007年，期间A股最大涨幅606.76%。",
                     "这一轮A股历史上的最大牛市，是多重因素共同推动的，如股权分置改革全面推行、国内经济高速增长、A股估值较低等。",
                     "这一轮上涨之后，A股到了1星级泡沫阶段，也透支了之后未来若干年的涨幅。"])

    img_page(prs, body, "第一轮熊市：2007-2008年", "gen_r1bear.png",
             points=["2008年，次贷危机爆发，加上A股当时整体处于1星级泡沫阶段，股市开始暴跌。",
                     "中证全指从最高5521.69点，下跌到1574.76点，跌去71.48%。"])

    img_page(prs, body, "第二轮牛市：2008-2009年", "gen_r2bull.png",
             points=["2009年，在4万亿计划刺激下，A股出现一波小牛市。中证全指上涨了142.42%。"])

    img_page(prs, body, "第二轮熊市：2009-2014年", "gen_r2bear.png",
             points=["这一轮熊市持续的时间比较长，和流动性收紧、上市公司基本面低迷有关系。",
                     "从2011-2012年，央行连续多次加息、提高存款准备金率，市场上的钱少了，股价自然下跌。同时2013年，A股上市公司盈利同比下降，基本面也比较低迷，推动出现了A股历史上最大的一轮熊市。"])

    img_page(prs, body, "第三轮牛市：2014-2015年", "gen_r3bull.png",
             points=["2014年下半年，在密集降息刺激下，A股开始迅速上涨。",
                     "利率下降+杠杆投资，共同催生出2015年上半年小盘股牛市。",
                     "中证全指从2734.65点，上涨到7984.43点，涨了191.97%。"])

    img_page(prs, body, "第三轮熊市：2015-2016年", "gen_r3bear.png",
             points=["2015年6月份，小盘股进入泡沫阶段，之后A股迎来连续大跌。",
                     "中证全指下跌了49.74%。"])

    img_page(prs, body, "第四轮牛市：2016-2017年", "gen_r4bull.png",
             points=["2016-2017年，大盘价值风格涨幅较好，当时银行保险地产等行业表现普遍不错。",
                     "2017年底，在大盘价值股的带领下，大盘再次涨到3星级。中证全指上涨了28.07%。"])

    img_page(prs, body, "第四轮熊市：2018年", "gen_r4bear.png",
             points=["2018年初，价值股也涨到较高水平，之后在贸易摩擦影响下，A股大跌。",
                     "到2018年底回到5星级，期间最大跌幅34.52%。"])

    img_page(prs, body, "第五轮牛市：2019-2021年", "gen_r5bull.png",
             points=["2019-2021年，A股整体反弹上涨。",
                     "2020年3月份因为疫情出现短期大跌，但之后修复。",
                     "到2020年底2021年初，市场达到3星级。期间中证全指上涨了76.47%。"])

    img_page(prs, body, "第五轮熊市：2021-2024年", "gen_r5bear.png",
             points=["这一轮熊市持续了三年多，主要原因是上市公司的盈利比较低迷，叠加美元加息流动性收紧。",
                     "走势图数据调整为2024年9月18日。"])

    img_page(prs, body, "第六轮牛市：2024年至今", "gen_r6bull.png",
             points=["这一轮牛市，受益于上市公司盈利情况的好转，以及美元降息。",
                     "展示中证全指2024年9月以来的走势，并将最低点和本轮最高点打红点。",
                     "最新点位和最大涨幅需与第一页表格最后一行保持一致。"],
             pending=True)

    img_page(prs, body, "（根据反馈新增）每次熊市出现的原因各不相同", "img14.png")

    E.box_page(prs, body, "（根据反馈新增）影响市场涨跌的三个周期",
               [("资金面", "受到利率的影响"),
                ("基本面", "取决于上市公司盈利情况"),
                ("情绪面", "投资者的情绪变化")],
               bubble="当其中1-2种处于熊市、牛市周期，也就带来了股票市场的涨涨跌跌。")

    img_page(prs, body, "市场涨涨跌跌，长期还是向上走的", "gen_longup.png",
             points=["346期第3页，数据和走势图待更新。",
                     "从2004年以来截至最新，中证全指以及全收益的点数、年化收益率、涨跌幅需要同步更新。",
                     "走势图右侧红/绿区间和红点位置需跟随最新数据调整。"],
             pending=True)

    img_page(prs, body, "市场底部在不断上涨", "gen_bottomup.png",
             points=["指数点数 = 估值 × 盈利。",
                     "上市公司的盈利长期上涨，是推动市场向上的根本动力。",
                     "452期第9页，数据待更新。"],
             pending=True)

    summary(prs, body, "总结", [
        "从过去20年的情况看，推动每次牛熊市出现的原因、涨跌幅、持续时间都不完全一样。",
        "我们没办法根据历史去预测未来。",
        "不过，牛熊总是交替出现的。牛市不必过于兴奋，熊市也不必过于恐慌。",
        "用好螺丝钉星级及指数估值表，做好「低估买入-正常估值持有-高估卖出」，好收益大概率就是水到渠成的。",
    ])

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"完成：{len(prs.slides)}页 -> {OUT}")


if __name__ == "__main__":
    build()
