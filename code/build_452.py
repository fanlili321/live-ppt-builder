# -*- coding: utf-8 -*-
"""452期：螺丝钉星级是什么，投资的时候如何用？

按已确认大纲 + 解析稿生成。大纲图已下载到 imgs_452。
数据待更新页先用大纲图/占位承接；组合净值页用本地投顾净值做原生图。
"""
import sys
from pathlib import Path

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E
from live_helpers import section, ph_note, summary, combo, pct, chart_pg

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


IMG = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_452")
IMG450 = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_450")
OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/452期直播PPT_根据解析稿生成.pptx"

ORG = RGBColor(0xF1, 0x76, 0x40)
PALE = RGBColor(0xFC, 0xE6, 0xD2)
BLUE = RGBColor(0x7C, 0xA7, 0xE0)
GREEN = RGBColor(0x62, 0xB5, 0x72)


def p(n):
    return str(IMG / f"v{n}.png")


def clear(sl):
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)


def newslide(prs, body, title):
    sl = prs.slides.add_slide(body)
    clear(sl)
    t_pt = E.title_fit(title)
    E.add_tb(sl, title, E.CONTENT_L, E.CONTENT_TOP, E.CONTENT_W,
             E.text_height_cm(title, t_pt, E.CONTENT_W) + 0.2,
             size=t_pt, bold=True, color=E.C_BLACK, align=PP_ALIGN.CENTER)
    return sl


def img_page(prs, body, title, path, points=None, pending=True):
    if path and Path(path).exists():
        return E.img_page(prs, body, title, str(path), points=points, pending=pending)
    return ph_note(prs, body, title, "参考图/数据待补充；已按解析稿先占位。")


def strategy_relation(prs, body):
    title = "螺丝钉星级，跟哪些指数/策略比较相关？"
    active = combo("螺丝钉金钉宝主动优选C", 2, "2024-09-18")
    enhanced = combo("螺丝钉金钉宝指数增强C", 2, "2024-09-18")
    if active and enhanced:
        chart_pg(
            prs, body, title, "（3）指数增强策略（对应中证800）、主动优选策略，也跟螺丝钉星级有比较强关系。",
            [pct("主动优选C", active), pct("指数增强C", enhanced)],
            "左侧和右侧分别展示成立以来主动优选、指数增强策略叠加星级走势图；星级数据待补充后叠加。"
        )
    else:
        ph_note(prs, body, title, "主动优选/指数增强净值或星级历史待补充。")


def current_strategy(prs, body):
    title = "当下适合投资什么？"
    sl = newslide(prs, body, title)
    y = E.CONTENT_TOP + E.text_height_cm(title, E.title_fit(title), E.CONTENT_W) + 1.2
    rows = [
        ("5-5.9星级", "增加股票基金策略的投资金额", "主动优选、指数增强"),
        ("4-4.9星级", "股票基金金额比5星级少", "同时配置月薪宝、全天候等稳健策略"),
        ("3-3.9星级", "出现止盈机会", "再转化为月薪宝、365天、全天候等稳健策略"),
    ]
    for i, (star, action, target) in enumerate(rows):
        top = y + i * 3.0
        box = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(E.CONTENT_L + 2), Cm(top), Cm(E.CONTENT_W - 4), Cm(2.2))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0xFF, 0xF7, 0xEF) if i != 2 else RGBColor(0xF1, 0xF5, 0xFF)
        box.line.color.rgb = ORG if i != 2 else BLUE
        box.line.width = Pt(1.5)
        box.shadow.inherit = False
        E.add_tb(sl, star, E.CONTENT_L + 2.5, top + 0.45, 5.5, 1.1,
                 size=20, bold=True, color=ORG if i != 2 else BLUE, align=PP_ALIGN.LEFT)
        E.add_tb(sl, action, E.CONTENT_L + 8.3, top + 0.45, 9.5, 1.1,
                 size=18, color=E.C_DARK, align=PP_ALIGN.LEFT)
        E.add_tb(sl, target, E.CONTENT_L + 18.0, top + 0.45, 11.0, 1.1,
                 size=18, color=E.C_DARK, align=PP_ALIGN.LEFT)
    return sl


def build():
    prs = Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")

    E.cover_page(prs, cover, "螺丝钉星级是什么，\n投资的时候如何用？")

    img_page(prs, body, "螺丝钉星级：判断市场整体估值", p(4),
             points=["参考439期第13页，右侧图片待更新；标题改成如上，删除估值表那句话。"])
    img_page(prs, body, "2012年以来，A股的牛熊市和星级的对应关系", p(5),
             points=["删除图例和批注；绿色圆圈标注2012-2014、2018-2019、2022-2024熊市。",
                     "红色圆圈标注2014-2015、2016-2017、2020-2021、2024年至今牛市。"])
    img_page(prs, body, "如何查看当前星级及2012年以来的历史星级", p(6),
             points=["参考444期第16页，截图待更新；删除黑体文字。"])
    img_page(prs, body, "（根据反馈修改）螺丝钉星级，跟哪些指数/策略比较相关？", p(10),
             points=["（1）星级跟中证全指、万得全A相关性最强。",
                     "展示2012年初以来，中证全指及全收益、万得全A及全收益，叠加星级走势。"])
    img_page(prs, body, "（根据反馈修改）螺丝钉星级，跟哪些指数/策略比较相关？", p(7),
             points=["（2）沪深300、中证800、中证A500、中证偏股基金指数，跟螺丝钉星级也相关性比较强。",
                     "展示2012年初以来，沪深300全收益、中证800全收益、中证A500全收益与星级走势；中证偏股基金是否纳入待确认。"])
    strategy_relation(prs, body)
    img_page(prs, body, "同样的星级下，收益一样吗？", p(17),
             points=["参考444期第3页，数据待更新，并补充中证全指全收益走势图。"])
    img_page(prs, body, "（根据反馈新增）指数基金的收益来源", p(11),
             points=["参考450期第7页，删除“宽基”两个字。"])
    img_page(prs, body, "（根据反馈新增）收益来源一：盈利增长", p(12),
             points=["参考446期第8页，走势图数据待更新。"])
    img_page(prs, body, "（根据反馈新增）收益来源二：估值提升", p(13),
             points=["参考450期第9页，改为中证全指数据；两条虚线标注位置待确认。"])
    img_page(prs, body, "（根据反馈新增）收益来源三：股息率（分红收益）", p(14),
             points=["参考450期第13页，改为中证全指数据。"])
    img_page(prs, body, "（根据反馈新增）如何结合星级，投资不同策略？", p(15),
             points=["参考5.30高净值PPT第48页。"], pending=False)
    img_page(prs, body, "（根据反馈新增）不同策略的收益&波动风险对比", p(16),
             points=["参考5.30高净值PPT第49页，数据待更新。"])
    img_page(prs, body, "（根据反馈新增）过去10年，出现的投资&止盈机会", p(17),
             points=["参考5.30高净值PPT第50页，数据待更新。"])
    img_page(prs, body, "（根据反馈新增）回测数据不同星级买入，持有24个月，对应的收益和风险", p(18),
             pending=False)
    img_page(prs, body, "（根据反馈新增）回测数据", p(19),
             points=["参考5.30高净值PPT第53页。"], pending=False)
    current_strategy(prs, body)
    img_page(prs, body, "当下适合投资什么？", p(20),
             points=["参考5.30高净值PPT第56页。"], pending=False)
    summary(prs, body, "总结", [
        "螺丝钉星级，5星级代表股市估值较低，投资价值较高。1星级代表股市存在泡沫，投资价值较低。",
        "5-5.9星级，增加股票基金策略（如主动优选、指数增强）的投资金额。",
        "4-4.9星级，股票基金策略的投资金额比5星级少，同时配置月薪宝、全天候等稳健策略。",
        "3-3.9星级，出现止盈机会，再转化为稳健策略。",
    ])

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"完成：{len(prs.slides)}页 -> {OUT}")


if __name__ == "__main__":
    build()
