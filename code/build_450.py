# -*- coding: utf-8 -*-
"""450期：按已确认解析稿+飞书批注生成直播PPT。

批注落实：
- “（1）中证800”“（2）中证全指”作为页内小标题加粗。
- “这是新组合，不是指数增强...”是给制作人的提醒，不进入PPT。
- 缺数据/缺最新版截图的页面先占位。
"""
import sys
from pathlib import Path
import openpyxl

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


IMG = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_450")
COMBO = Path("/Users/fanlili/Downloads/投顾组合净值数据_20200101_20260615.xlsx")
OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/450期直播PPT_批注确认版.pptx"
END = "2026-06-12"


def img(n: int) -> str:
    return str(IMG / f"v{n}.png")


def clear_placeholders(slide):
    for ph in list(slide.placeholders):
        ph.element.getparent().remove(ph.element)


def title_slide(prs, layout, title):
    sl = prs.slides.add_slide(layout)
    clear_placeholders(sl)
    t_pt = E.title_fit(title)
    t_h = E.text_height_cm(title, t_pt, E.CONTENT_W) + 0.2
    E.add_tb(sl, title, E.CONTENT_L, E.CONTENT_TOP, E.CONTENT_W, t_h,
             size=t_pt, bold=True, color=E.C_BLACK, align=PP_ALIGN.CENTER)
    return sl


def source(sl, text="数据来源：Wind"):
    E.add_tb(sl, text, E.CONTENT_L + E.CONTENT_W - 13.0, E.CONTENT_BOTTOM - 0.55,
             13.0, 0.45, size=11, color=E.C_GRAY, align=PP_ALIGN.RIGHT)


def placeholder(prs, layout, title, note, body=None):
    sl = title_slide(prs, layout, title)
    y = E.CONTENT_TOP + E.text_height_cm(title, E.title_fit(title), E.CONTENT_W) + 1.0
    if body:
        y = E.add_arrow_points(sl, body, y, pt=21) + 0.2
    box = sl.shapes.add_shape(
        E.MSO_SHAPE.ROUNDED_RECTANGLE, Cm(E.CONTENT_L + 2.0), Cm(max(y, 7.0)),
        Cm(E.CONTENT_W - 4.0), Cm(4.0)
    )
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFB, 0xEF, 0xC5)
    box.line.color.rgb = E.C_MAIN
    box.line.width = Pt(2)
    box.shadow.inherit = False
    E.add_tb(sl, note, E.CONTENT_L + 3.0, max(y, 7.0) + 0.7, E.CONTENT_W - 6.0, 2.6,
             size=22, color=E.C_DARK, align=PP_ALIGN.CENTER)
    return sl


def clean_section(prs, layout, title):
    sl = prs.slides.add_slide(layout)
    clear_placeholders(sl)
    # 过渡页不用母版占位符，避免“单击此处”等提示文字残留。
    band = sl.shapes.add_shape(E.MSO_SHAPE.RECTANGLE, Cm(0), Cm(0), Cm(33.87), Cm(4.2))
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor(0xE9, 0x7C, 0x2D)
    band.line.fill.background()
    card = sl.shapes.add_shape(E.MSO_SHAPE.ROUNDED_RECTANGLE, Cm(3.2), Cm(6.6), Cm(27.4), Cm(5.2))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    card.line.color.rgb = RGBColor(0xEE, 0xDD, 0xCC)
    card.shadow.inherit = False
    E.add_tb(sl, title, 4.5, 8.0, 24.8, 2.0, size=40, bold=True,
             color=E.C_BLACK, align=PP_ALIGN.LEFT)
    return sl


def table_page(prs, layout, title, headers, rows, src=None):
    sl = title_slide(prs, layout, title)
    t_h = E.text_height_cm(title, E.title_fit(title), E.CONTENT_W) + 0.2
    top = E.CONTENT_TOP + t_h + 0.8
    height = E.CONTENT_BOTTOM - top - 0.7
    tbl = sl.shapes.add_table(len(rows) + 1, len(headers),
                              Cm(E.CONTENT_L), Cm(top), Cm(E.CONTENT_W), Cm(height)).table
    widths = [4.0, 4.0, 12.0, 6.0, 6.0]
    total = sum(widths)
    for c, w in enumerate(widths):
        tbl.columns[c].width = Cm(w / total * E.CONTENT_W)
    for r in range(len(rows) + 1):
        tbl.rows[r].height = Cm(height / (len(rows) + 1))

    def put(r, c, text, bold=False, fill=None, color=E.C_DARK, align=PP_ALIGN.CENTER):
        cell = tbl.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill or RGBColor(0xFF, 0xFF, 0xFF)
        cell.margin_top = Cm(0.04)
        cell.margin_bottom = Cm(0.04)
        cell.margin_left = Cm(0.08)
        cell.margin_right = Cm(0.08)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.size = Pt(17)
        run.font.bold = bold
        run.font.color.rgb = color
        E.set_run_font(run)

    for c, h in enumerate(headers):
        put(0, c, h, True, E.C_MAIN, RGBColor(0xFF, 0xFF, 0xFF))
    for r, row in enumerate(rows, 1):
        bg = RGBColor(0xF5, 0xF5, 0xF5) if r % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
        for c, text in enumerate(row):
            put(r, c, text, c == 0, bg, align=PP_ALIGN.LEFT if c == 2 else PP_ALIGN.CENTER)
    if src:
        source(sl, src)
    return sl


def combo_series(name, col=2, start=None):
    if not COMBO.exists():
        return []
    wb = openpyxl.load_workbook(COMBO, read_only=True, data_only=True)
    ws = wb.active
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[1] != name or r[col] in (None, ""):
            continue
        d = str(r[0])[:10]
        if start and d < start:
            continue
        try:
            out.append({"date": d, "close": float(r[col])})
        except Exception:
            pass
    out.sort(key=lambda x: x["date"])
    return out


def to_pct_series(name, data):
    if not data:
        return None
    base = data[0]["close"]
    return {"name": name, "dates": [x["date"] for x in data],
            "values": [(x["close"] / base - 1) * 100 for x in data]}


def fetch(code, start, end=END, freq="W"):
    try:
        return E.fetch_kline(code, start, end, freq)
    except Exception as exc:
        print(f"取数失败 {code}: {exc}")
        return []


def chart_multi_or_placeholder(prs, layout, title, series, points=None, src="数据来源：Wind"):
    series = [s for s in series if s]
    if not series:
        return placeholder(prs, layout, title, "此页数据待补充，已按批注先占位。", points)
    sl = E.multi_chart_page(prs, layout, title, series, points=points, y_pct=True)
    # engine默认来源为Wind；内部组合页额外覆盖一行说明
    if src != "数据来源：Wind":
        source(sl, src)
    return sl


def add_bold_subtitle(sl, text, y):
    E.add_tb(sl, text, E.CONTENT_L, y, E.CONTENT_W, 0.8,
             size=23, bold=True, color=E.C_DARK, align=PP_ALIGN.LEFT)
    return y + 0.9


def p5_csi800(prs, layout):
    title = "投资宽基指数，选哪个更合适？"
    d1 = fetch("000906.CSI", "2004-12-31", END, "W")
    d2 = fetch("H00906.CSI", "2004-12-31", END, "W")
    if not d1 or not d2:
        sl = E.img_page(prs, layout, title, img(4), pending=True)
        y = E.CONTENT_TOP + E.text_height_cm(title, E.title_fit(title), E.CONTENT_W) + 0.4
        add_bold_subtitle(sl, "（1）中证800", y)
        return sl
    s1 = to_pct_series("中证800", d1)
    s2 = to_pct_series("中证800全收益", d2)
    sl = chart_multi_or_placeholder(
        prs, layout, title, [s1, s2],
        points=["中证800=沪深300+中证500",
                "在2021年，沪深300、中证500的基金数量比较多，中证1000基金数量很少，中证2000基金还没有。",
                "指数增强组合，以中证800指数为核心。"],
    )
    add_bold_subtitle(sl, "（1）中证800", E.CONTENT_TOP + 1.45)
    return sl


def p6_csi_all(prs, layout):
    title = "投资宽基指数，选哪个更合适？"
    combo = combo_series("螺丝钉金钉宝全市场指数C1", 2, "2024-02-28")
    bench = combo_series("螺丝钉金钉宝全市场指数C1", 3, "2024-02-28")
    index = fetch("000985.CSI", "2024-02-28", END, "D")
    series = [
        to_pct_series("金钉宝全市场指数C1", combo),
        to_pct_series("组合业绩基准", bench),
        to_pct_series("中证全指", index),
    ]
    sl = chart_multi_or_placeholder(
        prs, layout, title, series,
        points=["沪深300+中证500+中证1000+中证2000，约等于中证全指。",
                "2025年以来，首次出现了中证全指的指数基金。",
                "中证1000、中证2000基金逐渐丰富，也可以做中证全指的组合策略。"],
        src="数据来源：Wind、螺丝钉投顾组合净值"
    )
    add_bold_subtitle(sl, "（2）中证全指", E.CONTENT_TOP + 1.45)
    return sl


def build():
    prs = Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")

    # 1 封面
    E.cover_page(prs, cover, "定投宽基指数，该如何做：\n螺丝钉指数增强组合定投复盘")

    # 2-4 信息图/表格
    E.img_page(prs, body, "A股常见的四类指数", img(1))
    E.img_page(prs, body, "什么是宽基指数？", img(2),
               points=["右下方图新增：中证800：大中盘股代表；中证全指：市场整体代表。"])
    table_page(
        prs, body, "A股市场上比较有代表性的宽基指数",
        ["指数", "含义", "成分股", "市值平均数（亿）", "市值中位数（亿）"],
        [
            ["沪深300", "大盘股", "A股市场上规模最大的第1-300只股票", "待补", "待补"],
            ["中证500", "中盘股", "A股市场上规模最大的第301-800只股票", "待补", "待补"],
            ["中证800", "大中盘股", "A股市场上规模最大的第1-800只股票", "待补", "待补"],
            ["中证1000", "小盘股", "A股市场上规模最大的第801-1800只股票", "待补", "待补"],
            ["中证2000", "小微盘股", "A股市场上规模最大的第1801-3800只股票", "待补", "待补"],
            ["中证全指", "市场整体", "覆盖沪深300+中证500+中证1000+中证2000", "待补", "待补"],
        ],
        "数据来源：待当当提供，当前先占位"
    )

    # 5-6 原生图
    p5_csi800(prs, body)
    p6_csi_all(prs, body)

    # 7-14 参考/待更新页
    E.img_page(prs, body, "宽基指数基金的收益来源", img(3),
               points=["将原图“红利指数基金”改为“宽基指数基金的收益来源”。"])
    E.img_page(prs, body, "收益来源一：盈利增长", img(4), pending=True,
               points=["指数改为中证800以及全收益指数；样式不变，气泡文字按新数据调整。"])
    E.img_page(prs, body, "收益来源二：估值提升", img(5), pending=True,
               points=["更新中证800 PE数据；PE线条上标注15、19两条横虚线。"])
    E.img_page(prs, body, "投资指增组合，也遵循「低估买入、正常估值持有、高估止盈」", img(6))
    E.img_page(prs, body, "投资指增组合，也遵循「低估买入、正常估值持有、高估止盈」", img(7), pending=True,
               points=["参考449期第21页更新数据；最右侧改为黄色“持有”。"])
    E.img_page(prs, body, "「低估买入、正常估值持有、高估止盈」的结果", img(8), pending=True,
               points=["来自组合季报，数据待更新；按批注先占位。"])
    placeholder(prs, body, "「低估买入、正常估值持有、高估止盈」的结果",
                "此页找胡亮确认是否需要更新；按批注先占位。")
    E.img_page(prs, body, "收益来源三：股息率（分红收益）", img(10), pending=True,
               points=["将参考页指数改成中证800，样式不变。"])

    # 15 过渡
    clean_section(prs, body, "常见疑问解答")

    # 16-23 问答页
    E.img_page(prs, body, "Q1：宽基指数和星级是啥关系呢？", img(9), pending=True,
               points=["星级代表市场整体估值。",
                       "跟星级关系最紧密的是全市场指数，也就是中证全指。",
                       "沪深300、中证800覆盖A股大多数市值，也跟星级关系密切。"])

    active = combo_series("螺丝钉金钉宝主动优选C", 2, "2024-09-18")
    enhanced = combo_series("螺丝钉金钉宝指数增强C", 2, "2024-09-18")
    chart_multi_or_placeholder(
        prs, body, "Q2：指数增强组合和主动优选组合，哪个收益高？",
        [to_pct_series("主动优选C", active), to_pct_series("指数增强C", enhanced)],
        points=["这两个收益来源不同，涨跌强势阶段也有区别。"],
        src="数据来源：螺丝钉投顾组合净值"
    )

    E.img_page(prs, body, "Q3：什么时候会打开股票基金组合的申购？", img(11), pending=True,
               points=["当前市场在3点几星级，主动优选和指数增强组合还是关闭申购，保证投资者不会追涨买入。",
                       "等之后市场下跌，回到4.1星级后，主动优选和指数增强会打开申购。",
                       "此图主要体现关闭申购、帮助逆市投资，不体现止盈；只更新走势图数据和背景颜色，气泡批注不改。"])
    E.img_page(prs, body, "Q4：股票基金组合，如何止盈？", img(12))
    E.img_page(prs, body, "Q4：股票基金组合，如何止盈？", img(34) if (IMG / "v34.png").exists() else img(12),
               points=["（1）调仓止盈"])
    E.img_page(prs, body, "Q4：股票基金组合，如何止盈？", img(14),
               points=["（2）高估止盈",
                       "以指数增强组合为例，适合参考中证800的估值。",
                       "15倍市盈率以下进入低估；19倍市盈率以上进入高估。"])
    E.img_page(prs, body, "Q5：止盈后出来的钱，适合投资哪些品种？", img(15), pending=True,
               points=["参考449期第22页，数据待更新；将序号（2）删除，只删除序号，后面文字保留。"])
    E.img_page(prs, body, "Q6：未来什么时候转回到股票基金组合？", img(16),
               points=["3点几星级，比较适合考虑相对稳健的品种，比如月薪宝、365天、全天候等。",
                       "之后市场下跌回到4-5星级，可以考虑转回指数、主动等股票基金组合。"])

    # 24 总结
    E.arrows_page(prs, body, "总结", [
        "定投宽基指数基金，可以考虑中证800、中证全指等比较有代表性的指数。",
        "宽基指数基金的收益来源，主要有三个：盈利增长、估值提升、股息率（分红收益）。",
        "在投资时，可以参考估值，低估买入、正常估值持有、高估卖出。",
        "螺丝钉对指数增强组合的投资，也是用的这个思路。",
    ], pt=25, spread=True)

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"完成：{len(prs.slides)}页 -> {OUT}")


if __name__ == "__main__":
    build()
