# -*- coding: utf-8 -*-
"""451期：牛市进入中后期了吗，我们该如何应对？

基于已确认解析稿生成 PPT。数据待更新页必须重新取数并做原生可编辑图表；
历史素材只作为样式参考，缺最新版截图处先占位。
"""
import sys
import json
import subprocess
from pathlib import Path

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E
from live_helpers import section, ph_note, bubble, summary

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


IMG449 = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_449")
IMG450 = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_450")
OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/451期直播PPT_根据解析稿生成.pptx"

ORG = RGBColor(0xF1, 0x76, 0x40)
PALE = RGBColor(0xFC, 0xE6, 0xD2)
BLUE = RGBColor(0x7C, 0xA7, 0xE0)
GREEN = RGBColor(0x62, 0xB5, 0x72)
RED = RGBColor(0xFF, 0x27, 0x00)
WIND_DIR = "/Users/fanlili/.agents/skills/wind-mcp-skill"
DATA_DIR = Path("/Users/fanlili/Downloads/live-ppt-tool/current")
P2_CACHE = {
    "885189.WI": DATA_DIR / "451_p2_885189_WI.json",
    "SPX.GI": DATA_DIR / "451_p2_SPX_GI.json",
    "H00985.CSI": DATA_DIR / "451_p2_H00985_CSI.json",
    "HSI.HI": DATA_DIR / "451_p2_HSI_HI.json",
}


def clear(sl):
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)


def newslide(prs, body, title):
    sl = prs.slides.add_slide(body)
    clear(sl)
    t_pt = E.title_fit(title)
    t_h = E.text_height_cm(title, t_pt, E.CONTENT_W) + 0.2
    E.add_tb(sl, title, E.CONTENT_L, E.CONTENT_TOP, E.CONTENT_W, t_h,
             size=t_pt, bold=True, color=E.C_BLACK, align=PP_ALIGN.CENTER)
    return sl


def btop(title):
    return E.CONTENT_TOP + E.text_height_cm(title, E.TITLE_PT, E.CONTENT_W) + 0.2 + 0.8


def img_page(prs, body, title, path, points=None, pending=True):
    if path and Path(path).exists():
        return E.img_page(prs, body, title, str(path), points=points, pending=pending)
    return ph_note(prs, body, title, "参考图/最新版截图待补充；已按解析稿先占位。")


def clean_section(prs, body, title):
    section(prs, body, title)


def source(sl, text="数据来源：待更新"):
    E.add_tb(sl, text, E.CONTENT_L + E.CONTENT_W - 12.0, E.CONTENT_BOTTOM - 0.55,
             12.0, 0.45, size=11, color=E.C_GRAY, align=PP_ALIGN.RIGHT)


def wind_kline(code, start="20240902", end="20251226"):
    cache = P2_CACHE.get(code)
    if cache and cache.exists():
        outer = json.loads(cache.read_text(encoding="utf-8"))
        if outer.get("ok") is False:
            raise RuntimeError(f"{cache.name} 不是有效行情数据，请先重新取数。")
    else:
        params = json.dumps({"windcode": code, "begin_date": start, "end_date": end, "period": "11"})
        r = subprocess.run(
            ["node", "scripts/cli.mjs", "call", "index_data", "get_index_kline", params],
            cwd=WIND_DIR, text=True, capture_output=True, timeout=90
        )
        if r.returncode != 0:
            raise RuntimeError(r.stderr or r.stdout)
        outer = json.loads(r.stdout)
    text = outer["content"][0]["text"]
    inner = json.loads(text)
    rows = inner["data"]["rows"]
    out = []
    for row in rows:
        try:
            out.append({"date": row[9][:4] + "-" + row[9][4:6] + "-" + row[9][6:8],
                        "close": float(row[2])})
        except Exception:
            continue
    return out


def aligned_pct_series(raw_items):
    """按共同交易日对齐，避免多市场假期导致断线。"""
    maps = [(name, {x["date"]: x["close"] for x in data}) for name, data in raw_items]
    common = set(maps[0][1])
    for _, dm in maps[1:]:
        common &= set(dm)
    dates = [d for d in sorted(common)]
    series = []
    for name, dm in maps:
        base = dm[dates[0]]
        series.append({"name": name, "dates": dates,
                       "values": [(dm[d] / base - 1) * 100 for d in dates]})
    return series


def aligned_two_series(a, b):
    ma = {x["date"]: x["close"] for x in a}
    mb = {x["date"]: x["close"] for x in b}
    dates = sorted(set(ma) & set(mb))
    return dates, [ma[d] for d in dates], [mb[d] for d in dates]


def global_market_page(prs, body):
    title = "近1年多，A股港股领涨全球"
    pairs = [
        ("全球股票市场", DATA_DIR / "451_p2_885189_WI_daily.json"),
        ("美股全市场", DATA_DIR / "451_p2_SPX_GI_daily.json"),
        ("中证全指全收益", DATA_DIR / "451_p2_H00985_CSI_daily.json"),
        ("港股恒生指数全收益", DATA_DIR / "451_p2_HSIRH_HI_daily.json"),
    ]
    raw = [(name, json.loads(path.read_text(encoding="utf-8"))) for name, path in pairs]
    series = aligned_pct_series(raw)
    sl = newslide(prs, body, title)
    y = btop(title)
    E.add_tb(sl, "参考449期第2页，数据已从Wind重新抓取；日频点数换算为区间累计涨跌幅。",
             E.CONTENT_L, y, E.CONTENT_W, 0.8, size=21, color=E.C_DARK, align=PP_ALIGN.LEFT)
    chart_l, chart_t, chart_w, chart_h = E.CONTENT_L + 1.0, y + 1.2, E.CONTENT_W - 2.0, E.CONTENT_BOTTOM - y - 1.7
    E.add_multi_line_chart(sl, chart_l, chart_t, chart_w, chart_h, series, y_pct=True)
    # 右侧终值标签，和参考图保持接近；图表仍为原生可编辑。
    x = chart_l + chart_w - 5.0
    top = chart_t + 1.0
    ranked = sorted([(i, s) for i, s in enumerate(series)], key=lambda z: z[1]["values"][-1], reverse=True)
    for row_i, (orig_i, s) in enumerate(ranked):
        val = s["values"][-1]
        box = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(top + row_i * 1.05), Cm(4.6), Cm(0.82))
        box.fill.solid(); box.fill.fore_color.rgb = E._pal(orig_i)
        box.line.fill.background(); box.shadow.inherit = False
        tb = E.add_tb(sl, f"{s['name']}\n上涨{val:.2f}%", x + 0.12, top + row_i * 1.05 + 0.03,
                      4.35, 0.75, size=10, bold=True, color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    source(sl, "数据来源：Wind")
    return sl


def usd_rate_page(prs, body):
    title = "原因一：美元降息周期，美元利率汇率贬值"
    usd = json.loads((DATA_DIR / "451_p3_usd_index_daily.json").read_text(encoding="utf-8"))
    us10y = json.loads((DATA_DIR / "451_p3_us10y_daily.json").read_text(encoding="utf-8"))
    dates, usd_vals, yld_vals = aligned_two_series(usd, us10y)
    sl = newslide(prs, body, title)
    y = btop(title)
    E.add_tb(sl, "美元指数及美国10年期国债收益率已重新取数；降息节点沿用参考页。",
             E.CONTENT_L, y, E.CONTENT_W, 0.75, size=21, color=E.C_DARK, align=PP_ALIGN.LEFT)

    cut_labels = {
        "2024-09-18": "-0.50%",
        "2024-11-07": "-0.25%",
        "2024-12-18": "-0.25%",
    }
    marker_vals = [None] * len(dates)
    label_map = {}
    for d, lab in cut_labels.items():
        if d in dates:
            idx = dates.index(d)
            marker_vals[idx] = usd_vals[idx]
            label_map[idx] = lab

    chart_l, chart_t = E.CONTENT_L + 1.0, y + 1.15
    chart_w, chart_h = E.CONTENT_W - 2.0, E.CONTENT_BOTTOM - y - 3.05
    E.add_overlay_dual_axis(
        sl, chart_l, chart_t, chart_w, chart_h, dates,
        "美元指数（左轴）", usd_vals,
        "美国10年期国债收益率（右轴）", yld_vals,
        marker_series=[("降息", marker_vals, "2BA84A", label_map)],
        bottom_kind="area",
        area_color="C7DBF0",
        source="数据来源：Wind"
    )
    bubble(sl, "近期的波动，也和美联储新任主席偏鹰派有关。市场担心降息节奏放缓甚至加息。",
           E.CONTENT_L + 2.0, chart_t + chart_h + 0.75, E.CONTENT_W - 4.0)
    return sl


def drawdown_stats(data, start_date):
    sub = [x for x in data if x["date"] >= start_date]
    peak = sub[0]
    max_dd = 0
    max_pair = (sub[0], sub[0])
    for x in sub:
        if x["close"] > peak["close"]:
            peak = x
        dd = x["close"] / peak["close"] - 1
        if dd < max_dd:
            max_dd = dd
            max_pair = (peak, x)
    return max_pair[0], max_pair[1], max_dd


def market_still_bull_page(prs, body):
    title = "近期虽然有所波动，但还在牛市"
    data = json.loads((DATA_DIR / "451_p5_000985_CSI_daily.json").read_text(encoding="utf-8"))
    hi = max(data, key=lambda x: x["close"])
    sep_low = min([x for x in data if "2024-09-01" <= x["date"] <= "2024-09-30"], key=lambda x: x["close"])
    rise = hi["close"] / sep_low["close"] - 1
    dd_peak, dd_trough, max_dd = drawdown_stats(data, "2026-02-01")
    last = data[-1]
    pullback = last["close"] / hi["close"] - 1

    sl = newslide(prs, body, title)
    y = btop(title)
    chart_l, chart_t = E.CONTENT_L + 0.8, y + 0.15
    chart_w, chart_h = E.CONTENT_W - 1.6, 8.8
    E.add_line_chart(sl, data, chart_l, chart_t, chart_w, chart_h,
                     hi_pt=hi, lo_pt=sep_low, series_name="中证全指")

    E.add_tb(sl, "数据来源：Wind", chart_l + chart_w - 6.5, chart_t + chart_h + 0.08,
             6.5, 0.45, size=11, color=E.C_GRAY, align=PP_ALIGN.RIGHT)

    stats = [
        ("本轮最高点", f"{hi['date']}\n{hi['close']:,.0f}点"),
        ("2024年9月低点至最高点", f"上涨{rise*100:.0f}%"),
        ("2026年2月至今最大回撤", f"{dd_peak['date']}至{dd_trough['date']}\n{max_dd*100:.0f}%"),
        ("截至2026年5月底回调", f"{pullback*100:.0f}%"),
    ]
    gap = 0.45
    box_w = (E.CONTENT_W - 3 * gap) / 4
    top = chart_t + chart_h + 0.75
    for i, (h, v) in enumerate(stats):
        x = E.CONTENT_L + i * (box_w + gap)
        box = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(top), Cm(box_w), Cm(2.05))
        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFF, 0xF7, 0xEF)
        box.line.color.rgb = ORG; box.line.width = Pt(1.2); box.shadow.inherit = False
        E.add_tb(sl, h, x + 0.2, top + 0.18, box_w - 0.4, 0.55,
                 size=13, bold=True, color=ORG, align=PP_ALIGN.CENTER)
        tb = E.add_tb(sl, v, x + 0.2, top + 0.78, box_w - 0.4, 1.05,
                      size=15, bold=True, color=E.C_DARK, align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return sl


def signal_boxes(prs, body):
    title = "现在到牛市中后期了么？"
    sl = newslide(prs, body, title)
    y = btop(title)
    labels = [
        ("星级", "接近上一轮牛市最高点"),
        ("估值", "宽基指数到历史80%-90%"),
        ("申购", "单日申购超百亿股票基金"),
        ("融资", "A股融资规模5月创新高"),
        ("ETF", "出现大额资金净流出"),
    ]
    gap = 0.55
    w = (E.CONTENT_W - 4 * gap) / 5
    for i, (a, b) in enumerate(labels):
        x = E.CONTENT_L + i * (w + gap)
        box = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y + 0.5), Cm(w), Cm(5.3))
        box.fill.solid()
        box.fill.fore_color.rgb = PALE if i % 2 == 0 else RGBColor(0xF8, 0xF1, 0xE8)
        box.line.color.rgb = ORG
        box.line.width = Pt(1.6)
        box.shadow.inherit = False
        E.add_tb(sl, a, x, y + 1.1, w, 1.1, size=24, bold=True, color=ORG, align=PP_ALIGN.CENTER)
        tb = E.add_tb(sl, b, x + 0.35, y + 2.65, w - 0.7, 2.0, size=17,
                      color=E.C_DARK, align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    bubble(sl, "一些信号还没到历史最高值，但整体进入了牛市中后期。\n牛市中后期，在市场贪婪情绪推动下，也会持续一段时间。",
           E.CONTENT_L + 2.2, y + 7.0, E.CONTENT_W - 4.4)
    return sl


def low_strategy(prs, body):
    title = "我们该如何应对？"
    sl = newslide(prs, body, title)
    y = btop(title)
    E.add_tb(sl, "（3）低估的品种", E.CONTENT_L, y, E.CONTENT_W, 0.8,
             size=23, bold=True, color=ORG, align=PP_ALIGN.LEFT)
    points = [
        "低估品种可以少量定投。",
        "下方序号改成三角，删除原图中“低估的品种”字样。",
        "不适合一次性投入大笔资金到股票基金。",
    ]
    E.add_arrow_points(sl, points, y + 1.1, pt=23)
    bubble(sl, "不适合一次性投入大笔资金到股票基金",
           E.CONTENT_L + 4.0, y + 6.7, E.CONTENT_W - 8.0)
    return sl


def q2_page(prs, body):
    title = "Q2：牛市中后期，还可以投资股票基金吗？"
    sl = newslide(prs, body, title)
    y = btop(title)
    boxes = [
        ("高估品种", "不追涨；已有持仓可考虑分批止盈"),
        ("正常估值", "耐心持有，等待估值变化"),
        ("低估品种", "可以少量定投，保留未来加仓资金"),
    ]
    gap = 1.0
    w = (E.CONTENT_W - 2 * gap) / 3
    for i, (h, txt) in enumerate(boxes):
        x = E.CONTENT_L + i * (w + gap)
        shp = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y + 1.0), Cm(w), Cm(6.5))
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(0xFF, 0xF7, 0xEF)
        shp.line.color.rgb = ORG
        shp.line.width = Pt(1.6)
        shp.shadow.inherit = False
        E.add_tb(sl, h, x + 0.4, y + 1.7, w - 0.8, 1.1, size=24, bold=True,
                 color=ORG, align=PP_ALIGN.CENTER)
        tb = E.add_tb(sl, txt, x + 0.7, y + 3.4, w - 1.4, 2.5, size=19,
                      color=E.C_DARK, align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    source(sl, "根据已确认大纲整理")
    return sl


def build():
    prs = Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")

    E.cover_page(prs, cover, "牛市进入中后期了吗，\n我们该如何应对？")

    global_market_page(prs, body)
    usd_rate_page(prs, body)
    img_page(prs, body, "原因二：股市基本面复苏，盈利同比恢复增长", IMG449 / "v10.png",
             points=["参考449期第5页，走势图数据待更新。"])
    market_still_bull_page(prs, body)
    signal_boxes(prs, body)
    img_page(prs, body, "牛市中后期的信号", IMG449 / "v16.png",
             points=["（1）星级接近上一轮牛市最高点；估值表图片待更新，后续文字取消加粗。"])
    img_page(prs, body, "牛市中后期的信号", IMG450 / "v5.png",
             points=["（2）宽基指数估值到了历史80%-90%的位置；参考450期第9页，数据待更新。"])
    img_page(prs, body, "牛市中后期的信号", None,
             points=["（3）出现单日申购超百亿的股票基金。"], pending=True)
    img_page(prs, body, "牛市中后期的信号", IMG449 / "v25.png",
             points=["（4）A股融资规模5月份创下历史新高；融资保证金比例从80%提升到100%；信号板图片待替换为最新。"])
    img_page(prs, body, "牛市中后期的信号", None,
             points=["（5）ETF出现大额资金净流出。",
                     "展示2024年9月以来沪深300ETF份额走势图 vs 沪深300走势。",
                     "气泡：ETF净流出主要说明机构在减仓，是牛市进入中后期的典型信号，并不代表市场接下来要下跌。"], pending=True)
    img_page(prs, body, "牛市中后期，为何波动大？", None, pending=True)
    img_page(prs, body, "牛市中后期，为何波动大？", None, pending=True)
    img_page(prs, body, "我们该如何应对？", IMG450 / "v14.png",
             points=["参考441期第17页；下方补充小气泡：坚持低估买入，正常估值持有，高估卖出。"])
    img_page(prs, body, "我们该如何应对？", IMG449 / "v26.png",
             points=["参考430期第8页，图片待更新。"])
    img_page(prs, body, "我们该如何应对？", IMG450 / "v26.png",
             points=["（2）高估的品种，考虑分批止盈；数据待更新。"])
    low_strategy(prs, body)
    img_page(prs, body, "我们该如何应对？", IMG449 / "v26.png",
             points=["（4）止盈资金、更多存量资金，考虑稳健品种，降低波动；参考449期第12页，数据待更新。"])

    clean_section(prs, body, "常见疑问解答")
    img_page(prs, body, "Q1：一般牛市中后期会持续多久？担心没来得及止盈就下跌了", None,
             points=["参考430期第16页，右侧图片待更新。"], pending=True)
    q2_page(prs, body)
    summary(prs, body, "总结", [
        "近1年多，A股港股领涨全球，当前已到牛市中后期了。出现一些波动是很正常的。",
        "我们能做的是以不变应万变，坚持低买高卖，跌出机会买入，涨出机会卖出，其他时间耐心等待。",
        "对于高估品种，可以分批止盈。",
        "低估品种，可以少量定投，同时确保回到4-5星级时还有资金加仓。",
        "对于止盈的资金、存量资金，可以转入稳健类品种，降低波动风险。",
    ])

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"完成：{len(prs.slides)}页 -> {OUT}")


if __name__ == "__main__":
    build()
