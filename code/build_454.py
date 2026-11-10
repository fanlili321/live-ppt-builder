# -*- coding: utf-8 -*-
"""454期：指数调仓，对我们投资有啥影响呢？"""
import sys
from pathlib import Path

sys.path.insert(0, "/Users/fanlili/Downloads/live-ppt-tool")
import ppt_engine as E
from live_helpers import combo, pct, chart_pg, section, summary, table_page

from pptx import Presentation


IMG = Path("/Users/fanlili/Downloads/live-ppt-tool/imgs_454")
OUT = "/Users/fanlili/Desktop/范丽丽./直播PPT成品/454期直播PPT_根据解析稿生成.pptx"


def p(n):
    return str(IMG / f"v{n:02d}.png")


def img_page(prs, body, title, n, points=None, pending=False):
    path = Path(p(n))
    if path.exists():
        return E.img_page(prs, body, title, str(path), points=points, pending=pending)
    return E.img_page(prs, body, title, p(1), points=["参考图缺失，待补充。"], pending=True)


def valuation_table(prs, body):
    rows = [
        ["宽基", "沪深300", "19", "19.74", "1.62", "38.31↑", "5.88↑"],
        ["宽基", "中证500", "50", "23.12", "1.89", "34.98↑", "2.96↑"],
        ["宽基", "中证1000", "100", "30.68", "1.99", "43.60↑", "3.17↑"],
        ["宽基", "中证2000", "182", "40.86", "2.72", "34.74↓", "2.26↓"],
        ["策略", "红利低波100", "18", "12.77", "1.09", "12.71↓", "0.97↓"],
        ["策略", "300价值", "6", "22.75", "1.28", "11.93↓", "1.30↑"],
        ["策略", "300成长", "20", "17.41", "1.98", "30.26↑", "3.99↑"],
        ["策略", "中证A500", "34", "29.14", "2.14", "34.20↑", "3.65↑"],
    ]
    table_page(
        prs, body, "本次定期调仓，对各个指数估值的影响",
        ["指数类型", "指数名称", "调仓数量", "调出PE", "调出PB", "调入PE", "调入PB"],
        rows, [1.1, 2.2, 1.2, 1.15, 1.15, 1.15, 1.15],
        src="数据来源：Wind，截至2026年6月定期调仓"
    )


def enhanced_chart(prs, body):
    title = "（根据建议新增）指数增强组合的表现"
    data = combo("螺丝钉金钉宝指数增强C", 2, "2024-09-18")
    if data:
        chart_pg(
            prs, body, title,
            "展示2024年9月以来指数增强组合表现；中证800叠加线待数据补充。",
            [pct("指数增强组合", data)],
            "申购状态标注：2026年1月12日关闭申购；2026年3月23日打开申购；2026年4月14日关闭申购。"
        )
    else:
        E.img_page(prs, body, title, p(8),
                   points=["指数增强组合净值、中证800数据待补充。"], pending=True)


def build():
    prs = Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, "自定义版式")
    cover = E.get_layout(prs, 4, "标题幻灯片")

    E.cover_page(prs, cover, "指数调仓，对我们投资有啥影响呢？")

    img_page(prs, body, "什么是指数调仓？", 1)
    img_page(prs, body, "指数调仓的分类", 2)
    img_page(prs, body, "指数调仓的分类", 3,
             points=["（2）定期调仓。"])
    img_page(prs, body, "指数定期调仓的频率及时间", 4,
             points=["参考421期第5页；右下角气泡里的12月12日需要调整为6月12日。"],
             pending=True)
    valuation_table(prs, body)
    img_page(prs, body, "定期调仓，暴露市值选股型指数的缺点：追涨杀跌", 5)
    img_page(prs, body, "解决方法1：考虑策略指数，不基于市值选股，避免追涨杀跌", 6)
    img_page(prs, body, "红利&价值指数，调入调出股票的估值差别", 7,
             points=["中证红利内容不用变；删除300价值2024年12月数据，补充2026年6月数据。"])
    img_page(prs, body, "解决方法2：考虑指数增强基金", 8)
    enhanced_chart(prs, body)
    img_page(prs, body, "市值选股型指数的优点：能容纳的资金量比较大", 9,
             points=["参考421期第11页，指数基金规模数据待更新。"], pending=True)
    section(prs, body, "常见疑问解答")
    img_page(prs, body, "Q1：调仓会让指数基金净值发生变化吗？", 10)
    img_page(prs, body, "Q2：调仓需要我们投资者做什么吗？", 11)
    img_page(prs, body, "Q3：价值风格指数，调仓后估值降低，投资性价比会变高吗？", 12)
    summary(prs, body, "总结", [
        "本次定期调仓对各指数估值影响不同。",
        "宽基中沪深300、中证500、中证1000调仓后估值提高，中证2000下降。",
        "策略指数中300价值等价值风格估值降低，300成长、中证A500等成长指数估值提高。",
        "投资者不需要对调仓额外操作，基金经理后台完成，投资者耐心持有即可。",
    ])

    for _ in range(8):
        E.remove_slide(prs, 0)
    prs.save(OUT)
    print(f"完成：{len(prs.slides)}页 -> {OUT}")


if __name__ == "__main__":
    build()
