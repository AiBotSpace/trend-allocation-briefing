# -*- coding: utf-8 -*-
"""Generator for 趋势与全天候配置简报 · 第1期 (2026-11-17), 修订版 v2.
v2 changes (2026-11-17): +市场情绪 (三) +趋势系统与仓位 (七); 中证500/中证1000/
上证50收盘点位补入; 英伟达财报时间修正为美东11-17盘后 (北京时间11-18早间).
All figures verified against public sources on 2026-11-17 (see 来源 section).
Regenerate from this script; never edit the .docx directly.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SCRATCH = "/home/hatch/workspace/trend-allocation-briefing"
OUT = f"{SCRATCH}/editions/briefing-01.docx"

doc = Document()
doc.core_properties.title = "趋势与全天候配置简报 · 第1期（2026年11月17日）"
doc.core_properties.author = "贾维斯"

# --- base styles ---
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.15
for h in ("Heading 1", "Heading 2"):
    hs = doc.styles[h]
    hs.font.name = "Calibri"
    hs.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
doc.styles["Heading 1"].font.size = Pt(15)
doc.styles["Heading 1"].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
doc.styles["Heading 2"].font.size = Pt(12)
doc.styles["Heading 2"].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
for s in ("List Number", "List Bullet"):
    ls = doc.styles[s]
    ls.font.name = "Calibri"
    ls.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    ls.paragraph_format.space_after = Pt(4)


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(10)
        shade(hdr[i], "D9E2F3")
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            r = p.add_run(v)
            r.font.size = Pt(10)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t


def para(text, bold=False, italic=False, size=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    return p


def bullet(text):
    doc.add_paragraph(text, style="List Bullet")


def _new_num_id():
    """Clone the 'List Number' numbering definition under a fresh numId so a
    new list restarts at 1 instead of continuing the previous list."""
    from copy import deepcopy
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    style = doc.styles["List Number"]
    numPr = style.element.find(qn("w:pPr")).find(qn("w:numPr"))
    base_id = int(numPr.find(qn("w:numId")).get(qn("w:val")))
    nums = numbering.findall(qn("w:num"))
    new_id = max(int(n.get(qn("w:numId"))) for n in nums) + 1
    for n in nums:
        if int(n.get(qn("w:numId"))) == base_id:
            new_num = deepcopy(n)
            new_num.set(qn("w:numId"), str(new_id))
            # Force the new instance to restart at 1 (LibreOffice merges
            # same-abstractNum instances into one continuous list otherwise).
            lvl_override = OxmlElement("w:lvlOverride")
            lvl_override.set(qn("w:ilvl"), "0")
            start_override = OxmlElement("w:startOverride")
            start_override.set(qn("w:val"), "1")
            lvl_override.append(start_override)
            new_num.append(lvl_override)
            numbering.append(new_num)
            break
    return new_id


def _set_num_id(para, num_id):
    pPr = para._p.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = OxmlElement("w:numPr")
        pPr.append(numPr)
    el = numPr.find(qn("w:numId"))
    if el is None:
        el = OxmlElement("w:numId")
        numPr.append(el)
    el.set(qn("w:val"), str(num_id))


def numbered(text, num_id=None):
    p = doc.add_paragraph(text, style="List Number")
    if num_id is not None:
        _set_num_id(p, num_id)
    return p


# ===== Title =====
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = t.add_run("趋势与全天候配置简报 · 第1期")
r.bold = True
r.font.size = Pt(22)
r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

para("2026年11月17日 ｜ 数据截至2026年11月16日收盘（A股）／11月16日美东收盘（美股）；11月17日早盘快照截至09:35",
     italic=True, size=10)
para("本期为基线版：首次建立对照基线，下期起对照本周点位只写变化与新增信号，不重复完整背景。事实、推断、未知严格分开；未核实处明确标注未知。",
     italic=True, size=10)
para("本版更新（11月17日）：新增市场情绪（第三节）与趋势系统仓位（第七节）；补入中证500／中证1000／上证50收盘点位；修正英伟达财报时间为美东11月17日盘后（北京时间11月18日早间）。",
     italic=True, size=10)

# ===== 一、结论 =====
doc.add_heading("一、结论", level=1)
numbered("A股调整未止，成长领跌。11月16日沪深300收3,979.80点（-0.91%），创2025年7月以来新低；创业板指收2,320.02点（-4.37%），再创年内新低；科创50跌5.16%。11月初券商“温和触底后反弹”的预期，尚未在价格上兑现。")
numbered("风格极端分化：银行等权重抗跌（工商银行等多只银行股续创历史新高），算力硬件重挫，全市场超4,300只个股下跌。沪指-0.87%的温和跌幅掩盖了成长板块的实际杀伤。")
numbered("市场情绪处于冰点：情绪温度18°，上涨1,136家对下跌4,344家（连续第三天超4,300家下跌），涨停约33家、跌停约63家，炸板率33.3%，两融余额单日减少163亿元，成交缩量一成——存量出清，无增量资金。")
numbered("全天候参考组合股债双弱：A股下跌的同时，中国10年期国债收益率上行至1.63%（债券价格下跌），30年国债期货跌超0.5%；黄金小幅回落；原油年内大涨是组合中唯一的强对冲方向。")
numbered("外围同步走弱：美股三大指数齐跌（纳指-2.49%），VIX升至约23；美元指数创2025年4月以来新高，市场对12月美联储加息的定价约90%。英伟达Q3财报（美东11月17日盘后，北京时间11月18日早间）是本周第一个关键验证点。")
numbered("趋势系统信号为防御：主要指数均处空头排列且破位确认，A股权益仓位建议0%–5%（占总资产）；加仓需等待收复MA20、广度修复与放量三者组合确认（第七节）。")

# ===== 二、事实：A股 =====
doc.add_heading("二、事实：A股主要指数（2026年11月16日收盘）", level=1)
table(
    ["指数", "收盘点位", "日涨跌", "备注"],
    [
        ["上证指数", "3,602.57", "-0.87%", "盘中跌破3,600点，险守关口；沪市成交约7,062亿元"],
        ["深证成指", "10,690.78", "-2.53%", "深市成交约7,783亿元"],
        ["沪深300", "3,979.80", "-0.91%", "创2025年7月7日以来新低（约16个月新低）；52周区间3,958.14–5,064.27"],
        ["创业板指", "2,320.02", "-4.37%", "再创年内新低；失守2,400点；盘中最低2,308.90"],
        ["科创50", "1,213.68", "-5.16%", "跌破1,250点；宽基中跌幅最大"],
        ["中证500", "6,297.18", "-2.60%", ""],
        ["中证1000", "6,164.62", "-3.03%", ""],
        ["上证50", "2,766.57", "-0.30%", "宽基中最抗跌"],
        ["北证50", "未核实", "-4.54%", "收盘点位未核实"],
    ],
    widths=[2.2, 2.0, 1.6, 9.0],
)
bullet("两市成交约1.48万亿元，较上个交易日减少约1,712亿元（约-10%）；沪深京三市合计约1.494万亿元，缩量10.38%。")
bullet("涨跌家数：沪市上涨589只／下跌1,718只；深市上涨537只／下跌2,354只；全市场超250股跌超9%。")
bullet("风格：银行等权重股抗跌，工商银行等多只银行股续创历史新高；算力硬件股重挫。")
bullet("上周五（11月13日）收盘对照：沪深300 4,016.29（-0.68%）、创业板指2,426.08（-3.32%）、上证指数3,634.35（-0.59%）、上证50 2,775.02（+0.11%）。")
bullet("11月17日早盘快照（09:35，盘中数据，非收盘）：沪深300 ETF -0.87%，科创50 ETF -4.92%，中证500 ETF -3.08%——早盘继续走弱。")

doc.add_heading("ETF区间回报（Finnhub，截至2026年11月17日09:35）", level=2)
table(
    ["ETF（跟踪指数）", "近5日", "近1月", "年初至今", "52周低点"],
    [
        ["沪深300 ETF", "-1.72%", "-4.12%", "-9.79%", "1.363（10月23日）"],
        ["科创50 ETF", "-2.21%", "-9.8%", "-4.39%", "11.22港元（11月3日）"],
        ["中证500 ETF", "-4.65%", "-6.85%", "-11.79%", "7.133（10月30日）"],
    ],
    widths=[3.5, 1.8, 1.8, 2.0, 3.5],
)

# ===== 三、事实：市场情绪 =====
doc.add_heading("三、事实：市场情绪（2026年11月16日）", level=1)
table(
    ["情绪指标", "数值（截至11-16收盘，除标注外）", "说明"],
    [
        ["情绪周期", "冰点期，温度18°（连板网，午间）", "情绪极端低迷"],
        ["上涨／下跌家数", "1,136／4,344（沪深京）", "连续第三天超4,300家下跌"],
        ["涨停／跌停", "约32–33家／约60–63家", "两来源口径略异；午间跌停仅14家，午后扩散"],
        ["连板", "最高4板（北纬科技）；连板9家；晋级率27%", "高标有限，接力弱"],
        ["炸板率", "33.3%（真实炸板）", "打板成功率低"],
        ["昨日涨停溢价", "+0.48%（红盘率44.4%，大面率16.7%）", "昨日涨停今日基本无溢价"],
        ["全A指数", "1,645.564（-2.37%）", "全市场普跌"],
        ["轮动池广度（91只ETF）", "23涨／1平／67跌，中位-1.28%；涨超2%仅1只（银行ETF +2.23%），跌超2%达35只；站上MA20仅21/91", "广度极差，红利是唯一避风港"],
        ["新高／新低（63日，午间）", "82／621", "新低占优"],
        ["两融余额（截至11-13）", "22,865.25亿（-163.40亿）；融资余额22,571.97亿（-166.34亿）；融券293.28亿（+2.94亿）", "杠杆资金净流出"],
        ["成交", "沪深京1.494万亿，缩量10.38%", "无增量资金，存量出清"],
        ["北向资金", "当日净买入自2024年8月起不再披露", "本期无该数据，见未知"],
        ["尾盘", "跌停66→63、上涨950→1,136", "尾盘30分钟首次承接试探，但缩量说明无增量资金"],
    ],
    widths=[3.2, 6.5, 5.5],
)
bullet("主线缺失：最强题材AI应用仅3家涨停，无持续主线；资金在防御方向抱团，工商银行、建设银行、南京银行、成都银行、杭州银行、江苏银行、青岛银行、美的集团盘中创历史新高。")

# ===== 四、10月 vs 11月 =====
doc.add_heading("四、事实：10月调整与11月“企稳”对照", level=1)
doc.add_heading("10月复盘（招商证券11月观点）", level=2)
bullet("10月A股涨少跌多：大盘价值与消费龙头逆势上涨，科创50、创业板指与科技龙头领跌；食品饮料、银行等红利方向涨幅领先，电子、通信领跌。")
bullet("10月主旋律是流动性冲击：海外美债利率高位徘徊，对全球权益资产产生虹吸效应；国内微观流动性面临赎回等资金负反馈，整体偏弱运行。")
bullet("9月底以来市场处于第二轮去杠杆过程（中信建投）；10月以来稳市资金陆续入场；10月可跟踪资金转为净流入，ETF净流入对冲融资净流出。")
doc.add_heading("11月初预期（十大券商，11月2日）", level=2)
bullet("多数券商认为，经历前期深度调整后，短期估值性价比修复，超跌反弹的条件进一步改善，当前位置不宜过分恐慌。")
bullet("招商证券：11月市场将在温和触底后迎来反弹，需要“先守后攻，静待花开”；围绕超跌反弹、周期稳健、红利优选均衡配置。")
bullet("申万宏源：超跌反弹随时可能触发，一个重要的中长期低点可能在岁末年初出现。")
doc.add_heading("11月16日现实", level=2)
para("截至11月16日收盘，沪深300创约16个月新低，创业板指再创年内新低，科创50跌破1,250点。“企稳”尚未在价格上兑现；11月上半月延续调整，而非触底反弹。（事实陈述：以收盘价对照券商11月初观点。）")

# ===== 五、外围与宏观 =====
doc.add_heading("五、事实：外围与宏观", level=1)
doc.add_heading("美股（11月16日美东收盘）", level=2)
table(
    ["指数", "收盘", "日涨跌", "备注"],
    [
        ["道琼斯工业平均", "51,608.17", "-1.41%", "五大指数中最抗跌"],
        ["标普500", "7,416.59", "-1.92%", "连续第三日下跌；11个板块9跌2涨"],
        ["纳斯达克综合", "24,509.76", "-2.49%", "连续第三日下跌"],
        ["纳斯达克100", "27,197.15", "-2.67%", "主要指数中跌幅最大"],
        ["罗素2000", "2,742.94", "-1.70%", ""],
    ],
    widths=[3.0, 2.2, 1.6, 8.0],
)
bullet("板块：科技板块-3.24%领跌；必需消费（+0.27%）与公用事业（+0.18%）收涨。费城半导体指数-5.28%。VIX约22.9（+18.6%）。")
bullet("触发因素（公开报道）：英伟达人事动荡持续发酵（CEO辞任），NVDA收跌4.12%至156.14美元；Q3财报将于美东11月17日盘后公布（北京时间11月18日早间）。美光 -7.6%，AMD -5.6%，博通 -3.9%。另据公开报道，海外机构大幅下修2028–2030年全球AI加速器资本开支预期（最高接近四成）。")
bullet("亚太11月16日：日经225 -2.42%（59,662.02点）；韩国KOSPI -3.59%（5,100.50点），SK海力士-7.89%，三星电子-4.47%。")
doc.add_heading("利率与汇率", level=2)
bullet("美国10年期国债收益率5.117%（11月16日，下跌3.6个基点）；市场对12月加息25个基点的定价约90%（CME FedWatch，11月16日）。")
bullet("中国10年期国债收益率1.63%（11月16日，上行2个基点）；2年期1.24%（上行1个基点）；3个月1.17%（持平）。11月11–13日，10年国开债收益率累计上行7.5个基点；11月16日30年国债期货跌超0.5%，10年期国债期货跌0.13%。")
bullet("背景：11月11日央行发布三季度货币政策执行报告，提及国债免征所得税问题，市场对政府债利息所得税免税是否取消的讨论升温；是否取消仍停留在讨论阶段，并非已经实施。")
bullet("美元兑人民币中间价6.7141（11月16日，较前一交易日升值11基点，连续3日升值；年初以来升值4.687%）；在岸即期约6.67。美元指数约103.5，创2025年4月以来新高，连续第三周上涨。")
bullet("中国10月经济数据（公开报道）：社会消费品零售总额同比+0.4%，规模以上工业增加值同比+4.8%，均低于预期。")

# ===== 六、全天候 =====
doc.add_heading("六、全天候参考组合对照", level=1)
para("参考组合定义（公开口径，本简报系列固定使用，非个人持仓）：股票30%／长期债券40%／中期债券15%／黄金7.5%／大宗商品7.5%（经典全天候结构；股票锚定沪深300，债券锚定中国国债，黄金与商品锚定国际期货报价）。")
table(
    ["资产", "参考权重", "本周信号（截至11-16）", "数据锚点"],
    [
        ["股票（A股）", "30%", "下跌，成长领跌", "沪深300当日-0.91%；近1月约-4.1%（ETF口径）"],
        ["长期债券（中国国债）", "40%", "收益率上行，价格下跌", "10年期1.63%（+2bp）；30年国债期货跌超0.5%"],
        ["中期债券", "15%", "收益率小幅上行", "2年期1.24%（+1bp）"],
        ["黄金", "7.5%", "小幅回落", "COMEX黄金期货约3,864美元/盎司（-0.56%；年初至今约-11%）"],
        ["大宗商品", "7.5%", "原油强势，铜小幅回落", "WTI约89.12美元/桶（+0.63%；年初至今约+55%）；布伦特约96.26美元/桶（+0.44%；年初至今约+58%）；铜约6.02美元/磅（-0.99%；年初至今约+7%）"],
    ],
    widths=[3.0, 1.8, 3.5, 8.5],
)
para("组合层面：股债双弱，黄金未提供对冲，原油是主要正贡献方向。组合层面精确回报需以指数序列按固定规则计算，本期未建立统一口径，标注未知；下期起以本周数据为基线跟踪漂移。股票部分的仓位建议见第七节趋势系统输出；本表权重为参考基准，非当前建议仓位。")

# ===== 七、趋势系统与仓位 =====
doc.add_heading("七、趋势系统与仓位", level=1)
para("本简报采用的趋势系统（规则公开，每期沿用同一口径，不做个股推荐）：", bold=True)
bullet("趋势过滤：指数收盘价 vs MA20／MA60／MA250（年线）。多头排列 = 价格＞MA20＞MA60＞MA250且均线向上；空头排列反之。")
bullet("趋势强度：价格在52周高低区间中的位置；近1月／近6月涨跌幅。")
bullet("广度确认：上涨家数占比、站上MA20的ETF占比、涨停－跌停差。")
bullet("波动率环境：VIX及近期振幅。")
doc.add_heading("趋势系统读数（截至11月16日收盘）", level=2)
table(
    ["维度", "读数", "信号"],
    [
        ["上证 vs 均线", "收3,602.57，跌破MA5 3,647.42／MA10 3,653.13／MA20 3,662.20；位于MA60 3,812.29下方5.50%", "空头"],
        ["创业板 vs 均线", "收2,320.02，位于MA20 2,532.75下方8.40%、MA60 3,030.78下方23.45%；跌破45日低点2,376.28（下方2.37%）", "空头，破位确认"],
        ["沪深300 vs 均线", "收3,979.80；11月9日已跌破MA20 4,101.27／MA60 4,394.20／MA250 4,629.37，三线方向全部向下（11月16日当日均线值未核实，见未知）", "空头"],
        ["52周位置", "沪深300距52周低3,958.14仅+0.55%；沪深300 ETF（11-16收1.385）距52周低1.363仅+1.6%", "接近年内低点"],
        ["动量", "沪深300 ETF近1月-4.12%、近6月-15.25%；科创50 ETF近1月-9.8%、近6月-26.2%", "动量向下"],
        ["广度", "下跌4,344家；站上MA20仅21/91；MA20高于MA60仅11/91", "广度恶化"],
        ["波动率", "VIX约22.9（+18.6%）", "波动放大"],
    ],
    widths=[2.6, 9.0, 3.2],
)
doc.add_heading("仓位映射（A股权益仓位，占总资产；参考权重30%为满配基准）", level=2)
table(
    ["系统状态", "判定条件", "仓位建议"],
    [
        ["强多头", "多头排列 + 广度确认 + 动量向上", "25%–30%"],
        ["偏多", "价格站上MA20，广度修复中", "15%–25%"],
        ["中性／震荡", "方向不明，均线纠缠", "10%–15%"],
        ["偏空", "空头排列，但未破位", "5%–10%"],
        ["防御", "空头排列 + 破位确认 + 广度恶化", "0%–5%"],
    ],
    widths=[2.4, 7.0, 2.6],
)
para("本期信号：防御 → A股权益仓位建议0%–5%（占总资产；相当于30%参考权重的0–17%）。", bold=True)
para("交叉验证：独立第三方系统（A股ETF罗盘，11月17日盘前版）同日输出仓位0%–10%，方向一致。")
para("加仓触发条件（需组合确认，单一条件不足以改变信号）：", bold=True)
bullet("沪深300／上证收复MA20并站稳3个交易日。")
bullet("上涨家数占比回升至50%以上，且跌停家数回落至20家以下。")
bullet("两市成交额回升至1.8万亿以上（放量确认）。")
bullet("创业板指收复45日低点2,376.28。")
para("保持防御／进一步减仓的触发：沪指有效跌破3,600点；沪深300跌破52周低3,958.14；跌停家数继续扩散。", bold=True)
para("说明：以上为规则系统的信号输出，非个性化投资建议；仓位以占总资产比例计；最终决策由用户自行判断。", italic=True, size=10)

# ===== 八、推断 =====
doc.add_heading("八、推断（非事实，仅供参考，不构成投资建议）", level=1)
_infer_id = _new_num_id()
numbered("成长指数跌幅扩大（创业板-4.37%、科创50-5.16%）且成交缩量约10%，调整仍在进行中，尚未出现量价层面的止跌信号。", _infer_id)
numbered("银行股续创新高与成长指数创年内新低并存，表明资金在防御方向抱团，而非全面企稳；沪指的相对抗跌主要来自权重股，难以代表市场整体。", _infer_id)
numbered("股债双弱削弱了全天候组合的分散效果：本期组合中仅原油提供主要对冲，与“股债负相关”的常态假设偏离。", _infer_id)
numbered("外围风险（英伟达事件、美联储加息预期、美元走强）与A股调整同向共振，短期波动来源以外生冲击为主。", _infer_id)
numbered("情绪冰点＋缩量出清是典型的底部酝酿特征之一，但趋势系统尚未给出修复确认信号（均线、广度、成交三者均未确认），冰点本身不是买入信号。", _infer_id)

# ===== 九、未知 =====
doc.add_heading("九、未知", level=1)
_unknown_id = _new_num_id()
numbered("北证50指数11月16日收盘点位未核实。", _unknown_id)
numbered("全天候参考组合层面回报：统一计算口径未建立，本期不提供组合回报数字。", _unknown_id)
numbered("原油报价不同来源差异较大（布伦特96–106美元/桶区间），本期以Yahoo Finance期货近月报价为准，绝对水平需谨慎对待。", _unknown_id)
numbered("11月17日全天收盘尚未公布；英伟达Q3财报（美东11月17日盘后，北京时间11月18日早间）与美国10月零售销售（11月17日21:30北京时间）尚未发布。", _unknown_id)
numbered("美联储12月是否加息尚未决定；政府债利息所得税免税是否取消仍停留在市场讨论阶段。", _unknown_id)
numbered("10月各主要指数月度涨跌幅未逐项核实，本期仅引用券商复盘定性结论。", _unknown_id)
numbered("沪深300的11月16日当日MA20／MA60／MA250值未核实（本期引用11月9日值；均线移动缓慢，推断11月16日收盘仍在三线下方）。", _unknown_id)
numbered("北向资金当日净买入自2024年8月起不再披露，本期无该数据；某聚合站声称的净买入额与披露口径矛盾，不予采用。", _unknown_id)

# ===== 十、本周关注 =====
doc.add_heading("十、本周关注", level=1)
bullet("美东11月17日盘后（北京时间11月18日早间）：英伟达Q3财报（临时CEO首份财报与指引）。")
bullet("11月17日21:30（北京时间）：美国10月零售销售（预期环比+0.2%）。")
bullet("11月17–18日：APEC工商领导人峰会（深圳前海）。")
bullet("美联储官员表态与12月加息预期变化。")
bullet("趋势系统观察点：沪指3,600点关口、沪深300的52周低3,958.14、MA20收复情况、跌停家数与成交量变化。")

# ===== 来源 =====
doc.add_heading("数据来源", level=1)
for s in [
    "A股收盘：新浪财经收评、TradingView（路透）、21财经（新华财经）、TekStocks七维观察、a.finews.fyi A股日报，均2026年11月16日。",
    "沪深300点位：钜亨网、closelook.net，2026年11月16日。",
    "ETF区间回报：Finnhub，截至2026年11月17日09:35。",
    "市场情绪与均线数据：A股ETF罗盘（etf.peekabo.cc，2026年11月17日盘前版）；连板网情绪快照（2026年11月16日午间）；指股网11.16涨停复盘。",
    "两融余额：同花顺数据（据21经济网2026年11月16日报道）。",
    "沪深300均线（11月9日）：joinchao大盘晚间策略报告（2026年11月9日）。",
    "11月13日收盘：gudaobang.com A股收评（2026年11月13日）。",
    "10月复盘与11月观点：招商证券《先守后攻，静待花开——A股2026年11月观点及配置建议》；澎湃新闻十大券商观点汇总（2026年11月2日，新浪财经转载）。",
    "美股收盘：The Edge US Market Wrap、cappnotes Closing Look、daily-finance摘要，2026年11月16日。",
    "利率：中国债券信息网中债收益率曲线（据金融界2026年11月16日报道）；美国10年期国债收益率据finanzen.at，2026年11月16日。",
    "汇率：人民币中间价据中国外汇交易中心（MoneyDJ、TradingView转载），2026年11月16日；美元指数据公开报道。",
    "商品：Yahoo Finance期货近月报价（据fininfostation.com，2026年11月16日14:33台北更新）。",
    "北向资金披露口径：沪深交易所自2024年8月起调整北向资金披露方式（据aikanpan.top、hhxg.top口径说明）。",
]:
    bullet(s)
para("本简报为公开信息整理与研究参考，不构成投资建议；数据以官方发布为准。", italic=True, size=10)

doc.save(OUT)
print("saved", OUT)
