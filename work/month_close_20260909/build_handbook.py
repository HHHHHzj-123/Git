from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

BASE = Path(r"C:\Users\HZJ\Desktop\Git\work\month_close_20260909")
OUT = BASE / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
DOCX = OUT / "M02-001-总账会计年度工作与月结实战手册.docx"
IMG = BASE / "月结全景框架.png"

NAVY = "17365D"
BLUE = "DCE6F1"
PALE = "F4F7FB"
GRAY = "D9DEE7"
TEXT = "20242C"
RED = "A61B29"
GREEN = "548235"


def font_path(names):
    roots = [Path(r"C:\Windows\Fonts")]
    for n in names:
        for root in roots:
            p = root / n
            if p.exists():
                return str(p)
    return None


FONT = font_path(["msyh.ttc", "msyh.ttf", "simhei.ttf"])
FONT_BOLD = font_path(["msyhbd.ttc", "msyhbd.ttf", "simhei.ttf"]) or FONT


def make_overview():
    w, h = 2200, 1450
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    f_title = ImageFont.truetype(FONT_BOLD, 66)
    f_main = ImageFont.truetype(FONT_BOLD, 42)
    f_sub = ImageFont.truetype(FONT, 30)
    f_small = ImageFont.truetype(FONT, 25)
    d.text((1100, 60), "总账会计年度工作与一次月结全景", font=f_title, fill="#111111", anchor="ma")
    # central flow
    flow = [
        ("计划", "关账日历\nRACI\n口径与阈值"),
        ("截止", "收货 发货\n服务验收\n费用与资产"),
        ("入账", "子模块接口\n计提摊销\n成本与税费"),
        ("核对", "账账 账实\n账单 账税\n内部往来"),
        ("报告", "报表勾稽\n波动解释\n管理分析"),
        ("关账", "复核签批\n锁定期间\n问题复盘"),
    ]
    x0, y0, boxw, boxh, gap = 90, 250, 300, 230, 48
    for i, (head, body) in enumerate(flow):
        x = x0 + i * (boxw + gap)
        d.rounded_rectangle((x, y0, x+boxw, y0+boxh), radius=28, fill="#F4F7FB", outline="#17365D", width=4)
        d.rectangle((x, y0, x+boxw, y0+65), fill="#17365D")
        d.text((x+boxw/2, y0+32), head, font=f_main, fill="white", anchor="mm")
        for j, line in enumerate(body.split("\n")):
            d.text((x+boxw/2, y0+102+j*45), line, font=f_sub, fill="#20242C", anchor="mm")
        if i < len(flow)-1:
            ax = x+boxw+8
            ay = y0+boxh/2
            d.line((ax, ay, ax+28, ay), fill="#548235", width=8)
            d.polygon([(ax+28, ay-13), (ax+47, ay), (ax+28, ay+13)], fill="#548235")
    # three supporting layers
    layers = [
        ("业务与证据", "合同 订单 收发货 验收 发票 银行回单 工时 盘点记录", "#EAF2F8"),
        ("系统与控制", "采购 销售 仓储 生产 费用 薪酬 资金 → ERP子模块 → 总账", "#EEF6EC"),
        ("判断与输出", "权责发生制 截止性 估计与减值 税会差异 → 报表 管理分析 审计税务资料", "#FDF2E9"),
    ]
    for i, (head, body, fill) in enumerate(layers):
        y = 580 + i*160
        d.rounded_rectangle((160, y, 2040, y+115), radius=20, fill=fill, outline="#B8C2CE", width=3)
        d.text((210, y+56), head, font=f_main, fill="#111111", anchor="lm")
        d.text((560, y+56), body, font=f_sub, fill="#20242C", anchor="lm")
    # annual cadence
    d.text((160, 1110), "年度节奏", font=f_main, fill="#111111")
    months = [("每月", "D-5至D+7月结"), ("季度", "季末加深复核与所得税预缴"), ("年中", "盘点 预测 半年复盘"), ("年末", "截止 盘点 减值 审计 汇算")]
    for i, (a, b) in enumerate(months):
        x = 160 + i*480
        d.rounded_rectangle((x, 1180, x+410, 1325), radius=20, fill="#FFFFFF", outline="#17365D", width=3)
        d.text((x+205, 1218), a, font=f_main, fill="#17365D", anchor="mm")
        d.text((x+205, 1277), b, font=f_small, fill="#20242C", anchor="mm")
    im.save(IMG, quality=95)


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v)); node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def set_cell_border(cell, color=GRAY, sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders"); tcPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "start" if edge == "left" else "end" if edge == "right" else edge
        e = borders.find(qn(f"w:{tag}"))
        if e is None:
            e = OxmlElement(f"w:{tag}"); borders.append(e)
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), sz); e.set(qn("w:color"), color)


def set_run_font(run, size=10.5, bold=False, color=TEXT, east="Microsoft YaHei"):
    run.font.name = east; run.font.size = Pt(size); run.bold = bold; run.font.color.rgb = RGBColor.from_string(color)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "1F4E79"); rPr.append(color)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    fonts = OxmlElement("w:rFonts"); fonts.set(qn("w:eastAsia"), "Microsoft YaHei"); rPr.append(fonts)
    r.append(rPr); t = OxmlElement("w:t"); t.text = text; r.append(t); h.append(r); paragraph._p.append(h)


def add_table(doc, headers, rows, widths=None, font_size=8.4):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for j, h in enumerate(headers):
        cell = hdr.cells[j]; cell.text = h; set_cell_shading(cell, NAVY); set_cell_margins(cell); set_cell_border(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs: set_run_font(r, font_size, True, "FFFFFF")
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j, val in enumerate(row):
            cells[j].text = str(val)
            cells[j].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cells[j]); set_cell_border(cells[j])
            if i % 2: set_cell_shading(cells[j], PALE)
            for p in cells[j].paragraphs:
                p.paragraph_format.space_after = Pt(0); p.paragraph_format.line_spacing = 1.08
                if len(str(val)) <= 9: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs: set_run_font(r, font_size)
    if widths:
        for row in table.rows:
            for j, width in enumerate(widths):
                row.cells[j].width = Cm(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


def add_body(doc, text, bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6); p.paragraph_format.line_spacing = 1.35
    if bold_lead and text.startswith(bold_lead):
        r = p.add_run(bold_lead); set_run_font(r, 10.5, True)
        r = p.add_run(text[len(bold_lead):]); set_run_font(r, 10.5)
    else:
        r = p.add_run(text); set_run_font(r, 10.5)
    return p


def add_bullets(doc, items, level=0):
    for item in items:
        p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
        p.paragraph_format.space_after = Pt(3); p.paragraph_format.line_spacing = 1.25
        r = p.add_run(item); set_run_font(r, 10.2)


def add_steps(doc, items):
    for idx, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.left_indent = Cm(0.55)
        p.paragraph_format.first_line_indent = Cm(-0.55)
        r = p.add_run(f"{idx}.  {item}"); set_run_font(r, 10.2)


def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.keep_with_next = True
    return p


def add_page_break(doc):
    doc.add_page_break()


make_overview()
doc = Document()
sec = doc.sections[0]
sec.page_height = Cm(29.7); sec.page_width = Cm(21)
sec.top_margin = Cm(1.8); sec.bottom_margin = Cm(1.7); sec.left_margin = Cm(1.9); sec.right_margin = Cm(1.9)

# styles
styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Microsoft YaHei"; normal.font.size = Pt(10.5); normal.font.color.rgb = RGBColor.from_string(TEXT)
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
for name, size, color in [("Title", 25, "000000"), ("Subtitle", 12, "555555"), ("Heading 1", 17, "000000"), ("Heading 2", 13, "000000"), ("Heading 3", 11, "000000")]:
    st = styles[name]; st.font.name = "Microsoft YaHei"; st.font.size = Pt(size); st.font.color.rgb = RGBColor.from_string(color)
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    if name != "Subtitle": st.font.bold = True
# Remove any theme-provided border under the Title style.
title_ppr = styles["Title"]._element.get_or_add_pPr()
for node in list(title_ppr):
    if node.tag == qn("w:pBdr"):
        title_ppr.remove(node)
styles["Heading 1"].paragraph_format.space_before = Pt(18); styles["Heading 1"].paragraph_format.space_after = Pt(8)
styles["Heading 2"].paragraph_format.space_before = Pt(13); styles["Heading 2"].paragraph_format.space_after = Pt(5)
styles["Heading 3"].paragraph_format.space_before = Pt(9); styles["Heading 3"].paragraph_format.space_after = Pt(4)

# footer page numbers
for section in doc.sections:
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("M02-001  总账月结实战手册    ")
    set_run_font(run, 8.5, False, "666666")
    fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE")
    footer._p.append(fld)

# cover
p = doc.add_paragraph(style="Title"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(110); p.add_run("总账会计年度工作与月结实战手册")
p = doc.add_paragraph(style="Subtitle"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("M02-001  从年度节奏到一次完整月结")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(28)
r = p.add_run("适用读者  从审计转向企业总账会计或财务主管"); set_run_font(r, 11, False, "333333")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(150)
r = p.add_run("版本 2026年9月  基准场景为自然年度 一般纳税人 非上市企业 月度结账"); set_run_font(r, 9.5, False, "666666")
add_page_break(doc)

add_heading(doc, "使用说明", 1)
add_body(doc, "这份手册回答两个问题：一年之中总账会计和财务主管在什么时候做什么，以及一次月结怎样从业务截止推进到报表锁账。结论是，月结不是总账在月末集中做分录，而是一项跨部门、跨系统、以证据和截止性为核心的项目管理工作。")
add_body(doc, "全文采用四类标签区分来源。法律与准则是强制边界；行业常见实务是多数企业为提高及时性和完整性采用的做法；企业核算口径需要管理层批准并保持一致；案例假设只用于演示，不能直接替代具体企业的合同、系统配置和税务判断。")
add_table(doc, ["标识", "含义", "使用方法"], [
    ["法律或准则", "法律法规、企业会计准则或统一会计制度的明确要求", "按适用主体和生效日期执行"],
    ["行业常见实务", "常见的关账组织和控制方法", "结合规模、系统和报表时限调整"],
    ["企业核算口径", "关账天数、重要性阈值、暂估方法、审批层级等内部政策", "形成书面政策并保持一致"],
    ["案例假设", "为展示业务链而设定的数字和时间", "先理解链条，再替换成所在企业参数"],
], [2.7, 6.3, 7.0])
add_heading(doc, "阅读路径", 2)
add_bullets(doc, [
    "先看第1章全景图和第2章年度节奏，建立时间感。",
    "再看第3至第5章，掌握关账治理、日历和每一步的完成标准。",
    "第6至第10章分别对应采购暂估、成本、往来、税务和报表，是月结的五个重点模块。",
    "第11章把各模块放入一个制造企业案例；第14章可直接用于面试表达。",
    "配套Excel用于分工、打勾、记录差异和形成关账证据。",
])
add_page_break(doc)

add_heading(doc, "1 总账月结全景", 1)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(str(IMG), width=Cm(17.0))
cap = doc.add_paragraph("图1  总账月结全景框架"); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
for r in cap.runs: set_run_font(r, 8.5, False, "666666")
add_body(doc, "完整月结由六个连续环节构成：计划、截止、入账、核对、报告、关账。任何一个环节缺失，都可能出现账已经锁定但报表仍不完整的假关账。财务主管关注任务依赖、例外事项和复核证据；总账会计把各子模块结果汇入总账，完成判断、调整、勾稽与报表。")
add_table(doc, ["环节", "核心问题", "最低完成证据", "常见失败"], [
    ["计划", "谁在何时交付什么", "关账日历 RACI 上月问题清单", "只有口头通知 没有依赖关系"],
    ["截止", "本月业务是否完整且属于本期", "收发货 服务验收 盘点 银行截止清单", "把收到发票误当成确认费用的唯一条件"],
    ["入账", "子模块和调整分录是否完整", "接口批次控制表 分录清单 审批记录", "重复导入 漏提费用 手工分录无依据"],
    ["核对", "账账 账实 账单 账税能否解释", "余额调节表 对账单 差异跟踪", "只核金额 不核账龄性质和长期挂账"],
    ["报告", "报表能否勾稽并解释经营变化", "试算平衡表 报表 勾稽表 波动分析", "报表平了但业务逻辑不合理"],
    ["关账", "是否完成复核并控制后续改动", "签批 锁账记录 重开审批 复盘清单", "随意反结账或静默修改历史数据"],
], [2.1, 4.3, 5.2, 5.0])

add_heading(doc, "2 一年之中的工作安排", 1)
add_heading(doc, "2.1 三层时间结构", 2)
add_bullets(doc, [
    "日常层：审核凭证、处理接口异常、银行及往来核对、发票与税务资料、异常事项跟踪。日常不清，月末一定拥堵。",
    "月度层：执行关账日历，完成截止、计提、成本、对账、报表、分析和锁账。一般企业的高峰是D-3至D+7。",
    "年度与季度层：季末增加所得税预缴、资产减值和管理复盘；年末增加盘点、函证、审计、年度决算和汇算清缴。",
])
add_heading(doc, "2.2 年度工作日历和忙碌程度", 2)
annual = [
    ["1月", "极忙", "上年12月及年度硬关账 年初开账", "年末截止 盘点差异 奖金及费用计提 减值 关联方核对 审计PBC", "上年财务报表 年结底稿 期初余额"],
    ["2月", "中高", "审计与春节压缩月结", "审计询证及抽凭 年终奖个税 短工作月 月结日历提前", "审计资料 短周期月结包"],
    ["3月", "高", "年度审计 汇算准备 一季度预结", "税会差异 研发费用 资产损失 关联申报资料 一季度预测", "汇算底稿初稿 Q1预结"],
    ["4月", "极忙", "一季度关账与预缴 年报集中期", "季度深度复核 企业所得税预缴 审计调整落账 报表披露支持", "Q1报表 预缴申报勾稽"],
    ["5月", "极忙", "企业所得税年度汇算清缴收尾", "纳税调整 优惠及留存备查 申报表与财务报表核对", "年度纳税申报 汇算清缴档案"],
    ["6月", "高", "半年预结 盘点与预测", "存货和固定资产抽盘 减值复核 半年预计 全年滚动预测", "半年关账准备 预测更新"],
    ["7月", "极忙", "半年关账与二季度预缴", "半年报表 管理复盘 所得税预缴 审计或集团审阅", "半年报表 Q2申报 管理分析"],
    ["8月", "中", "清理半年遗留和流程改进", "长期挂账 暂估未冲 主数据 故障接口 自动化优化", "问题关闭清单 流程改进方案"],
    ["9月", "高", "三季度预结和下一年度预算准备", "预算假设 资本开支 人员计划 税负预测 年末风险清单", "Q3预结 预算底稿"],
    ["10月", "极忙", "三季度关账与预算", "Q3深度复核 所得税预缴 预算编制 年末盘点计划", "Q3报表 预算初稿 盘点计划"],
    ["11月", "高", "预算定稿和年结预演", "年末计提 减值测试 关联方确认 函证地址 PBC清单", "预算定稿 年结检查清单"],
    ["12月", "极忙", "年度截止和盘点", "收发货及服务截止 存货盘点 资产清查 奖金计提 减值 关账沟通", "截止证据 盘点资料 年结预报"],
]
add_table(doc, ["时间", "忙度", "主任务", "具体工作", "关键输出"], annual, [1.3, 1.5, 3.5, 7.2, 4.0], 8.0)
add_body(doc, "忙度是行业常见实务判断，不是统一规定。集团合并、上市披露、IPO、外资集团快报、出口退税、项目制收入、年度预算周期都会改变高峰。春节所在月份还会因工作日减少而提前截止。")
add_heading(doc, "2.3 法定期限与公司关账期限不要混淆", 2)
add_table(doc, ["事项", "法定基线", "月结中的安排", "性质"], [
    ["企业所得税预缴", "月度或季度终了后15日内申报预缴", "季结后先完成税会利润和调整测算 留出复核与缴款时间", "法律明确"],
    ["企业所得税年度汇算", "年度终了后5个月内完成申报和结清税款", "通常1至5月滚动准备 不能等5月才开始", "法律明确"],
    ["个人所得税扣缴", "次月15日内缴库并报送扣缴申报表", "工资关账后核对应付职工薪酬 银行和申报数据", "法律明确"],
    ["增值税", "纳税期限按主管税务机关核定 月季纳税通常期满后15日内申报", "月结同步完成销项 进项 账载税额 发票和申报表核对", "法律明确且以核定税期为准"],
    ["公司D+5或D+7关账", "法律通常不规定企业内部必须第几天关账", "根据报表使用时间和组织能力设置并经批准", "企业核算管理口径"],
], [3.0, 5.3, 6.4, 2.0], 8.2)

add_heading(doc, "3 月结开始前的组织设计", 1)
add_heading(doc, "3.1 先定义什么叫关完", 2)
add_body(doc, "财务主管应把关账完成定义成可验证的条件，而不是凭证都做完。最低标准是：业务截止完成；全部接口批次完整；重大计提、摊销、折旧、成本与税费入账；重点科目差异在阈值内或已有批准处理；报表勾稽通过；重大波动有解释；复核签批完成；期间被锁定。")
add_heading(doc, "3.2 关账日历的六个字段", 2)
add_table(doc, ["字段", "必须回答的问题", "示例"], [
    ["任务", "具体做到什么程度", "完成4月采购暂估并取得采购经理确认"],
    ["责任人", "谁制作 谁复核 谁批准", "应付会计制作 总账复核 财务主管批准"],
    ["截止时间", "按工作日还是自然日", "D+2 12:00前"],
    ["输入", "依赖谁交付什么", "未开票入库清单 服务验收单 合同"],
    ["输出证据", "完成后留下什么", "暂估明细 对账差异 分录号 审批记录"],
    ["例外路径", "晚到数据如何处理", "超过阈值升级 估计入账 次月真冲调整"],
], [2.3, 6.2, 8.0])
add_heading(doc, "3.3 财务主管与总账会计的分工", 2)
add_table(doc, ["角色", "主要责任", "不能只做什么"], [
    ["财务主管", "制定日历和口径 协调跨部门 复核重大判断 控制重开期间 对外报告负责", "不能只催进度不审质量"],
    ["总账会计", "汇总子模块 完成调整分录 科目核对 报表勾稽 差异跟踪 形成关账包", "不能只把分录录进系统"],
    ["业务及子模块负责人", "按截止时间提供真实完整数据 解释差异 确认暂估和业务状态", "不能把发票是否收到当成唯一判断"],
    ["复核人", "检查依据 计算 期间 科目 税务和报表影响 追问异常", "不能只在封面签字"],
], [2.7, 8.3, 5.5])

add_heading(doc, "4 一次标准月结日历", 1)
add_body(doc, "以下采用自然月末D、内部D+7锁账。D+表示月末后的工作日。它是可执行范本，不是法律统一期限。小型企业可能D+10，上市或集团快报可能D+3至D+5。")
timeline = [
    ["D-5至D-3", "预结", "发关账通知 清理上月问题 核对未处理接口 预收集暂估和计提资料", "关账日历 未决事项 责任人确认"],
    ["D-2至D", "业务截止", "冻结关键收发货时点 取得服务验收 盘点现金存货 采集工资资产和银行截止数据", "截止清单 盘点记录 验收证据"],
    ["D+1", "子模块完成", "采购销售费用资金薪酬固定资产库存完成入账 接口控制总额核对", "子模块关账确认 接口批次表"],
    ["D+2", "计提和调整", "暂估采购及费用 计提薪酬利息税费 折旧摊销 外币重估 收入截止", "调整分录清单及依据"],
    ["D+3", "成本与模块对账", "完成生产成本分配 完工入库 销售成本结转 核对子模块与总账", "成本计算表 模块对账表"],
    ["D+4", "全科目核对", "完成银行 往来 存货 资产 税费 贷款 关联方等余额核对", "科目余额调节表 差异台账"],
    ["D+5", "试算与分析", "出试算平衡和初版报表 检查异常余额 同比环比 预算差异 毛利税负", "初版报表 波动解释"],
    ["D+6", "报告复核", "完成三表勾稽 现金流量表 附注或集团包 管理层复核重大事项", "报表勾稽表 复核意见"],
    ["D+7", "签批锁账", "清空重大未决事项 完成签批 锁定会计期间 发布正式报表", "关账签批 锁账记录 正式报表"],
    ["D+8至D+10", "关账后", "经营分析 复盘晚报和手工调整 更新长期问题和下月行动", "管理分析 关账复盘"],
]
add_table(doc, ["时间", "阶段", "核心动作", "完成证据"], timeline, [2.1, 2.2, 8.2, 5.0], 8.3)

add_heading(doc, "5 月结逐步操作", 1)
add_heading(doc, "5.1 D-5至D-3 预结", 2)
add_steps(doc, [
    "复制上月关账清单并检查未关闭问题，不能每月从空白开始。",
    "发出关账通知，明确仓库、采购、销售、人力、业务负责人和IT的交付时间。",
    "预跑应收账龄、应付未清项、GR IR或暂估、负库存、在建工程、长期待摊和税费余额，提前暴露异常。",
    "确定当月重大事项：大额合同、退货、诉讼、资产投产、融资、并购、关联交易和一次性政策变化。",
    "确认汇率、成本分配参数、折旧期间、税码和会计期间状态，避免结账日才修主数据。",
])
add_heading(doc, "5.2 D-2至D0 业务截止", 2)
add_body(doc, "截止性的关键是经济业务属于哪个期间。发票只是证据之一。货物已验收入库但未收票，通常仍应确认存货和负债；服务已在本月接受且能可靠估计，通常应计入本月费用。反过来，提前收到发票并不当然意味着全部属于本月。")
add_bullets(doc, [
    "采购：取得最后收货单号、未开票入库、在途物资、服务验收、退货及价格差异。",
    "销售：取得最后发货单号、客户签收或验收、退货、折让、寄售和未开票履约情况。",
    "存货：记录停单时间，控制倒签单据，盘点差异由仓库和财务共同确认。",
    "资金：保存月末银行余额、未达账、受限资金和未入账费用。",
    "薪酬资产：锁定人头、考勤、奖金依据、投产日期、处置和停用状态。",
])
add_heading(doc, "5.3 D+1至D+3 入账和成本", 2)
add_body(doc, "先让子模块在一致的截止点完成，再向总账传输。每个接口至少核对批次数、记录数、借贷金额和错误日志。对手工分录实行制作与复核分离，大额、非常规和跨期分录附完整依据。")
add_bullets(doc, [
    "自动类：折旧、摊销、外币重估、标准接口、周期性分录。自动生成不等于无需复核。",
    "估计类：采购暂估、服务费、水电费、奖金、返利、利息、质量索赔和预计信用损失。记录假设、数据来源和次月冲回方式。",
    "成本类：先完成生产订单或成本中心归集，再分配制造费用、计算完工与在产品，最后结转销售成本。不要先凭毛利目标倒挤成本。",
])
add_heading(doc, "5.4 D+4 全科目核对", 2)
add_table(doc, ["科目模块", "对什么", "重点检查", "月结结论"], [
    ["货币资金", "银行对账单 网银 现金盘点", "未达账长期挂账 受限资金 重复付款", "账实一致或差异已解释"],
    ["应收及合同余额", "销售系统 客户对账 回款", "负数余额 账龄 逾期 争议 折让", "余额性质与减值合理"],
    ["应付及采购暂估", "采购入库 发票 供应商对账", "重复暂估 已收票未冲 长期借方", "负债完整且无重复"],
    ["存货与成本", "仓储 生产 成本系统 盘点", "负库存 呆滞品 成本异常 盘点差异", "数量金额和成本逻辑一致"],
    ["固定资产与在建工程", "资产台账 项目清单 实物", "已投产未转固 闲置 报废 折旧起点", "台账总账一致并复核减值"],
    ["职工薪酬", "工资表 社保公积金 个税 银行", "应付余额异常 人员差异 奖金依据", "计提 发放 申报可勾稽"],
    ["税费", "发票平台 申报表 完税凭证", "销项漏计 进项异常 税会差异", "账票表税款一致或有调节"],
    ["借款及利息", "合同 银行函证 还款计划", "利息漏提 短长分类 受限条款", "本金利息与合同一致"],
    ["权益及关联方", "董事会文件 内部对账", "未分配利润衔接 内部差异 长期挂账", "批准依据完整 双边一致"],
], [3.0, 4.0, 6.2, 4.3], 8.1)
add_heading(doc, "5.5 D+5至D+7 报表 复核和锁账", 2)
add_steps(doc, [
    "从最终试算平衡表生成报表，检查借贷平衡、资产等于负债加所有者权益以及报表映射完整。",
    "做环比、同比、预算和单位指标分析。总账会计先解释数字，财务主管再挑战业务合理性。",
    "完成现金流量表和附注或集团报表包的勾稽，确保使用的是同一版总账数据。",
    "对重大估计、异常余额、未决差异和截止例外取得书面批准。未决事项应有金额、影响、责任人和完成日。",
    "完成制单、复核、批准签批后锁定期间。需要重开时走正式申请，说明原因、影响、分录和重新出表范围。",
])

add_heading(doc, "6 采购暂估", 1)
add_heading(doc, "6.1 业务链", 2)
add_body(doc, "业务发生 → 采购订单和合同 → 收货或服务验收 → 采购系统或ERP收货 → 判断控制权和义务是否已发生 → 未收票暂估 → 次月发票匹配和冲销 → 应付与总账核对 → 月结检查未开票和长期暂估 → 存货 负债 成本及税务分别反映。")
add_heading(doc, "6.2 制造企业案例", 2)
add_body(doc, "案例假设：4月28日收到原材料，合同不含税价100万元，适用税率假设为13%，材料已验收入库并可使用；4月未收到发票，5月10日收到合规发票，6月付款。企业采用不含税暂估、次月初红冲的内部政策。")
add_table(doc, ["时点", "原始单据与系统", "会计判断与分录", "税务与月结"], [
    ["4月收货", "合同 PO 送货单 验收单 入库单 ERP收货记录", "货物已经验收入库并形成付款义务。借 原材料 1,000,000；贷 应付账款 暂估 1,000,000", "未取得可抵扣凭证时不在本例确认进项抵扣。月末以未开票收货清单逐项确认数量价格"],
    ["5月收票", "专用发票 发票查验记录 PO收货发票三单匹配", "先按政策冲回暂估，再按发票和匹配结果正式入账。借 原材料 1,000,000；借 应交税费 应交增值税 进项税额 130,000；贷 应付账款 1,130,000", "是否抵扣还要满足现行增值税法及配套规则、凭证和用途条件。核对发票平台与进项税额"],
    ["6月付款", "付款申请 审批 银行回单 供应商对账", "借 应付账款 1,130,000；贷 银行存款 1,130,000", "核对供应商未清项和银行流水，保存审批及付款证据"],
], [2.4, 5.0, 6.2, 4.2], 8.0)
add_heading(doc, "6.3 常见差异和处理", 2)
add_bullets(doc, [
    "数量差异：按实际验收数量确认，采购和仓库查明短溢。",
    "价格差异：先按合同、最近采购价或经批准估计暂估；收票后把差额按企业政策计入存货、成本或价格差异科目。",
    "已耗用后收票：差额是否追溯调整存货和主营业务成本，取决于金额、库存去向和企业成本政策，不宜一律进当期费用。",
    "重复暂估：当月暂估、次月冲回、正式发票三者必须有唯一匹配键，如PO加行项目加收货单。",
    "所得税：会计确认和税前扣除凭证是两个问题。年度汇算时按税法和税前扣除凭证规定检查取得凭证时点和更正路径。",
])

add_heading(doc, "7 成本归集和结转", 1)
add_heading(doc, "7.1 正确顺序", 2)
add_steps(doc, [
    "核对原材料、在制品、产成品数量，处理负库存、倒冲和跨期领退料。",
    "把直接材料、直接人工和制造费用归集到产品、订单、工序或成本中心。",
    "检查制造费用分配基础是否仍合理，区分正常产能和异常停工等事项。",
    "确认完工数量、在产品完工程度和约当产量等估计，形成复核底稿。",
    "结转完工产品，再依据销售出库结转销售成本，最后分析单位成本和毛利异常。",
])
add_heading(doc, "7.2 数字案例", 2)
add_body(doc, "4月期初在产品50万元；本月投入直接材料320万元、直接人工50万元、制造费用82万元；期末在产品经盘点和完工程度测算为62万元。可归集成本为502万元，完工产品成本为440万元。")
add_table(doc, ["项目", "金额万元", "计算或分录"], [
    ["期初在产品", "50", "上月生产成本期末余额"],
    ["本月直接材料", "320", "领料单和生产订单归集"],
    ["本月直接人工", "50", "工资和工时分配"],
    ["本月制造费用", "82", "100万元电费中制造部分10万元 折旧22万元 其他制造费用50万元"],
    ["可归集成本", "502", "50加320加50加82"],
    ["减 期末在产品", "62", "盘点数量乘完工程度及成本方法"],
    ["完工产品成本", "440", "借 库存商品 440；贷 生产成本 440"],
    ["本月销售成本", "280", "假设已售产品成本280。借 主营业务成本 280；贷 库存商品 280"],
], [5.3, 2.6, 8.0], 8.5)
add_body(doc, "风险点在于成本不是为了得到目标毛利而倒挤出来。财务应核对投入产出、单位耗用、工时、产量、在产品估计和分配率。ERP成本运行失败时，应先查主数据、订单状态和负库存，不能用一笔总额手工分录掩盖系统差异。")

add_heading(doc, "8 往来清理", 1)
add_heading(doc, "8.1 月结不只是核对余额", 2)
add_body(doc, "应收应付的余额可能在数学上与明细账一致，但仍存在负数余额、客户供应商串户、预收预付未重分类、已核销未清项、长期无业务挂账和争议款。总账会计要同时检查金额、账龄、性质、交易对手和可收可付性。")
add_table(doc, ["现象", "可能原因", "核查证据", "处理路径"], [
    ["应收出现贷方", "预收款 未匹配收款 红字发票", "合同 银行回单 客户明细", "匹配或按报表口径重分类 不直接冲收入"],
    ["应付出现借方", "预付款 重复付款 退货未结算", "付款审批 采购订单 供应商对账", "重分类并追踪业务状态"],
    ["长期暂估", "发票未到 订单关闭失败 供应商争议", "PO 收货 发票 合同 对账函", "逐项确认义务 取消无效项目 补票或调整"],
    ["关联方差异", "截止不同 汇率不同 对方错账", "双边明细 交易编号 汇率", "逐笔匹配并在集团截止前确认"],
    ["小额长期挂账", "员工离职 尾差 缺资料", "原始凭证 人员信息 审批", "按权限清理 不能长期用小额规避判断"],
], [3.2, 5.0, 5.2, 5.2])
add_heading(doc, "8.2 清理优先级", 2)
add_bullets(doc, [
    "金额大或账龄长；",
    "负数余额和性质异常；",
    "关联方和内部交易；",
    "可能涉及收入、成本、税费或现金流量表分类；",
    "需要业务、法务或管理层作判断的争议款。",
])

add_heading(doc, "9 账税核对", 1)
add_heading(doc, "9.1 账 票 表 款四层核对", 2)
add_table(doc, ["层次", "核对内容", "常见差异", "完成证据"], [
    ["账", "总账和明细账中的销项 进项 转出 已交 未交税额", "科目用错 跨期 漏转", "税费科目调节表"],
    ["票", "开具和取得发票 发票平台状态 用途确认 红字发票", "已开票未入账 未开票已纳税 异常凭证", "发票台账 平台导出"],
    ["表", "纳税申报表及附表", "申报口径与会计口径不同", "申报表到账务桥接表"],
    ["款", "实际缴税 退税 抵减 留抵", "缴款未入账 退税在途", "完税凭证 银行回单"],
], [2.0, 6.0, 5.0, 4.5])
add_body(doc, "账税一致不是要求每个数字机械相等，而是要求差异有合法原因、有计算过程、可追溯到单据和申报表。收入确认时点、增值税纳税义务时点、企业所得税税前扣除和会计费用确认可能不同，需要建立桥接表。")
add_heading(doc, "9.2 月结税务检查", 2)
add_bullets(doc, [
    "增值税：收入与开票、未开票收入、销项税额、进项用途、转出、红字、留抵和缴款。",
    "企业所得税：季度税会利润、纳税调整、弥补亏损、优惠条件和预计税负；年度1至5月滚动汇算。",
    "个人所得税：工资表、应付职工薪酬、银行发放额、人员信息和扣缴申报勾稽。",
    "印花税：应税凭证台账、计税金额、税率和申报周期；按次、按季或按年事项分别管理。",
    "其他税费：房产税、城镇土地使用税、附加税费等按企业所在地、资产和主管税务机关口径管理，不能套用一张全国固定日历。",
])

add_heading(doc, "10 报表勾稽和分析", 1)
add_heading(doc, "10.1 基础勾稽", 2)
add_table(doc, ["检查", "应达到的结果", "不通过时先查什么"], [
    ["试算平衡", "总账借方发生额和贷方发生额平衡", "接口遗漏 重复导入 分录方向"],
    ["资产负债表", "资产等于负债加所有者权益", "科目映射 未分配利润 报表取数版本"],
    ["利润衔接", "本期利润与利润表及权益变动衔接", "结转期间 报表公式 调整分录"],
    ["现金流量表", "期末现金及现金等价物与资产负债表口径调节后衔接", "受限资金 外币折算 非现金事项 分类"],
    ["明细与附注", "明细合计等于报表项目", "重分类 抵销 映射 更新版本"],
    ["单体与集团包", "同一口径下可桥接", "会计政策差异 集团调整 内部交易 汇率"],
], [3.0, 7.7, 7.0])
add_heading(doc, "10.2 分析性复核", 2)
add_body(doc, "分析性复核不是在报表后附一句收入上升。应把异常定位到业务驱动：销量、售价、产品结构、材料价格、产量和产能利用率、人员、汇率、信用损失、一次性项目。无法解释的重大波动必须回到账务和业务证据，不要为了让波动消失而调账。")
add_bullets(doc, [
    "收入：数量乘价格，再看渠道、客户和履约截止。",
    "毛利：售价、材料、人工、制造费用分摊、产品结构和存货跌价。",
    "费用：人数、薪酬、项目、合同期限、计提冲回和资本化。",
    "营运资金：应收周转、逾期、库存天数、采购信用期和现金转换周期。",
    "税负：会计收入与应税销售额、利润与应纳税所得额的桥接。",
])

add_heading(doc, "11 制造企业4月完整月结案例", 1)
add_heading(doc, "11.1 场景", 2)
add_body(doc, "甲制造公司采用D+7关账。4月发生以下事项：月末原材料已入库未收票100万元；本月电费预计12万元，其中制造10万元、管理2万元；工资80万元，其中生产50万元、管理18万元、销售12万元；折旧30万元，其中制造22万元、管理5万元、销售3万元；期初在产品50万元，本月直接材料320万元，其他制造费用50万元，期末在产品62万元；销售商品控制权已转移，不含税收入400万元，假设适用13%税率；对应销售成本280万元；利息5万元；银行手续费0.2万元。")
add_heading(doc, "11.2 从单据到报表", 2)
case_rows = [
    ["采购暂估", "合同 PO 入库单 未开票清单", "ERP收货 暂估匹配", "借原材料100 贷应付暂估100", "进项抵扣另按凭证和用途判断", "应付暂估与收货清单核对"],
    ["电费计提", "电表 合同 历史单价 部门确认", "周期性计提分录", "借制造费用10 管理费用2 贷应付或其他应付款12", "收票与扣除凭证分开管理", "次月真冲并分析估计差"],
    ["工资计提", "工资表 人员和工时 审批", "薪酬模块至总账", "借生产成本50 管理18 销售12 贷应付职工薪酬80", "与个税扣缴申报 银行发放核对", "人员 金额 申报 付款四项勾稽"],
    ["折旧", "资产卡片 投产和使用状态", "固定资产模块折旧", "借制造22 管理5 销售3 贷累计折旧30", "税法折旧形成的差异另建台账", "台账与总账 增减变动核对"],
    ["成本结转", "领料 工时 产量 在产品盘点", "成本模块计算", "完工成本440；销售成本280", "企业所得税关注成本真实性和凭证", "成本模块与总账 库存与仓库核对"],
    ["销售", "合同 发货 签收或验收", "销售出库 开票或未开票税务处理", "借应收452 贷收入400 销项税52", "示例假设纳税义务已发生", "应收 销售 发票 销项四方核对"],
    ["利息和手续费", "借款合同 计息表 银行流水", "手工或资金模块", "借财务费用5.2 贷应付利息5 银行存款0.2", "扣除条件和凭证另核", "合同 计提 付款 银行对账"],
]
add_table(doc, ["事项", "原始单据", "系统操作", "会计分录 万元", "税务", "月结核对"], case_rows, [2.2, 3.4, 3.0, 4.7, 3.6, 4.0], 7.4)
add_heading(doc, "11.3 D+7如何组织", 2)
add_steps(doc, [
    "D-3由总账发关账通知，仓库、采购、生产、人力和销售分别确认最后业务时点。",
    "D0仓库导出最后收货和发货序号，盘点在产品；业务部门确认已接受但未开票的服务。",
    "D+1核对各子模块批次、记录数和金额，修复失败接口。",
    "D+2完成采购暂估、电费、工资、折旧、利息和销售截止分录。",
    "D+3完成成本计算，核对440万元完工成本、62万元在产品和280万元销售成本的系统逻辑。",
    "D+4完成银行、应收、应付、存货、资产、薪酬、税费和借款对账，并把未解决差异登记到问题台账。",
    "D+5生成报表，检查毛利和费用波动。案例数据并非完整公司账套，因此不能仅用列示分录推导完整净利润。",
    "D+6财务主管复核重大暂估、在产品估计、收入截止、税务桥接和报表勾稽。",
    "D+7签批并锁账；如果5月收到4月原材料发票，按内部政策冲回暂估、正式入账并核对差额。",
])

add_heading(doc, "12 行业和企业差异", 1)
add_table(doc, ["类型", "月结重点", "常见额外判断", "系统依赖"], [
    ["制造业", "存货 生产成本 在产品 完工和销售成本", "正常与异常损耗 产能分配 呆滞跌价", "仓储 生产 MES 成本模块"],
    ["科技和项目制企业", "合同履约 项目收入 研发费用 人工分摊", "履约义务 验收 资本化条件 项目预计总成本", "项目 工时 合同和研发系统"],
    ["贸易企业", "采购入库 销售出库 在途 商品成本和信用", "物流费用归属 返利 退货 渠道库存", "采购 仓储 销售 物流"],
    ["集团企业", "内部往来 交易抵销 统一口径和报表包", "合并范围 会计政策 汇率 少数股东", "合并系统和集团主数据"],
], [3.0, 5.2, 5.4, 4.4])
add_body(doc, "企业应把关账天数、重要性阈值、暂估方法、汇率来源、成本分配、在产品估计、报表重分类和重开期间审批写入制度或关账手册。改变口径前评估原因和影响，经适当批准后记录生效期间，避免为了当期结果随意变化。")

add_heading(doc, "13 ERP和自动化", 1)
add_heading(doc, "13.1 数据流和关账顺序", 2)
add_body(doc, "典型顺序是业务系统完成截止，采购、销售、仓储、薪酬、固定资产等子模块先关账，再运行成本和总账，最后生成报表。SAP中常见FI、MM、SD、AA、CO之间存在依赖；金蝶、用友也有相似的业务单据到凭证和子模块结账逻辑。具体菜单、期间控制和反结账规则取决于版本与配置。")
add_bullets(doc, [
    "接口控制：批次号、记录数、金额控制总额、错误日志、重跑标识。",
    "期间控制：业务日期、过账日期、子模块期间和总账期间保持一致。",
    "权限控制：制单、复核、过账、修改主数据和重开期间分离。",
    "审计轨迹：手工分录来源、修改记录、审批、导入文件和报表版本可追溯。",
])
add_heading(doc, "13.2 哪些适合自动化", 2)
add_table(doc, ["工作", "适合工具", "自动化边界", "人工复核"], [
    ["多实体或多银行数据合并", "Power Query Python RPA", "字段和主数据标准化后自动合并", "检查完整性 余额和重复项"],
    ["子模块对总账核对", "Excel Power Query Python", "规则清晰 唯一键稳定的匹配", "判断性质异常和长期差异"],
    ["周期性分录", "ERP自动分录 RPA", "固定规则且经批准的折旧摊销", "参数 期间 异常波动"],
    ["关账催办和状态汇总", "流程平台 AI Agent", "读取任务状态和发送提醒", "升级例外和资源冲突"],
    ["波动检测", "Power BI Python AI", "发现异常和生成候选原因", "结合业务证据作最终解释"],
    ["会计估计和税务判断", "AI辅助检索和底稿", "不得直接替代专业判断和批准", "合同 事实 法源 重要性和责任"],
], [3.5, 3.8, 6.0, 4.7])

add_heading(doc, "14 面试表达", 1)
add_heading(doc, "14.1 面试官为什么问月结", 2)
add_body(doc, "这个问题同时检验六项能力：是否理解业务到总账的数据链；能否做截止和会计估计；能否组织跨部门交付；能否核对全科目；能否从试算平衡走到三张报表；能否控制关账后的修改。只有分录清单，通常不足以证明能独立负责月结。")
add_heading(doc, "14.2 两分钟合格答案", 2)
add_body(doc, "我会先根据报表提交日倒排关账日历，明确每项任务的制作人、复核人、输入、输出和升级路径。月末前完成预结和业务截止，重点盯收发货、服务验收、工资、资产和库存。D+1至D+3让采购、销售、费用、薪酬、固定资产和库存子模块完成并核对接口控制总额，同时做暂估、计提、摊销、折旧、外币重估、收入截止和成本结转。D+4完成银行、往来、存货、资产、税费、借款和关联方等全科目核对；D+5出试算平衡和初版报表，做同比、环比、预算和业务指标分析；D+6完成三表及集团包勾稽和主管复核；D+7清理重大未决事项、签批并锁账。发现差异时我会记录金额、期间、根因、报表税务影响、责任人和完成日，重大事项及时升级，重开期间走正式审批。")
add_heading(doc, "14.3 追问时要体现的能力", 2)
add_table(doc, ["追问", "合格答案重点", "容易失分"], [
    ["发票没到为什么能入账", "权责发生制和截止性 业务证据 暂估方法 次月匹配 税务分开", "把发票当唯一入账条件"],
    ["成本算不出来怎么办", "查负库存 订单状态 主数据 分配基础 接口错误 保留审计轨迹", "直接倒挤成本或手工覆盖"],
    ["往来对不上怎么办", "逐笔匹配 截止 汇率 串户 未达 差异责任和时限", "只说让业务确认"],
    ["报表平了是否能关账", "还需分析异常 重大估计 税务勾稽 复核和锁账", "把借贷平衡等同于正确"],
    ["业务晚交资料怎么办", "预结 催办 升级 估计 重要性和例外审批 次月复盘", "不留记录地等资料或直接漏记"],
], [4.0, 8.5, 4.2])

add_heading(doc, "15 关账完成标准和常见红旗", 1)
add_heading(doc, "15.1 完成标准", 2)
add_bullets(doc, [
    "任务清单全部完成或例外已批准；",
    "所有子模块与总账的差异为零或在已批准阈值内且有解释；",
    "重大手工分录有制单、复核、依据和反转安排；",
    "报表勾稽通过，重大波动有业务证据；",
    "税务申报数据可从账务和发票资料桥接；",
    "锁账、报表版本和关账包归档完成。",
])
add_heading(doc, "15.2 需要立即追查的红旗", 2)
add_bullets(doc, [
    "用一笔大额手工分录把子模块差异清零；",
    "月底集中倒签收货、发货或验收日期；",
    "长期暂估持续增加，正式发票入账时未冲原暂估；",
    "负库存、负数应收应付和关联方差异长期无人负责；",
    "收入、毛利、税负或现金流异常却没有业务解释；",
    "锁账后频繁无审批修改，或不同报表使用不同试算平衡版本；",
    "AI或脚本自动生成分录但没有规则版本、日志和人工复核。",
])

add_heading(doc, "16 学习和实操路线", 1)
add_table(doc, ["阶段", "目标", "练习", "通过标准"], [
    ["第1周", "看懂年度和月结全景", "用Excel把一家假设公司的责任人和日期补齐", "能说明六个环节和主要依赖"],
    ["第2周", "掌握暂估和往来", "完成采购暂估案例及5类往来异常处理", "能区分会计确认 发票 税务和付款"],
    ["第3周", "掌握成本与存货", "复算制造案例并解释在产品和毛利变化", "能从数量 工时 费用和分配率查差异"],
    ["第4周", "完成模拟关账", "填完关账清单 科目核对 问题跟踪和报表勾稽", "能在10分钟内汇报进度 风险和结论"],
], [2.5, 5.0, 6.5, 4.5])

add_heading(doc, "17 官方依据和版本", 1)
add_body(doc, "以下法源用于确定会计核算、报表责任和税务期限的强制边界，核验日期为2026年9月9日。具体业务仍需结合交易事实、主体、主管税务机关核定税期及后续政策。")
sources = [
    ("中华人民共和国会计法 2024年", "https://kjs.mof.gov.cn/zt/kjfxcgc/kjfqw/202408/t20240814_3941788.htm", "实际经济业务 会计资料 财务报告责任 内部控制"),
    ("企业会计准则 基本准则", "https://www.mof.gov.cn/zcsjtsgb/gztsgb/201407/t20140723_3578373.htm", "权责发生制 可靠性 一致性 实质重于形式 重要性 谨慎性"),
    ("企业会计准则第14号 收入", "https://kjs.mof.gov.cn/zt/kjzzss/kuaijizhunzeshishi/201709/t20170907_2694006.htm", "收入确认与履约判断"),
    ("企业会计准则第28号 会计政策会计估计变更和差错更正", "https://kjs.mof.gov.cn/zt/kjzzss/kuaijizhunzeshishi/200806/t20080618_46220.htm", "关账后差错与前期差错"),
    ("中华人民共和国企业所得税法", "https://fgk.chinatax.gov.cn/zcfgk/c100009/c5193018/content.html", "纳税年度 预缴和年度汇算期限"),
    ("中华人民共和国个人所得税法", "https://www.chinatax.gov.cn/n810219/n810744/n3752930/n3752974/c3970366/content.html", "扣缴申报和汇算期限"),
    ("中华人民共和国增值税法", "https://shanghai.chinatax.gov.cn/sjtax/ztzl/yshj/ldjj/202412/t474700.html", "2026年起增值税基本制度和申报期限"),
    ("中华人民共和国增值税法实施条例", "https://fgk.chinatax.gov.cn/zcfgk/c100010/c5246349/content.html", "2026年起配套实施规则"),
    ("中华人民共和国印花税法", "https://fgk.chinatax.gov.cn/zcfgk/c100009/c5193058/content.html", "应税凭证和申报期限"),
    ("企业所得税税前扣除凭证管理办法", "https://fgk.chinatax.gov.cn/zcfgk/c100012/c5194804/content.html", "税前扣除凭证和补救路径"),
]
for title, url, use in sources:
    p = doc.add_paragraph(style="List Bullet")
    add_hyperlink(p, title, url)
    r = p.add_run("  " + use); set_run_font(r, 9.5)

add_heading(doc, "18 知识库归属", 1)
add_body(doc, "主模块：M02 总账与月结季结年结。直接关联M03全科目核对与往来清理、M04采购应付与费用、M05销售应收与收款、M06存货与成本、M07薪酬资产与专项核算、M08资金与融资、M09发票与税务、M10财务报表与合并报表、M11预算经营分析与财务BP、M12 ERP内控与AI自动化。")
add_body(doc, "后续深化顺序建议为：采购暂估 → 成本结转 → 往来清理 → 账税核对 → 报表勾稽。每个专题继续沿业务发生、原始单据、系统操作、会计判断、分录、税务、对账、月结、报表影响和风险点展开。")

# document settings to refresh fields
settings = doc.settings._element
upd = OxmlElement("w:updateFields"); upd.set(qn("w:val"), "true"); settings.append(upd)
doc.core_properties.title = "总账会计年度工作与月结实战手册"
doc.core_properties.subject = "M02-001 总账月结"
doc.core_properties.author = "企业财务实操知识库"
doc.save(DOCX)
print(DOCX)
