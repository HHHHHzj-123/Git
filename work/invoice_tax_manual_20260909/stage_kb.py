from pathlib import Path

base = Path(r"C:\Users\HZJ\Desktop\Git\work\invoice_tax_manual_20260909")
stage = base / "kb_stage"
topic = stage / "09-发票与税务" / "01-企业发票管理与税务申报"
topic.mkdir(parents=True, exist_ok=True)

readme = """# M09-001 企业发票管理与税务申报

建立日期：2026-09-09。政策核验截止：2026-09-09。

## 学习目标

本专题服务于从审计转向总账会计或财务主管的实操训练，围绕“业务—单据—发票—入账—应交税费—月结—申报—缴税—核对—归档”建立完整工作链条。

## 文件与阅读顺序

1. `企业发票与税务申报实操.xmind`：先看三张工作表，建立税务全景、账票税核对和异常闭环框架。
2. `企业发票管理与税务申报实操手册.docx`：按章节学习。第1—7章建立税种、岗位和发票基础；第8—10章深入采购进项、销售销项和增值税申报；第12—15章处理企业所得税、印花税与个税；第17—24章覆盖核对、异常、系统、自动化和完整月份案例。
3. `企业税务月结与申报检查清单.xlsx`：用于执行练习。先维护参数和税务日历，再按月度、季度、年度清单完成事项，使用账票税核对表、暂估台账和异常整改台账形成闭环。

## 案例口径

贯穿案例为一般纳税人制造业/硬件科技企业“远澜智能设备（苏州）有限公司”，以2026年6月和第二季度为教学期间。案例金额主要以万元列示；城市维护建设税及教育费附加的合计比例仅作为案例参数，实际工作应按企业所在地、纳税人身份及当期有效政策确认。

## 使用边界

手册明确区分企业会计准则要求、税法要求、行业常见实务和企业内部政策。申报期限、优惠资格、扣除凭证和地方办理口径可能变化，实际申报前应再次核对主管税务机关和电子税务局当期提示。

## 建议练习

以Excel中的2026年6月案例为底稿，逐笔追溯销项46.80万元、取得进项21.58万元、进项转出2.60万元、可抵扣进项18.98万元、期初留抵3.00万元和应纳增值税24.82万元的来源；再从累计会计利润170.00万元追溯至第二季度应补企业所得税18.05万元。
"""
(topic / "00-阅读说明.md").write_text(readme, encoding="utf-8")

nav_src = Path(r"C:\Users\HZJ\Desktop\财务管理实操\00-知识库与学习导航\README.md")
nav = nav_src.read_text(encoding="utf-8-sig")
entry = "- [M09-001 企业发票管理与税务申报](../09-发票与税务/01-企业发票管理与税务申报/00-阅读说明.md)：全年税务日历、发票全生命周期、采购进项、销售销项、增值税申报、所得税、印花税、个税、账票税核对、异常整改、ERP与AI自动化。\n"
anchor = "- [M12-001 会计职业风险]"
if "M09-001 企业发票管理与税务申报" not in nav:
    pos = nav.index(anchor)
    nav = nav[:pos] + entry + nav[pos:]

section_anchor = "**M12-001：会计职业风险。**"
topic_para = "**M09-001：企业发票管理与税务申报。** 资料已建立，包含3张工作表、971个主题的XMind导图，51页Word手册，以及11页签Excel执行清单；覆盖全年税务日历、数电发票、采购与进项、销售与销项、增值税完整申报、企业所得税、印花税、个税、账票税核对、15类异常整改、ERP数据流、AI自动化和制造业/硬件科技企业完整月份案例。当前状态为资料已建立、待按案例复算与模拟申报。\n\n"
if "**M09-001：企业发票管理与税务申报。**" not in nav:
    pos = nav.index(section_anchor)
    nav = nav[:pos] + topic_para + nav[pos:]

start = nav.index("## 下一学习主题")
end = nav.index("## 维护方式", start)
next_section = """## 下一学习主题

**M09-002：增值税申报与账票税核对专项练习。**

拟使用M09-001的Excel底稿，从销售台账、开票数据、未开票收入、收入明细账和销项税开始，再连接收票、用途判断、勾选确认、进项转出、留抵和申报表，完成一套可复算、可追溯的模拟申报。

"""
nav = nav[:start] + next_section + nav[end:]
(stage / "README.md").write_text(nav, encoding="utf-8")

ledger_src = Path(r"C:\Users\HZJ\Desktop\财务管理实操\00-知识库与学习导航\02-学习与问题台账.md")
ledger = ledger_src.read_text(encoding="utf-8-sig")
hist_row = "| 2026-09-09 | 建立M09-001企业发票管理与税务申报专题 | 3张工作表、971个主题的XMind；51页Word手册；11页签Excel；全年税务日历、发票全生命周期、采购进项、销售销项、增值税申报、所得税、印花税、个税、15类异常整改、ERP与AI、完整月份案例 | 政策核验截至2026-09-09；地区口径、申报期限和优惠资格需在实际办理前复核；金额以万元为教学口径 |\n"
if "建立M09-001企业发票管理与税务申报专题" not in ledger:
    marker = "| 2026-09-09 | 建立M02-002"
    pos = ledger.index(marker)
    ledger = ledger[:pos] + hist_row + ledger[pos:]

progress_row = "| M09-001 | 企业发票管理与税务申报 | M09 | M02、M03、M04、M05、M07、M08、M12 | 资料已建立，待案例复算与模拟申报 | XMind全景导图、Word实操手册、11页签税务检查清单、完整月份案例 | 先看全景图和第10章，再复算Excel增值税、企业所得税及异常整改 |\n"
if "| M09-001 | 企业发票管理与税务申报 |" not in ledger:
    marker = "| M02-002 | 企业业务循环"
    pos = ledger.index(marker)
    ledger = ledger[:pos] + progress_row + ledger[pos:]

pending = "\nM09-001已按2026-09-09可检索的现行官方资料核验。涉及地方附加、印花税税源采集、税收优惠和具体申报操作时，应结合纳税人主管税务机关、所属期及电子税务局当期规则再次确认。\n"
marker = "M12-001已按2026-09-08"
if pending.strip() not in ledger:
    pos = ledger.index(marker)
    ledger = ledger[:pos] + pending + "\n" + ledger[pos:]
(stage / "02-学习与问题台账.md").write_text(ledger, encoding="utf-8")

print(topic)
print(stage / "README.md")
print(stage / "02-学习与问题台账.md")
