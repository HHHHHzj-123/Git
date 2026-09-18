from __future__ import annotations

import csv
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent
OUT = BASE / "deliverables"
PACK = OUT / "财务自动化实战项目包"

COMPANY = {
    "name": "华辰智能设备有限公司",
    "profile": "中国大陆增值税一般纳税人，约300名员工，生产工业智能终端及配套耗材。",
    "systems": "ERP覆盖采购、销售、存货、成本和总账，但银行、发票、月结与审计资料仍大量依赖Excel。",
    "volumes": "月销售订单约1,200笔、采购及费用发票约1,800张、8个银行账户、月银行流水约3,500笔。",
}

TOOLS = [
    ["Excel", "少量到中等数据、明确公式、需要人工复核", "预提、报表勾稽、检查清单", "复杂组合匹配和大量文件"],
    ["Power Query", "多文件导入、清洗、追加、合并、重复刷新", "银行流水、发票、应收应付", "复杂算法和系统动作"],
    ["Power BI", "持续更新的指标模型和分析展示", "应收、库存、月结Dashboard", "正式记账和审批"],
    ["Python", "大数据、复杂匹配、批量文件和规则扫描", "组合对账、TB异常、成本波动", "几十行数据且Excel已经稳定"],
    ["RPA", "无API但界面步骤固定", "定期下载、上传、回执归档", "验证码、高风险操作、界面频繁变化"],
    ["ERP原生功能", "系统已有标准模块和权限控制", "自动折旧、三单匹配、凭证接口", "主数据和配置不可靠"],
    ["API", "系统开放稳定接口并获正式授权", "ERP、银企、OA、发票平台连接", "厂商不支持或缺少权限"],
    ["AI", "文本理解、语义分类、初步解释", "合同摘要、邮件分类、异常说明初稿", "精确计算、自动审批和最终判断"],
    ["AI Agent", "成熟流程上的多步骤编排", "月结进度、Issue汇总、资料检查", "基础数据、规则、权限和日志尚未稳定"],
]

PANORAMA = [
    ["采购与应付", "三单匹配、重复付款、账户变更", "高", "高", "中", "高", "Power Query/Python", "付款审批与供应商真实性"],
    ["销售与应收", "回款认领、账龄、核销建议", "高", "中高", "中", "高", "Power Query/Python", "折扣、争议、坏账判断"],
    ["费用报销", "标准、日期、重复票、预算", "高", "中高", "中高", "中高", "规则/OCR/AI辅助", "招待、礼品和私人消费"],
    ["发票", "字段、验真、重复、税率", "高", "高", "中", "高", "数电票/规则/OCR", "业务真实性与抵扣资格"],
    ["银行与资金", "流水匹配、未达、余额调节", "高", "高", "中", "高", "Power Query/Python", "未知款和组合匹配"],
    ["工资薪酬", "花名册、工资、个税、实发核对", "高", "高", "中", "高", "Power Query", "个人信息权限和异常确认"],
    ["存货", "负库存、库龄、无移动、盘点差异", "高", "高", "中", "高", "Power Query/BI", "呆滞原因与减值"],
    ["成本", "BOM、耗用、单位成本、毛利波动", "中高", "中", "高", "中高", "Python/BI", "经营原因与口径判断"],
    ["固定资产", "折旧、卡片GL、闲置、转固提醒", "中", "高", "中", "中高", "ERP/Power Query", "转固时点和减值"],
    ["税务", "销进项整理、账票税核对、资料归档", "高", "中高", "高", "中高", "ERP/Power Query", "政策、申报复核与提交"],
    ["全科目核对", "方向、账龄、余额和模块差异", "高", "高", "中高", "高", "Excel/Python", "异常实质与调账判断"],
    ["月结", "任务、依赖、证据、延期与Issue", "高", "中高", "中", "高", "Excel/协同系统", "重大事项升级与关闭"],
    ["财务报表", "映射、公式、滚动与勾稽", "高", "高", "中", "高", "Excel/Python", "列报和重大调整"],
    ["财务分析", "刷新、同比、预算、PVM和Bridge", "中高", "中高", "高", "中高", "PQ/BI/Python", "经营解释与行动建议"],
    ["资料归档", "命名、分类、清单和PBC", "高", "高", "低", "高", "Python/RPA", "权限和资料完整性"],
    ["审计支持", "PBC状态、追问、未决事项", "高", "高", "中", "高", "Excel/协同系统", "缺失资料与审计判断"],
]

PROJECTS = [
    ("AUT-01", "银行流水自动对账", "银行与资金", "Power Query＋Excel，复杂组合匹配再用Python", 2, 28, 10, "一对一、跨日、手续费、利息、尾差和组合匹配"),
    ("AUT-02", "客户回款自动匹配", "销售与应收", "Power Query/Python", 3, 36, 12, "付款户名、金额、备注、客户主数据和应收明细评分"),
    ("AUT-03", "应付重复付款检查", "采购与应付", "Excel/Power Query", 2, 14, 4, "发票号、供应商、金额、账户及付款批次重复"),
    ("AUT-04", "发票重复识别", "发票", "Power Query/规则引擎", 2, 14, 5, "发票号码、数电票号码、金额、报销人和影像哈希"),
    ("AUT-05", "暂估自动匹配", "采购与应付", "Power Query/Python", 3, 34, 10, "PO、入库、供应商、物料、数量、金额和期间匹配"),
    ("AUT-06", "预提费用自动计算", "费用与月结", "Excel/Power Query", 1, 12, 4, "合同期间、履约进度、已入账、已开票和已付款滚动"),
    ("AUT-07", "应收账龄自动分析", "销售与应收", "Power Query/Power BI", 2, 16, 6, "到期日、逾期天数、信用期、争议和责任销售"),
    ("AUT-08", "长期挂账自动扫描", "全科目核对", "Power Query/Excel", 2, 18, 6, "30/90/180/365/730天、长期不动、方向和跨科目"),
    ("AUT-09", "TB异常余额检查", "全科目核对", "Excel/Python", 3, 28, 8, "科目方向、阈值、零余额、长期余额和规则版本"),
    ("AUT-10", "子模块与GL自动对账", "月结", "Power Query", 2, 22, 7, "AR、AP、存货、固定资产、薪酬、税务和银行"),
    ("AUT-11", "固定资产异常检查", "固定资产", "Excel/Power Query", 2, 16, 3, "卡片GL、停用、报废、闲置、转固和折旧异常"),
    ("AUT-12", "负库存扫描", "存货", "Power Query", 2, 10, 3, "仓库、物料、批次、日期和负数原因"),
    ("AUT-13", "存货库龄分析", "存货", "Power Query/Power BI", 3, 26, 7, "最后移动、最近领用、最近销售和责任部门"),
    ("AUT-14", "BOM异常分析", "成本", "Python/Power BI", 3, 44, 8, "标准BOM、实际领料、替代料、损耗和版本"),
    ("AUT-15", "单位成本波动分析", "成本", "Power Query/Python/BI", 3, 46, 10, "材料、人工、制造费用、产量、价格和BOM桥接"),
    ("AUT-16", "财务报表自动勾稽", "财务报表", "Excel/Python", 2, 18, 5, "资产负债平衡、利润、现金流、权益滚动和TB映射"),
    ("AUT-17", "月结Checklist追踪", "月结", "Excel/协同系统，后期Agent", 2, 20, 4, "任务、Owner、截止、依赖、证据、Issue和升级"),
    ("AUT-18", "财务资料自动归档", "资料归档", "Python/RPA", 3, 28, 6, "公司、期间、类型、版本、命名和目录规则"),
    ("AUT-19", "审计PBC跟踪", "审计支持", "Excel/协同系统", 2, 18, 12, "清单、责任人、资料路径、提交、追问和未决"),
    ("AUT-20", "财务分析自动刷新", "财务分析", "Power Query/Power BI", 3, 40, 12, "数据刷新、同比环比、预算、PVM、毛利和营运资金"),
]

SECURITY = [
    ["公开数据", "L1", "已公开年报、公开政策", "可以", "核对来源与日期", "一般业务权限", "核实来源"],
    ["内部一般", "L2", "脱敏科目表、流程模板", "经授权", "删除公司及个人标识", "按项目最小权限", "抽样复核"],
    ["敏感财务", "L3", "未公开报表、交易明细、合同", "原则上限制外部AI", "字段最小化、金额或名称替换", "本地处理、访问日志", "逐项复核关键数字"],
    ["个人信息", "L4", "工资、身份证、电话、银行卡", "严格限制", "掩码、匿名化、加密", "限定HR与薪酬人员", "完整性和权限复核"],
    ["高敏感凭证", "L5", "网银密码、U盾、税务凭证、管理员密码", "禁止", "不得交给AI或普通脚本", "仅受控系统和授权人员", "双人及职责分离"],
]

def write_csv(path: Path, headers, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)


def generate_samples():
    samples = {}
    # 1 bank reconciliation
    p = PACK / "01-银行自动对账"
    bank = [
        ["B001", "2026-08-31", 0, 113000, "华东经销商回款 SO1001", "华东商贸"],
        ["B002", "2026-08-31", 50000, 0, "付A供应商 PO2001", "A电子"],
        ["B003", "2026-08-31", 50, 0, "银行手续费", "开户银行"],
        ["B004", "2026-09-01", 0, 22600, "客户B回款", "客户B"],
        ["B005", "2026-08-31", 0, 120, "结息", "开户银行"],
        ["B006", "2026-08-31", 30000, 0, "供应商C第一笔", "供应商C"],
        ["B007", "2026-08-31", 20000, 0, "供应商C第二笔", "供应商C"],
    ]
    erp = [
        ["E001", "2026-08-31", 113000, "收客户华东SO1001", "华东商贸"],
        ["E002", "2026-08-31", -50000, "付PO2001", "A电子"],
        ["E003", "2026-08-31", 22600, "客户B回款", "客户B"],
        ["E004", "2026-08-31", -50000, "支付供应商C", "供应商C"],
    ]
    expected = [
        ["B001", "E001", "精确匹配", 1.00, "金额、日期、户名和摘要一致"],
        ["B002", "E002", "精确匹配", 1.00, "金额、日期和PO一致"],
        ["B003", "", "银行单边", 0.00, "建议补记手续费"],
        ["B004", "E003", "跨日候选", 0.85, "金额户名一致，日期相差1日"],
        ["B005", "", "银行单边", 0.00, "建议补记利息收入"],
        ["B006+B007", "E004", "多对一候选", 0.80, "两笔付款合计等于ERP金额，需人工确认"],
    ]
    write_csv(p/"输入数据"/"银行流水.csv", ["bank_id","date","debit","credit","summary","counterparty"], bank)
    write_csv(p/"输入数据"/"ERP银行明细.csv", ["erp_id","date","amount","summary","counterparty"], erp)
    write_csv(p/"预期结果"/"预期匹配结果.csv", ["bank_id","erp_id","match_type","confidence","reason"], expected)
    samples["bank"] = {"bank": bank, "erp": erp, "expected": expected}

    # 2 accrual matching
    p = PACK / "02-暂估自动匹配"
    accrual = [
        ["A001","2026-08","V001","PO1001","GR1001","M001",100,1000,100000,"未冲"],
        ["A002","2026-08","V002","PO1002","GR1002","M002",50,800,40000,"未冲"],
        ["A003","2026-05","V003","PO1003","GR1003","M003",10,5000,50000,"未冲"],
        ["A004","2026-08","V004","PO1004","GR1004","M004",20,2000,40000,"未冲"],
    ]
    inv = [
        ["I001","2026-09-05","V001","PO1001","M001",100,1000,100000,13000],
        ["I002","2026-09-06","V002","PO1002","M002",50,820,41000,5330],
        ["I003","2026-09-07","V999","PO1004","M004",20,2000,40000,5200],
    ]
    exp = [
        ["A001","I001","可冲销",0,"PO/供应商/物料/数量/金额一致"],
        ["A002","I002","金额差异",1000,"单价由800变为820，需采购确认"],
        ["A003","","长期未到票",50000,"超过90天，向采购和供应商追踪"],
        ["A004","I003","供应商不一致",0,"不得自动冲销，核实代开或主数据错误"],
    ]
    write_csv(p/"输入数据"/"暂估明细.csv", ["accrual_id","period","vendor","po","gr","material","qty","unit_price","amount","status"], accrual)
    write_csv(p/"输入数据"/"发票明细.csv", ["invoice_id","invoice_date","vendor","po","material","qty","unit_price","amount","tax"], inv)
    write_csv(p/"预期结果"/"预期匹配结果.csv", ["accrual_id","invoice_id","result","amount_diff","action"], exp)
    samples["accrual"]={"accrual":accrual,"invoice":inv,"expected":exp}

    # 3 long balance scan
    p = PACK / "03-长期挂账扫描"
    balances = [
        ["AR001","应收账款","C001","2026-08-15",113000,0,"销售部"],
        ["AR002","应收账款","C002","2025-12-01",226000,0,"销售部"],
        ["AP001","应付账款","V001","2026-07-20",0,50000,"采购部"],
        ["PP001","预付账款","V009","2024-06-01",180000,0,"采购部"],
        ["OR001","其他应收款","EMP008","2025-01-01",30000,0,"行政部"],
        ["AR003","应收账款","C003","2026-06-30",0,10000,"销售部"],
    ]
    exp = [
        ["AR001","30天以内","低","正常跟踪"],
        ["AR002","180-365天","高","确认回款计划、争议和减值迹象"],
        ["AP001","30-90天","低","核对付款条件"],
        ["PP001","2年以上","高","核实交付、退款和资金占用"],
        ["OR001","1-2年","高","员工借款催收并检查个税或损失风险"],
        ["AR003","90-180天","高","应收贷方异常，先查业务实质再决定重分类"],
    ]
    write_csv(p/"输入数据"/"往来明细.csv", ["item_id","account","party","business_date","debit_balance","credit_balance","owner_dept"], balances)
    write_csv(p/"预期结果"/"预期异常结果.csv", ["item_id","aging_bucket","priority","action"], exp)
    samples["aging"]={"balances":balances,"expected":exp}

    # 4 TB anomaly
    p = PACK / "04-全科目异常检查"
    tb = [
        ["1122.01","应收账款-客户C001","借",113000,"C001"],
        ["1122.02","应收账款-客户C003","贷",10000,"C003"],
        ["2202.01","应付账款-供应商V002","借",25000,"V002"],
        ["1405.01","库存商品-A","贷",8000,"WH01"],
        ["5101.01","制造费用-折旧","借",12000,"P01"],
        ["1606.01","固定资产清理-设备X","借",50000,"FA009"],
        ["2221.01.01","应交增值税-进项税额","借",130000,""],
    ]
    rules = [
        ["R001","1122","贷",1,"应收贷方","查预收、退款、错记；不得机械认定合同负债"],
        ["R002","2202","借",1,"应付借方","查预付、退货、重复付款或错记"],
        ["R003","1405","贷",1,"负库存","查截止、接口、单位换算和领料"],
        ["R004","5101","非零",1,"制造费用期末余额","检查成本结转是否完成"],
        ["R005","1606","非零",30000,"固定资产清理长期或大额","检查处置流程和结转"],
    ]
    exp = [["1122.02","R001","高","先查款项实质"],["2202.01","R002","中","核对供应商明细"],["1405.01","R003","高","冻结相关出库并查数据"],["5101.01","R004","高","成本模块未关闭"],["1606.01","R005","中","检查处置结转"]]
    write_csv(p/"输入数据"/"TB.csv", ["account_code","account_name","direction","balance","dimension"], tb)
    write_csv(p/"输入数据"/"异常规则.csv", ["rule_id","account_prefix","condition","threshold","rule_name","action"], rules)
    write_csv(p/"预期结果"/"预期异常结果.csv", ["account_code","rule_id","priority","action"], exp)
    samples["tb"]={"tb":tb,"rules":rules,"expected":exp}

    # 5 financial statement checks
    p = PACK / "05-财务报表自动勾稽"
    fs = [
        ["资产负债表","资产总计",5000000], ["资产负债表","负债总计",3000000], ["资产负债表","所有者权益总计",1990000],
        ["利润表","净利润",250000], ["现金流量表","现金净增加额",100000], ["资产负债表","期初未分配利润",600000],
        ["利润分配","本期分红",50000], ["资产负债表","期末未分配利润",800000], ["资产负债表","期初货币资金",900000],
        ["资产负债表","期末货币资金",990000], ["TB","未映射科目金额",12000],
    ]
    exp = [
        ["CHK01","资产=负债+权益",10000,"未通过","5000000-3000000-1990000"],
        ["CHK02","未分配利润滚动",0,"通过","600000+250000-50000-800000"],
        ["CHK03","现金流与货币资金变动",10000,"未通过","990000-900000-100000"],
        ["CHK04","TB未映射科目",12000,"未通过","应为0"],
    ]
    write_csv(p/"输入数据"/"报表数据.csv", ["statement","item","amount"], fs)
    write_csv(p/"预期结果"/"预期校验结果.csv", ["check_id","check_name","difference","status","formula_logic"], exp)
    samples["fs"]={"data":fs,"expected":exp}
    return samples


def write_scripts_and_readmes():
    common = """# 使用说明\n\n本项目用于华辰智能设备有限公司模拟数据。先阅读输入字段和规则，再运行工具。自动化输出是匹配建议或异常清单，不能直接替代记账、核销、付款、申报或管理层批准。\n\n## 标准步骤\n1. 备份原始数据，保持原始文件只读。\n2. 将当期文件放入“输入数据”。\n3. 检查字段名、数据类型、日期和金额方向。\n4. 运行Excel/Power Query或Python方案。\n5. 将结果与“预期结果”比较并完成UAT。\n6. 对未匹配和高风险项目进行人工判断。\n7. 保存输入版本、规则版本、输出、复核人和运行日志。\n\n## 运行Python\n在项目目录执行：`python 脚本名称.py --input ... --output ...`。建议使用公司批准的Python环境。\n\n## 失败回退\n脚本、查询或接口失败时，使用输入数据和规则表按原人工流程完成关键工作，并记录失败原因。不要因为工具失败而延误月结。\n"""
    for i, name in enumerate(["银行自动对账","暂估自动匹配","长期挂账扫描","全科目异常检查","财务报表自动勾稽"],1):
        p=PACK/f"{i:02d}-{name}"
        p.mkdir(parents=True, exist_ok=True)
        (p/"README.md").write_text(common+f"\n## 本项目\n{name}。先运行确定性规则，高风险和无法解释的结果必须转人工。\n",encoding="utf-8")

    (PACK/"01-银行自动对账"/"PowerQuery方案.pq").write_text('''let\n  Bank = Csv.Document(File.Contents("银行流水.csv"),[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv]),\n  BankHeaders = Table.PromoteHeaders(Bank,[PromoteAllScalars=true]),\n  BankTypes = Table.TransformColumnTypes(BankHeaders,{{"date", type date},{"debit", type number},{"credit", type number}}),\n  BankAmount = Table.AddColumn(BankTypes,"amount", each [credit]-[debit], type number),\n  ERP = Csv.Document(File.Contents("ERP银行明细.csv"),[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv]),\n  ERPHeaders = Table.PromoteHeaders(ERP,[PromoteAllScalars=true]),\n  ERPTypes = Table.TransformColumnTypes(ERPHeaders,{{"date", type date},{"amount", type number}}),\n  Join = Table.NestedJoin(BankAmount,{"date","amount"},ERPTypes,{"date","amount"},"ERP",JoinKind.LeftOuter),\n  Expand = Table.ExpandTableColumn(Join,"ERP",{"erp_id","counterparty"},{"erp_id","erp_counterparty"}),\n  Status = Table.AddColumn(Expand,"status", each if [erp_id]=null then "未匹配" else "精确候选")\nin Status''',encoding="utf-8")
    (PACK/"02-暂估自动匹配"/"PowerQuery方案.pq").write_text('''let\n  A = Excel.CurrentWorkbook(){[Name="AccrualTable"]}[Content],\n  I = Excel.CurrentWorkbook(){[Name="InvoiceTable"]}[Content],\n  J = Table.NestedJoin(A,{"vendor","po","material"},I,{"vendor","po","material"},"Invoice",JoinKind.LeftOuter),\n  E = Table.ExpandTableColumn(J,"Invoice",{"invoice_id","qty","amount"},{"invoice_id","invoice_qty","invoice_amount"}),\n  D = Table.AddColumn(E,"amount_diff",each [invoice_amount]-[amount],type number),\n  S = Table.AddColumn(D,"result",each if [invoice_id]=null then "未到票" else if [amount_diff]=0 and [invoice_qty]=[qty] then "可冲销" else "差异待查")\nin S''',encoding="utf-8")
    (PACK/"03-长期挂账扫描"/"PowerQuery方案.pq").write_text('''let\n  S = Excel.CurrentWorkbook(){[Name="BalanceTable"]}[Content],\n  T = Table.TransformColumnTypes(S,{{"business_date",type date},{"debit_balance",type number},{"credit_balance",type number}}),\n  Age = Table.AddColumn(T,"age_days",each Duration.Days(#date(2026,9,10)-[business_date]),Int64.Type),\n  Bucket = Table.AddColumn(Age,"aging_bucket",each if [age_days]<=30 then "30天以内" else if [age_days]<=90 then "30-90天" else if [age_days]<=180 then "90-180天" else if [age_days]<=365 then "180-365天" else if [age_days]<=730 then "1-2年" else "2年以上"),\n  Flag = Table.AddColumn(Bucket,"direction_flag",each if Text.StartsWith([account],"应收") and [credit_balance]>0 then "方向异常" else "")\nin Flag''',encoding="utf-8")
    (PACK/"04-全科目异常检查"/"PowerQuery方案.pq").write_text('''let\n  TB = Excel.CurrentWorkbook(){[Name="TBTable"]}[Content],\n  Rules = Excel.CurrentWorkbook(){[Name="RuleTable"]}[Content],\n  Prefix = Table.AddColumn(TB,"prefix4",each Text.Start([account_code],4)),\n  Join = Table.NestedJoin(Prefix,{"prefix4"},Rules,{"account_prefix"},"Rule",JoinKind.LeftOuter),\n  Expand = Table.ExpandTableColumn(Join,"Rule",{"rule_id","condition","threshold","rule_name","action"},{"rule_id","condition","threshold","rule_name","action"}),\n  Flag = Table.AddColumn(Expand,"is_exception",each if [condition]="贷" and [direction]="贷" then true else if [condition]="借" and [direction]="借" then true else if [condition]="非零" and Number.Abs([balance])>[threshold] then true else false)\nin Flag''',encoding="utf-8")

    scripts = {
      "01-银行自动对账/银行对账.py": r'''from pathlib import Path
import argparse, pandas as pd

def main():
    ap=argparse.ArgumentParser(description="银行与ERP一对一及跨日候选匹配")
    ap.add_argument("--bank",default="输入数据/银行流水.csv")
    ap.add_argument("--erp",default="输入数据/ERP银行明细.csv")
    ap.add_argument("--output",default="运行结果.xlsx")
    a=ap.parse_args(); b=pd.read_csv(a.bank); e=pd.read_csv(a.erp)
    b["date"]=pd.to_datetime(b["date"],errors="coerce"); e["date"]=pd.to_datetime(e["date"],errors="coerce")
    b["amount"]=b["credit"].fillna(0)-b["debit"].fillna(0); used=set(); out=[]
    for _,r in b.iterrows():
        c=e[(e.amount.round(2)==round(r.amount,2)) & (~e.erp_id.isin(used))].copy()
        if len(c):
            c["days"]=(c.date-r.date).abs().dt.days; c=c[c.days<=1].sort_values(["days","erp_id"])
        if len(c):
            x=c.iloc[0]; used.add(x.erp_id); typ="精确匹配" if x.days==0 else "跨日候选"
            out.append([r.bank_id,x.erp_id,typ,1.0 if x.days==0 else .85,"需复核户名和摘要"])
        else: out.append([r.bank_id,"","未匹配",0,"检查手续费、利息、组合或漏记"])
    result=pd.DataFrame(out,columns=["bank_id","erp_id","match_type","confidence","reason"])
    with pd.ExcelWriter(a.output) as w:
        result.to_excel(w,sheet_name="匹配结果",index=False); e[~e.erp_id.isin(used)].to_excel(w,sheet_name="ERP未匹配",index=False)
if __name__=="__main__": main()
''',
      "02-暂估自动匹配/暂估匹配.py": r'''import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--accrual",default="输入数据/暂估明细.csv"); ap.add_argument("--invoice",default="输入数据/发票明细.csv"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    x=pd.read_csv(a.accrual); y=pd.read_csv(a.invoice)
    z=x.merge(y,on=["vendor","po","material"],how="left",suffixes=("_accrual","_invoice"))
    z["amount_diff"]=z["amount_invoice"]-z["amount_accrual"]
    z["qty_diff"]=z["qty_invoice"]-z["qty_accrual"]
    z["result"]=z.apply(lambda r:"未到票" if pd.isna(r.invoice_id) else ("可冲销" if abs(r.amount_diff)<.01 and abs(r.qty_diff)<.001 else "差异待查"),axis=1)
    z.to_excel(a.output,index=False)
if __name__=="__main__": main()
''',
      "03-长期挂账扫描/长期挂账扫描.py": r'''import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input",default="输入数据/往来明细.csv"); ap.add_argument("--as-of",default="2026-09-10"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    x=pd.read_csv(a.input); x["business_date"]=pd.to_datetime(x.business_date,errors="coerce"); x["age_days"]=(pd.Timestamp(a.as_of)-x.business_date).dt.days
    x["aging_bucket"]=pd.cut(x.age_days,[-1,30,90,180,365,730,10**9],labels=["30天以内","30-90天","90-180天","180-365天","1-2年","2年以上"])
    x["direction_flag"]=((x.account.str.startswith("应收"))&(x.credit_balance>0))|((x.account.str.startswith("应付"))&(x.debit_balance>0))
    x["priority"]=x.apply(lambda r:"高" if r.age_days>180 or r.direction_flag else ("中" if r.age_days>90 else "低"),axis=1)
    x.sort_values(["priority","age_days"],ascending=[True,False]).to_excel(a.output,index=False)
if __name__=="__main__": main()
''',
      "04-全科目异常检查/TB异常检查.py": r'''import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--tb",default="输入数据/TB.csv"); ap.add_argument("--rules",default="输入数据/异常规则.csv"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    tb=pd.read_csv(a.tb,dtype={"account_code":str}); rules=pd.read_csv(a.rules,dtype={"account_prefix":str}); out=[]
    for _,r in tb.iterrows():
        for _,q in rules.iterrows():
            if not r.account_code.startswith(q.account_prefix): continue
            hit=(q.condition==r.direction) or (q.condition=="非零" and abs(r.balance)>q.threshold)
            if hit: out.append([r.account_code,r.account_name,r.balance,q.rule_id,q.rule_name,q.action])
    pd.DataFrame(out,columns=["account_code","account_name","balance","rule_id","rule_name","action"]).to_excel(a.output,index=False)
if __name__=="__main__": main()
''',
      "05-财务报表自动勾稽/报表勾稽.py": r'''import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input",default="输入数据/报表数据.csv"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    x=pd.read_csv(a.input); v=dict(zip(x.item,x.amount))
    checks=[("资产=负债+权益",v["资产总计"]-v["负债总计"]-v["所有者权益总计"]),("未分配利润滚动",v["期初未分配利润"]+v["净利润"]-v["本期分红"]-v["期末未分配利润"]),("现金流与货币资金变动",v["期末货币资金"]-v["期初货币资金"]-v["现金净增加额"]),("TB未映射科目",v["未映射科目金额"])]
    out=pd.DataFrame(checks,columns=["check_name","difference"]); out["status"]=out.difference.abs().le(.01).map({True:"通过",False:"未通过"}); out.to_excel(a.output,index=False)
if __name__=="__main__": main()
'''
    }
    for rel, text in scripts.items():
        path=PACK/rel; path.write_text(text,encoding="utf-8")

    (PACK/"00-项目总说明").mkdir(parents=True,exist_ok=True)
    (PACK/"00-项目总说明"/"README.md").write_text("""# 财务自动化实战项目包\n\n第一批项目集中在银行对账、暂估匹配、长期挂账、TB异常和报表勾稽。它们与总账月结紧密相关，规则相对清晰，适合先用Excel和Power Query建立稳定流程，再按数据量增加Python。\n\n所有输出均为建议或异常清单。正式入账、核销、付款、申报和重大判断必须由授权人员复核批准。\n""",encoding="utf-8")
    (PACK/"00-项目总说明"/"手工Fallback流程.md").write_text("""# 手工Fallback流程\n\n1. 确认自动化失败范围和最后成功时间。\n2. 冻结当次输出，避免继续使用不完整结果。\n3. 使用已归档的原始数据和上期模板恢复人工处理。\n4. 优先完成月结关键路径、资金和申报事项。\n5. 记录错误、影响、临时措施和责任人。\n6. 修复后重新运行，并与人工结果进行并行核对。\n7. 由复核人确认后恢复正式使用。\n""",encoding="utf-8")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    samples=generate_samples(); write_scripts_and_readmes()
    data={"company":COMPANY,"tools":TOOLS,"panorama":PANORAMA,"projects":[{"id":x[0],"name":x[1],"process":x[2],"tool":x[3],"level":x[4],"hours":x[5],"saved":x[6],"rules":x[7]} for x in PROJECTS],"security":SECURITY,"samples":samples}
    (BASE/"content.json").write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps({"projects":len(PROJECTS),"sample_projects":len(samples)},ensure_ascii=False))

if __name__=="__main__": main()
