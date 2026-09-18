from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
import json, zipfile, hashlib, base64

OUT=Path(r'C:\Users\HZJ\Desktop\Git\work\fpna-output')
DOCX=OUT/'企业财务分析与经营分析实战手册.docx'
XMIND=OUT/'企业财务与经营分析全景.xmind'
NAVY='17365D'; BLUE='D9EAF7'; PALE='F5F8FB'; GRAY='D9D9D9'; RED='C00000'

def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=tcPr.find(qn('w:shd'))
    if shd is None: shd=OxmlElement('w:shd'); tcPr.append(shd)
    shd.set(qn('w:fill'),fill)
def margins(cell, top=100, start=120, bottom=100, end=120):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr(); tcMar=tcPr.first_child_found_in('w:tcMar')
    if tcMar is None: tcMar=OxmlElement('w:tcMar'); tcPr.append(tcMar)
    for tag,val in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        el=tcMar.find(qn('w:'+tag))
        if el is None: el=OxmlElement('w:'+tag); tcMar.append(el)
        el.set(qn('w:w'),str(val)); el.set(qn('w:type'),'dxa')
def repeat_header(row):
    trPr=row._tr.get_or_add_trPr(); el=OxmlElement('w:tblHeader'); el.set(qn('w:val'),'true'); trPr.append(el)
def set_cell_text(cell,text,bold=False,color='000000',size=8.5,align=None):
    cell.text=''; p=cell.paragraphs[0]
    if align is not None: p.alignment=align
    r=p.add_run(str(text)); r.bold=bold; r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(size); r.font.color.rgb=RGBColor.from_string(color)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; margins(cell)
def add_table(doc, headers, rows, widths=None):
    t=doc.add_table(rows=1, cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False
    for i,h in enumerate(headers): set_cell_text(t.rows[0].cells[i],h,True,'FFFFFF',8.5,WD_ALIGN_PARAGRAPH.CENTER); shade(t.rows[0].cells[i],NAVY)
    repeat_header(t.rows[0])
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        for i,v in enumerate(row): set_cell_text(cells[i],v,False,'000000',8.3,WD_ALIGN_PARAGRAPH.CENTER if len(str(v))<16 else WD_ALIGN_PARAGRAPH.LEFT); shade(cells[i], 'FFFFFF' if ri%2 else 'EEF5FA')
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Cm(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)
    return t
def add_p(doc,text,bold_lead=None):
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(5); p.paragraph_format.line_spacing=1.25
    if bold_lead and text.startswith(bold_lead):
        r=p.add_run(bold_lead); r.bold=True; r.add_text(text[len(bold_lead):])
    else:p.add_run(text)
    return p
def bullets(doc,items):
    for x in items:
        p=doc.add_paragraph(style='List Bullet'); p.paragraph_format.space_after=Pt(2); p.add_run(x)
def heading(doc,text,level=1):
    p=doc.add_heading(text,level=level); p.paragraph_format.keep_with_next=True; return p
def page_break(doc): doc.add_page_break()

doc=Document(); sec=doc.sections[0];sec.top_margin=Cm(2.2);sec.bottom_margin=Cm(2);sec.left_margin=Cm(2.2);sec.right_margin=Cm(2.0)
styles=doc.styles
styles['Normal'].font.name='Microsoft YaHei';styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');styles['Normal'].font.size=Pt(10);styles['Normal'].font.color.rgb=RGBColor(0,0,0)
for n,size in [('Title',26),('Heading 1',18),('Heading 2',14),('Heading 3',11.5)]:
    st=styles[n];st.font.name='Microsoft YaHei';st._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');st.font.size=Pt(size);st.font.color.rgb=RGBColor(0,0,0);st.font.bold=True
styles['Heading 1'].paragraph_format.space_before=Pt(16);styles['Heading 1'].paragraph_format.space_after=Pt(8)
styles['Heading 2'].paragraph_format.space_before=Pt(10);styles['Heading 2'].paragraph_format.space_after=Pt(5)
styles['Heading 3'].paragraph_format.space_before=Pt(7);styles['Heading 3'].paragraph_format.space_after=Pt(3)

# Cover
p=doc.add_paragraph();p.paragraph_format.space_before=Pt(110);p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('企业财务分析与经营分析实战手册');r.bold=True;r.font.name='Microsoft YaHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');r.font.size=Pt(28);r.font.color.rgb=RGBColor(0,0,0)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('从财务报表结果到业务驱动和管理行动');r.font.size=Pt(15);r.font.name='Microsoft YaHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑')
doc.add_paragraph();p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run('案例公司  深圳拓维三维科技有限公司  模拟').italic=True
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run('版本日期  2026年9月10日')
page_break(doc)

heading(doc,'使用说明',1)
add_p(doc,'本手册面向具备审计和报表基础、准备进入企业总账、财务主管、FP&A或财务BP岗位的学习者。重点不是重复计算财务比率，而是把财务结果拆到产品、客户、渠道、订单、产量、价格、材料、良率、产能和回款。')
add_p(doc,'任何分析结论都应经过三次校验：数字能否勾稽，业务信息能否证明原因，建议能否转化为责任人、期限和目标值。数据不足时，应明确列出待验证事项，不把推测写成事实。')
add_table(doc,['标识','含义','使用要求'],[['事实数据','来自财务系统或经营系统的原始结果','注明期间、单位和口径'],['计算结果','由公式得到的差异、比率或Bridge','必须设置勾稽'],['业务信息','业务部门确认的事件','记录来源和确认人'],['分析判断','基于事实形成的专业判断','说明证据和限制'],['待验证推测','可能原因但证据不足','列明需要的数据和责任部门'],['管理建议','可以执行的改进措施','设置责任人、期限和目标值']],[3,5,9])
heading(doc,'目录',1)
p=doc.add_paragraph();run=p.add_run();fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'TOC \\o "1-3" \\h \\z \\u');run._r.addnext(fld)
add_p(doc,'在Word或WPS中打开后，如目录未显示页码，请右键目录并选择更新域。')
page_break(doc)

chapters=[
('第1章 财务分析在企业里的真实作用','财务分析的价值在于改变未来经营，而不是复述已经发生的数字。财务需要把利润表、资产负债表和现金流量表连成同一条经营因果链。',[
('四类工作的边界',['核算回答业务如何进入账簿；报表分析回答结果发生了什么；经营分析回答业务为何形成这些结果；财务BP推动业务采取行动。','总账会计应保证基础数据完整并识别异常。财务主管要把异常分解到责任中心，并组织业务验证。']),
('管理层真正关心的问题',['目标是否完成，未完成发生在哪里；利润是否能转化为现金；增长是否可持续；哪些风险会影响未来；公司现在可以采取什么行动。']),
('分析交付标准',['结论必须对应数字；原因必须对应业务证据；建议必须对应责任人；下月必须复盘上月行动。'])]),
('第2章 企业深度财务与经营分析全景','完整分析遵循结果识别、差异定位、业务验证、原因量化和行动跟踪。跳过任何一步都可能产生错误结论。',[
('九步分析法',['确定结果和比较基准。','量化变化金额和比例。','按产品、客户、渠道、地区、部门定位。','拆解数量、价格、结构、效率和成本。','向业务系统取得验证数据。','量化每项原因。','判断一次性、短期或结构性。','评估利润、现金和未来影响。','形成行动计划并持续跟踪。']),
('证据链',['财务总账说明最终结果；ERP业务模块说明交易构成；CRM说明客户和订单；MES说明产量、良率和工时；WMS说明库存；银行和AR说明回款。'])]),
('第3章 事实 判断 推测和建议','财务分析最常见的质量问题，是把时间上的同时发生写成因果关系。',[
('规范表达',['事实：综合毛利率同比下降3.5个百分点。','拆解：材料价格、产品结构和良率分别影响1.0、1.2和0.4个百分点。','业务信息：采购确认核心控制板价格上涨；生产确认树脂线良率下降。','判断：材料因素可能持续，良率因素具有改善空间。','待验证：渠道促销是否真正形成终端销量。','行动：销售取得Sell-out，质量部门完成缺陷Pareto。']),
('证据不足时',['不要写“市场竞争导致收入下降”。应写“收入下降集中在A产品经销渠道；需要销售提供订单流失原因、竞品价格和终端库存验证”。'])]),
('第4章 财务分析常用比较方法','比较方法回答的问题不同。同比用于识别季节性后的变化，环比用于识别近期拐点，预算差异用于检验经营承诺。',[
('比较矩阵',['同比：适合季节性业务，但注意基数。','环比：适合观察最近变化，但易受季节性干扰。','预算与实际：适合责任评价，但必须先评估预算是否合理。','Forecast与实际：适合检验预测能力和最新信息。','滚动12个月：降低单月波动。','同业比较：用于提出问题，不能直接下结论。']),
('组合解释',['收入同比增长20%但环比下降15%，可能同时说明公司仍高于去年低基数，但近期订单、交付或季节性正在转弱。必须进一步比较历史季节曲线和订单数据。','预算完成105%不一定经营良好。如果预算过低、收入来自低毛利产品、应收和库存同步增加，完成率会掩盖质量问题。'])]),
('第5章 差异分析方法','深度差异分析要求所有影响因素能够勾稽到总差异，并为无法解释部分设置残差。',[
('方法选择',['金额Bridge用于解释利润额变化；毛利率Bridge用于解释盈利结构；PVM用于解释收入；材料价格和用量差异用于解释生产成本；人数和人均成本用于解释人员费用。']),
('残差',['残差不是可以随意使用的“其他”。残差非零通常意味着分类不互斥、基期口径变化、新产品未处理或公式顺序不一致。'])]),
('第6章 收入分析','收入变化应至少拆到数量、价格、产品结构、客户、渠道、地区、汇率和一次性事项。',[
('PVM公式',['收入等于总销量×产品结构×产品价格。销量影响使用总销量变化和基期结构；结构影响使用本期总销量和结构变化；价格影响使用本期销量和价格变化。三项必须与实际收入变化一致。']),
('连续案例',['E1桌面机销量增长但价格下降；P1专业机销量增长；耗材随装机量扩大。模型计算基期收入10,802.29万元、本期12,766.48万元，变化1,964.19万元，PVM残差为0。']),
('新客户和新产品',['新增客户不是天然独立因素。如果新增客户销量已经包含在销量影响中，再单列会重复。只有先把客户分组设计成互斥维度，才能建立客户Bridge。']),
('取数与沟通',['销售运营提供订单、出货、价格、折扣、客户和渠道；财务将不含税收入与总账勾稽；物流或交付团队验证截止。'])]),
('第7章 毛利与毛利率分析','综合毛利率是各产品毛利率按收入结构加权的结果。毛利率下降不能只归因于成本增长。',[
('原因树',['收入端：价格、产品、客户、渠道、地区、返利、汇率。','材料端：采购价格、供应商、BOM、替代料、单耗、报废。','生产端：良率、人工效率、利用率、固定费用吸收、外协、新品爬坡。']),
('毛利率Bridge',['深圳拓维的38.0%下降至34.5%，分解为价格-0.8个百分点、结构-1.2、材料-1.0、良率-0.4、产能利用率-0.6、新品和其他+0.5，合计-3.5个百分点。正式报告必须检查期初加各因素等于期末。']),
('产能影响',['产能利用率下降不会改变固定制造费用总额，却会提高单位合格产品吸收的固定成本。应结合标准工时、有效产能、停机和实际产量分析。']),
('良率影响',['良率下降会增加合格品单位材料投入、报废、返工人工和设备工时。财务应取得MES投产量、合格量、报废原因和标准单耗。'])]),
('第8章 期间费用深度分析','费用分析需要同时评价金额、费用率、固定变动属性、责任部门和投入产出。',[
('销售费用',['销售费用增长30%、收入增长50%时，费用率可能改善，但仍需判断新增费用是否对应可持续收入、健康毛利和现金回收。']),
('人员费用',['拆解为平均人数×平均薪酬×在岗月份，再分析职级、奖金、社保、股份支付和招聘时间。']),
('研发费用',['按项目、阶段、人员和外包拆解，同时跟踪技术里程碑。投入增长而里程碑持续延期，说明费用增加尚未形成相应经营产出。']),
('财务费用',['拆解借款余额、利率、票据贴现、汇兑和资金占用。营运资金恶化往往通过借款和利息再次影响利润。'])]),
('第9章 从营业利润到净利润','营业利润之外的项目需要区分经营性、一次性和会计估计因素。',[
('减值',['信用减值应与客户信用、逾期和期后回款联系；存货跌价应与售价、库龄、订单和预计处置费用联系。']),
('政府补助和投资收益',['分析时说明是否重复发生、是否依赖特定项目、是否影响现金，以及剔除后主营业务盈利水平。']),
('有效税率',['从法定税率出发，解释优惠税率、研发加计扣除、永久性差异、以前年度调整和递延所得税。'])]),
('第10章 资产负债表深度分析','资产负债表反映企业过去经营决策积累形成的资源和义务。余额增加本身不是好坏结论。',[
('应收',['拆销售增长、客户结构、账期、逾期、争议和收入确认。应收增长快于收入时，优先分析DSO和大客户。']),
('存货',['先拆原材料、在产品、产成品、耗材和发出商品，再结合订单、备货、S&OP、工单和库龄判断。']),
('固定资产和在建工程',['分析资本开支是否形成产能、实际利用率、投产节奏和投资回报。扩产早于需求会同时增加折旧和现金压力。']),
('负债',['应付增加可能来自采购增长、账期改善或延迟付款。必须区分经营改善和供应商关系恶化。'])]),
('第11章 营运资金与现金转换周期','CCC等于DSO加DIO减DPO，反映从付出采购现金到收回销售现金的时间。',[
('三项联动',['DSO上升可能来自账期延长或逾期；DIO上升可能来自备货、销量下降或生产停滞；DPO下降可能来自供应商收紧信用。']),
('改善原则',['销售负责信用条款和回款；供应链负责预测和库存；采购负责供应商账期；财务统一口径、量化现金影响并推动行动。']),
('避免错误',['不能单纯延迟供应商付款来改善现金，这可能损害交付、折扣和信用。应评估整体经济影响。'])]),
('第12章 现金流与自由现金流','净利润不等于现金流，因为收入确认、成本结转和现金收付的时间不同。',[
('净利润到现金',['深圳拓维净利润1,500万元，加回折旧1,300万元；应收增加占用3,700万元，存货增加占用4,600万元，应付增加释放1,300万元，其他经营项目净增加4,700万元，经营现金流为500万元。']),
('自由现金流',['经营现金流减维持和扩张性资本开支形成自由现金流。分析时应区分维持现有经营和扩张产能的资本开支。']),
('核查路径',['从间接法Bridge定位到具体资产负债科目，再下钻客户、SKU、供应商和项目。'])]),
('第13章 制造企业经营指标体系','经营指标是因，财务指标往往是果。指标定义必须由业务和财务共同确认。',[
('销售与订单',['订单额、订单转收入率、Backlog、销量、ASP、终端销量、退货率。']),
('生产',['产量、理论产能、有效产能、标准工时、实际工时、良率、报废、返工、停机。']),
('供应链',['采购价格、交付周期、供应商集中度、安全库存、预测准确率、库龄。']),
('售后',['活跃设备、维修率、质保成本、耗材复购和用户留存。'])]),
('第14章 业务指标到财务结果的映射','经营指标只有与财务结果建立稳定映射，才能成为领先指标。',[
('映射示例',['良率下降→单位材料上升→单位成本上升→毛利率下降。','客户账期延长→应收增加→DSO上升→现金流下降→利息增加。','产能利用率下降→单位固定制造费用上升→存货和销售成本上升。','退货率上升→收入冲减、退款负债和质保成本增加。']),
('责任',['业务部门拥有经营指标，财务拥有口径、勾稽、影响量化和跨部门闭环。'])]),
('第15章 硬件加耗材商业模式','硬件加耗材模式需要同时分析单次设备毛利和客户生命周期价值。',[
('核心指标',['装机量、活跃设备、激活率、耗材月均用量、复购率、流失率、兼容耗材渗透率、设备和耗材毛利。']),
('低硬件毛利逻辑',['只有新增设备能够形成高概率、高毛利、持续的耗材复购，降低硬件毛利才可能合理。分析应比较硬件补贴与未来贡献毛利的现值。']),
('风险',['渠道压货不等于有效装机；设备销量增长但激活率和耗材复购下降时，长期利润逻辑可能失效。'])]),
('第16章 不同行业的核心分析指标','同一指标不能机械用于所有商业模式。',[
('行业差异',['SaaS关注ARR、续费率、净收入留存和获客回收期；零售关注同店增长、坪效、库存周转和损耗；工程施工关注订单、履约进度、合同资产和现金回款；半导体制造关注利用率、良率和折旧。']),
('金融行业',['银行、保险和证券不建议使用普通毛利率比较。银行关注净息差和成本收入比，保险关注综合成本率和新业务价值，证券关注收入结构、杠杆和ROE。'])]),
('第17章 行业毛利率与可比公司','行业基准用于提出问题，不直接证明公司经营好坏。',[
('数据库口径',['官方工业行业毛利率由营业收入和营业成本计算；上市公司样本应按同一期间和同一合并口径提取。']),
('Peer筛选',['控制产业链位置、产品结构、客户、渠道、地区、规模、自制外包和会计政策。']),
('统计方法',['同时展示样本数、低位、中位数和高位。样本不足时不输出区间。异常值保留但单独标记。'])]),
('第18章 财务异常问题诊断树','诊断树用于指导取数，不用于提前认定原因。',[
('收入下降',['先拆销量、价格和结构，再看产品、客户、渠道、地区；随后核查订单、交付、退货和竞争信息。']),
('毛利下降',['先区分价格与成本，再拆材料、人工、制造费用、良率、产能和新品。']),
('现金流下降',['从利润到现金Bridge定位应收、库存、预付、应付和税费，再下钻业务对象。'])]),
('第19章 月度财务与经营分析报告','月度报告应让管理层迅速知道发生了什么、为什么、未来影响及下一步行动。',[
('报告结构',['核心结论、收入、毛利、费用、营运资金、现金流、资产负债重大变化、经营KPI、风险、未来影响、行动计划和上月复盘。']),
('写作要求',['每段采用事实、拆解、业务原因、判断、风险和行动顺序。图表标题直接说明问题。']),
('优秀示例',['销售费用增加600万元但费用率下降。新增费用主要用于渠道扩张；渠道收入增长50%，但DSO和退货率上升。建议把渠道奖励从发货额调整为终端销量与回款。'])]),
('第20章 分析图表选择','图表服务于管理问题，而不是装饰。',[
('选择规则',['趋势用折线；预算实际用柱状；结构用堆积柱；因素贡献用瀑布；产品增长和毛利用散点；账龄库龄用分层柱。']),
('图表控制',['一张图回答一个问题；类别过多时只展示重要项目；坐标、期间、单位和口径必须明确。'])]),
('第21章 财务分析常见错误','浅层分析通常来自口径不清、没有下钻或没有验证。',[
('高频错误',['只同比不拆原因；只看金额不看结构；只看利润不看现金；把相关性当因果；忽略基数、季节、汇率和政策变化；预算本身不合理却机械分析；用期末余额计算周转率；Bridge没有残差；建议没有责任人。']),
('纠正方法',['在分析底稿中强制设置口径栏、来源栏、待验证栏和行动闭环栏。'])]),
('第22章 贯穿案例经营会议报告','深圳拓维2026年表现为收入增长、毛利下降、库存和应收增加、现金流恶化，同时耗材占比和研发投入提高。',[
('核心判断',['收入增长质量弱于表面增速。耗材增长有长期价值，但硬件降价、树脂良率和产能吸收压低当期毛利。库存与应收共同造成现金缺口。']),
('短期因素',['新品爬坡、树脂良率、部分战略备货，可通过生产改善和库存消化缓解。']),
('结构性因素',['渠道账期延长、桌面机价格压力、扩产节奏快于需求，需要调整商业政策和产能计划。']),
('行动计划',['销售按终端销量和回款重设渠道奖励；生产改善前三大缺陷；供应链冻结无订单慢动SKU采购；研发按里程碑分级项目；财务每月复盘毛利Bridge和CCC。'])])
]

for ci,(ct,intro,subs) in enumerate(chapters):
    if ci: page_break(doc)
    heading(doc,ct,1);add_p(doc,intro)
    for st,paras in subs:
        heading(doc,st,2)
        for para in paras:add_p(doc,para)
    if ct.startswith('第4章'):
        add_table(doc,['比较方法','适合回答','主要限制'],[['同比','相对去年同期是否改善','基数和业务口径变化'],['环比','近期是否出现拐点','季节性'],['预算vs实际','是否完成经营承诺','预算质量'],['Forecast vs实际','预测是否准确','信息更新时间'],['滚动12个月','持续趋势','掩盖近期变化'],['同业比较','提出差异问题','可比性']],[3.2,6.5,6.5])
    if ct.startswith('第6章'):
        add_table(doc,['产品','基期销量','基期价格','本期销量','本期价格','变化含义'],[['E1桌面机','8,726','4,200','9,862','3,990','销量增长但降价'],['P1专业机','2,994','9,800','3,714','9,604','销量贡献为主'],['I1工业机','435','65,000','472','65,650','小批量高单价'],['耗材','149,538','92','227,993','95','装机和复购共同增长']],[3.4,2.4,2.4,2.4,2.4,4.8])
    if ct.startswith('第7章'):
        add_table(doc,['毛利率因素','影响百分点','证据','管理责任'],[['基期毛利率','38.0%','2025产品毛利','财务'],['价格','-0.8','折扣和成交价格','销售'],['产品渠道结构','-1.2','产品和渠道收入','销售'],['材料','-1.0','采购价和BOM','采购/研发'],['良率报废','-0.4','MES和报废单','生产/质量'],['产能利用率','-0.6','有效产能和工时','生产'],['新品及其他','+0.5','试产工单','研发/生产'],['预测毛利率','34.5%','Bridge勾稽','财务']],[4,3,6.3,4])
    if ct.startswith('第11章'):
        add_table(doc,['指标','2025','2026预测','变化','主要原因'],[['DSO','65天','83天','+18天','经销商账期及逾期'],['DIO','82天','121天','+39天','备货、销量和慢动库存'],['DPO','55天','63天','+8天','采购增长和部分延付'],['CCC','92天','141天','+49天','应收和库存共同占用']],[3.2,2.7,3,2.5,7])
    if ct.startswith('第16章'):
        add_table(doc,['商业模式','收入驱动','成本驱动','核心指标'],[['制造业','销量、ASP、Mix','材料、人工、制造费用','良率、利用率、单位成本'],['SaaS','客户、席位、续费、提价','云资源、实施、支持','ARR、NRR、CAC回收期'],['零售','门店、客流、客单、同店','采购、损耗、租金','同店、坪效、周转'],['半导体制造','晶圆量、制程、价格','折旧、材料、良率','利用率、良率、ASP'],['工程施工','订单、履约进度','分包、材料、人工','Backlog、合同资产、现金'],['银行','贷款、存款、息差','资金和信用成本','净息差、不良率、成本收入比']],[3.2,4.5,4.7,5.5])
    if ct.startswith('第19章'):
        add_table(doc,['模块','数据','推荐图','必须回答'],[['核心结论','收入利润现金风险','KPI卡','最重要的三到五件事'],['收入','PVM及结构','瀑布/堆积柱','增长来自哪里'],['毛利','价格结构成本','瀑布','下降因素各多少'],['营运资金','DSO DIO DPO','趋势/分层柱','资金在哪里'],['现金流','净利润到现金','瀑布','利润为何未变现'],['行动','责任人期限目标','跟踪表','下月如何复盘']],[3,5,3.5,6.5])

page_break(doc);heading(doc,'附录A 月度分析取数责任矩阵',1)
add_table(doc,['数据','系统','责任部门','财务核对','更新频率'],[['收入与成本','ERP/GL','财务','与总账和产品明细勾稽','月'],['订单和客户','CRM','销售运营','订单转收入及客户口径','周/月'],['产量良率工时','MES','生产/质量','与完工入库和成本核对','日/月'],['库存和库龄','WMS/ERP','供应链','与总账存货及盘点核对','周/月'],['采购价格','采购系统','采购','与发票和标准成本核对','月'],['回款和逾期','AR/银行','销售/财务','与银行及客户核销','周/月']],[4,3,3.5,6.5,2.5])
heading(doc,'附录B 数据与公式检查清单',1)
bullets(doc,['所有金额使用同一币种和含税/不含税口径。','销量、价格和收入能够相乘勾稽。','PVM、毛利Bridge和现金流Bridge残差为零。','周转率优先使用平均余额和匹配的收入、成本或采购额。','新产品、退出产品和并购范围变化单独处理。','事实、业务信息、判断和待验证事项分开记录。','每项建议设置责任人、截止日期、目标值和复盘证据。'])
heading(doc,'附录C 主要资料来源',1)
add_p(doc,'国家统计局：《2025年全国规模以上工业企业利润增长0.6%》，用于工业行业营业收入、营业成本和宏观周转指标。https://www.stats.gov.cn/zwfwck/sjfb/202601/t20260127_1962382.html')
add_p(doc,'上海证券交易所和深圳证券交易所上市公司定期报告，用于后续维护可比公司样本。https://www.sse.com.cn/disclosure/listedinfo/periodic/ ；https://www.szse.cn/disclosure/listed/fixed/index.html')
add_p(doc,'证监会上市公司行业分类及年度报告披露规则，用于行业分类和披露口径。https://www.csrc.gov.cn/csrc/c100028/c7520295/content.shtml')

page_break(doc);heading(doc,'附录D 财务异常诊断实战卡',1)
add_p(doc,'实战卡用于决定下一步查什么，不直接替代业务访谈和数据验证。每个问题先核对财务口径，再按责任对象下钻。')
diagnostics=[
('收入同比下降','产品/客户/渠道/地区/PVM','订单、取消、出货、签收、终端库存','销售运营、物流、CRM','区分需求下降、交付延迟和收入截止'),
('收入增长但毛利下降','价格、Mix、材料、良率、产能','成交价、BOM、采购价、良率、工时','销售、采购、生产、质量','量化每项对毛利率的百分点影响'),
('毛利额增长但毛利率下降','销量贡献与单位盈利','产品贡献毛利和结构','销售、成本会计','判断规模增长能否弥补单位盈利下降'),
('销售费用率下降','费用增长与收入质量','新增渠道收入、回款、退货、费用归属','销售、财务','费用率改善不代表投入产出健康'),
('研发费用快速增长','人员、项目、外包、设备','工时、里程碑、版本、项目预算','研发PMO、人力','区分战略投入、低效项目和进度延误'),
('应收增长快于收入','账期、逾期、客户结构、截止','合同信用期、逐户账龄、期后回款','销售、AR、法务','量化DSO和逾期客户贡献'),
('存货增长快于销售','原料、在品、成品、发出商品','S&OP、订单、工单、库龄、BOM需求','供应链、生产、销售','区分战略库存和需求预测失误'),
('经营现金流低于净利润','应收、存货、预付、应付、税费','现金流Bridge和科目明细','财务及各业务部门','查清其他经营项目，禁止大额黑箱'),
('产能利用率下降','有效产能、实际产量、停机、换线','标准工时、排产、设备日志','生产、设备、计划','理论产能不能直接作为分母'),
('良率下降','产品、产线、工序、缺陷类型','MES、报废单、返工、材料批次','质量、生产、采购','量化材料、人工和产能三类影响'),
('DPO上升','采购增长、账期、延迟付款','供应商账龄、合同、逾期清单','采购、AP、资金','避免把违约付款当作管理改善'),
('耗材收入占比提高','装机、激活、复购、价格、渠道','设备序列号、活跃设备、耗材订单','客户成功、销售、售后','识别真实复购和渠道压货'),
]
add_table(doc,['异常','分析路径','验证数据','沟通部门','完成标准'],diagnostics[:6],[3,4.2,5.2,3.2,5])
add_table(doc,['异常','分析路径','验证数据','沟通部门','完成标准'],diagnostics[6:],[3,4.2,5.2,3.2,5])

heading(doc,'附录E 月度经营会议报告样稿',1)
heading(doc,'一 核心结论',2)
add_p(doc,'【事实】2026年预测收入4.50亿元，低于预算3,000万元；综合毛利率34.5%，低于预算4.5个百分点；期末存货1.18亿元、应收1.02亿元，经营现金流仅500万元。')
add_p(doc,'【判断】收入缺口并非所有业务同时转弱，主要集中在桌面机和部分经销商。耗材和专业机仍增长。利润压力来自降价、结构、材料、良率和产能利用率。现金压力同时来自库存和应收。')
add_p(doc,'【风险】若渠道去库存和新品低良率延续两个季度，公司将同时面临毛利下降、存货减值和短期融资增加。')
add_p(doc,'【行动】销售按终端销量调整促销，质量改善树脂线前三大缺陷，供应链冻结无订单慢动SKU采购，财务按周跟踪TOP10逾期客户。')
heading(doc,'二 收入与毛利',2)
add_table(doc,['事项','事实及计算','业务解释','待验证','行动'],[['收入','预测4.50亿元，低于预算6.3%','桌面机低于预算，耗材高于预算','渠道库存和终端销量','按Sell-out调整发货'],['毛利率','38.0%降至34.5%','价格-0.8、结构-1.2、材料-1.0、良率-0.4、利用率-0.6、新品其他+0.5个百分点','材料合同与树脂缺陷','采购谈判和良率改善'],['耗材','收入占比22%升至29%','装机和复购增长','活跃设备与兼容耗材','建立设备队列分析']],[3,5,5,4,4])
heading(doc,'三 营运资金与现金',2)
add_table(doc,['指标','2025','2026预测','变化','管理含义'],[['DSO','65天','83天','+18天','客户信用和逾期共同占用现金'],['DIO','82天','121天','+39天','战略备货中混有需求不足和慢动库存'],['DPO','55天','63天','+8天','部分来自采购增长，部分来自延迟付款'],['CCC','92天','141天','+49天','增长所需资金明显上升'],['经营现金流','4,500万元','500万元','-4,000万元','利润没有形成现金，融资需求增加']],[3,3,3,3,8])
heading(doc,'四 行动跟踪',2)
add_table(doc,['行动','责任人','期限','目标值','复盘证据'],[['TOP10客户回款计划','销售总监/财务经理','2026-09-30','逾期余额下降20%','银行到账及应收核销'],['树脂线缺陷改善','质量总监','2026-10-10','良率达到95%','MES良率及报废原因'],['慢动库存采购冻结','供应链总监','2026-09-18','180天以上库存下降15%','库龄和采购订单'],['促销经济性复盘','销售运营/财务BP','2026-09-20','促销增量贡献毛利为正','价格销量和退货']],[5,4,3,4,6])

heading(doc,'附录F 财务分析岗位面试题',1)
questions=[
('收入增长但经营现金流下降，你怎么分析？','考察三张报表联动和业务下钻','从净利润到现金Bridge定位应收、库存、预付、应付，再拆客户和SKU，并区分增长占用与异常占用。'),
('毛利率下降如何定位原因？','考察成本和经营理解','先做价格与结构，再做材料、人工、制造费用、良率、利用率和新品；所有影响勾稽到总变化。'),
('预算完成率105%是否说明经营良好？','考察批判性判断','先评估预算质量，再看收入结构、毛利、现金、一次性订单和未来透支。'),
('存货增加40%是否一定不好？','考察供应链分析','拆原料、在品、成品和库龄，结合订单、战略备货、产能爬坡和滞销判断。'),
('如何证明你的分析原因不是猜测？','考察证据意识','明确事实、业务信息和判断；列取数来源、业务确认人、量化贡献和待验证事项。'),
]
add_table(doc,['问题','面试官为什么问','合格答案应体现'],questions,[5.5,5.5,10])

# headers, page numbers and update fields
for sec in doc.sections:
    hp=sec.header.paragraphs[0];hp.text='企业财务分析与经营分析实战手册';hp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    for r in hp.runs:r.font.size=Pt(8);r.font.color.rgb=RGBColor(100,100,100)
    fp=sec.footer.paragraphs[0];fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=fp.add_run('第 ');fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');run._r.addnext(fld);fp.add_run(' 页')
settings=doc.settings._element;upd=OxmlElement('w:updateFields');upd.set(qn('w:val'),'true');settings.append(upd)
doc.core_properties.title='企业财务分析与经营分析实战手册';doc.core_properties.subject='财务报表分析 经营分析 财务BP';doc.core_properties.author='企业财务实操知识库'
OUT.mkdir(parents=True,exist_ok=True);doc.save(DOCX)

# XMind two-sheet structure
def sid(x):return hashlib.sha1(x.encode('utf-8')).hexdigest()[:26]
def topic(t,path,children=None,root=False):
    d={'id':sid(path),'class':'topic','title':t}
    if root:d['structureClass']='org.xmind.ui.map.unbalanced'
    if children:d['children']={'attached':children}
    return d
def branch(title,leaves,path):return topic(title,path,[topic(x,path+'/'+x) for x in leaves])
overview=[
 branch('1 分析全景',['财务结果','异常识别','指标拆解','业务验证','原因量化','短期长期判断','管理行动','持续跟踪'],'o/1'),
 branch('2 利润表',['收入PVM','毛利Bridge','期间费用','利润质量','有效税率'],'o/2'),
 branch('3 资产负债表',['应收与账龄','存货与库龄','固定资产和产能','应付与供应链','借款和偿债'],'o/3'),
 branch('4 营运资金现金流',['DSO','DIO','DPO','CCC','净利润到现金','自由现金流'],'o/4'),
 branch('5 经营指标',['订单','销量','ASP','利用率','良率','单位成本','交付','复购','质保'],'o/5'),
 branch('6 行业商业模式',['制造业','硬件科技','SaaS','零售','半导体','汽车新能源','医药','工程物流','公用事业','金融不适用毛利率'],'o/6'),
 branch('7 分析输出',['事实','计算','业务信息','判断','待验证','风险','行动建议'],'o/7'),
 branch('8 模拟案例',['3D打印机和耗材','收入增长毛利下降','库存应收增加','现金流恶化','研发增长','经营会议报告'],'o/8')]
detail=[]
for ct,_,subs in chapters:
    detail.append(topic(ct,'d/'+ct,[branch(st,paras,'d/'+ct+'/'+st) for st,paras in subs]))
def sheet(title,root,children,key):return {'id':sid('sheet/'+key),'class':'sheet','title':title,'rootTopic':topic(root,'root/'+key,children,True),'topicPositioning':'fixed'}
content=[sheet('01 财务经营分析全景','企业财务与经营分析全景',overview,'overview'),sheet('02 手册详细目录','企业财务分析与经营分析实战手册',detail,'detail')]
metadata={'creator':{'name':'Xmind','version':'24.0'},'activeSheetId':content[0]['id']};manifest={'file-entries':{'content.json':{},'metadata.json':{},'Thumbnails/thumbnail.png':{}}}
png=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/W9n7WQAAAABJRU5ErkJggg==')
with zipfile.ZipFile(XMIND,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('content.json',json.dumps(content,ensure_ascii=False,separators=(',',':')));z.writestr('metadata.json',json.dumps(metadata,ensure_ascii=False,separators=(',',':')));z.writestr('manifest.json',json.dumps(manifest,ensure_ascii=False,separators=(',',':')));z.writestr('Thumbnails/thumbnail.png',png)
with zipfile.ZipFile(XMIND) as z:
    assert z.testzip() is None;json.loads(z.read('content.json'))
print(DOCX);print(XMIND);print('paragraphs',len(doc.paragraphs),'tables',len(doc.tables))
