import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const base = "C:/Users/HZJ/Desktop/Git/work/month_close_20260909";
const outDir = path.join(base, "deliverables");
const wb = Workbook.create();
const navy="#17365D", blue="#DCE6F1", green="#E2F0D9", amber="#FFF2CC", red="#FCE8E8", line="#D9DEE7", textColor="#20242C";

function title(s, t, sub, endCol){
  s.showGridLines=false;
  s.getRange(`A1:${endCol}1`).merge(); s.getRange("A1").values=[[t]];
  s.getRange(`A1:${endCol}1`).format.fill=navy;
  s.getRange("A1").format.font={name:"Microsoft YaHei",size:16,bold:true,color:"#FFFFFF"};
  s.getRange(`A1:${endCol}1`).format.rowHeight=30;
  s.getRange(`A2:${endCol}2`).merge(); s.getRange("A2").values=[[sub]];
  s.getRange("A2").format.font={name:"Microsoft YaHei",size:10,color:"#566171"};
  s.getRange(`A2:${endCol}2`).format.rowHeight=34; s.getRange(`A2:${endCol}2`).format.wrapText=true;
}
function header(r){r.format.fill=navy;r.format.font={name:"Microsoft YaHei",size:9,bold:true,color:"#FFFFFF"};r.format.wrapText=true;r.format.verticalAlignment="center";r.format.horizontalAlignment="center";r.format.rowHeight=28;}
function body(r){r.format.font={name:"Microsoft YaHei",size:9,color:textColor};r.format.wrapText=true;r.format.verticalAlignment="top";r.format.borders={preset:"all",style:"thin",color:line};}
function table(s,address,name){const t=s.tables.add(address,true,name);t.showFilterButton=true;return t;}

// 01 年度工作框架
{
 const s=wb.worksheets.add("01年度工作框架");s.tabColor=navy;
 title(s,"总账会计与财务主管年度工作框架","基准：自然年度、非上市企业、一般纳税人、月度结账。忙度属于常见实务判断；税务期限须结合主管税务机关核定税期和当年节假日安排。","H");
 s.getRange("A4:H4").values=[["时间","忙度","核心任务","月结与核算","税务与外部事项","管理与内控","关键输出","行业或企业差异"]];header(s.getRange("A4:H4"));
 const rows=[
 ["1月","极忙","上年12月及年度硬关账 年初开账","截止 盘点差异 奖金费用计提 减值 期初余额","季度或月度申报 审计资料","年度问题清单 关账复盘","年度报表 年结底稿","集团合并和上市报告更紧"],
 ["2月","中高","审计与春节压缩月结","短周期截止 审计调整跟踪","常规申报 年终奖个税核对","调整关账日历","审计资料 月结包","春节月份通常提前交付"],
 ["3月","高","审计 汇算准备 Q1预结","税会差异 研发 资产损失","汇算底稿 关联资料准备","季度预测","汇算初稿 Q1预结","研发和项目制企业工作量更大"],
 ["4月","极忙","Q1关账和预缴 年报集中期","季度深度复核 审计调整落账","企业所得税预缴等","经营复盘","Q1报表 申报勾稽","上市及集团披露更紧"],
 ["5月","极忙","企业所得税汇算收尾","年度报表与纳税申报核对","年度终了后5个月内完成汇算","税务档案归档","年度申报 汇算档案","有优惠和境外关联交易更复杂"],
 ["6月","高","半年预结 盘点 预测","存货资产抽盘 减值预判","常规申报","滚动预测","半年预结 预测更新","制造业重点在库存和产能"],
 ["7月","极忙","半年关账 Q2预缴","半年报表 深度科目复核","企业所得税预缴等","半年经营复盘","半年报表 Q2申报","集团可能需半年审阅"],
 ["8月","中","清理遗留 流程改进","长期挂账 暂估 主数据 接口","税务问题整改","自动化和制度优化","问题关闭 改进方案","适合补课和流程项目"],
 ["9月","高","Q3预结 次年预算准备","年末风险清单","税负预测","预算假设和资本开支","Q3预结 预算底稿","业务波动影响预算节奏"],
 ["10月","极忙","Q3关账和预算","季度深度复核","企业所得税预缴等","预算编制 盘点计划","Q3报表 预算初稿","制造和贸易开始年末备货检查"],
 ["11月","高","预算定稿 年结预演","减值 关联方 函证 PBC","年度税会差异预估","年结演练","预算定稿 年结清单","复杂集团宜做试关账"],
 ["12月","极忙","年度截止和盘点","收发货 服务截止 盘点 资产 奖金 减值","年度税务资料封存","审计和盘点组织","截止证据 盘点资料","出口 项目制和多地点盘点更复杂"]
 ];
 s.getRange(`A5:H${4+rows.length}`).values=rows;body(s.getRange(`A5:H${4+rows.length}`));table(s,`A4:H${4+rows.length}`,"AnnualCloseCalendar");
 for(let i=0;i<rows.length;i++){const r=5+i,lv=rows[i][1];s.getRange(`B${r}`).format.fill=lv==="极忙"?red:lv==="高"?amber:lv==="中高"?blue:green;s.getRange(`A${r}:H${r}`).format.rowHeight=54;}
 s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(2);[9,9,24,32,28,26,28,30].forEach((w,i)=>s.getRangeByIndexes(0,i,17,1).format.columnWidth=w);
}

// 02 月结日历
{
 const s=wb.worksheets.add("02月结日历");s.tabColor="#548235";
 title(s,"D-5至D+10月结日历","公司内部范本，可按实际改为D+3、D+5或D+10。状态、实际完成日、问题编号留给使用者维护。","L");
 const heads=["时间","阶段","任务","输入及依赖","制作人","复核人","输出证据","完成标准","状态","实际完成日","问题编号","备注"];
 s.getRange("A4:L4").values=[heads];header(s.getRange("A4:L4"));
 const rows=[
 ["D-5","预结","发布关账通知 更新任务依赖","报表提交日 上月问题","财务主管","财务负责人","关账日历","责任人和截止时间已确认","未开始","","",""],
 ["D-4","预结","预跑账龄 暂估 负库存 资产和税费异常","ERP明细 上月台账","总账会计","财务主管","预结异常清单","高风险事项已指派","未开始","","",""],
 ["D-3","预结","收集未开票收货和服务计提","采购 仓库 业务确认","应付会计","总账会计","暂估初稿","逐项有合同验收和金额依据","未开始","","",""],
 ["D-2","截止","确认最后收发货和验收序号","仓库 销售 采购","库存和收入会计","总账会计","截止清单","单据连续且倒签受控","未开始","","",""],
 ["D-1","截止","现金存货抽盘 工资资产数据锁定","盘点表 工资表 资产卡片","各模块会计","总账会计","盘点和锁定证据","差异已记录","未开始","","",""],
 ["D0","截止","完成当月业务单据并保存截止快照","业务系统和ERP","各模块会计","财务主管","月末快照","关键系统截止一致","未开始","","",""],
 ["D+1","子模块","完成AP AR 费用 资金 薪酬 资产 库存接口","子模块完成确认","各模块会计","总账会计","接口控制表","批次数 记录数 金额 错误为零","未开始","","",""],
 ["D+2","调整","暂估 计提 摊销 折旧 利息 税费 外币重估","估计底稿 合同 台账","总账会计","财务主管","调整分录清单","重大分录有依据和审批","未开始","","",""],
 ["D+3","成本","成本归集 分配 完工和销售成本结转","领料 工时 产量 在产品","成本会计","总账会计","成本计算表","库存成本总账可勾稽","未开始","","",""],
 ["D+4","核对","全科目核对和差异跟踪","明细账 外部对账 业务台账","总账会计","财务主管","余额调节表","差异为零或获批","未开始","","",""],
 ["D+5","分析","试算平衡 初版报表 异常和波动分析","最终凭证 预算 业务指标","总账会计","财务主管","初版报表 分析","重大波动有证据","未开始","","",""],
 ["D+6","报告","三表 附注 集团包及账税勾稽复核","试算平衡和调节表","总账会计","财务主管","报表勾稽表","版本一致 勾稽通过","未开始","","",""],
 ["D+7","锁账","签批 发布报表 锁定期间","关账包 未决事项","财务主管","财务负责人","签批和锁账记录","重大事项关闭或批准","未开始","","",""],
 ["D+8至D+10","关账后","管理分析 复盘 更新下月行动","正式报表 问题台账","财务主管","财务负责人","管理分析 复盘","根因和改进责任明确","未开始","","",""]
 ];
 s.getRange(`A5:L${4+rows.length}`).values=rows;body(s.getRange(`A5:L${4+rows.length}`));table(s,`A4:L${4+rows.length}`,"MonthlyCloseCalendar");
 s.getRange(`I5:I${4+rows.length}`).dataValidation={rule:{type:"list",values:["未开始","进行中","已完成","阻塞","不适用"]}};
 s.getRange(`J5:J${4+rows.length}`).setNumberFormat("yyyy-mm-dd hh:mm");s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(3);
 [10,11,34,34,16,16,30,32,12,17,13,26].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);s.getRange(`A5:L${4+rows.length}`).format.rowHeight=58;
}

// 03 模块检查清单
{
 const s=wb.worksheets.add("03模块检查清单");s.tabColor="#5B9BD5";
 title(s,"月结模块检查清单","沿业务发生、原始单据、系统、会计判断、分录税务、对账、报表和风险闭环。","M");
 const heads=["编号","模块","检查事项","原始单据","系统或数据","会计判断或分录","税务处理","对账对象","报表影响","风险点","建议日","状态","证据位置"];
 s.getRange("A4:M4").values=[heads];header(s.getRange("A4:M4"));
 const rows=[
 ["C01","资金","银行余额及未达账","对账单 回单 现金盘点","网银 资金模块 GL","补记手续费 利息 未达不直接调账","利息和手续费凭证","银行对账单","货币资金 现金流","长期未达 受限资金漏披露","D+1至D+4","未开始",""],
 ["C02","采购应付","未开票收货暂估","合同 PO 入库 验收","采购库存 AP GL","按已发生义务暂估 次月匹配冲回","进项抵扣与会计确认分开","收货 发票 供应商","存货 应付 成本","漏提 重复暂估 价格差","D+2","未开始",""],
 ["C03","销售应收","收入和发货截止","合同 发货 签收 验收","销售 AR GL 发票","按履约和控制权判断收入","核对纳税义务和开票","销售 发票 回款 客户","收入 应收 税费","提前收入 未开票漏税 退货","D0至D+2","未开始",""],
 ["C04","费用","服务费用和预付摊销","合同 验收 账单 预算","费用 OA AP GL","已接受服务计提 跨期费用摊销","税前扣除凭证另核","业务台账 合同","期间费用 应付 预付","只按发票入账 计提无依据","D+2","未开始",""],
 ["C05","薪酬","工资奖金社保个税","工资表 人员 考勤 审批","HR 薪酬 GL 个税","按受益对象计提和分配","次月个税扣缴申报核对","工资 银行 个税 社保","成本费用 应付职工薪酬","人员差异 奖金无依据","D+2至D+4","未开始",""],
 ["C06","固定资产","增减变动折旧转固减值","验收 投产 处置 盘点","资产模块 项目 GL","确定可使用状态 折旧和减值","税法折旧差异台账","资产台账 实物 GL","资产 折旧 费用","已投产未转固 闲置未评估","D+2至D+4","未开始",""],
 ["C07","存货成本","成本归集完工在产品销售成本","领料 工时 产量 盘点","仓储 生产 成本 GL","按政策分配并复核在产品","成本凭证与扣除条件","数量 金额 成本模块 GL","存货 成本 毛利","负库存 倒挤成本 异常损耗","D+3至D+5","未开始",""],
 ["C08","往来","账龄性质和清理","对账单 合同 回款付款","AR AP GL","重分类 减值 核销须有审批","坏账和核销税务资料","客户供应商关联方","往来 减值 披露","负数 长期挂账 串户","D+4","未开始",""],
 ["C09","税务","账票表款四层核对","发票 申报表 完税凭证","税务平台 发票平台 GL","建立差异桥接 不机械求相等","增值税 所得税 个税 印花税","发票 申报 银行 GL","税费 利润 现金流","漏报 错期 异常凭证","D+4至申报前","未开始",""],
 ["C10","融资外币","本金利息汇兑和分类","合同 对账单 汇率","资金 外币重估 GL","计提利息 外币重估 长短期分类","利息扣除条件","合同 银行 GL","借款 财务费用","漏提利息 汇率错误","D+2至D+4","未开始",""],
 ["C11","报表","试算映射三表附注勾稽","最终TB 报表映射","GL 报表 合并","重分类 抵销 调整有依据","申报表与报表桥接","TB 报表 集团包","三表和附注","版本不一致 公式错","D+5至D+6","未开始",""],
 ["C12","锁账归档","签批锁账和更正","关账包 审批","ERP期间 档案系统","重开需说明原因影响","更正影响税务时同步评估","分录 报表 审批","全部报表","静默改账 证据缺失","D+7","未开始",""]
 ];
 s.getRange(`A5:M${4+rows.length}`).values=rows;body(s.getRange(`A5:M${4+rows.length}`));table(s,`A4:M${4+rows.length}`,"ModuleChecklist");
 s.getRange(`L5:L${4+rows.length}`).dataValidation={rule:{type:"list",values:["未开始","进行中","已完成","阻塞","不适用"]}};
 s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(3);[9,13,28,28,25,36,28,27,25,28,15,12,30].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);s.getRange(`A5:M${4+rows.length}`).format.rowHeight=68;
}

// 04 科目对账与报表勾稽
{
 const s=wb.worksheets.add("04科目与报表勾稽");s.tabColor="#8064A2";
 title(s,"科目对账与报表勾稽底稿","先明确账面余额、外部或子系统余额和可解释调整。未解释差异不要直接用分录抹平。","N");
 const heads=["科目或项目","期末账面余额","对账来源余额","时间性差异","错误差异","调整后对方余额","未解释差异","阈值","结论","报表项目","制作人","复核人","状态","证据位置"];
 s.getRange("A4:N4").values=[heads];header(s.getRange("A4:N4"));
 const subjects=["库存现金","银行存款","应收账款","合同资产或负债","其他应收款","原材料","在产品","库存商品","固定资产及累计折旧","在建工程","应付账款及暂估","其他应付款","应付职工薪酬","应交税费","短期及长期借款","关联方往来","未分配利润","现金及现金等价物"];
 const rows=subjects.map(x=>[x,"","","","","","","","待核对","","","","未开始",""]);
 s.getRange(`A5:N${4+rows.length}`).values=rows;body(s.getRange(`A5:N${4+rows.length}`));table(s,`A4:N${4+rows.length}`,"AccountReconciliation");
 // formulas: adjusted counterparty = source + timing; unexplained = book - adjusted counterparty - error
 for(let r=5;r<5+rows.length;r++){s.getRange(`F${r}`).formulas=[[`=IF(COUNT(C${r}:D${r})=0,"",C${r}+D${r})`]];s.getRange(`G${r}`).formulas=[[`=IF(COUNT(B${r}:F${r})=0,"",B${r}-F${r}-E${r})`]];s.getRange(`I${r}`).formulas=[[`=IF(G${r}="","待核对",IF(ABS(G${r})<=IF(H${r}="",0,H${r}),"通过","需处理"))`]];}
 s.getRange(`B5:H${4+rows.length}`).setNumberFormat("#,##0.00;[Red]-#,##0.00");
 s.getRange(`M5:M${4+rows.length}`).dataValidation={rule:{type:"list",values:["未开始","进行中","已完成","阻塞","不适用"]}};
 s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(2);[25,16,16,16,16,18,16,12,14,22,15,15,12,30].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);s.getRange(`A5:N${4+rows.length}`).format.rowHeight=34;
}

// 05 问题跟踪
{
 const s=wb.worksheets.add("05问题跟踪");s.tabColor="#C65911";
 title(s,"月结差异和问题跟踪","每个未决事项必须有金额、影响、责任人和完成日。已锁账事项若需更正，应关联重开或次期调整审批。","O");
 const heads=["问题编号","期间","模块","问题描述","金额","优先级","会计影响","税务影响","临时处理","根因","整改措施","责任人","计划完成日","状态","关闭证据"];
 s.getRange("A4:O4").values=[heads];header(s.getRange("A4:O4"));
 const rows=Array.from({length:20},(_,i)=>[`ISS-${String(i+1).padStart(3,"0")}`,"","","","","中","","","","","","","","未开始",""]);
 s.getRange("A5:O24").values=rows;body(s.getRange("A5:O24"));table(s,"A4:O24","CloseIssueLog");
 s.getRange("F5:F24").dataValidation={rule:{type:"list",values:["高","中","低"]}};s.getRange("N5:N24").dataValidation={rule:{type:"list",values:["未开始","进行中","待复核","已关闭","不适用"]}};s.getRange("M5:M24").setNumberFormat("yyyy-mm-dd");s.getRange("E5:E24").setNumberFormat("#,##0.00;[Red]-#,##0.00");
 s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(4);[13,12,14,34,14,10,27,27,30,30,30,15,15,12,30].forEach((w,i)=>s.getRangeByIndexes(0,i,24,1).format.columnWidth=w);s.getRange("A5:O24").format.rowHeight=42;
}

wb.recalculate();
await fs.mkdir(outDir,{recursive:true});
const inspect=await wb.inspect({kind:"workbook,sheet,table,formula",maxChars:12000,tableMaxRows:6,tableMaxCols:8,tableMaxCellChars:100});
await fs.writeFile(path.join(base,"workbook_inspect.txt"),inspect.ndjson??String(inspect),"utf8");
const previews=[
 ["01年度工作框架","A1:H16","preview-01.png"],
 ["02月结日历","A1:L18","preview-02.png"],
 ["03模块检查清单","A1:M16","preview-03.png"],
 ["04科目与报表勾稽","A1:N12","preview-04.png"],
 ["05问题跟踪","A1:O12","preview-05.png"]
];
for(const [sheetName,range,file] of previews){const blob=await wb.render({sheetName,range,scale:1,format:"png"});await fs.writeFile(path.join(base,file),new Uint8Array(await blob.arrayBuffer()));}
const xlsx=await SpreadsheetFile.exportXlsx(wb);await xlsx.save(path.join(outDir,"M02-001-年度工作与月结执行清单.xlsx"));
console.log(JSON.stringify({sheets:previews.map(x=>x[0]),output:path.join(outDir,"M02-001-年度工作与月结执行清单.xlsx")}));
