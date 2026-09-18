import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const base = new URL("./", import.meta.url).pathname.slice(1);
const outDir = `${base}deliverables`;
const qaDir = `${base}qa_sheets`;
await fs.mkdir(outDir,{recursive:true}); await fs.mkdir(qaDir,{recursive:true});
const data=JSON.parse(await fs.readFile(`${base}content.json`,"utf8"));
const navy="#17365D", blue="#2F75B5", pale="#EAF2F8", gray="#F3F6F9", line="#D9E1E8", amber="#FFF2CC", green="#E2F0D9", red="#FCE4D6", font="Microsoft YaHei";
function col(n){let s="";while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function title(sh,text,sub,endCol){sh.showGridLines=false;sh.getRange(`A1:${col(endCol)}1`).merge();sh.getRange("A1").values=[[text]];sh.getRange("A1").format={font:{name:font,size:16,bold:true,color:"#000000"},rowHeight:30,verticalAlignment:"center"};sh.getRange(`A2:${col(endCol)}2`).merge();sh.getRange("A2").values=[[sub]];sh.getRange("A2").format={font:{name:font,size:9,italic:true,color:"#666666"},rowHeight:24,wrapText:true};}
function header(sh,addr){sh.getRange(addr).format={fill:navy,font:{name:font,size:9,bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,rowHeight:34,borders:{preset:"all",style:"thin",color:"#FFFFFF"}}}
function body(sh,addr,size=9){sh.getRange(addr).format={font:{name:font,size:Math.round(size)},verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:line}}}
function widths(sh,map){for(const [c,w] of Object.entries(map)) sh.getRange(`${c}:${c}`).format.columnWidth=w}
function statusValidation(r){r.dataValidation={rule:{type:"list",values:["未开始","学习中","待复练","已掌握","不适用"]}}}
function statusCF(r){r.conditionalFormats.add("containsText",{text:"已掌握",format:{fill:green,font:{color:"#548235",bold:true}}});r.conditionalFormats.add("containsText",{text:"未开始",format:{fill:amber,font:{color:"#9C6500"}}});r.conditionalFormats.add("containsText",{text:"待复练",format:{fill:red,font:{color:"#C00000"}}});}
function makeTableSheet(wb,name,text,sub,headers,rows,widthMap,tableName){const sh=wb.worksheets.add(name);title(sh,text,sub,headers.length);sh.getRange(`A5:${col(headers.length)}5`).values=[headers];header(sh,`A5:${col(headers.length)}5`);if(rows.length){sh.getRange(`A6:${col(headers.length)}${5+rows.length}`).values=rows;body(sh,`A6:${col(headers.length)}${5+rows.length}`);sh.tables.add(`A5:${col(headers.length)}${5+rows.length}`,true,tableName).style="TableStyleMedium2";}widths(sh,widthMap);sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(1);return sh;}

// Workbook 1: training
const w1=Workbook.create();
const names=["00-学习看板","01-能力地图","02-20项实操训练","03-训练步骤与证据","04-凭证科目影响","05-模块对账矩阵","06-月结SOP","07-学习资源","08-个人错题复练"];
for(const n of names) w1.worksheets.add(n);
{
 const sh=w1.worksheets.getItem(names[0]);title(sh,"ERP财务实操训练清单","面向总账会计和财务主管；案例公司：华创智联设备有限公司；版本日期：2026-09-10",10);
 sh.getRange("A5:D5").values=[["学习指标","当前值","目标","说明"]];header(sh,"A5:D5");
 sh.getRange("A6:A11").values=[["案例总数"],["已掌握案例"],["学习中案例"],["待复练案例"],["完成率"],["未开始高优先级"]];
 sh.getRange("B6:B11").formulas=[["=COUNTA('02-20项实操训练'!A6:A25)"],["=COUNTIF('02-20项实操训练'!Q6:Q25,\"已掌握\")"],["=COUNTIF('02-20项实操训练'!Q6:Q25,\"学习中\")"],["=COUNTIF('02-20项实操训练'!Q6:Q25,\"待复练\")"],["=IF(B6=0,0,B7/B6)"],["=COUNTIFS('02-20项实操训练'!P6:P25,\"最高\",'02-20项实操训练'!Q6:Q25,\"未开始\")"]];
 sh.getRange("C6:C11").values=[[20],[20],[0],[0],[1],[0]];sh.getRange("D6:D11").values=[["固定20个贯穿案例"],["全部能够独立解释和检查"],["应逐步归零"],["复练后再标已掌握"],["已掌握/案例总数"],["优先处理GL、月结、AP、AR、成本和对账"]];body(sh,"A6:D11");sh.getRange("B10:C10").format.numberFormat="0%";
 sh.getRange("F5:J5").values=[["模块","案例数","已掌握","完成率","优先级"]];header(sh,"F5:J5");
 const mods=["主数据与P2P","采购","采购与库存","AP与月结","AP","AP与成本","资金与AP","生产与库存","薪酬与成本","成本","销售与AR","销售、AR与税务","资金与AR","OA与费用","FA","GL与月结","异常处理"];
 const mend=5+mods.length;sh.getRange(`F6:F${mend}`).values=mods.map(x=>[x]);sh.getRange(`G6:G${mend}`).formulas=mods.map((x,i)=>[`=COUNTIF('02-20项实操训练'!C$6:C$25,F${i+6})`]);sh.getRange(`H6:H${mend}`).formulas=mods.map((x,i)=>[`=COUNTIFS('02-20项实操训练'!C$6:C$25,F${i+6},'02-20项实操训练'!Q$6:Q$25,\"已掌握\")`]);sh.getRange(`I6:I${mend}`).formulas=mods.map((x,i)=>[`=IF(G${i+6}=0,0,H${i+6}/G${i+6})`]);sh.getRange(`J6:J${mend}`).values=mods.map(x=>[/(GL|成本|AP|AR|库存)/.test(x)?"最高":"其次"]);body(sh,`F6:J${mend}`);sh.getRange(`I6:I${mend}`).format.numberFormat="0%";
 widths(sh,{A:23,B:14,C:12,D:38,E:3,F:22,G:11,H:11,I:12,J:11});sh.freezePanes.freezeRows(5);
}
{
 const sh=w1.worksheets.getItem(names[1]);title(sh,"总账会计和财务主管ERP能力地图","先理解业务事件和对账逻辑，再学习具体按钮",10);
 const hdr=["能力编号","模块","岗位能力","优先级","应理解业务","应掌握系统逻辑","关键报表","核心对账","面试表达","状态"];
 const rows=[
 ["CAP-01","GL","识别四类凭证来源","最高","各模块何时形成会计影响","自动、接口、手工、月结凭证","凭证来源统计、科目明细","来源系统控制总数与GL","能从凭证追到源单和接口批次","未开始"],
 ["CAP-02","月结","组织模块关闭顺序","最高","库存生产成本等依赖","期间、任务依赖和阻断条件","月结状态表、未完成任务","模块状态与GL期间","能解释为什么成本未结不能最终关账","未开始"],
 ["CAP-03","AP","管理收货暂估和财务应付","最高","收货、发票、付款","三单匹配、未清项和清账","未开票收货、供应商余额","AP与应付/暂估GL","能说明暂估到票和价差","未开始"],
 ["CAP-04","AR","区分订单交付收入开票收款","最高","履约、开票、收款","信用、应收未清项和认领","账龄、收入开票桥","AR与应收GL","能说明开票不等于收入","未开始"],
 ["CAP-05","库存","控制数量和价值完整性","最高","收发存、盘点、计价","数量账、价值账和凭证生成","收发存、库存余额","库存与存货GL","能排查负库存零成本","未开始"],
 ["CAP-06","成本","完成制造成本月结","最高","BOM、工单、领料、报工、费用","在制、完工、分配和结算","工单成本、差异分析","成本模块与GL","能拆解单位成本异常","未开始"],
 ["CAP-07","FA","管理卡片折旧处置","最高","验收、可使用状态、处置","类别、折旧运行和凭证","资产价值表","FA与GL","能排查已使用未转固","未开始"],
 ["CAP-08","对账","建立子模块GL核对矩阵","最高","模块范围和截止","控制科目、维度和过账状态","模块余额和GL","八大模块对账","能由汇总差异追至单据","未开始"],
 ["CAP-09","费用","审核电子报销和服务采购","其次","真实性、预算、归属","OA接口、OCR、付款核销","费用接口、员工往来","OA/AP/GL","能区分系统校验和专业判断","未开始"],
 ["CAP-10","资金","识别银行状态和未达","其次","付款、收款、退汇","银企状态、回单和清账","银行调节表","银行与GL","能处理未知状态而不重复付款","未开始"],
 ["CAP-11","薪酬","核对HR工资财务银行","其次","工资社保个税和费用归属","接口批次和成本中心映射","工资汇总、银行发薪","薪酬与GL","能解释工资不全进管理费用","未开始"],
 ["CAP-12","税务接口","解释账票税差异","其次","开票收票进销项申报","发票平台、ERP和电子税务局","账票税桥接","Tax与GL","能说明合理差异和异常差异","未开始"],
 ["CAP-13","主数据","控制关键主数据质量","再次","客户供应商物料BOM科目","创建变更冻结和日志","变更清单","主数据至交易","能解释垃圾输入导致自动错账","未开始"],
 ["CAP-14","接口","排查漏传重复和延迟","再次","源系统到目标系统","唯一键、控制总数、重跑","接口日志","源总数与目标总数","能界定财务和IT分工","未开始"],
 ["CAP-15","权限","复核职责分离和紧急权限","再次","主数据付款关账","角色、审批和日志","权限冲突报告","人员角色与操作日志","能提出小团队补偿控制","未开始"]];
 sh.getRange("A5:J5").values=[hdr];header(sh,"A5:J5");sh.getRange("A6:J20").values=rows;body(sh,"A6:J20");statusValidation(sh.getRange("J6:J20"));statusCF(sh.getRange("J6:J20"));widths(sh,{A:12,B:12,C:28,D:11,E:30,F:32,G:26,H:25,I:34,J:11});sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(2);sh.tables.add("A5:J20",true,"CapabilityMap").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[2]);title(sh,"20项ERP贯穿案例训练","黄色列为学习者输入；系统名称和按钮应按所在企业版本补充",20);
 const hdr=["案例编号","案例名称","模块","业务背景","原始单据","系统操作","会计分录","系统结果","总账检查","常见错误","凭证触发点","子模块","GL影响","对账对象","是否影响月结","优先级","掌握状态","练习次数","最近练习日期","个人备注"];
 const rows=data.cases.map(c=>[c.id,c.title,c.module,c.background,c.docs,c.actions,c.journal,c.result,c.check,c.errors,"由业务事件及企业配置决定",c.module,"见会计分录","源单/子模块/GL",/(主数据|销售订单)/.test(c.title)?"视情况":"是",/(GL|月结|AP|成本|库存|应付|应收|固定)/.test(c.module)?"最高":"其次","未开始",0,null,""]);
 sh.getRange("A5:T5").values=[hdr];header(sh,"A5:T5");sh.getRange("A6:T25").values=rows;body(sh,"A6:T25",7.5);sh.getRange("Q6:T25").format.fill=amber;statusValidation(sh.getRange("Q6:Q25"));statusCF(sh.getRange("Q6:Q25"));sh.getRange("R6:R25").dataValidation={rule:{type:"wholeNumber",operator:"between",formula1:0,formula2:99}};sh.getRange("S6:S25").format.numberFormat="yyyy-mm-dd";widths(sh,{A:12,B:24,C:18,D:40,E:34,F:46,G:48,H:36,I:36,J:34,K:22,L:18,M:22,N:24,O:13,P:11,Q:12,R:11,S:14,T:30});sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(3);sh.tables.add("A5:T25",true,"ERP20Cases").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[3]);title(sh,"训练步骤与证据","每个案例按七步记录输入、动作、预期和实际结果",12);
 const hdr=["案例编号","步骤号","步骤类型","操作目的","输入数据或资料","系统动作","预期结果","实际结果","截图或证据位置","差异","处理记录","完成状态"];
 const types=["业务与单据","主数据","系统操作","会计凭证","子模块结果","总账和对账","异常复盘"];
 let rows=[];for(const c of data.cases){types.forEach((t,i)=>rows.push([c.id,i+1,t,["理解业务实质","确认基础数据","完成系统动作","验证会计影响","检查模块状态","完成总账核对","总结错误和预防"][i],i===0?c.docs:i===1?"组织、客户/供应商/物料、科目、税码和期间":i===2?c.actions:i===3?c.journal:i===4?c.result:i===5?c.check:c.errors,"由学习者按实际软件填写",i===3?"凭证金额、科目和维度正确":i===5?"源单、子模块和GL可勾稽":"完成并保留证据","","","","","未开始"]));}
 sh.getRange("A5:L5").values=[hdr];header(sh,"A5:L5");sh.getRange(`A6:L${5+rows.length}`).values=rows;body(sh,`A6:L${5+rows.length}`,7.5);sh.getRange(`H6:L${5+rows.length}`).format.fill=amber;statusValidation(sh.getRange(`L6:L${5+rows.length}`));statusCF(sh.getRange(`L6:L${5+rows.length}`));widths(sh,{A:12,B:9,C:16,D:26,E:38,F:42,G:34,H:28,I:28,J:24,K:28,L:11});sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(2);sh.tables.add(`A5:L${5+rows.length}`,true,"TrainingEvidence").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[4]);title(sh,"业务事件与凭证科目影响","分录为案例示范；实际科目和触发点取决于企业政策及系统配置",12);
 const hdr=["案例编号","业务事件","来源单据","凭证来源","是否必然入账","示例分录","关键辅助核算","冲销方式","子模块","核对报表","判断提示","复核状态"];
 const rows=data.cases.map(c=>[c.id,c.title,c.docs,"自动/接口/手工取决于系统设计",/(新增|采购1,000|创建客户)/.test(c.title)?"通常否":"通常是或需判断",c.journal,"公司、客户/供应商、物料、成本中心、项目、税码","优先从原单据执行可追溯冲销",c.module,c.check,"先确认真实业务，再判断会计事件","未开始"]);
 sh.getRange("A5:L5").values=[hdr];header(sh,"A5:L5");sh.getRange("A6:L25").values=rows;body(sh,"A6:L25",7.6);statusValidation(sh.getRange("L6:L25"));statusCF(sh.getRange("L6:L25"));widths(sh,{A:12,B:25,C:36,D:25,E:14,F:52,G:36,H:30,I:18,J:38,K:30,L:11});sh.freezePanes.freezeRows(5);sh.tables.add("A5:L25",true,"JournalImpact").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[5]);title(sh,"ERP子模块与总账核对矩阵","先统一组织、币种、截止日、过账状态和科目范围，再比较余额",12);
 const hdr=["子模块","GL范围","子模块报表","总账依据","统一口径","常见差异","排查路径","责任","本月子模块金额","本月GL金额","差异","复核状态"];
 const rows=data.recons.map((r,i)=>[...r,0,0,null,"未开始"]);sh.getRange("A5:L5").values=[hdr];header(sh,"A5:L5");sh.getRange("A6:J13").values=rows.map(r=>r.slice(0,10));sh.getRange("K6:K13").formulas=rows.map((r,i)=>[`=I${i+6}-J${i+6}`]);sh.getRange("L6:L13").values=rows.map(r=>[r[11]]);body(sh,"A6:L13",7.6);sh.getRange("I6:J13").format.fill=amber;sh.getRange("I6:K13").format.numberFormat="#,##0.00;[Red](#,##0.00);-";statusValidation(sh.getRange("L6:L13"));statusCF(sh.getRange("L6:L13"));widths(sh,{A:12,B:34,C:38,D:30,E:34,F:38,G:42,H:24,I:16,J:16,K:14,L:11});sh.freezePanes.freezeRows(5);sh.tables.add("A5:L13",true,"SubledgerGLRecon").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[6]);title(sh,"ERP月结SOP","日期为相对月末；技术上可关账不代表业务和会计已经完整关闭",13);
 const hdr=["序号","阶段","关键动作","计划时间","责任人","完成证据","主要风险","前置任务","是否阻断后续","执行状态","实际完成日","异常及处理","复核人"];
 const rows=data.close_steps.map((r,i)=>[...r,i===0?"无":data.close_steps[i-1][1],i<11?"是":"最终关闭","未开始",null,"","财务主管"]);
 sh.getRange("A5:M5").values=[hdr];header(sh,"A5:M5");sh.getRange("A6:M17").values=rows;body(sh,"A6:M17",7.7);sh.getRange("J6:L17").format.fill=amber;statusValidation(sh.getRange("J6:J17"));statusCF(sh.getRange("J6:J17"));sh.getRange("K6:K17").format.numberFormat="yyyy-mm-dd";widths(sh,{A:8,B:15,C:45,D:12,E:18,F:38,G:34,H:18,I:14,J:11,K:14,L:32,M:14});sh.freezePanes.freezeRows(5);sh.tables.add("A5:M17",true,"MonthEndSOP").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[7]);title(sh,"官方学习资源与版本追踪","实际照界面操作前再次核验产品线、版本和企业配置",10);
 const hdr=["软件产品","资料名称","版本或时间","适合学习内容","官方链接","是否真实界面来源","版本过时风险","观看/阅读状态","个人笔记","复核日期"];
 const rows=data.sources.map(r=>[...r,"只有页面提供版本明确界面时采用","中到高；按页面版本判断","未开始","",null]);sh.getRange("A5:J5").values=[hdr];header(sh,"A5:J5");sh.getRange(`A6:J${5+rows.length}`).values=rows;body(sh,`A6:J${5+rows.length}`,7.8);statusValidation(sh.getRange(`H6:H${5+rows.length}`));statusCF(sh.getRange(`H6:H${5+rows.length}`));sh.getRange(`J6:J${5+rows.length}`).format.numberFormat="yyyy-mm-dd";widths(sh,{A:24,B:38,C:25,D:46,E:70,F:28,G:30,H:12,I:36,J:14});sh.freezePanes.freezeRows(5);sh.tables.add(`A5:J${5+rows.length}`,true,"LearningSources").style="TableStyleMedium2";
}
{
 const sh=w1.worksheets.getItem(names[8]);title(sh,"个人错题与复练记录","保留错误现象、原因和证据，复练通过后再标记已掌握",12);
 const hdr=["记录编号","日期","案例编号","错误现象","错误原因","正确逻辑","是否涉及会计判断","正确处理层级","复练日期","复练结果","证据位置","备注"];
 const rows=Array.from({length:40},(_,i)=>[`ERR-${String(i+1).padStart(3,"0")}`,null,"","","","","","",null,"未开始","",""]);
 sh.getRange("A5:L5").values=[hdr];header(sh,"A5:L5");sh.getRange("A6:L45").values=rows;body(sh,"A6:L45",8);sh.getRange("B6:B45").format.numberFormat="yyyy-mm-dd";sh.getRange("I6:I45").format.numberFormat="yyyy-mm-dd";sh.getRange("G6:G45").dataValidation={rule:{type:"list",values:["是","否","不确定"]}};sh.getRange("H6:H45").dataValidation={rule:{type:"list",values:["业务源单","主数据配置","接口批次","子模块","GL会计判断","报表映射"]}};statusValidation(sh.getRange("J6:J45"));statusCF(sh.getRange("J6:J45"));widths(sh,{A:13,B:13,C:12,D:32,E:34,F:42,G:18,H:18,I:13,J:11,K:28,L:30});sh.freezePanes.freezeRows(5);sh.tables.add("A5:L45",true,"PersonalErrors").style="TableStyleMedium2";
}
w1.recalculate();
const inspect1=await w1.inspect({kind:"table",range:"00-学习看板!A1:J23",include:"values,formulas",tableMaxRows:30,tableMaxCols:12});await fs.writeFile(`${qaDir}/training_inspect.ndjson`,inspect1.ndjson??String(inspect1));
const err1=await w1.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:300},summary:"training formula scan"});await fs.writeFile(`${qaDir}/training_errors.ndjson`,err1.ndjson??String(err1));
for(const n of names){const b=await w1.render({sheetName:n,autoCrop:"all",scale:0.8,format:"png"});await fs.writeFile(`${qaDir}/training_${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f1=await SpreadsheetFile.exportXlsx(w1);await f1.save(`${outDir}/ERP财务实操训练清单.xlsx`);

// Workbook 2: exceptions
const w2=Workbook.create();const enames=["00-异常看板","01-69项异常总表","02-采购与应付","03-销售与应收","04-库存生产成本","05-费用资产薪酬","06-银行与税务","07-总账月结报表","08-接口主数据权限","09-问题处理记录"];
for(const n of enames)w2.worksheets.add(n);
const map={MD:"主数据",P2P:"采购与应付",O2C:"销售与应收",COST:"库存生产成本",EFP:"费用资产薪酬",BT:"银行与税务",IF:"接口批处理",GL:"总账月结报表",AUTH:"权限内控"};
{
 const sh=w2.worksheets.getItem(enames[0]);title(sh,"财务ERP常见异常排查看板","69项异常；处理顺序：业务事实和源单 → 主数据/接口 → 子模块/凭证 → GL/报表",9);
 sh.getRange("A5:E5").values=[["异常类别","异常数量","会阻断月结的判断","第一责任","排查原则"]];header(sh,"A5:E5");const cats=Object.keys(map);sh.getRange("A6:A14").values=cats.map(c=>[`${c} ${map[c]}`]);sh.getRange("B6:B14").formulas=cats.map((c,i)=>[`=COUNTIF('01-69项异常总表'!B$6:B$74,LEFT(A${i+6},FIND(\" \" ,A${i+6})-1))`]);sh.getRange("C6:C14").values=cats.map(c=>[c==="AUTH"?"视权限事件是否影响数据和支付":"影响本期完整性、成本、银行、税务或关键对账时阻断"]);sh.getRange("D6:D14").values=cats.map(c=>[c==="IF"?"IT与模块会计":c==="AUTH"?"系统负责人和财务主管":"业务负责人和模块会计"]);sh.getRange("E6:E14").values=cats.map(c=>["先确定业务事实和来源单据，不用无依据GL凭证调平"]);body(sh,"A6:E14");widths(sh,{A:25,B:13,C:42,D:28,E:48,F:3,G:3,H:3,I:3});sh.freezePanes.freezeRows(5);
}
function exceptionRows(list){return list.map(e=>[e.id,e.category,e.symptom,e.cause,e.check,e.owner,e.finance,e.it,e.impact,e.close,e.solution,e.prevent,"未处理","","",null]);}
function fillExceptionSheet(sh,name,sub,list,tableName){title(sh,name,sub,16);const hdr=["异常编号","类别","异常现象","可能原因","检查数据","处理人","财务处理","IT处理","会计影响","是否影响月结","解决方案","预防控制","状态","异常金额","责任人","关闭日期"];sh.getRange("A5:P5").values=[hdr];header(sh,"A5:P5");const rows=exceptionRows(list);sh.getRange(`A6:P${5+rows.length}`).values=rows;body(sh,`A6:P${5+rows.length}`,7.2);sh.getRange(`M6:P${5+rows.length}`).format.fill=amber;sh.getRange(`M6:M${5+rows.length}`).dataValidation={rule:{type:"list",values:["未处理","处理中","待复核","已关闭","不适用"]}};sh.getRange(`M6:M${5+rows.length}`).conditionalFormats.add("containsText",{text:"已关闭",format:{fill:green,font:{color:"#548235",bold:true}}});sh.getRange(`M6:M${5+rows.length}`).conditionalFormats.add("containsText",{text:"未处理",format:{fill:red,font:{color:"#C00000"}}});sh.getRange(`N6:N${5+rows.length}`).format.numberFormat="#,##0.00;[Red](#,##0.00);-";sh.getRange(`P6:P${5+rows.length}`).format.numberFormat="yyyy-mm-dd";widths(sh,{A:13,B:10,C:28,D:36,E:42,F:28,G:46,H:44,I:38,J:38,K:45,L:38,M:11,N:14,O:16,P:13});sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(3);sh.tables.add(`A5:P${5+rows.length}`,true,tableName).style="TableStyleMedium2";}
fillExceptionSheet(w2.worksheets.getItem(enames[1]),"69项ERP常见异常总表","可按模块、状态、月结影响和责任人筛选；IT修复后必须由财务验证",data.exceptions,"AllERPExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[2]),"采购与应付异常","供应商、PO、收货、发票、应付、付款和核销",data.exceptions.filter(e=>e.category==="P2P"),"P2PExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[3]),"销售与应收异常","信用、交货、出库、收入、开票、收款和核销",data.exceptions.filter(e=>e.category==="O2C"),"O2CExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[4]),"库存生产和成本异常","数量、价值、BOM、工单、报工、在制和成本结算",data.exceptions.filter(e=>e.category==="COST"),"CostExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[5]),"费用资产和薪酬异常","费控、员工往来、资产卡片、折旧和工资接口",data.exceptions.filter(e=>e.category==="EFP"),"EFPExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[6]),"银行与税务异常","银企状态、流水、调节表、进销项和应交税费",data.exceptions.filter(e=>e.category==="BT"),"BTExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[7]),"总账月结和报表异常","凭证、期间、控制科目、报表映射和结转",data.exceptions.filter(e=>e.category==="GL"),"GLExceptions");
fillExceptionSheet(w2.worksheets.getItem(enames[8]),"接口 主数据和权限异常","基础数据、接口批次、系统任务和职责分离",data.exceptions.filter(e=>["IF","MD","AUTH"].includes(e.category)),"IMAPExceptions");
{
 const sh=w2.worksheets.getItem(enames[9]);title(sh,"ERP问题处理记录","实际问题按唯一编号登记；关闭前保留源单、日志、调整和复核证据",18);const hdr=["问题编号","发现日期","异常编号","系统模块","业务单据号","凭证号","现象","影响金额","影响期间","是否阻断月结","业务原因","系统原因","临时控制","最终方案","财务处理人","IT处理人","复核人","状态"];
 const rows=Array.from({length:60},(_,i)=>[`INC-${String(i+1).padStart(4,"0")}`,null,"","","","","",0,"","待判断","","","","","","","","未处理"]);sh.getRange("A5:R5").values=[hdr];header(sh,"A5:R5");sh.getRange("A6:R65").values=rows;body(sh,"A6:R65",7.8);sh.getRange("B6:B65").format.numberFormat="yyyy-mm-dd";sh.getRange("H6:H65").format.numberFormat="#,##0.00;[Red](#,##0.00);-";sh.getRange("J6:J65").dataValidation={rule:{type:"list",values:["是","否","待判断"]}};sh.getRange("R6:R65").dataValidation={rule:{type:"list",values:["未处理","处理中","待复核","已关闭"]}};widths(sh,{A:13,B:13,C:13,D:18,E:18,F:18,G:32,H:14,I:14,J:14,K:30,L:30,M:30,N:36,O:15,P:15,Q:15,R:11});sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(3);sh.tables.add("A5:R65",true,"IncidentLog").style="TableStyleMedium2";
}
w2.recalculate();
const inspect2=await w2.inspect({kind:"table",range:"00-异常看板!A1:E14",include:"values,formulas",tableMaxRows:20,tableMaxCols:8});await fs.writeFile(`${qaDir}/exception_inspect.ndjson`,inspect2.ndjson??String(inspect2));
const err2=await w2.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:300},summary:"exception formula scan"});await fs.writeFile(`${qaDir}/exception_errors.ndjson`,err2.ndjson??String(err2));
for(const n of enames){const b=await w2.render({sheetName:n,autoCrop:"all",scale:0.75,format:"png"});await fs.writeFile(`${qaDir}/exception_${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f2=await SpreadsheetFile.exportXlsx(w2);await f2.save(`${outDir}/财务ERP常见异常排查表.xlsx`);
console.log(JSON.stringify({training:`${outDir}/ERP财务实操训练清单.xlsx`,exceptions:`${outDir}/财务ERP常见异常排查表.xlsx`,trainingSheets:names.length,exceptionSheets:enames.length,cases:data.cases.length,exceptionCount:data.exceptions.length},null,2));
