import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const base = "C:/Users/HZJ/Desktop/Git/work/finance_risk_20260908";
const data = JSON.parse(await fs.readFile(path.join(base, "content.json"), "utf8"));
const outDir = path.join(base, "deliverables");
const wb = Workbook.create();
const navy = "#17365D", blue = "#DCE6F1", line = "#D9DEE7", text = "#20242C";
const fills = { 红: "#FCE8E8", 橙: "#FCEFD9", 黄: "#FFF7CC", 蓝: "#E6F0FA" };
const fontColors = { 红: "#A61B29", 橙: "#9A4A00", 黄: "#6B5900", 蓝: "#1F4E79" };

function title(sheet, title, subtitle, endCol) {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format.font = { name: "Microsoft YaHei", size: 16, bold: true, color: "#FFFFFF" };
  sheet.getRange(`A1:${endCol}1`).format.fill = navy;
  sheet.getRange(`A1:${endCol}1`).format.rowHeight = 30;
  sheet.getRange(`A2:${endCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format.font = { name: "Microsoft YaHei", size: 10, color: "#566171" };
  sheet.getRange(`A2:${endCol}2`).format.rowHeight = 34;
  sheet.getRange(`A2:${endCol}2`).format.wrapText = true;
}

function header(sheet, range) {
  range.format.fill = navy;
  range.format.font = { name: "Microsoft YaHei", size: 10, bold: true, color: "#FFFFFF" };
  range.format.wrapText = true;
  range.format.verticalAlignment = "center";
  range.format.horizontalAlignment = "center";
  range.format.rowHeight = 28;
}

function body(sheet, range) {
  range.format.font = { name: "Microsoft YaHei", size: 10, color: text };
  range.format.verticalAlignment = "top";
  range.format.wrapText = true;
  range.format.borders = { preset: "all", style: "thin", color: line };
}

function addTable(sheet, address, name) {
  const t = sheet.tables.add(address, true, name);
  t.showFilterButton = true;
  return t;
}

// 01 总览框架
{
  const s = wb.worksheets.add("01总览框架"); s.tabColor = navy;
  title(s, "会计职业风险｜大框架", "颜色表示建议行动速度，不是司法等级或犯罪概率。法律核验截至2026-09-08；适用具体事项时仍需复核行为发生日与最新规定。", "F");
  s.getRange("A3:F3").merge();
  s.getRange("A3").values = [["重点红线：虚开增值税发票。专用发票见R05，普通发票及其他发票见R06；法律边界见手册第5.3—5.4节，官方案例见C01—C03。"]];
  s.getRange("A3:F3").format.fill = "#FCE8E8";
  s.getRange("A3").format.font = { name: "Microsoft YaHei", size: 10, bold: true, color: "#A61B29" };
  s.getRange("A3:F3").format.rowHeight = 28;
  s.getRange("A4:F4").values = [["行动级", "如何理解", "典型动作", "目录场景数", "是否等于犯罪", "使用提示"]]; header(s, s.getRange("A4:F4"));
  const lv = [
    ["红", "明知造假、侵占挪用、销毁或协助违法", "暂停涉疑动作；保全、报告、专业介入", data.risks.filter(x=>x.level==="红").length, "否；仍按具体要件认定", "不能用金额小或老板指令作为安全线"],
    ["橙", "真实性、授权、资料或身份有重大疑点", "先核实再处理；确认故意违法则升级", data.risks.filter(x=>x.level==="橙").length, "否", "避免在事实未清时签字、付款、申报"],
    ["黄", "常规差错或程序遗漏", "及时更正、补报、复核与修复控制", data.risks.filter(x=>x.level==="黄").length, "通常不是", "仍需区分重大前期差错和行政后果"],
    ["蓝", "有依据的专业判断或能力提升", "形成底稿、取得复核、持续训练", data.risks.filter(x=>x.level==="蓝").length, "否", "发现实质错误后转差错路径"]
  ];
  s.getRange("A5:F8").values = lv; body(s,s.getRange("A5:F8"));
  for (let i=0;i<4;i++){s.getRange(`A${5+i}`).format.fill=fills[lv[i][0]];s.getRange(`A${5+i}`).format.font={name:"Microsoft YaHei",size:10,bold:true,color:fontColors[lv[i][0]]};}
  s.getRange("A10:F10").values = [["序号", "大框架", "核心分支", "对应手册", "最重要的工作结论", "知识库关系"]]; header(s,s.getRange("A10:F10"));
  s.getRange(`A11:F${10+data.framework.length}`).values = data.framework.map(x=>[x[0],x[1],x[2],x[3],x[4],"M12-001；关联M01/M02/M03/M06/M08/M09/M10/M11"]); body(s,s.getRange(`A11:F${10+data.framework.length}`));
  addTable(s,`A10:F${10+data.framework.length}`,"OverviewFramework");
  s.freezePanes.freezeRows(4);
  [8,20,38,18,48,30].forEach((w,i)=>s.getRangeByIndexes(0,i,20,1).format.columnWidth=w);
  s.getRange("A5:F18").format.rowHeight=42;
}

// 02 风险明细
{
  const s=wb.worksheets.add("02风险明细"); s.tabColor="#A61B29";
  title(s,"26个会计职业风险场景｜详细框架","可按流程和行动级筛选。每一行沿着“业务发生→单据→系统→判断→法源→补救→对账月结→影响”阅读；完整解释见HTML手册第4章。","N");
  const heads=["编号","流程","行动级","风险场景","易违规触发","原始单据","系统操作与控制","会计及法律判断","法源编号","事前预防","发现后补救","对账与月结","报表及责任影响","案例索引"];
  s.getRange("A4:N4").values=[heads];header(s,s.getRange("A4:N4"));
  const rows=data.risks.map(r=>[r.id,r.group,r.level,r.title,r.trigger,r.docs,r.system,r.judgment,r.law,r.prevent,r.remedy,r.close,r.impact,r.cases]);
  s.getRange(`A5:N${4+rows.length}`).values=rows;body(s,s.getRange(`A5:N${4+rows.length}`));
  addTable(s,`A4:N${4+rows.length}`,"RiskDetailTable");
  for(let i=0;i<rows.length;i++){
    const lev=rows[i][2], row=5+i;
    s.getRange(`C${row}`).format.fill=fills[lev];s.getRange(`C${row}`).format.font={name:"Microsoft YaHei",size:10,bold:true,color:fontColors[lev]};
    if(rows[i][0]==="R05"||rows[i][0]==="R06"){
      s.getRange(`D${row}`).format.fill="#FCE8E8";
      s.getRange(`D${row}`).format.font={name:"Microsoft YaHei",size:10,bold:true,color:"#A61B29"};
    }
    s.getRange(`A${row}:N${row}`).format.rowHeight=76;
  }
  s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(4);
  [9,16,8,24,34,32,36,43,31,38,42,37,43,14].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);
}

// 03 行动清单
{
  const s=wb.worksheets.add("03行动清单");s.tabColor="#548235";
  title(s,"入职、月结、应急、离职行动清单","前6列为知识库建议，后5列留给你按公司实际填写。建议时间不替代税务、诉讼或监管文书中的法定期限。","K");
  const heads=["阶段","控制或处置事项","建议责任人","应形成的证据","建议时间","检查重点","实际负责人","计划完成日","状态","证据位置","备注"];
  s.getRange("A4:K4").values=[heads];header(s,s.getRange("A4:K4"));
  const rows=data.actions.map(a=>[a.stage,a.task,a.owner,a.evidence,a.timing,a.check,"","","未开始","",""]);
  s.getRange(`A5:K${4+rows.length}`).values=rows;body(s,s.getRange(`A5:K${4+rows.length}`));
  addTable(s,`A4:K${4+rows.length}`,"ActionChecklistTable");
  s.getRange(`I5:I${4+rows.length}`).dataValidation={rule:{type:"list",values:["未开始","进行中","已完成","阻塞","不适用"]}};
  s.getRange(`H5:H${4+rows.length}`).setNumberFormat("yyyy-mm-dd");
  s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(2);
  [11,35,23,30,20,40,18,14,12,28,30].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);
  s.getRange(`A5:K${4+rows.length}`).format.rowHeight=58;
}

// 04 官方案例
{
  const s=wb.worksheets.add("04官方案例");s.tabColor="#8064A2";
  title(s,"9个官方案例｜事实、结果、启示与边界","案例用于识别事实和控制，不用于预测另一个案件的定罪或量刑。公开报道不是完整案卷。","H");
  const heads=["编号","案例","公开时间／程序","关键事实","公开结果","可学的控制点","不能照搬的边界","官方链接"];
  s.getRange("A4:H4").values=[heads];header(s,s.getRange("A4:H4"));
  const rows=data.cases.map(c=>[c.id,c.title,c.date,c.facts,c.result,c.lesson,c.boundary,c.url]);
  s.getRange(`A5:H${4+rows.length}`).values=rows;body(s,s.getRange(`A5:H${4+rows.length}`));addTable(s,`A4:H${4+rows.length}`,"OfficialCasesTable");
  s.freezePanes.freezeRows(4);[9,28,22,43,36,40,46,48].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);s.getRange(`A5:H${4+rows.length}`).format.rowHeight=86;
}

// 05 法源索引
{
  const s=wb.worksheets.add("05法源索引");s.tabColor="#5B9BD5";
  title(s,"官方法源与版本索引","优先财政部、国家税务总局、全国人大、最高法等官方来源；部分为政府机关官网转载。使用具体条款时核验行为发生日与后续修订。","F");
  const heads=["编号","官方法源","版本／施行信息","本手册使用范围","核验日期","官方链接"];
  s.getRange("A4:F4").values=[heads];header(s,s.getRange("A4:F4"));
  const rows=data.sources.map(x=>[x.id,x.title,x.date,x.scope,data.date,x.url]);
  s.getRange(`A5:F${4+rows.length}`).values=rows;body(s,s.getRange(`A5:F${4+rows.length}`));addTable(s,`A4:F${4+rows.length}`,"OfficialSourcesTable");
  s.freezePanes.freezeRows(4);[9,35,25,55,13,60].forEach((w,i)=>s.getRangeByIndexes(0,i,rows.length+4,1).format.columnWidth=w);s.getRange(`A5:F${4+rows.length}`).format.rowHeight=64;
}

wb.recalculate();
await fs.mkdir(outDir,{recursive:true});
const inspection=await wb.inspect({kind:"workbook,sheet,table",maxChars:8000,tableMaxRows:4,tableMaxCols:6,tableMaxCellChars:100});
await fs.writeFile(path.join(base,"workbook_inspect.txt"),inspection.ndjson??String(inspection),"utf8");
const previews=[
  ["01总览框架","A1:F18","preview-01-overview.png"],
  ["02风险明细","A1:N12","preview-02-risks.png"],
  ["03行动清单","A1:K13","preview-03-actions.png"],
  ["04官方案例","A1:H13","preview-04-cases.png"],
  ["05法源索引","A1:F13","preview-05-sources.png"]
];
for(const [sheetName,range,file] of previews){const blob=await wb.render({sheetName,range,scale:1,format:"png"});await fs.writeFile(path.join(base,file),new Uint8Array(await blob.arrayBuffer()));}
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path.join(outDir,"02-风险框架与行动清单-v1.1.xlsx"));
console.log(JSON.stringify({sheets:previews.map(x=>x[0]),risks:data.risks.length,cases:data.cases.length,sources:data.sources.length,actions:data.actions.length}));
