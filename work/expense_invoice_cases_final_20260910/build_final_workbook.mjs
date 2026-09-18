import { Workbook, SpreadsheetFile } from "file:///C:/Users/HZJ/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
import fs from "fs/promises"; import { fileURLToPath } from "url";
const root=new URL(".",import.meta.url);const data=JSON.parse(await fs.readFile(new URL("./cases.json",root),"utf8"));
await fs.mkdir(new URL("./deliverables/",root),{recursive:true});await fs.mkdir(new URL("./qa/",root),{recursive:true});
const outFile=fileURLToPath(new URL("./deliverables/异常发票与费用处理速查表.xlsx",root));const wb=Workbook.create();
const names=["00-使用说明","01-141案例速查","02-老板与特殊物品","03-员工报销","04-员工福利","05-招待礼品推广","06-发票与无票","07-主体与资产","08-红线与法律后果","09-异常处理登记","10-老板指令SOP","11-沟通模板","12-法规索引"];
const sh={};for(const n of names)sh[n]=wb.worksheets.add(n);
const NAVY="#17365D",BLUE="#2F75B5",PALE="#EAF2F8",GRAY="#F2F4F7",RED="#FCE4D6",GREEN="#E2F0D9",AMBER="#FFF2CC",WHITE="#FFFFFF",font="Microsoft YaHei";
function col(n){let s="";while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function title(s,t,sub,cols){s.mergeCells(`A1:${col(cols)}1`);s.getRange("A1").values=[[t]];s.getRange("A1").format={fill:NAVY,font:{name:font,size:18,bold:true,color:WHITE},rowHeight:30};s.mergeCells(`A2:${col(cols)}2`);s.getRange("A2").values=[[sub]];s.getRange("A2").format={fill:PALE,font:{name:font,size:9,italic:true,color:"#555555"},rowHeight:24};}
function header(s,r){s.getRange(r).format={fill:NAVY,font:{name:font,bold:true,color:WHITE},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,rowHeight:34,borders:{top:{style:"thin",color:"#9EADBA"},bottom:{style:"thin",color:"#9EADBA"},left:{style:"thin",color:"#9EADBA"},right:{style:"thin",color:"#9EADBA"}}};}
function body(s,r){s.getRange(r).format={font:{name:font,size:9},verticalAlignment:"center",wrapText:true,borders:{top:{style:"hair",color:"#D9E1E8"},bottom:{style:"hair",color:"#D9E1E8"},left:{style:"hair",color:"#D9E1E8"},right:{style:"hair",color:"#D9E1E8"}}};}
function widths(s,obj){for(const [c,w] of Object.entries(obj))s.getRange(`${c}:${c}`).format.columnWidth=w}
const lawText=c=>c.split(/[、；]/).map(k=>{const x=data.sources[k];return x?`${k} ${x[0]} ${x[1]}`:k}).join("；");
const account=c=>c.entries.split("。")[0].replace("建议科目：","");
const ynVat=c=>c.vat.includes("不得抵扣")?"否/需分拆":c.vat.includes("可")?"符合条件可":"逐案判断";
const ynCit=c=>c.cit.includes("不得扣")?"否/需调增":c.cit.includes("可扣")?"符合条件可":"逐案判断";
const ynIit=c=>c.iit.includes("不构成")||c.iit.includes("不涉及")?"通常否":c.iit.includes("可能")||c.iit.includes("工资")||c.iit.includes("红利")?"可能/需判断":"逐案判断";
const h=["案例编号","类别","业务场景","关键词","会计科目","是否可入账","是否可抵扣增值税","是否可税前扣除","是否涉及个税","风险等级","需要资料","处理方式","法规依据","沟通建议","拒不整改后果","可能责任主体","财务升级/留痕"];
const rows=data.cases.map(c=>[c.id,c.category,c.title,c.tags,account(c),c.status,ynVat(c),ynCit(c),ynIit(c),c.risk,c.docs,c.operation,lawText(c.laws),c.comm,c.consequences,c.responsible,"OA/邮件记录事实、判断、替代方案、审批与关闭证据；橙色/红色提交财务负责人、法务或治理层"]);
function caseSheet(s,t,subtitle,filter){const rr=rows.filter(filter);title(s,t,subtitle,17);s.getRange("A5:Q5").values=[h];header(s,"A5:Q5");s.getRange(`A6:Q${5+rr.length}`).values=rr;body(s,`A6:Q${5+rr.length}`);s.getRange(`A6:Q${5+rr.length}`).format.rowHeight=74;for(let r=6;r<=5+rr.length;r++){const risk=rr[r-6][9];if(risk.startsWith("红"))s.getRange(`A${r}:Q${r}`).format.fill=RED;else if(risk.startsWith("橙"))s.getRange(`A${r}:Q${r}`).format.fill=AMBER;else if(risk.startsWith("蓝"))s.getRange(`A${r}:Q${r}`).format.fill=GREEN;else if(r%2===0)s.getRange(`A${r}:Q${r}`).format.fill="#F7FAFC"}widths(s,{A:13,B:18,C:32,D:18,E:24,F:28,G:17,H:17,I:16,J:17,K:42,L:48,M:44,N:48,O:64,P:48,Q:48});s.freezePanes.freezeRows(5);s.freezePanes.freezeColumns(3);s.tables.add(`A5:Q${5+rr.length}`,true,`T${s.name.replace(/[^0-9A-Za-z\u4e00-\u9fa5]/g,"").slice(0,15)}`).style="TableStyleMedium2";}
{
 const s=sh[names[0]];title(s,"异常发票与费用处理速查表","141个案例｜政策核验日：2026-09-10｜先筛选，再回Word阅读完整判断",8);
 s.getRange("A5:H5").values=[["使用步骤","问题","输出","颜色","财务动作","拒不整改后果","责任判断","提醒"]];header(s,"A5:H5");s.getRange("A6:H12").values=[
 ["1 还原事实","真实业务是什么","业务事实卡","蓝/黄","取得单据","退回或补件","申请人与审批人","发票真实≠交易真实"],["2 查受益人","谁最终获益","受益人结论","黄/橙","分拆公私","追回/个税","老板/员工/股东","公司支付≠公司费用"],["3 穿透主体","合同履约票款是否匹配","主体图","黄/橙","补代付委托","暂停付款抵扣","各交易主体","不能只看票款一致"],["4 税会判断","会计/VAT/EIT/IIT","四套结论","各色","入账+税务台账","调账补税滞纳金","财务经办复核","四套口径分开"],["5 风险升级","是否虚构或强令造假","升级记录","红","停止执行","行政/职业/刑事","单位负责人和直接责任人员","刑事须满足构成要件"],["6 关闭","是否完成整改","关闭证据","绿","复核归档","不得虚假关闭","复核人","季度更新法规"],["检索","按编号/关键词/风险筛选","对应案例","—","回Word看全文","—","—","XMind看关系"]];body(s,"A6:H12");widths(s,{A:18,B:30,C:24,D:14,E:28,F:40,G:34,H:34});s.freezePanes.freezeRows(5);
}
caseSheet(sh[names[1]],"141案例总速查","必备字段+拒不整改后果+责任主体",()=>true);
caseSheet(sh[names[2]],"老板/股东与特殊物品","公司支付不等于公司费用；奇怪物品按用途和受益人判断",r=>["老板/股东私人支出","公司购买特殊物品"].includes(r[1]));
caseSheet(sh[names[3]],"员工费用报销","时间地点金额异常是调查入口，不是自动违法结论",r=>r[1]==="员工费用报销");
caseSheet(sh[names[4]],"员工福利","联动应付职工薪酬、福利费限额、进项限制和个税",r=>r[1]==="员工福利");
caseSheet(sh[names[5]],"客户招待、礼品与市场推广","区分招待、广告宣传、折扣返利、样品赠送",r=>r[1]==="客户招待、礼品与市场推广");
caseSheet(sh[names[6]],"发票异常与真实无票支出","会计入账、增值税抵扣和所得税扣除分别判断",r=>["真实业务但无发票","发票异常"].includes(r[1]));
caseSheet(sh[names[7]],"主体不一致与资产费用边界","穿透六主体；按控制、受益期和可使用状态判断资本化",r=>["主体不一致与代收代付","资产与费用边界"].includes(r[1]));
caseSheet(sh[names[8]],"高风险红线与法律后果","红线必须停止执行；行政处罚、税收处罚和刑事责任分别判断",r=>r[1]==="高风险红线");
{
 const s=sh[names[9]];title(s,"异常处理登记","发现异常→保全资料→量化→调账/更正→缴税→复核→关闭",18);const hh=["编号","发现日","案例编号","申请人","实际受益人","异常事实","风险等级","涉及金额","是否暂停付款","是否暂停抵扣","会计调整","更正申报","补税","滞纳金","行政/法律意见","责任人","复核人","状态"];s.getRange("A5:R5").values=[hh];header(s,"A5:R5");s.getRange("A6:R35").values=Array(30).fill(Array(18).fill(null));body(s,"A6:R35");s.getRange("B6:B35").setNumberFormat("yyyy-mm-dd");s.getRange("H6:H35").setNumberFormat("#,##0.00");for(const c of ["I","J","L"])s.getRange(`${c}6:${c}35`).dataValidation={rule:{type:"list",values:["是","否","待判断"]}};s.getRange("R6:R35").dataValidation={rule:{type:"list",values:["未开始","调查中","待整改","待复核","已关闭","升级法务"]}};widths(s,{A:13,B:13,C:13,D:16,E:18,F:40,G:16,H:14,I:15,J:15,K:32,L:16,M:14,N:14,O:38,P:14,Q:14,R:14});s.freezePanes.freezeRows(5);s.tables.add("A5:R35",true,"AbnormalRegister").style="TableStyleMedium2";
}
{
 const s=sh[names[10]];title(s,"财务面对不合理老板指令SOP","明确事实→法规→替代方案→书面留痕→升级→红线拒绝",7);s.getRange("A5:G5").values=[["步骤","要做什么","输出","可继续条件","停止条件","升级对象","禁止动作"]];header(s,"A5:G5");s.getRange("A6:G13").values=[["1","还原业务和最终受益人","事实卡","事实清楚","疑似虚构","业务负责人","猜测事实"],["2","穿透合同履约票款","主体图","法律关系成立","无业务/资金回流","财务主管","只看发票"],["3","分别判断会计与三税","判断备忘","存在合法处理","明显违法","税务负责人","倒挤结论"],["4","引用有效法规","法源清单","官方依据明确","政策不确定","外部税务顾问","引用失效文件"],["5","提出合法替代方案","选项清单","管理层选择合法方案","坚持造假","财务负责人/CFO","迎合改名"],["6","OA/邮件书面留痕","审批记录","授权完整","口头强令","治理层/法务","补造证据"],["7","红线停止付款制单抵扣","暂停工单","独立调查后解除","继续虚构","审计委员会/股东会","删除记录"],["8","整改、复核、归档","关闭证据","复核人签署","假关闭","CFO/法务","覆盖历史"]];body(s,"A6:G13");widths(s,{A:9,B:34,C:24,D:30,E:28,F:26,G:30});s.freezePanes.freezeRows(5);
}
{
 const scripts=["这笔款公司可以先支付，但真实受益人是个人，账务上需要挂个人往来并约定归还，不能直接计入公司费用。","发票验真只能证明票据存在，尚不能证明公司实际取得了货物或服务；请补合同、验收和付款资料。","会计上应如实反映这笔支出，企业所得税是否扣除需要另做判断，我会同步登记纳税调整。","这项用途依法不得抵扣进项税，发票可以入账，但税额需要计入成本费用。","当前合同、履约、发票和付款主体不一致，请先说明法律关系并补代付/委托资料。","这张票已在系统中出现过，需先确认是否重复报销；核实前暂不付款。","周末或异地发票不是自动违规，但需要行程、对象和业务目的形成闭环。","私人部分请从本次申请中剔除，公司只承担能证明与公务相关的部分。","如果公司最终承担这项个人利益，需要按工资薪金或股息红利评估个税，不能只改费用科目。","供应商不能及时开票不等于业务不能入账；我们先按真实义务确认，同时跟踪税前扣除凭证。","这项支出达到资产确认条件，应登记资产并按受益期折旧/摊销，不能为了当期利润任意费用化。","请让开票方按真实品名和金额红冲重开，财务不能在发票上手工修改。","这笔招待费可以入账，但所得税扣除受双限额约束，需有对象、目的和审批资料。","员工福利需要名单、制度和签收，同时判断福利费限额、进项抵扣限制与个税。","母公司可以代付款，但费用、进项和资产应归实际购买并使用的主体，双方同步挂往来。","目前资料不足以支持抵扣和付款，我先把事项登记为异常，不会把事实不清的业务硬做进账。","如果继续按虚假业务处理，风险会从内部退单升级为补税、滞纳金、罚款，达到法定条件还可能涉及刑事责任。","我不能制作虚假合同、验收或会议记录；可以协助按真实业务寻找合法的账务和税务处理。","请把指令和业务事实写入OA，由有权负责人复核；财务会同时保留专业判断和替代方案。","该事项已经触及红线，我将停止制单、付款或抵扣，并按公司授权矩阵提交财务负责人和法务。"];
 const s=sh[names[11]];title(s,"20条专业沟通模板","专业、克制、提供合法替代方案并保留判断",4);s.getRange("A5:D5").values=[["序号","适用场景","可直接使用的表达","后续动作"]];header(s,"A5:D5");s.getRange("A6:D25").values=scripts.map((x,i)=>[i+1,i<5?"事实/主体":i<10?"税会/个税":i<15?"发票/福利/资产":"升级/红线",x,"记录回复、责任人和关闭日期"]);body(s,"A6:D25");widths(s,{A:10,B:18,C:90,D:34});s.freezePanes.freezeRows(5);s.tables.add("A5:D25",true,"CommunicationScripts").style="TableStyleMedium2";
}
{
 const s=sh[names[12]];title(s,"法规与官方资料索引","法规名称+条款/文件号+核心要求+官方链接",6);s.getRange("A5:F5").values=[["代码","法规名称","条款/文件号","核心要求","官方链接","核验日"]];header(s,"A5:F5");const rr=Object.entries(data.sources).map(([k,x])=>[k,x[0],x[1],x[2],x[3],new Date("2026-09-10")]);s.getRange(`A6:F${5+rr.length}`).values=rr;body(s,`A6:F${5+rr.length}`);s.getRange(`F6:F${5+rr.length}`).setNumberFormat("yyyy-mm-dd");widths(s,{A:10,B:48,C:34,D:72,E:70,F:14});s.freezePanes.freezeRows(5);s.tables.add(`A5:F${5+rr.length}`,true,"LawIndex").style="TableStyleMedium2";
}
for(const s of Object.values(sh)){const u=s.getUsedRange();u.format.font.name=font;u.format.verticalAlignment="center"}
wb.recalculate();const err=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:100},summary:"formula scan"});console.log(err.ndjson);
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(outFile);
for(const n of ["00-使用说明","01-141案例速查","08-红线与法律后果","09-异常处理登记","12-法规索引"]){const img=await wb.render({sheetName:n,range:"A1:Q18",scale:1,format:"png"});await fs.writeFile(new URL(`./qa/${n}.png`,root),new Uint8Array(await img.arrayBuffer()))}
console.log(outFile);
