import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL("./deliverables/", import.meta.url).pathname.slice(1);
await fs.mkdir(outputDir, { recursive: true });
const outFile = `${outputDir}/企业税务月结与申报检查清单.xlsx`;
const wb = Workbook.create();
const names = ["00-概览与参数","01-全年税务日历","02-月度税务清单","03-季度税务清单","04-年度税务清单","05-增值税账票税","06-企业所得税","07-发票异常","08-未到票暂估","09-税务异常整改","10-数据字典"];
const sheets = Object.fromEntries(names.map(n => [n, wb.worksheets.add(n)]));
const navy="#17365D", blue="#2F75B5", pale="#EAF2F8", gray="#F2F4F7", line="#D9E1E8", amber="#FFF2CC", red="#FCE4D6", green="#E2F0D9", font="Arial";

function col(n){let s=""; while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function title(sh, text, subtitle, endCol=11){
  sh.showGridLines=false;
  sh.getRange(`A1:${col(endCol)}1`).merge(); sh.getRange("A1").values=[[text]];
  sh.getRange("A1").format={font:{name:font,size:15,bold:true,color:"#000000"},rowHeight:28,verticalAlignment:"center"};
  sh.getRange(`A2:${col(endCol)}2`).merge(); sh.getRange("A2").values=[[subtitle]];
  sh.getRange("A2").format={font:{name:font,size:9,italic:true,color:"#666666"},rowHeight:22};
  sh.getRange(`A3:${col(endCol)}3`).format.borders={bottom:{style:"thin",color:blue}};
}
function header(sh, addr){
  const r=sh.getRange(addr); r.format={fill:navy,font:{name:font,size:9,bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,rowHeight:30,borders:{preset:"all",style:"thin",color:"#FFFFFF"}};
}
function body(sh, addr, size=9){
  const r=sh.getRange(addr); r.format={font:{name:font,size},verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:line}};
}
function widths(sh, map){for(const [c,w] of Object.entries(map)) sh.getRange(`${c}:${c}`).format.columnWidth=w}
function statusValidation(r){r.dataValidation={rule:{type:"list",values:["未开始","进行中","待复核","已完成","不适用"]}}}
function statusCF(r){
  r.conditionalFormats.add("containsText",{text:"已完成",format:{fill:green,font:{color:"#548235",bold:true}}});
  r.conditionalFormats.add("containsText",{text:"未开始",format:{fill:amber,font:{color:"#9C6500"}}});
  r.conditionalFormats.add("containsText",{text:"进行中",format:{fill:pale,font:{color:blue}}});
}
function commonChecklist(sh, titleText, subtitle, rows){
  title(sh,titleText,subtitle,11);
  const hdr=["事项","责任人","完成期限","数据来源","核对对象","是否完成","异常金额 万元","异常原因","处理方案","复核人","关闭日期"];
  sh.getRange("A5:K5").values=[hdr]; header(sh,"A5:K5");
  sh.getRange(`A6:K${5+rows.length}`).values=rows; body(sh,`A6:K${5+rows.length}`);
  statusValidation(sh.getRange(`F6:F${5+rows.length}`)); statusCF(sh.getRange(`F6:F${5+rows.length}`));
  sh.getRange(`C6:C${5+rows.length}`).setNumberFormat("yyyy-mm-dd"); sh.getRange(`K6:K${5+rows.length}`).setNumberFormat("yyyy-mm-dd");
  sh.getRange(`G6:G${5+rows.length}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  widths(sh,{A:24,B:12,C:13,D:22,E:22,F:11,G:15,H:24,I:26,J:12,K:13}); sh.freezePanes.freezeRows(5); sh.freezePanes.freezeColumns(1);
  sh.tables.add(`A5:K${5+rows.length}`,true,`${sh.name.replaceAll("-","").replaceAll("与","")}Table`).style="TableStyleMedium2";
}

// 00 概览与参数
{
 const sh=sheets[names[0]]; title(sh,"企业税务月结与申报检查清单","远澜智能设备（苏州）有限公司　案例期间 2026年6月　政策核验日 2026年9月9日",12);
 sh.getRange("A5:D5").values=[["关键参数","数值","适用说明","维护责任"]]; header(sh,"A5:D5");
 sh.getRange("A6:D15").values=[
  ["政策核验日",new Date("2026-09-09"),"使用前检查是否有新政策","税务会计"],
  ["案例所属期",new Date("2026-06-30"),"用于本工作簿示例","总账会计"],
  ["一般货物增值税率",0.13,"按具体货物和交易复核","税务会计"],
  ["一般服务增值税率",0.06,"咨询等案例参数","税务会计"],
  ["企业所得税率",0.25,"案例公司不适用特殊低税率","税务会计"],
  ["城市维护建设税率",0.07,"案例位于市区，实际按所在地核验","税务会计"],
  ["教育费附加率",0.03,"案例参数","税务会计"],
  ["地方教育附加率",0.02,"江苏案例参数，实际核验","税务会计"],
  ["差异复核阈值 万元",1,"超过阈值必须书面解释","财务主管"],
  ["内部申报复核提前天数",3,"早于法定截止日完成","财务主管"]
 ]; body(sh,"A6:D15"); sh.getRange("B6:B7").setNumberFormat("yyyy-mm-dd"); sh.getRange("B8:B14").setNumberFormat("0.00%"); sh.getRange("B14:B15").setNumberFormat("0.00");
 sh.getRange("F5:I5").values=[["当前工作状态","公式结果","说明","负责人"]]; header(sh,"F5:I5");
 sh.getRange("F6:F11").values=[["月度清单未完成"],["季度清单未完成"],["年度清单未完成"],["发票异常未关闭"],["暂估未关闭"],["税务整改未关闭"]];
 sh.getRange("G6:G11").formulas=[["=COUNTIF('02-月度税务清单'!F6:F25,\"<>已完成\")-COUNTIF('02-月度税务清单'!F6:F25,\"不适用\")"],["=COUNTIF('03-季度税务清单'!F6:F19,\"<>已完成\")-COUNTIF('03-季度税务清单'!F6:F19,\"不适用\")"],["=COUNTIF('04-年度税务清单'!F6:F21,\"<>已完成\")-COUNTIF('04-年度税务清单'!F6:F21,\"不适用\")"],["=COUNTIF('07-发票异常'!N6:N25,\"<>已关闭\")-COUNTBLANK('07-发票异常'!A6:A25)"],["=COUNTIF('08-未到票暂估'!N6:N25,\"<>已关闭\")-COUNTBLANK('08-未到票暂估'!A6:A25)"],["=COUNTIF('09-税务异常整改'!P6:P25,\"<>已关闭\")-COUNTBLANK('09-税务异常整改'!A6:A25)"]];
 sh.getRange("H6:I11").values=[["0表示全部完成或不适用","总账会计"],["季度末使用","税务会计"],["年度汇算使用","税务经理"],["按异常编号追踪","AP或AR"],["重点关注账龄","AP"],["重大问题主管审批","财务主管"]]; body(sh,"F6:I11"); sh.getRange("G6:G11").format.font={name:font,size:11,bold:true,color:"#000000"};
 sh.getRange("A18:D18").values=[["2026年申报期限","截止日","适用范围","来源"]]; header(sh,"A18:D18");
 const months=[["1月",new Date("2026-01-20")],["2月",new Date("2026-02-24")],["3月",new Date("2026-03-16")],["4月",new Date("2026-04-20")],["5月",new Date("2026-05-22")],["6月",new Date("2026-06-15")],["7月",new Date("2026-07-15")],["8月",new Date("2026-08-17")],["9月",new Date("2026-09-15")],["10月",new Date("2026-10-26")],["11月",new Date("2026-11-16")],["12月",new Date("2026-12-15")]];
 sh.getRange("A19:D30").values=months.map(x=>[x[0],x[1],"实行按月或按季度期满后15日内申报的税种","国家税务总局2026年度申报期限通知"]); body(sh,"A19:D30"); sh.getRange("B19:B30").setNumberFormat("yyyy-mm-dd");
 sh.getRange("F18:I18").values=[["官方来源","用途","有效性说明","链接"]]; header(sh,"F18:I18");
 sh.getRange("F19:I30").values=[
  ["中华人民共和国增值税法","税率、抵扣、义务时间","2026-01-01施行","https://www.npc.gov.cn/npc/c2/c30834/202412/t20241225_442015.html"],
  ["增值税法实施条例","进项分摊与转出","2026-01-01施行","https://fgk.chinatax.gov.cn/zcfgk/c100010/c5246349/content.html"],
  ["税总公告2024年第11号","数电票和红字流程","全文有效","https://fgk.chinatax.gov.cn/zcfgk/c100012/c5236067/content.html"],
  ["发票管理办法","发票基础制度","全文有效","https://fgk.chinatax.gov.cn/zcfgk/c100010/c5195084/content.html"],
  ["发票管理办法实施细则","电子发票与数据","2024年修订","https://fgk.chinatax.gov.cn/zcfgk/c100011/c5221006/content.html"],
  ["企业所得税法","预缴与汇算","现行","https://fgk.chinatax.gov.cn/zcfgk/c100009/c5193018/content.html"],
  ["税前扣除凭证管理办法","扣除凭证","全文有效","https://fgk.chinatax.gov.cn/zcfgk/c100012/c5194804/content.html"],
  ["研发费用公告2023年第7号","100%加计扣除","制度性安排","https://fgk.chinatax.gov.cn/zcfgk/c102416/c5201978/content.html"],
  ["印花税法","合同税源","现行","https://fgk.chinatax.gov.cn/zcfgk/c100009/c5193058/content.html"],
  ["个税扣缴申报管理办法","全员全额申报","全文有效","https://fgk.chinatax.gov.cn/zcfgk/c100012/c5194838/content.html"],
  ["税收征管法","逾期与滞纳金","现行","https://fgk.chinatax.gov.cn/zcfgk/c100009/c5195081/content.html"],
  ["2026年度申报期限通知","年度征期","2026年度","https://fgk.chinatax.gov.cn/zcfgk/c102424/c5245729/content.html"]
 ]; body(sh,"F19:I30",8); widths(sh,{A:25,B:15,C:38,D:18,E:3,F:29,G:25,H:20,I:58}); sh.freezePanes.freezeRows(3); sh.tabColor=navy;
}

// 01 全年日历
{
 const sh=sheets[names[1]]; title(sh,"企业全年税务工作日历","日常、月度、季度和年度任务应同时包含编制、复核、申报、缴款和归档",9);
 sh.getRange("A5:I5").values=[["层次","时间节点","事项","责任人","完成标准","数据来源","复核人","状态","说明"]]; header(sh,"A5:I5");
 const rows=[
  ["日常","业务发生时","合同涉税识别","业务和税务","税目主体时点留痕","合同系统","财务主管","进行中","重大条款签署前复核"],
  ["日常","每日","销项开票和交付","AR","发票与申请一致","销售和开票平台","总账","进行中","异常进入跟踪表"],
  ["日常","每日","进项收票验真查重","AP","发票池状态完整","税务数字账户","总账","进行中","未审核不得直接抵扣"],
  ["周度","每周","清理待开票未到票红字","AR和AP","清单有责任人期限","业务台账","总账","未开始","关注高金额和账龄"],
  ["月度","D-5至D-1","销售截止和税额预测","税务","形成销项预测","销售验收","主管","未开始","包含未开票收入"],
  ["月度","D+1至D+3","暂估和进项池整理","AP总账","暂估进项状态清晰","采购收货发票","主管","未开始","暂估不确认进项"],
  ["月度","D+4至D+5","账票税核对","总账税务","差异已解释","ERP和申报底稿","主管","未开始","保留差异台账"],
  ["月度","征期内","申报缴款和归档","税务资金","回执完税回单齐全","电子税务局银行","主管","未开始","申报成功后检查扣款"],
  ["季度","季末","企业所得税预缴","总账税务","累计数据可复算","利润和调整台账","主管","未开始","扣除前期已缴"],
  ["季度","依核定","印花税等财产行为税","税务","税源与合同资产一致","合同资产台账","主管","未开始","地区差异"],
  ["年度","10月至12月","汇算前资料清理","总账税务","缺失资料有行动计划","全年台账","主管","未开始","提前清理无票和损失"],
  ["年度","次年1月至5月","企业所得税汇算","税务","申报补退税归档","年度财务和税务资料","负责人","未开始","一般5月31日前"],
  ["年度","完成后","税务档案和交接更新","税务","资料可检索可复算","申报档案","主管","未开始","保存版本和审批"]
 ]; sh.getRange(`A6:I${5+rows.length}`).values=rows; body(sh,`A6:I${5+rows.length}`); statusValidation(sh.getRange(`H6:H${5+rows.length}`)); statusCF(sh.getRange(`H6:H${5+rows.length}`)); widths(sh,{A:10,B:16,C:28,D:16,E:28,F:24,G:12,H:11,I:28}); sh.freezePanes.freezeRows(5); sh.tables.add(`A5:I${5+rows.length}`,true,"AnnualTaxCalendar").style="TableStyleMedium2";
}

const monthRows=[
 ["锁定申报税种与截止日","税务会计",new Date("2026-07-01"),"税种核定和年度日历","电子税务局待办","未开始",0,"","核对本月适用税种","财务主管",null],
 ["销售截止和未开票收入","AR",new Date("2026-07-02"),"订单发货签收验收","收入和销项税","未开始",0,"","形成未开票清单","总账",null],
 ["蓝字红字作废发票汇总","AR",new Date("2026-07-02"),"开票平台","销项发票与销售","未开始",0,"","按税率汇总","总账",null],
 ["采购收货和未到票暂估","AP",new Date("2026-07-03"),"采购仓库","AP暂估和存货","未开始",60,"第二批材料未到票","登记并催票","总账",null],
 ["进项发票验真查重","AP",new Date("2026-07-03"),"税务数字账户","进项台账","未开始",0,"","异常票冻结","税务会计",null],
 ["进项用途确认","税务会计",new Date("2026-07-06"),"进项台账","用途确认数据","未开始",0,"","确认可抵扣18.98","财务主管",null],
 ["进项转出检查","总账",new Date("2026-07-06"),"领用和费用明细","进项转出","未开始",2.6,"福利领料","转出并留存领料单","税务会计",null],
 ["销项税测算","税务会计",new Date("2026-07-06"),"销售和开票数据","销项税明细","未开始",0,"","确认46.80","财务主管",null],
 ["应交税费明细核对","总账",new Date("2026-07-07"),"总账和子模块","增值税底稿","未开始",0,"","解释差异","财务主管",null],
 ["增值税申报表编制","税务会计",new Date("2026-07-08"),"账票税底稿","申报表","未开始",0,"","应纳24.82","财务主管",null],
 ["附加税费测算","税务会计",new Date("2026-07-08"),"增值税应纳额","附加申报","未开始",0,"","案例2.9784","财务主管",null],
 ["个税工资五方核对","薪酬总账",new Date("2026-07-08"),"HR工资银行","个税申报","未开始",0,"","代扣3.00","财务主管",null],
 ["主管复核批准","财务主管",new Date("2026-07-10"),"申报批准包","所有申报数据","未开始",0,"","签署复核结论","财务负责人",null],
 ["正式申报","税务会计",new Date("2026-07-13"),"批准申报表","电子税务局回执","未开始",0,"","保存回执","财务主管",null],
 ["税款缴纳","资金",new Date("2026-07-14"),"申报结果","银行和完税证明","未开始",0,"","确认扣款成功","税务会计",null],
 ["缴税账务处理","总账",new Date("2026-07-14"),"银行回单","应交税费余额","未开始",0,"","录入缴税凭证","财务主管",null],
 ["申报资料归档","税务会计",new Date("2026-07-15"),"申报回执完税回单","税务档案","未开始",0,"","按税种所属期归档","财务主管",null],
 ["异常事项关闭","总账",new Date("2026-07-15"),"异常台账","整改证据","未开始",60,"长期项目继续跟踪","更新状态和期限","财务主管",null],
 ["下月税额预测","税务会计",new Date("2026-07-20"),"滚动销售采购预测","资金计划","未开始",0,"","提供资金部门","财务主管",null],
 ["权限和待办检查","税务会计",new Date("2026-07-20"),"电子税务局","办税权限","未开始",0,"","处理人员变化","财务主管",null]
];
commonChecklist(sheets[names[2]],"月度税务工作清单","案例公司2026年6月所属期，日期为内部计划",monthRows);
const quarterRows=[
 ["季度利润表累计核对","总账",new Date("2026-07-06"),"总账利润表","A类预缴申报","未开始",0,"","确认累计利润170","财务主管",null],
 ["纳税调整台账更新","税务会计",new Date("2026-07-07"),"费用资产研发台账","应纳税所得额","未开始",0,"","调增13.2调减21","财务主管",null],
 ["研发费用资格和辅助账","研发财务",new Date("2026-07-07"),"项目工时材料成果","研发加计扣除","未开始",20,"","核验项目和归集","税务经理",null],
 ["业务招待费限额","税务会计",new Date("2026-07-07"),"招待费和营业收入","纳税调整","未开始",3.2,"","计算双限额","财务主管",null],
 ["资产折旧差异","资产总账",new Date("2026-07-08"),"资产卡片","税会差异","未开始",-1,"","更新会计税法折旧","财务主管",null],
 ["资产减值和损失","总账",new Date("2026-07-08"),"减值和处置资料","纳税调整","未开始",8,"","区分准备和实际损失","财务主管",null],
 ["以前年度亏损","税务会计",new Date("2026-07-08"),"历年申报表","亏损弥补","不适用",0,"","案例无可弥补亏损","财务主管",null],
 ["前期预缴核对","税务会计",new Date("2026-07-08"),"一季度申报完税","累计已缴","未开始",22.5,"","确认已缴22.5","财务主管",null],
 ["季度所得税测算","税务会计",new Date("2026-07-09"),"利润和调整台账","季度申报表","未开始",18.05,"","本季应补18.05","财务主管",null],
 ["印花税合同税源","法务税务",new Date("2026-07-09"),"合同台账","财产行为税申报","未开始",0,"","新增变更合同完整","财务主管",null],
 ["其他财产行为税","税务会计",new Date("2026-07-09"),"资产环保台账","税种核定","未开始",0,"","核对适用性","财务主管",null],
 ["主管批准和申报","财务主管",new Date("2026-07-10"),"季度批准包","电子税务局","未开始",0,"","保存批准和回执","财务负责人",null],
 ["季度缴款归档","资金税务",new Date("2026-07-14"),"回执和银行","应交税费","未开始",18.05,"","完税回单归档","财务主管",null],
 ["季度问题复盘","财务主管",new Date("2026-07-20"),"异常和差异台账","整改计划","未开始",0,"","责任期限明确","财务负责人",null]
]; commonChecklist(sheets[names[3]],"季度税务工作清单","企业所得税预缴、印花税和其他核定税费",quarterRows);
const annualRows=[
 ["政策和申报表版本更新","税务会计",new Date("2026-10-31"),"官方法规","年度工作底稿","未开始",0,"","建立版本表","财务主管",null],
 ["无票支出清理","AP总账",new Date("2026-12-15"),"未到票台账","税前扣除凭证","未开始",0,"","催票和补救资料","财务主管",null],
 ["研发辅助账全年复核","研发财务",new Date("2026-12-20"),"项目工时材料","加计扣除","未开始",0,"","项目级复算","税务经理",null],
 ["业务招待费广告费限额","税务会计",new Date("2027-01-20"),"费用和收入","纳税调整","未开始",0,"","计算本年及结转","财务主管",null],
 ["折旧摊销税会差异","资产总账",new Date("2027-01-20"),"资产卡片","纳税调整","未开始",0,"","逐项滚动","财务主管",null],
 ["资产损失资料","资产总账",new Date("2027-01-31"),"盘点处置诉讼资料","税前扣除","未开始",0,"","证据清单完整","财务主管",null],
 ["以前年度亏损","税务会计",new Date("2027-01-31"),"历年申报","弥补余额","未开始",0,"","核对到期年度","财务主管",null],
 ["税收优惠资格","税务经理",new Date("2027-02-15"),"资质和项目资料","优惠申报","未开始",0,"","条件逐项核验","财务负责人",null],
 ["关联交易和关联申报","税务经理",new Date("2027-02-28"),"关联方和交易台账","关联申报","未开始",0,"","范围金额一致","财务负责人",null],
 ["年度财务报表定稿","总账",new Date("2027-03-31"),"审计和关账","所得税申报","未开始",0,"","审计调整同步","财务主管",null],
 ["汇算清缴底稿","税务会计",new Date("2027-04-15"),"总账和全部台账","年度申报表","未开始",0,"","可追溯可复算","财务主管",null],
 ["风险扫描和更正","税务经理",new Date("2027-04-30"),"年度申报草表","历史申报","未开始",0,"","重大问题处理","财务负责人",null],
 ["批准年度申报","财务负责人",new Date("2027-05-15"),"申报批准包","年度申报表","未开始",0,"","签署批准","管理层",null],
 ["申报补退税","税务会计",new Date("2027-05-25"),"批准申报","完税退税","未开始",0,"","完成申报缴税","财务负责人",null],
 ["年度资料归档","税务会计",new Date("2027-06-15"),"申报回执底稿","税务档案","未开始",0,"","目录完整","财务主管",null],
 ["制度和控制更新","财务主管",new Date("2027-06-30"),"问题复盘","下一年度流程","未开始",0,"","完成改进","财务负责人",null]
]; commonChecklist(sheets[names[4]],"年度税务工作清单","企业所得税汇算清缴、优惠、资产损失和年度归档",annualRows);

// 05 增值税核对
{
 const sh=sheets[names[5]]; title(sh,"增值税账票税核对表","案例公司2026年6月，金额单位万元；黄色单元格为可更新输入",13);
 sh.getRange("A5:M5").values=[["类别","业务事项","业务金额 不含税","开票金额 不含税","账面金额 不含税","税率","发票税额","账面税额","用途确认税额","申报税额","差异","原因","处理"]]; header(sh,"A5:M5");
 const vat=[
  ["销项","正常销售",300,300,300,0.13,39,39,null,39,null,"",""],
  ["销项","未开票销售",80,0,80,0.13,0,10.4,null,10.4,null,"已验收取得收款权","申报未开票销售"],
  ["销项","退货红字",-20,-20,-20,0.13,-2.6,-2.6,null,-2.6,null,"销售退货","关联蓝字和退货单"],
  ["进项","材料专票",100,100,100,0.13,13,13,13,13,null,"",""],
  ["进项","设备专票",50,50,50,0.13,6.5,6.5,6.5,6.5,null,"",""],
  ["进项","咨询服务",10,10,10,0.06,0.6,0.6,0.6,0.6,null,"",""],
  ["进项","差旅凭证",2,2,2,0.09,0.18,0.18,0.18,0.18,null,"简化案例","逐票按凭证规则"],
  ["进项","生产水电",10,10,10,0.13,1.3,1.3,1.3,1.3,null,"",""],
  ["转出","福利领料",20,20,20,0.13,2.6,2.6,null,2.6,null,"改变用途","进项转出"]
 ]; sh.getRange("A6:J14").values=vat.map(r=>r.slice(0,10)); sh.getRange("K6:K14").formulas=vat.map((r,i)=>[`=H${i+6}-J${i+6}`]); sh.getRange("L6:M14").values=vat.map(r=>r.slice(11)); body(sh,"A6:M14");
 sh.getRange("C6:J14").format.fill=amber; sh.getRange("K6:K14").format.fill=gray; sh.getRange("F6:F14").setNumberFormat("0.00%"); sh.getRange("C6:K14").setNumberFormat("#,##0.00;[Red](#,##0.00);-");
 sh.getRange("A17:D17").values=[["增值税汇总","金额","计算","核对结果"]]; header(sh,"A17:D17");
 sh.getRange("A18:A23").values=[["销项税额"],["进项取得"],["进项转出"],["可抵扣进项"],["期初留抵"],["应纳增值税"]];
 sh.getRange("B18:B23").formulas=[["=SUMIF(A6:A14,\"销项\",J6:J14)"],["=SUMIF(A6:A14,\"进项\",J6:J14)"],["=SUMIF(A6:A14,\"转出\",J6:J14)"],["=B19-B20"],["=3"],["=B18-B21-B22"]];
 sh.getRange("C18:C23").values=[["39+10.4-2.6"],["13+6.5+0.6+0.18+1.3"],["福利用途"],["进项取得减转出"],["上期申报表"],["销项减可抵扣减留抵"]]; sh.getRange("D18:D23").formulas=[["=IF(ABS(B18-46.8)<0.001,\"一致\",\"检查\")"],["=IF(ABS(B19-21.58)<0.001,\"一致\",\"检查\")"],["=IF(ABS(B20-2.6)<0.001,\"一致\",\"检查\")"],["=IF(ABS(B21-18.98)<0.001,\"一致\",\"检查\")"],["=IF(B22=3,\"一致\",\"检查\")"],["=IF(ABS(B23-24.82)<0.001,\"一致\",\"检查\")"]]; body(sh,"A18:D23"); sh.getRange("B18:B23").setNumberFormat("#,##0.00"); statusCF(sh.getRange("D18:D23"));
 widths(sh,{A:11,B:24,C:17,D:17,E:17,F:10,G:14,H:14,I:16,J:14,K:12,L:24,M:26}); sh.freezePanes.freezeRows(5); sh.freezePanes.freezeColumns(2); sh.tables.add("A5:M14",true,"VATReconciliation").style="TableStyleMedium2";
}

// 06 所得税
{
 const sh=sheets[names[6]]; title(sh,"企业所得税检查表","2026年第二季度预缴案例，金额单位万元",10);
 sh.getRange("A5:J5").values=[["项目","会计金额","税务调整","调整方向","税务金额","数据来源","政策依据","是否完成","复核人","备注"]]; header(sh,"A5:J5");
 const cit=[
  ["利润总额",170,0,"基础",170,"累计利润表","企业所得税法","未开始","财务主管",""],
  ["业务招待费",8,3.2,"调增",11.2,"费用和收入台账","实施条例", "未开始","财务主管","扣除4.8"],
  ["行政罚款",2,2,"调增",4,"营业外支出","企业所得税法", "未开始","财务主管","不得扣除"],
  ["资产减值准备",8,8,"调增",16,"减值台账","资产损失规则", "未开始","财务主管","未实际损失"],
  ["折旧差异",0,-1,"调减",-1,"资产税会差异","实施条例", "未开始","财务主管","税收折旧较多"],
  ["研发费用加计扣除",20,-20,"调减",0,"研发辅助账","2023年第7号", "未开始","税务经理","额外100%"],
  ["以前年度亏损",0,0,"调减",0,"历年申报","企业所得税法", "不适用","财务主管","案例无余额"]
 ]; sh.getRange("A6:J12").values=cit; body(sh,"A6:J12"); statusValidation(sh.getRange("H6:H12")); statusCF(sh.getRange("H6:H12")); sh.getRange("B6:E12").setNumberFormat("#,##0.00;[Red](#,##0.00);-");
 sh.getRange("A15:D15").values=[["汇总","金额","公式","结果"]]; header(sh,"A15:D15"); sh.getRange("A16:A21").values=[["会计利润"],["纳税调整净额"],["应纳税所得额"],["累计应纳所得税"],["前期已预缴"],["本季应补"]]; sh.getRange("B16:B21").formulas=[["=B6"],["=SUM(C7:C12)"],["=B16+B17"],["=B18*0.25"],["=22.5"],["=B19-B20"]]; sh.getRange("C16:C21").values=[["利润表累计"],["调增与调减合计"],["利润加净调整"],["税率25%"],["一季度申报"],["累计税额减已缴"]]; sh.getRange("D16:D21").formulas=[["=IF(B16=170,\"一致\",\"检查\")"],["=IF(ABS(B17+7.8)<0.001,\"一致\",\"检查\")"],["=IF(ABS(B18-162.2)<0.001,\"一致\",\"检查\")"],["=IF(ABS(B19-40.55)<0.001,\"一致\",\"检查\")"],["=IF(B20=22.5,\"一致\",\"检查\")"],["=IF(ABS(B21-18.05)<0.001,\"一致\",\"检查\")"]]; body(sh,"A16:D21"); sh.getRange("B16:B21").setNumberFormat("#,##0.00");
 widths(sh,{A:28,B:15,C:15,D:12,E:15,F:24,G:24,H:11,I:12,J:28}); sh.freezePanes.freezeRows(5); sh.tables.add("A5:J12",true,"CITCheck").style="TableStyleMedium2";
}

// 07 发票异常
{
 const sh=sheets[names[7]]; title(sh,"发票异常跟踪表","所有异常必须有业务编号、处理动作、复核人和关闭证据",14);
 const hdr=["异常编号","发现日期","业务编号","发票号码","供应商或客户","异常类型","价税金额 万元","发现人","异常原因","系统处理","会计处理","税务处理","责任人","状态","复核人","关闭日期","证据位置"];
 sh.getRange("A5:Q5").values=[hdr]; header(sh,"A5:Q5");
 const rows=[
  ["INV-001",new Date("2026-06-13"),"T03","2606XXXX001","苏州精密材料","票面已核验",113,"AP","无异常","已匹配","已入账","用途确认","AP","已关闭","总账",new Date("2026-06-13"),"采购档案/T03"],
  ["INV-002",new Date("2026-06-28"),"T12","2606XXXX009","苏州精密材料","用途改变",22.6,"总账","材料转职工福利","标记转出","计入职工薪酬","进项转出2.6","总账","待复核","税务会计",null,"领料单/T12"],
  ["INV-003",new Date("2026-06-30"),"T04","","无锡电子组件","货到票未到",60,"AP","供应商未开票","暂估入库","原材料和暂估应付","不确认进项","采购","进行中","总账",null,"暂估台账/T04"]
 ]; sh.getRange("A6:Q25").values=[...rows,...Array(17).fill(Array(17).fill(null))]; body(sh,"A6:Q25"); sh.getRange("B6:B25").setNumberFormat("yyyy-mm-dd"); sh.getRange("P6:P25").setNumberFormat("yyyy-mm-dd"); sh.getRange("G6:G25").setNumberFormat("#,##0.00"); sh.getRange("F6:F25").dataValidation={rule:{type:"list",values:["抬头错误","税号错误","税率错误","金额错误","内容错误","重复开票","重复报销","跨期","票到货未到","货到票未到","用途改变","有票无业务","红字待处理","票面已核验"]}}; sh.getRange("N6:N25").dataValidation={rule:{type:"list",values:["未处理","进行中","待复核","已关闭"]}}; statusCF(sh.getRange("N6:N25")); widths(sh,{A:14,B:13,C:12,D:20,E:24,F:16,G:15,H:12,I:24,J:22,K:22,L:22,M:12,N:11,O:12,P:13,Q:26}); sh.freezePanes.freezeRows(5); sh.freezePanes.freezeColumns(3); sh.tables.add("A5:Q25",true,"InvoiceExceptions").style="TableStyleMedium2";
}

// 08 暂估
{
 const sh=sheets[names[8]]; title(sh,"未到票与暂估台账","暂估解决会计截止，不产生可抵扣进项税；收票后按实际发票和价差处理",14);
 const hdr=["记录编号","业务编号","供应商","业务类型","收货验收日","暂估金额 万元","预计税率","暂估税额 不入账","收到发票日","发票不含税金额","发票税额","价差","账龄天数","状态","责任人","预计关闭日","处理说明"];
 sh.getRange("A5:Q5").values=[hdr]; header(sh,"A5:Q5");
 sh.getRange("A6:Q25").values=[
  ["ACR-001","T02","苏州精密材料","原材料",new Date("2026-06-05"),100,0.13,13,new Date("2026-06-12"),100,13,0,7,"已关闭","AP",new Date("2026-06-12"),"冲暂估并按发票入账"],
  ["ACR-002","T04","无锡电子组件","原材料",new Date("2026-06-15"),60,0.13,7.8,null,null,null,null,15,"进行中","采购",new Date("2026-07-15"),"催票；税额仅供预测"],
  ...Array(18).fill(Array(17).fill(null))
 ]; body(sh,"A6:Q25"); sh.getRange("F6:H25").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sh.getRange("J6:L25").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sh.getRange("G6:G25").setNumberFormat("0.00%"); sh.getRange("E6:E25").setNumberFormat("yyyy-mm-dd"); sh.getRange("I6:I25").setNumberFormat("yyyy-mm-dd"); sh.getRange("P6:P25").setNumberFormat("yyyy-mm-dd"); sh.getRange("N6:N25").dataValidation={rule:{type:"list",values:["未开始","催票中","已收票待处理","待复核","已关闭","进行中"]}}; statusCF(sh.getRange("N6:N25")); sh.getRange("L6:L25").formulas=[["=IF(OR(F6=\"\",J6=\"\"),\"\",J6-F6)"]]; sh.getRange("L6:L25").fillDown(); widths(sh,{A:14,B:12,C:24,D:14,E:13,F:16,G:12,H:18,I:13,J:17,K:13,L:12,M:12,N:14,O:12,P:13,Q:30}); sh.freezePanes.freezeRows(5); sh.freezePanes.freezeColumns(3); sh.tables.add("A5:Q25",true,"AccrualTracking").style="TableStyleMedium2";
}

// 09 税务异常整改
{
 const sh=sheets[names[9]]; title(sh,"税务异常整改台账","按发现、量化、调账、更正、税款和控制整改形成闭环",16);
 const hdr=["整改编号","发现日期","税种","所属期","异常类型","发现异常","原因判断","资料保全","影响税额 万元","是否调账","是否更正申报","税款及滞纳金处理","防复发措施","责任人","复核人","状态","计划关闭日","实际关闭日"];
 sh.getRange("A5:R5").values=[hdr]; header(sh,"A5:R5");
 sh.getRange("A6:R25").values=[
  ["TAX-001",new Date("2026-07-03"),"增值税",new Date("2026-06-30"),"进项转出","福利领料已抵扣进项","用途改变","领料单和发票",2.6,"是","否","当期正常申报转出","增加用途检查","总账","税务主管","待复核",new Date("2026-07-10"),null],
  ["TAX-002",new Date("2026-07-05"),"增值税",new Date("2026-06-30"),"未开票收入","验收80万元未开票","AR开票与税务底稿未自动连接","合同验收和应收凭证",10.4,"否","否","当期按未开票销售申报","增加销售截止接口","税务会计","财务主管","进行中",new Date("2026-07-10"),null],
  ...Array(18).fill(Array(18).fill(null))
 ]; body(sh,"A6:R25"); sh.getRange("B6:B25").setNumberFormat("yyyy-mm-dd"); sh.getRange("D6:D25").setNumberFormat("yyyy-mm-dd"); sh.getRange("Q6:R25").setNumberFormat("yyyy-mm-dd"); sh.getRange("I6:I25").setNumberFormat("#,##0.00"); sh.getRange("J6:K25").dataValidation={rule:{type:"list",values:["是","否","待判断"]}}; sh.getRange("P6:P25").dataValidation={rule:{type:"list",values:["未处理","进行中","待复核","已关闭"]}}; statusCF(sh.getRange("P6:P25")); widths(sh,{A:14,B:13,C:12,D:13,E:16,F:26,G:26,H:24,I:15,J:12,K:14,L:28,M:28,N:12,O:12,P:11,Q:13,R:13}); sh.freezePanes.freezeRows(5); sh.freezePanes.freezeColumns(4); sh.tables.add("A5:R25",true,"TaxRemediation").style="TableStyleMedium2";
}

// 10 数据字典
{
 const sh=sheets[names[10]]; title(sh,"数据字典","用于统一状态、责任和异常分类；修改下拉口径时同步检查相关工作表",8);
 sh.getRange("A5:H5").values=[["类别","代码","名称","定义","是否需复核","默认责任人","关闭条件","备注"]]; header(sh,"A5:H5");
 const rows=[
  ["状态","S01","未开始","尚未执行","否","事项责任人","不适用",""],["状态","S02","进行中","已开始但未完成","否","事项责任人","全部步骤完成",""],["状态","S03","待复核","经办完成等待复核","是","复核人","复核通过",""],["状态","S04","已完成","任务全部完成","是","复核人","证据归档",""],["状态","S05","不适用","本期无触发业务","是","复核人","记录不适用原因",""],
  ["异常","E01","主体错误","合同付款发票主体不一致","是","AP或AR","更正及证据完成",""],["异常","E02","税率错误","税目税率与业务不符","是","税务会计","红冲重开及申报处理",""],["异常","E03","跨期","业务发票会计申报期间不同","是","总账","差异解释或更正完成",""],["异常","E04","重复","发票或报销重复","是","AP","冲销追回和控制整改",""],["异常","E05","遗漏","业务未进入账或申报","是","总账税务","补记更正缴税完成",""],
  ["税种","T01","增值税","一般纳税人月度申报案例","是","税务会计","申报缴款归档",""],["税种","T02","企业所得税","季度预缴年度汇算","是","税务会计","申报缴款归档",""],["税种","T03","个人所得税","全员全额扣缴申报","是","薪酬税务","申报缴款归档",""],["税种","T04","印花税","依税源和当地核定","是","税务会计","申报缴款归档",""],
  ["口径","P01","会计准则要求","确认计量列报","是","总账","政策判断留痕",""],["口径","P02","税法要求","纳税抵扣扣除申报","是","税务","政策判断留痕",""],["口径","P03","常见实务","成熟企业流程参考","否","流程负责人","企业批准",""],["口径","P04","内部政策","企业授权截止阈值","是","财务主管","制度发布","不代替法律"]
 ]; sh.getRange(`A6:H${5+rows.length}`).values=rows; body(sh,`A6:H${5+rows.length}`); widths(sh,{A:12,B:10,C:20,D:35,E:13,F:18,G:30,H:25}); sh.freezePanes.freezeRows(5); sh.tables.add(`A5:H${5+rows.length}`,true,"DataDictionary").style="TableStyleMedium2";
}

// 全局格式和导出
for (const sh of Object.values(sheets)) {
  const used=sh.getUsedRange(); used.format.font.name=font; used.format.verticalAlignment="center";
}
wb.recalculate();
const inspect = await wb.inspect({kind:"table",sheetId:"05-增值税账票税",range:"A17:D23",include:"values,formulas",tableMaxRows:10,tableMaxCols:6,maxChars:5000});
console.log(inspect.ndjson);
const errors = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:100},summary:"final formula error scan"});
console.log(errors.ndjson);
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outFile);
for (const sh of Object.values(sheets)) {
  const preview=await wb.render({sheetName:sh.name,autoCrop:"all",scale:1,format:"png"});
  await fs.writeFile(new URL(`./qa/${sh.name}.png`, import.meta.url),new Uint8Array(await preview.arrayBuffer()));
}
console.log(outFile);
