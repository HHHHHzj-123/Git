from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import json,uuid
ROOT=Path(__file__).parent;OUT=ROOT/'deliverables';OUT.mkdir(exist_ok=True)
DATA=json.loads((ROOT/'cases.json').read_text(encoding='utf-8')); FILE=OUT/'企业异常费用与发票问题.xmind'
def tid():return str(uuid.uuid4())
def topic(title,children=None,note=None):
    t={'id':tid(),'class':'topic','title':title}
    if children:t['children']={'attached':[topic(x[0],x[1] if len(x)>1 else None,x[2] if len(x)>2 else None) if isinstance(x,tuple) else topic(x) for x in children]}
    if note:t['notes']={'plain':{'content':note}}
    return t
def n(title,children):return(title,children)
def sheet(title,root,children):return {'id':tid(),'class':'sheet','title':title,'rootTopic':topic(root,children),'topicPositioning':'right','extensions':[]}

by={}
for c in DATA['cases']:by.setdefault(c['category'],[]).append(c)
branches=[]
for prefix,(name,coverage,count) in DATA['categories'].items():
    children=[]
    for c in by[name]:
        children.append(n(f"{c['id']} {c['title']}",[f"风险：{c['risk']}",f"结论：{c['status']}",f"科目：{c['entries'].split('。')[0]}",f"后果：{c['consequences'].split('。')[0]}",f"证据：{c['docs'].split('；')[0]}"]))
    branches.append(n(f'{name}（{count}题）',children))

decision=[n('11步判断',['还原业务','查最终受益人','穿透合同履约发票付款','会计确认','增值税','企业所得税','个人所得税','发票','证据','权限','拒不整改后果']),n('三个不同问题',['会计是否如实入账','增值税是否抵扣','企业所得税是否扣除']),n('四级风险',['蓝色 正常','黄色 补资料','橙色 暂停并升级','红色 拒绝执行']),n('主体六问',['谁签合同','谁履约','谁取得货物服务','谁承担付款义务','发票开给谁','资金最终流向谁'])]
conseq=[n('内部流程',['退回','拒付','追回','权限取消','依合法制度处分']),n('会计纠正',['冲销或重分类','更正报表','参与造假可能罚款/禁业']),n('税款处理',['进项扣减','纳税调增','更正申报','补税','日万分之五滞纳金']),n('行政处罚',['逃税：少缴税款50%至5倍罚款','虚开发票行政罚款','没收违法所得']),n('公司治理',['返还资金','赔偿损失','关联交易责任']),n('刑事责任',['须满足犯罪构成与追诉标准','单位/主管/直接责任人员','可能罚金拘役或有期徒刑','问题发票不自动等于坐牢'])]
sop=[n('老板指令SOP',['明确事实','判断后果','找法规','合法替代方案','书面说明','按权限升级','红线拒绝','保存原始记录']),n('财务留痕',['OA意见','邮件','异常工单','审批日志','更正记录']),n('不能做',['补造合同','虚构验收','删除证据','改开不实品名','借他人发票'])]
content=[sheet('141案例分类索引','企业异常费用与发票问题',branches),sheet('判断与后果','异常业务判断框架',decision+conseq),sheet('沟通与自我保护','财务面对不合理指令',sop)]
meta={'creator':{'name':'企业财务实操知识库','version':'1.0'},'activeSheetId':content[0]['id']};manifest={'file-entries':{'content.json':{},'metadata.json':{}}}
with ZipFile(FILE,'w',ZIP_DEFLATED) as z:
    z.writestr('content.json',json.dumps(content,ensure_ascii=False,separators=(',',':')));z.writestr('metadata.json',json.dumps(meta,ensure_ascii=False,separators=(',',':')));z.writestr('manifest.json',json.dumps(manifest,ensure_ascii=False,separators=(',',':')))
print(FILE)
