from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import json,uuid
r=Path(__file__).parent;o=r/'deliverables';o.mkdir(exist_ok=True);d=json.loads((r/'data.json').read_text(encoding='utf-8'));f=o/'财务报表编制与校验全景.xmind'
def t(x,ch=None):
 z={'id':str(uuid.uuid4()),'class':'topic','title':x}
 if ch:z['children']={'attached':[t(a[0],a[1]) if isinstance(a,tuple) else t(a) for a in ch]}
 return z
def n(x,ch):return(x,ch)
flow=n('报表形成流程',['业务与凭证','明细账与总账','调整前TB','AJ账务调整','AJ后TB','RJ报表重分类','最终报表TB','资产负债表','利润表','现金流量表','所有者权益变动表','四层勾稽','锁账归档'])
mapping=n('科目与报表映射',[n('映射类型',['直接','汇总','拆分','净额','条件']),n('判断维度',['正常方向','辅助核算','备抵','流动性','抵销条件','新增科目']),*[(x['name'],[x['map'],x['class']]) for x in d['accounts']]])
neg=n('负数余额',[n(x[0],[x[1],x[2],x[3],'负数≠错误','负数≠必须重分类']) for x in d['negative']])
adj=n('调整与重分类',[n('AJ 必须做凭证',[n(x['id']+' '+x['desc'],[x['debit_code'],x['credit_code'],x['profit'],x['tax']]) for x in d['adjustments']]),n('RJ 仅调整列报',[n(x['id']+' '+x['desc'],[x['to_item'],x['from_item'],x['warning']]) for x in d['reclasses']]),n('四种调整',['账务调整','报表重分类','审计调整来源','抵销/净额列报'])])
cf=n('现金流量表',[n('逐笔编制',[n(x['id']+' '+x['desc'],[x['class'],x['item'],x['evidence']]) for x in d['cashflows']]),n('核心桥',['期初现金900','净增加505','期末现金等价物1405','货币资金1525','受限保证金120']),n('30个易错场景',['预收预付','税费与补助','保证金和受限资金','资产与在建工程','利息租赁和分红','关联往来','票据保理','非现金交易','外币与未达'])])
checks=n('报表勾稽',[n('表内',['资产=负债+权益','利润层级','现金流加总','权益滚动']),n('表间',['净利润↔权益','现金↔资产负债表','未分配利润滚动','所得税']),n('表账',['主表↔最终TB','利润表↔发生额','现金流↔流水']),n('账模块',['AR','AP','Inventory','FA','Tax','Bank'])])
errors=n('50个报表错误',[n(x,['现象','定位','AJ/RJ判断','正确处理','校验']) for x in d['errors']])
main=[flow,mapping,n('资产负债表',['逐项目取数','明细方向','备抵净额','流动非流动','抵销限制']),n('利润表',['本月与累计','红字冲销','利得损失','所得税','本年利润']),cf,n('所有者权益变动表',['期初权益','净利润','其他综合收益','股东投入','利润分配','期末权益']),neg,adj,n('常见调整',['收入','成本','费用','薪酬','资产','往来','税务','减值','融资']),n('流动非流动',['长期借款300','租赁负债90','长期应收60','预计负债45','递延收益40']),n('减值与备抵',['应收-坏账','存货-跌价','固定资产-折旧-减值','无形资产-摊销-减值']),checks,errors,n('审计调整',['审计提出是来源','企业判断AJ/RJ','接受后入账','跨年期初衔接','不应年年表层挂账']),n('ERP报表逻辑',['业务模块','凭证','GL','TB','映射','报表模板','手工调整层','现金流辅助项目'])]
content=[{'id':str(uuid.uuid4()),'class':'sheet','title':'报表编制全景','rootTopic':t('企业财务报表编制',main),'topicPositioning':'right','extensions':[]},{'id':str(uuid.uuid4()),'class':'sheet','title':'TB到四表案例','rootTopic':t('远澜智能2026年8月',[flow,adj,cf,checks]),'topicPositioning':'right','extensions':[]},{'id':str(uuid.uuid4()),'class':'sheet','title':'异常与校验','rootTopic':t('报表异常处理',[neg,errors,checks]),'topicPositioning':'right','extensions':[]}]
with ZipFile(f,'w',ZIP_DEFLATED)as z:
 z.writestr('content.json',json.dumps(content,ensure_ascii=False,separators=(',',':')));z.writestr('metadata.json',json.dumps({'creator':{'name':'企业财务实操知识库'},'activeSheetId':content[0]['id']},ensure_ascii=False));z.writestr('manifest.json',json.dumps({'file-entries':{'content.json':{},'metadata.json':{}}}))
print(f)
