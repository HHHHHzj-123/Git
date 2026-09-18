from pathlib import Path
import shutil
r=Path(__file__).parent;kb=Path(r'C:\Users\HZJ\Desktop\财务管理实操');dest=kb/'10-财务报表与合并报表'/'01-企业财务报表编制调整重分类与校验';dest.mkdir(parents=True,exist_ok=True)
for x in ['00-阅读说明.md']:
 shutil.copy2(r/x,dest/x)
for x in ['企业财务报表编制、调整、重分类与校验实操手册.docx','财务报表编制与校验全景.xmind','财务报表编制与校验实操工具.xlsx']:
 shutil.copy2(r/'deliverables'/x,dest/x)
nav=kb/'00-知识库与学习导航'/'README.md';t=nav.read_text(encoding='utf-8');line='- [M10-001 企业财务报表编制、调整、重分类与校验](../10-财务报表与合并报表/01-企业财务报表编制调整重分类与校验/00-阅读说明.md)：负数余额穿透、AJ/RJ分层、调整前TB到四表、现金流逐笔分类、50个错误案例及四层勾稽。'
if 'M10-001 企业财务报表编制、调整、重分类与校验' not in t:nav.write_text(t.rstrip()+'\n\n'+line+'\n',encoding='utf-8')
log=kb/'00-知识库与学习导航'/'02-学习与问题台账.md';t=log.read_text(encoding='utf-8');line='| 2026-09-10 | 建立M10-001财务报表编制调整重分类与校验专题 | 178页Word、3张XMind工作表、22页签Excel；原始TB、11条AJ、8条RJ、四张报表、30个现金流场景和50个错误案例 | 案例公司2026年不提前执行财会〔2026〕11号；实际适用口径需按主体及期间复核 |'
if '建立M10-001财务报表编制调整重分类与校验专题' not in t:log.write_text(t.rstrip()+'\n\n'+line+'\n',encoding='utf-8')
for p in dest.iterdir():print(p.name,p.stat().st_size)
