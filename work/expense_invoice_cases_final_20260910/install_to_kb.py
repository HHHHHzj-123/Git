from pathlib import Path
import shutil
src=Path(__file__).parent
kb=Path(r'C:\Users\HZJ\Desktop\财务管理实操')
dest=kb/'09-发票与税务'/'02-企业发票报销与异常支出'/'01-正式版';dest.mkdir(parents=True,exist_ok=True)
for f in ['00-阅读说明.md','企业发票、报销与异常支出实战手册.docx','企业异常费用与发票问题.xmind','异常发票与费用处理速查表.xlsx']:
    p=src/f if f.endswith('.md') else src/'deliverables'/f
    shutil.copy2(p,dest/f)
nav=kb/'00-知识库与学习导航'/'README.md'
line='- [M09-002 企业发票、报销与异常支出](../09-发票与税务/02-企业发票报销与异常支出/01-正式版/00-阅读说明.md)：141个异常业务案例、拒不整改后果、责任主体、老板指令SOP、沟通模板及法规索引。'
t=nav.read_text(encoding='utf-8')
if 'M09-002 企业发票、报销与异常支出' not in t: nav.write_text(t.rstrip()+'\n\n'+line+'\n',encoding='utf-8')
ledger=kb/'00-知识库与学习导航'/'02-学习与问题台账.md'
record='| 2026-09-10 | 建立M09-002企业发票、报销与异常支出专题 | 141个案例；286页Word；3张XMind工作表；13页签Excel；新增拒不整改后果、责任主体、老板指令SOP和20条沟通模板 | 政策核验截至2026-09-10；地区和个案需按行为发生日、完整事实及主管机关口径复核；刑事责任仅在满足法定构成时成立 |'
t=ledger.read_text(encoding='utf-8')
if '建立M09-002企业发票、报销与异常支出专题' not in t: ledger.write_text(t.rstrip()+'\n\n'+record+'\n',encoding='utf-8')
for p in dest.iterdir(): print(p.name,p.stat().st_size)
