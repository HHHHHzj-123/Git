from pathlib import Path

base=Path(__file__).resolve().parent
p=base/'build_document.py'
s=p.read_text(encoding='utf-8')
s=s.replace("['风险层级','账龄参考','案例率','使用条件'],D['ecl_rates'],[2.2,2.2,2.0,2.2,7.3]","['风险层级','起始天数','截止天数','案例率','使用依据'],D['ecl_rates'],[2.4,2.0,2.0,2.0,7.3]")
p.write_text(s,encoding='utf-8')

p=base/'build_workbook.mjs'
s=p.read_text(encoding='utf-8')
s=s.replace("'02-应收明细'!N6:N20,\"逾期\"","'02-应收明细'!P6:P20,\"逾期\"")
s=s.replace("'02-应收明细'!L6:L20,\">90\"","'02-应收明细'!N6:N20,\">90\"")
s=s.replace("'02-应收明细'!O6:O20,\"是\"","'02-应收明细'!Q6:Q20,\"是\"")
s=s.replace("'业务状态','信用等级','逾期天数','账龄段','到期状态','争议','责任销售'","'业务状态','来源部门','信用等级','逾期天数','账龄段','到期状态','争议','责任销售'")
s=s.replace("K:18,L:12,M:13,N:14,O:13,P:10,Q:14","K:18,L:14,M:12,N:13,O:14,P:13,Q:10,R:14")
s=s.replace("s.getRange('P6:P20').format.fill=yellow;s.getRange('O6:O20').conditionalFormats", "s.getRange('Q6:Q20').format.fill=yellow;s.getRange('P6:P20').conditionalFormats")
s=s.replace("`='02-应收明细'!M${r}`,`='02-应收明细'!N${r}`", "`='02-应收明细'!N${r}`,`='02-应收明细'!O${r}`")
s=s.replace("return [...x,c[3],`=MAX(0,DATE(2026,8,31)-E${r})`", "return [x[0],x[1],x[2],new Date(x[3]),new Date(x[4]),...x.slice(5),c[3],`=MAX(0,DATE(2026,8,31)-E${r})`")
s=s.replace("`=MAX(0,DATE(2026,8,31)-DATEVALUE('${x[4]}'))`", "`=MAX(0,DATE(2026,8,31)-DATE(${x[4].slice(0,4)},${Number(x[4].slice(5,7))},${Number(x[4].slice(8,10))}))`")
s=s.replace(r"'02-应收明细'!N6:N20,\"逾期\"",r"'02-应收明细'!P6:P20,\"逾期\"")
s=s.replace(r"'02-应收明细'!L6:L20,\">90\"",r"'02-应收明细'!N6:N20,\">90\"")
s=s.replace(r"'02-应收明细'!O6:O20,\"是\"",r"'02-应收明细'!Q6:Q20,\"是\"")
s=s.replace('`=IF(L${r}=0,', '`=IF(N${r}=0,')
s=s.replace('IF(L${r}<=', 'IF(N${r}<=')
s=s.replace("'02-应收明细'!K6:K20", "'02-应收明细'!J6:J20")
s=s.replace("Array.from({length:60},(_,i)=>[`ISS-${String(i+1).padStart(3,'0')}`,'','','',0,'','','',null,'','','','','Open',''])", "[['ISS-001','北方自动化','逾期115天且存在验收争议','月结',904000,'现金、减值和收入','销售总监','完成客户验收争议方案',new Date('2026-09-15'),'冻结新增发货','个别评估ECL','销售与项目联合关闭','财务主管','Open',''],['ISS-002','安达经销商','超额度且逾期100天','信用检查',1017000,'现金和信用','渠道总监','取得回款并重评额度',new Date('2026-09-10'),'冻结发货','评估ECL','按回款恢复额度','财务经理','Open',''],['ISS-003','未知付款人','22.6万元回款未认领','银行对账',226000,'客户账龄和现金预测','资金主管','确认付款主体和用途',new Date('2026-09-02'),'暂挂未认领款','确认后核销','完善付款附言规则','财务主管','In Progress',''],...Array.from({length:57},(_,i)=>[`ISS-${String(i+4).padStart(3,'0')}`,'','','',0,'','','',null,'','','','','',''])]")
p.write_text(s,encoding='utf-8')
