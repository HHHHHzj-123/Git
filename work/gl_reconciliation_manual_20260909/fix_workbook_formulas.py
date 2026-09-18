from pathlib import Path
import re
p=Path(__file__).with_name('build_gl_workbook.mjs')
s=p.read_text(encoding='utf-8-sig')
sum_formula='=SUM('+','.join(f"ABS(\'03-异常余额清单\'!G{i})" for i in range(6,17))+')'
max_formula='=MAX('+','.join(f"ABS(\'03-异常余额清单\'!G{i})" for i in range(6,17))+')'
s=re.sub(r"=SUM\(ABS\('03-异常余额清单'!G6\).*?G30\)\)",sum_formula,s)
s=re.sub(r"=MAX\(ABS\('03-异常余额清单'!G6\).*?G30\)\)",max_formula,s)
s=s.replace(sum_formula,"=SUM('03-异常余额清单'!W6:W30)")
s=s.replace(max_formula,"=MAX('03-异常余额清单'!W6:W30)")
p.write_text(s,encoding='utf-8')
print(sum_formula)
print(max_formula)
