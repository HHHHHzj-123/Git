from pathlib import Path
p=Path(__file__).with_name('build_workbooks.mjs')
s=p.read_text(encoding='utf-8')
old="=SUM(ABS('02-银行账户与余额调节'!H6:H25))+SUM(ABS('03-AR与GL核对'!F6:F25))+SUM(ABS('04-AP与GL核对'!F6:F25))"
refs=[('02-银行账户与余额调节','H'),('03-AR与GL核对','F'),('04-AP与GL核对','F')]
new='+' .join(["="+f"ABS('{refs[0][0]}'!{refs[0][1]}6)"] + [f"ABS('{sh}'!{c}{r})" for sh,c in refs for r in range(6,26)][1:])
assert old in s
p.write_text(s.replace(old,new),encoding='utf-8')
