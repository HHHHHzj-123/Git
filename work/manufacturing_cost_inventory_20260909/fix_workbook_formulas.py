from pathlib import Path

p = Path(__file__).with_name("build_cost_workbook.mjs")
s = p.read_text(encoding="utf-8")
old = '''const fs=["=IF(AND(H6>=0,H6<=30),E6,0)","=IF(AND(H6>=31,H6<=90),E6,0)","=IF(AND(H6>=91,H6<=180),E6,0)","=IF(AND(H6>=181,H6<=365),E6,0)","=IF(AND(H6>=366,H6<=730),E6,0)","=IF(H6>730,E6,0)"]'''
new = '''const fs=["=IF(A6=\\\"\\\",\\\"\\\",IF(AND(H6>=0,H6<=30),E6,0))","=IF(A6=\\\"\\\",\\\"\\\",IF(AND(H6>=31,H6<=90),E6,0))","=IF(A6=\\\"\\\",\\\"\\\",IF(AND(H6>=91,H6<=180),E6,0))","=IF(A6=\\\"\\\",\\\"\\\",IF(AND(H6>=181,H6<=365),E6,0))","=IF(A6=\\\"\\\",\\\"\\\",IF(AND(H6>=366,H6<=730),E6,0))","=IF(A6=\\\"\\\",\\\"\\\",IF(H6>730,E6,0))"]'''
if old not in s:
    raise SystemExit("target formula block not found")
p.write_text(s.replace(old, new), encoding="utf-8")
print("patched")
