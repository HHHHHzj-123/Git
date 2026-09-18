from pathlib import Path
b=Path(__file__).resolve().parent
p=b/'build_data.py'; s=p.read_text('utf-8').replace('"inventory_total": 25860000','"inventory_total": 23271600').replace('"avg_inventory": 23800000','"avg_inventory": 22500000'); p.write_text(s,'utf-8')
p=b/'build_document.py'; s=p.read_text('utf-8').replace('总库存2,586万元','总库存2,327万元'); p.write_text(s,'utf-8')
p=b/'build_workbook.mjs'; s=p.read_text('utf-8').replace('2,586万元','2,327万元').replace(r'''=SUMIF('02-库存明细'!K6:K23,\"*呆滞*\",'02-库存明细'!H6:H23)''',r'''='02-库存明细'!H9+'02-库存明细'!H18''').replace(r'''=COUNTIF('24-Issue Log'!J6:J45,\"高\")''',r'''=COUNTIF('24-Issue Log'!E6:E45,\"高\")'''); p.write_text(s,'utf-8')
print('patched')
