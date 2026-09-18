import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--accrual",default="输入数据/暂估明细.csv"); ap.add_argument("--invoice",default="输入数据/发票明细.csv"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    x=pd.read_csv(a.accrual); y=pd.read_csv(a.invoice)
    z=x.merge(y,on=["vendor","po","material"],how="left",suffixes=("_accrual","_invoice"))
    z["amount_diff"]=z["amount_invoice"]-z["amount_accrual"]
    z["qty_diff"]=z["qty_invoice"]-z["qty_accrual"]
    z["result"]=z.apply(lambda r:"未到票" if pd.isna(r.invoice_id) else ("可冲销" if abs(r.amount_diff)<.01 and abs(r.qty_diff)<.001 else "差异待查"),axis=1)
    z.to_excel(a.output,index=False)
if __name__=="__main__": main()
