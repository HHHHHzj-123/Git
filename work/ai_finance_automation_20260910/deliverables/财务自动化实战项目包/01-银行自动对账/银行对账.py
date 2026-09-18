from pathlib import Path
import argparse, pandas as pd

def main():
    ap=argparse.ArgumentParser(description="银行与ERP一对一及跨日候选匹配")
    ap.add_argument("--bank",default="输入数据/银行流水.csv")
    ap.add_argument("--erp",default="输入数据/ERP银行明细.csv")
    ap.add_argument("--output",default="运行结果.xlsx")
    a=ap.parse_args(); b=pd.read_csv(a.bank); e=pd.read_csv(a.erp)
    b["date"]=pd.to_datetime(b["date"],errors="coerce"); e["date"]=pd.to_datetime(e["date"],errors="coerce")
    b["amount"]=b["credit"].fillna(0)-b["debit"].fillna(0); used=set(); out=[]
    for _,r in b.iterrows():
        c=e[(e.amount.round(2)==round(r.amount,2)) & (~e.erp_id.isin(used))].copy()
        if len(c):
            c["days"]=(c.date-r.date).abs().dt.days; c=c[c.days<=1].sort_values(["days","erp_id"])
        if len(c):
            x=c.iloc[0]; used.add(x.erp_id); typ="精确匹配" if x.days==0 else "跨日候选"
            out.append([r.bank_id,x.erp_id,typ,1.0 if x.days==0 else .85,"需复核户名和摘要"])
        else: out.append([r.bank_id,"","未匹配",0,"检查手续费、利息、组合或漏记"])
    result=pd.DataFrame(out,columns=["bank_id","erp_id","match_type","confidence","reason"])
    with pd.ExcelWriter(a.output) as w:
        result.to_excel(w,sheet_name="匹配结果",index=False); e[~e.erp_id.isin(used)].to_excel(w,sheet_name="ERP未匹配",index=False)
if __name__=="__main__": main()
