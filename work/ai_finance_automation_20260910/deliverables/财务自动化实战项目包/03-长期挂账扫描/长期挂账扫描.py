import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input",default="输入数据/往来明细.csv"); ap.add_argument("--as-of",default="2026-09-10"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    x=pd.read_csv(a.input); x["business_date"]=pd.to_datetime(x.business_date,errors="coerce"); x["age_days"]=(pd.Timestamp(a.as_of)-x.business_date).dt.days
    x["aging_bucket"]=pd.cut(x.age_days,[-1,30,90,180,365,730,10**9],labels=["30天以内","30-90天","90-180天","180-365天","1-2年","2年以上"])
    x["direction_flag"]=((x.account.str.startswith("应收"))&(x.credit_balance>0))|((x.account.str.startswith("应付"))&(x.debit_balance>0))
    x["priority"]=x.apply(lambda r:"高" if r.age_days>180 or r.direction_flag else ("中" if r.age_days>90 else "低"),axis=1)
    x.sort_values(["priority","age_days"],ascending=[True,False]).to_excel(a.output,index=False)
if __name__=="__main__": main()
