import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--tb",default="输入数据/TB.csv"); ap.add_argument("--rules",default="输入数据/异常规则.csv"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    tb=pd.read_csv(a.tb,dtype={"account_code":str}); rules=pd.read_csv(a.rules,dtype={"account_prefix":str}); out=[]
    for _,r in tb.iterrows():
        for _,q in rules.iterrows():
            if not r.account_code.startswith(q.account_prefix): continue
            hit=(q.condition==r.direction) or (q.condition=="非零" and abs(r.balance)>q.threshold)
            if hit: out.append([r.account_code,r.account_name,r.balance,q.rule_id,q.rule_name,q.action])
    pd.DataFrame(out,columns=["account_code","account_name","balance","rule_id","rule_name","action"]).to_excel(a.output,index=False)
if __name__=="__main__": main()
