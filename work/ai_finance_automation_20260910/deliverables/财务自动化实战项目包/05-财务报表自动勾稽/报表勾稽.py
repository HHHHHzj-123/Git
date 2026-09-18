import argparse, pandas as pd
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input",default="输入数据/报表数据.csv"); ap.add_argument("--output",default="运行结果.xlsx"); a=ap.parse_args()
    x=pd.read_csv(a.input); v=dict(zip(x.item,x.amount))
    checks=[("资产=负债+权益",v["资产总计"]-v["负债总计"]-v["所有者权益总计"]),("未分配利润滚动",v["期初未分配利润"]+v["净利润"]-v["本期分红"]-v["期末未分配利润"]),("现金流与货币资金变动",v["期末货币资金"]-v["期初货币资金"]-v["现金净增加额"]),("TB未映射科目",v["未映射科目金额"])]
    out=pd.DataFrame(checks,columns=["check_name","difference"]); out["status"]=out.difference.abs().le(.01).map({True:"通过",False:"未通过"}); out.to_excel(a.output,index=False)
if __name__=="__main__": main()
