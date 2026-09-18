let
  S = Excel.CurrentWorkbook(){[Name="BalanceTable"]}[Content],
  T = Table.TransformColumnTypes(S,{{"business_date",type date},{"debit_balance",type number},{"credit_balance",type number}}),
  Age = Table.AddColumn(T,"age_days",each Duration.Days(#date(2026,9,10)-[business_date]),Int64.Type),
  Bucket = Table.AddColumn(Age,"aging_bucket",each if [age_days]<=30 then "30天以内" else if [age_days]<=90 then "30-90天" else if [age_days]<=180 then "90-180天" else if [age_days]<=365 then "180-365天" else if [age_days]<=730 then "1-2年" else "2年以上"),
  Flag = Table.AddColumn(Bucket,"direction_flag",each if Text.StartsWith([account],"应收") and [credit_balance]>0 then "方向异常" else "")
in Flag