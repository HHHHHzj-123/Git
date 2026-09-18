let
  Bank = Csv.Document(File.Contents("银行流水.csv"),[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv]),
  BankHeaders = Table.PromoteHeaders(Bank,[PromoteAllScalars=true]),
  BankTypes = Table.TransformColumnTypes(BankHeaders,{{"date", type date},{"debit", type number},{"credit", type number}}),
  BankAmount = Table.AddColumn(BankTypes,"amount", each [credit]-[debit], type number),
  ERP = Csv.Document(File.Contents("ERP银行明细.csv"),[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv]),
  ERPHeaders = Table.PromoteHeaders(ERP,[PromoteAllScalars=true]),
  ERPTypes = Table.TransformColumnTypes(ERPHeaders,{{"date", type date},{"amount", type number}}),
  Join = Table.NestedJoin(BankAmount,{"date","amount"},ERPTypes,{"date","amount"},"ERP",JoinKind.LeftOuter),
  Expand = Table.ExpandTableColumn(Join,"ERP",{"erp_id","counterparty"},{"erp_id","erp_counterparty"}),
  Status = Table.AddColumn(Expand,"status", each if [erp_id]=null then "未匹配" else "精确候选")
in Status