let
  A = Excel.CurrentWorkbook(){[Name="AccrualTable"]}[Content],
  I = Excel.CurrentWorkbook(){[Name="InvoiceTable"]}[Content],
  J = Table.NestedJoin(A,{"vendor","po","material"},I,{"vendor","po","material"},"Invoice",JoinKind.LeftOuter),
  E = Table.ExpandTableColumn(J,"Invoice",{"invoice_id","qty","amount"},{"invoice_id","invoice_qty","invoice_amount"}),
  D = Table.AddColumn(E,"amount_diff",each [invoice_amount]-[amount],type number),
  S = Table.AddColumn(D,"result",each if [invoice_id]=null then "未到票" else if [amount_diff]=0 and [invoice_qty]=[qty] then "可冲销" else "差异待查")
in S