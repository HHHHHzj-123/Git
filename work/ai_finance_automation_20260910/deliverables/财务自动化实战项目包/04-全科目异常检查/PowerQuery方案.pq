let
  TB = Excel.CurrentWorkbook(){[Name="TBTable"]}[Content],
  Rules = Excel.CurrentWorkbook(){[Name="RuleTable"]}[Content],
  Prefix = Table.AddColumn(TB,"prefix4",each Text.Start([account_code],4)),
  Join = Table.NestedJoin(Prefix,{"prefix4"},Rules,{"account_prefix"},"Rule",JoinKind.LeftOuter),
  Expand = Table.ExpandTableColumn(Join,"Rule",{"rule_id","condition","threshold","rule_name","action"},{"rule_id","condition","threshold","rule_name","action"}),
  Flag = Table.AddColumn(Expand,"is_exception",each if [condition]="贷" and [direction]="贷" then true else if [condition]="借" and [direction]="借" then true else if [condition]="非零" and Number.Abs([balance])>[threshold] then true else false)
in Flag