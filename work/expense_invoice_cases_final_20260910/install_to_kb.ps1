$ErrorActionPreference='Stop'
$src=$PSScriptRoot
$kb='C:\Users\HZJ\Desktop\财务管理实操'
$dest=Join-Path $kb '09-发票与税务\02-企业发票报销与异常支出\01-正式版'
New-Item -ItemType Directory -Force -Path $dest | Out-Null
$files=@('00-阅读说明.md','deliverables\企业发票、报销与异常支出实战手册.docx','deliverables\企业异常费用与发票问题.xmind','deliverables\异常发票与费用处理速查表.xlsx')
foreach($f in $files){Copy-Item -LiteralPath (Join-Path $src $f) -Destination $dest -Force}
$nav=Join-Path $kb '00-知识库与学习导航\README.md'
$line='- [M09-002 企业发票、报销与异常支出](../09-发票与税务/02-企业发票报销与异常支出/01-正式版/00-阅读说明.md)：141个异常业务案例、拒不整改后果、责任主体、老板指令SOP、沟通模板及法规索引。'
if(-not (Select-String -LiteralPath $nav -SimpleMatch 'M09-002 企业发票、报销与异常支出' -Quiet)){Add-Content -LiteralPath $nav -Value "`r`n$line" -Encoding utf8}
$ledger=Join-Path $kb '00-知识库与学习导航\02-学习与问题台账.md'
$record='| 2026-09-10 | 建立M09-002企业发票、报销与异常支出专题 | 141个案例；286页Word；3张XMind工作表；13页签Excel；新增拒不整改后果、责任主体、老板指令SOP和20条沟通模板 | 政策核验截至2026-09-10；地区和个案需按行为发生日、完整事实及主管机关口径复核；刑事责任仅在满足法定构成时成立 |'
if(-not (Select-String -LiteralPath $ledger -SimpleMatch '建立M09-002企业发票、报销与异常支出专题' -Quiet)){Add-Content -LiteralPath $ledger -Value "`r`n$record" -Encoding utf8}
Get-ChildItem -LiteralPath $dest | Select-Object Name,Length,LastWriteTime
