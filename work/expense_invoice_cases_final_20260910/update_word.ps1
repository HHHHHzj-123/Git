$ErrorActionPreference = 'Stop'
$docx = (Resolve-Path -LiteralPath "$PSScriptRoot\deliverables\企业发票、报销与异常支出实战手册.docx").Path
$pdf = Join-Path $PSScriptRoot 'qa\企业发票、报销与异常支出实战手册.pdf'
$word = New-Object -ComObject Kwps.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  Start-Sleep -Seconds 2
  $doc = $word.Documents.Open($docx)
  foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
  foreach ($field in $doc.Fields) { $field.Update() | Out-Null }
  $doc.Save()
  $doc.ExportAsFixedFormat($pdf, 17)
  $pages = $doc.ComputeStatistics(2)
  $words = $doc.ComputeStatistics(0)
  Write-Output "pages=$pages words=$words pdf=$pdf"
  $doc.Close()
} finally {
  $word.Quit()
}
