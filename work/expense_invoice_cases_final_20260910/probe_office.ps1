$ids = @('Word.Application','Kwps.Application','Wps.Application')
foreach ($id in $ids) {
  try { $a=New-Object -ComObject $id; if($null -ne $a){ Write-Output "$id OK"; $a.Quit() } }
  catch { Write-Output "$id FAIL $($_.Exception.Message)" }
}
