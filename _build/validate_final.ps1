$ErrorActionPreference='Stop'
$out='c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try{
  $wb=$xl.Workbooks.Open($out)
  Write-Output ("ReadOnly? {0}" -f $wb.ReadOnly)
  Write-Output "--- Sheets (order) ---"
  foreach($s in $wb.Worksheets){ Write-Output ("  "+$s.Name) }

  # Solver models persist as hidden defined names solver_* scoped to each sheet
  Write-Output "--- Solver models present (per fit sheet) ---"
  $fitSheets=@('FS_NoVs111','FS_NoVp111','FS_NoVs100','FS_NoVp100',
               'Cook_NoVs111','Cook_NoVp111','Cook_NoVs100','Cook_NoVp100','Global_FS','Global_Cook')
  foreach($fs in $fitSheets){
    $cnt=0
    foreach($nm in $wb.Names){ if($nm.Name -like "*!solver_*" -and $nm.Name -like "*$fs*"){ $cnt++ } }
    Write-Output ("  {0,-13} solver names: {1}" -f $fs,$cnt)
  }

  # Check Summary has no error cells and literature populated
  $sm=$wb.Worksheets.Item('Summary')
  $errs=0
  for($r=4;$r -le 13;$r++){ for($c=2;$c -le 12;$c++){
    $v=$sm.Cells.Item($r,$c).Value2
    if($v -is [int] -and $v -lt -2000000000){ $errs++; Write-Output ("  ERROR cell R{0}C{1}={2}" -f $r,$c,$v) }
  }}
  Write-Output ("--- Summary error cells: {0} ---" -f $errs)
  Write-Output ("Katahara K0 (M10) = {0:N2}" -f [double]$sm.Range('M10').Value2)
  Write-Output ("Lenkkeri K0 (N10) = {0:N2}" -f [double]$sm.Range('N10').Value2)
  Write-Output ("Greiner  K0 (O10) = {0:N2}" -f [double]$sm.Range('O10').Value2)

  $wb.Close($false)
}catch{ Write-Output "ERR $_"; Write-Output $_.ScriptStackTrace }
finally{ $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
