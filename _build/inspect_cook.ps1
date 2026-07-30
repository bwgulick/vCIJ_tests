$ErrorActionPreference='Stop'
$out='c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try{
  $wb=$xl.Workbooks.Open($out)
  $ws=$wb.Worksheets.Item('Cook_NoVs111')
  Write-Output "row oil    P     f       Olen    C11m   C11M   C12m   C12M   C44m   C44M    dP"
  for($r=4;$r -le 25;$r++){
    Write-Output ("{0,2} {1,4} {2,6:N3} {3,7:N5} {4,7:N5} {5,7:N2} {6,7:N2} {7,7:N2} {8,7:N2} {9,6:N2} {10,6:N2} {11,7:N3}" -f `
      $r,[int]$ws.Cells.Item($r,1).Value2,[double]$ws.Cells.Item($r,2).Value2,[double]$ws.Cells.Item($r,18).Value2,
      [double]$ws.Cells.Item($r,15).Value2,
      [double]$ws.Cells.Item($r,25).Value2,[double]$ws.Cells.Item($r,29).Value2,
      [double]$ws.Cells.Item($r,26).Value2,[double]$ws.Cells.Item($r,30).Value2,
      [double]$ws.Cells.Item($r,27).Value2,[double]$ws.Cells.Item($r,31).Value2,
      [double]$ws.Cells.Item($r,32).Value2)
  }
  Write-Output "params:"
  $lbls=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
  for($i=0;$i -lt 6;$i++){ Write-Output ("  {0}={1:N4}" -f $lbls[$i],[double]$ws.Cells.Item(4+$i,42).Value2) }
  Write-Output ("  K0={0:N3} K0'={1:N3} misfit={2:N3}" -f [double]$ws.Cells.Item(11,42).Value2,[double]$ws.Cells.Item(12,42).Value2,[double]$ws.Cells.Item(18,42).Value2)
  $wb.Close($false)
}catch{Write-Output "ERR $_"}finally{$xl.Quit();[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null;[GC]::Collect()}
