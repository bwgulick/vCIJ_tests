$ErrorActionPreference = 'Stop'
$out = 'c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.ScreenUpdating=$false
function Load-Solver($xl) {
    try { $xl.Run('SolverReset') | Out-Null; return $true } catch {}
    foreach ($a in $xl.AddIns) { if ($a.Name -like 'SOLVER*') { try { $a.Installed = $true } catch {} } }
    try { $xl.Run('SolverReset') | Out-Null; return $true } catch {}
    $paths = @("$env:ProgramFiles\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "$env:ProgramFiles\Microsoft Office\Office16\Library\SOLVER\SOLVER.XLAM")
    foreach ($p in $paths) { if (Test-Path $p) { $xl.Workbooks.Open($p) | Out-Null; break } }
    try { $xl.Run('SolverReset') | Out-Null; return $true } catch { return $false }
}
try {
    $wb = $xl.Workbooks.Open($out)
    Load-Solver $xl | Out-Null
    $inp = $wb.Worksheets.Item('Inputs')
    $ws  = $wb.Worksheets.Item('FS_NoVs111')
    # set CutoffLowBar = 120
    $inp.Range('M23').Value2 = [double]120
    # reseed lengths to near 0.704 so solver restarts cleanly
    for($r=5;$r -le 25;$r++){ $oil=30*($r-4); $ws.Cells.Item($r,7).Value2=[double](0.704-($oil/630.0)*0.015) }
    # reseed params
    $sp=@(240.0,122.0,46.0,5.4,3.1,0.29); for($i=0;$i -lt 6;$i++){ $ws.Cells.Item(4+$i,35).Value2=[double]$sp[$i] }
    $wb.Application.Calculate()
    $ws.Activate()
    $target=$ws.Range('$AI$18'); $bychange=$ws.Range('$AI$4:$AI$9,$G$5:$G$25')
    $xl.Run('SolverReset')|Out-Null
    $xl.Run('SolverOk',$target,2,0,$bychange,1)|Out-Null
    try{$xl.Run('SolverOptions',0.000001,100,0.0001,$false,$false,2,1,1,0.075,$false,$false)|Out-Null}catch{}
    $res=$xl.Run('SolverSolve',$true); Write-Output "result=$res"
    try{$xl.Run('SolverFinish',1)|Out-Null}catch{}
    $lbls=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
    Write-Output "--- CutoffLow=120 (oil 120-420, matches source window) ---"
    for($i=0;$i -lt 6;$i++){ Write-Output ("{0,-6} = {1:N4}   (src {2})" -f $lbls[$i],[double]$ws.Cells.Item(4+$i,35).Value2, @('240.69','118.56','46.81','5.46','3.33','0.25')[$i]) }
    Write-Output ("K0     = {0:N4}   (src 159.3)" -f [double]$ws.Cells.Item(11,35).Value2)
    Write-Output ("misfit = {0:N4}   (src 2.23)" -f [double]$ws.Cells.Item(18,35).Value2)
    $wb.Save(); $wb.Close($true)
} catch { Write-Output "ERR $_"; Write-Output $_.ScriptStackTrace }
finally { $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
