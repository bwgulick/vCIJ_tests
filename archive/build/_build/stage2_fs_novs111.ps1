$ErrorActionPreference = 'Stop'
. 'c:\Users\bgulick\Downloads\vCIJ_tests\_build\fs_lib.ps1'
$out = 'c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

function Load-Solver($xl) {
    try { $xl.Run('SolverReset') | Out-Null; return $true } catch {}
    foreach ($a in $xl.AddIns) { if ($a.Name -like 'SOLVER*') { try { $a.Installed = $true } catch {} } }
    try { $xl.Run('SolverReset') | Out-Null; return $true } catch {}
    $paths = @(
      "$env:ProgramFiles\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "$env:ProgramFiles\Microsoft Office\Office16\Library\SOLVER\SOLVER.XLAM"
    )
    foreach ($p in $paths) { if (Test-Path $p) { $xl.Workbooks.Open($p) | Out-Null; break } }
    try { $xl.Run('SolverReset') | Out-Null; return $true } catch { return $false }
}

try {
    $wb = $xl.Workbooks.Open($out)
    $align = $wb.Worksheets.Item('Align')

    # remove existing FS_NoVs111 if present (idempotent)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq 'FS_NoVs111') { $s.Delete() } }

    $seed = @(240.0,122.0,46.0,5.4,3.1,0.29)
    $ws = Build-FSCombo -wb $wb -afterSheet $align -name 'FS_NoVs111' `
        -measC11 '=J{r}*L{r}^2' `
        -measC12 '=1.5*P{r}-0.5*R{r}-2*T{r}' `
        -measC44 '=J{r}*M{r}^2' `
        -seedParams $seed

    $wb.Save()

    # ----- run Solver via COM -----
    $solverOK = Load-Solver $xl
    Write-Output "Solver loaded: $solverOK"
    $preObj = [double]$ws.Cells.Item(18,35).Value2
    Write-Output ("pre-solve misfit = {0:N4}" -f $preObj)

    if ($solverOK) {
        $ws.Activate()
        $target  = $ws.Range('$AI$18')
        $bychange= $ws.Range('$AI$4:$AI$9,$G$5:$G$25')
        $xl.Run('SolverReset') | Out-Null
        # SolverOk(SetCell, MaxMinVal=2 min, ValueOf=0, ByChange, Engine=1 GRG)
        $xl.Run('SolverOk', $target, 2, 0, $bychange, 1) | Out-Null
        # SolverOptions: set convergence + central derivatives (Precision, ...); use SolverOptions for GRG
        try { $xl.Run('SolverOptions', 0.000001, 100, 0.0001, $false, $false, 2, 1, 1, 0.075, $false, $false) | Out-Null } catch { Write-Output "SolverOptions warn: $_" }
        $res = $xl.Run('SolverSolve', $true)
        Write-Output "SolverSolve result code: $res"
        try { $xl.Run('SolverFinish', 1) | Out-Null } catch {}
    }

    # read results
    $lbls = @('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
    Write-Output "----- FS_NoVs111 fitted params -----"
    for ($i=0;$i -lt 6;$i++){ Write-Output ("{0,-6} = {1:N4}" -f $lbls[$i], [double]$ws.Cells.Item(4+$i,35).Value2) }
    Write-Output ("K0     = {0:N4}" -f [double]$ws.Cells.Item(11,35).Value2)
    Write-Output ("K0'    = {0:N4}" -f [double]$ws.Cells.Item(12,35).Value2)
    Write-Output ("misfit = {0:N4}" -f [double]$ws.Cells.Item(18,35).Value2)
    Write-Output ("GH0    = {0:N4}" -f [double]$ws.Cells.Item(4,39).Value2)
    Write-Output "----- fitted lengths G4:G25 -----"
    $ls=@(); for($r=4;$r -le 25;$r++){ $ls += ("{0:N4}" -f [double]$ws.Cells.Item($r,7).Value2) }
    Write-Output ($ls -join ' ')
    Write-Output "----- sample rows: P, C11meas, C11mod, C44meas, C44mod -----"
    foreach($r in 4,9,14,18,25){
        Write-Output ("P={0,6:N3}  C11m={1,7:N2} C11M={2,7:N2}  C44m={3,6:N2} C44M={4,6:N2}  C12m={5,7:N2} C12M={6,7:N2}" -f `
            [double]$ws.Cells.Item($r,2).Value2,[double]$ws.Cells.Item($r,18).Value2,[double]$ws.Cells.Item($r,22).Value2,
            [double]$ws.Cells.Item($r,20).Value2,[double]$ws.Cells.Item($r,24).Value2,
            [double]$ws.Cells.Item($r,19).Value2,[double]$ws.Cells.Item($r,23).Value2)
    }

    $wb.Save()
    $wb.Close($true)
} catch {
    Write-Output "ERROR: $_"
    Write-Output $_.ScriptStackTrace
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
