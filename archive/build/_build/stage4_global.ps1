$ErrorActionPreference='Stop'
. 'c:\Users\bgulick\Downloads\vCIJ_tests\_build\global_lib.ps1'
$out='c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.ScreenUpdating=$false

function Load-Solver($xl){
    try{$xl.Run('SolverReset')|Out-Null;return $true}catch{}
    foreach($a in $xl.AddIns){ if($a.Name -like 'SOLVER*'){ try{$a.Installed=$true}catch{} } }
    try{$xl.Run('SolverReset')|Out-Null;return $true}catch{}
    $paths=@("$env:ProgramFiles\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "$env:ProgramFiles\Microsoft Office\Office16\Library\SOLVER\SOLVER.XLAM")
    foreach($p in $paths){ if(Test-Path $p){ $xl.Workbooks.Open($p)|Out-Null; break } }
    try{$xl.Run('SolverReset')|Out-Null;return $true}catch{return $false}
}
function Solve($xl,$ws,$targetAddr,$byAddr){
    $ws.Activate()
    $t=$ws.Range($targetAddr); $bc=$ws.Range($byAddr)
    $xl.Run('SolverReset')|Out-Null
    $xl.Run('SolverOk',$t,2,0,$bc,1)|Out-Null
    try{$xl.Run('SolverOptions',0.000001,100,0.0001,$false,$false,2,1,1,0.075,$false,$false)|Out-Null}catch{}
    $res=$xl.Run('SolverSolve',$true)
    try{$xl.Run('SolverFinish',1)|Out-Null}catch{}
    return $res
}

try{
    $wb=$xl.Workbooks.Open($out); Load-Solver $xl|Out-Null
    # remove stale global names + sheets
    foreach($nm in @('gC11_0','gC12_0','gC44_0','gC11p','gC12p','gC44p','gK0','gK0p','gC11a','gC12a','gC44a')){
        try{ $wb.Names.Item($nm).Delete() }catch{}
    }
    $existing=@(); foreach($s in $wb.Worksheets){ $existing+=$s.Name }
    foreach($t in @('Global_FS','Global_Cook')){ if($existing -contains $t){ $wb.Worksheets.Item($t).Delete() } }

    $seed=@(240.0,122.0,46.0,5.4,3.1,0.29)
    # place after last Cook sheet
    $anchor=$wb.Worksheets.Item('Cook_NoVp100')

    # ---- Global_FS ----
    $wsF=Build-GlobalFS -wb $wb -afterSheet $anchor -seedParams $seed
    $wb.Application.Calculate()
    $resF=Solve $xl $wsF '$AP$18' '$AP$4:$AP$9,$G$5:$G$25'

    # ---- Global_Cook ----
    $gc=Build-GlobalCook -wb $wb -afterSheet $wsF -seedParams $seed
    $wsC=$gc.ws; $pc=$gc.paramCol
    function ColL2([int]$c){ $s=''; while($c -gt 0){ $m=($c-1)%26; $s=[char](65+$m)+$s; $c=[int](($c-$m-1)/26) }; return $s }
    $PL=ColL2 $pc
    $wb.Application.Calculate()
    $resC=Solve $xl $wsC "`$$PL`$18" "`$$PL`$4:`$$PL`$9"

    Write-Output "Global_FS   solve=$resF"
    $lbls=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
    for($i=0;$i -lt 6;$i++){ Write-Output ("  {0,-6}={1:N4}" -f $lbls[$i],[double]$wsF.Cells.Item(4+$i,42).Value2) }
    Write-Output ("  K0={0:N3} K0'={1:N3} GH0={2:N3} misfit={3:N4}" -f `
      [double]$wsF.Cells.Item(11,42).Value2,[double]$wsF.Cells.Item(12,42).Value2,[double]$wsF.Cells.Item(14,46).Value2,[double]$wsF.Cells.Item(18,42).Value2)

    Write-Output "Global_Cook solve=$resC"
    for($i=0;$i -lt 6;$i++){ Write-Output ("  {0,-6}={1:N4}" -f $lbls[$i],[double]$wsC.Cells.Item(4+$i,$pc).Value2) }
    Write-Output ("  K0={0:N3} K0'={1:N3} misfit={2:N4}" -f `
      [double]$wsC.Cells.Item(11,$pc).Value2,[double]$wsC.Cells.Item(12,$pc).Value2,[double]$wsC.Cells.Item(18,$pc).Value2)

    $wb.Save(); $wb.Close($true)
}catch{ Write-Output "ERROR: $_"; Write-Output $_.ScriptStackTrace }
finally{ $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
