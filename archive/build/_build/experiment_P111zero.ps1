$ErrorActionPreference='Stop'
$src='c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'
$dst='c:\Users\bgulick\Downloads\vCIJ_tests\_test_P111zero.xlsx'
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Copy-Item $src $dst -Force

$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.ScreenUpdating=$false
function Load-Solver($xl){
    try{$xl.Run('SolverReset')|Out-Null;return}catch{}
    foreach($a in $xl.AddIns){ if($a.Name -like 'SOLVER*'){ try{$a.Installed=$true}catch{} } }
    try{$xl.Run('SolverReset')|Out-Null;return}catch{}
    $paths=@("$env:ProgramFiles\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM",
      "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\Library\SOLVER\SOLVER.XLAM")
    foreach($p in $paths){ if(Test-Path $p){ $xl.Workbooks.Open($p)|Out-Null; break } }
    try{$xl.Run('SolverReset')|Out-Null}catch{}
}
function SolveSheet($xl,$ws,$target,$by){
    $ws.Activate()
    $xl.Run('SolverReset')|Out-Null
    $xl.Run('SolverOk',$ws.Range($target),2,0,$ws.Range($by),1)|Out-Null
    try{$xl.Run('SolverOptions',0.000001,100,0.0001,$false,$false,2,1,1,0.075,$false,$false)|Out-Null}catch{}
    $r=$xl.Run('SolverSolve',$true); try{$xl.Run('SolverFinish',1)|Out-Null}catch{}; return $r
}
try{
    $wb=$xl.Workbooks.Open($dst); Load-Solver $xl
    # ---- THE CHANGE: force P111 at oil=0 to exactly 0 GPa ----
    $wb.Worksheets.Item('Align').Range('E4').Value2=[double]0
    $wb.Application.Calculate()

    $models=@(
      @('FS_NoVs111','$AI$18','$AI$4:$AI$9,$G$5:$G$25'),
      @('FS_NoVp111','$AI$18','$AI$4:$AI$9,$G$5:$G$25'),
      @('FS_NoVs100','$AI$18','$AI$4:$AI$9,$G$5:$G$25'),
      @('FS_NoVp100','$AI$18','$AI$4:$AI$9,$G$5:$G$25'),
      @('Cook_NoVs111','$AP$18','$AP$4:$AP$9'),
      @('Cook_NoVp111','$AP$18','$AP$4:$AP$9'),
      @('Cook_NoVs100','$AP$18','$AP$4:$AP$9'),
      @('Cook_NoVp100','$AP$18','$AP$4:$AP$9'),
      @('Global_FS','$AP$18','$AP$4:$AP$9,$G$5:$G$25'),
      @('Global_Cook','$CW$18','$CW$4:$CW$9')
    )
    foreach($m in $models){ $ws=$wb.Worksheets.Item($m[0]); $r=SolveSheet $xl $ws $m[1] $m[2]; Write-Output ("solved {0,-13} code={1}" -f $m[0],$r) }
    $wb.Save()

    $sm=$wb.Worksheets.Item('Summary'); $wb.Application.Calculate()
    Write-Output ""
    Write-Output "==== SUMMARY  with P111(oil0)=0  ===="
    $cols=@('nVs111FS','nVs111CK','nVp111FS','nVp111CK','nVs100FS','nVs100CK','nVp100FS','nVp100CK','GLOB_FS','GLOB_CK')
    $mets=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'",'K0',"K0'",'G_Hill','misfit')
    $h='metric     '; foreach($c in $cols){ $h+=('{0,10}' -f $c) }; Write-Output $h
    for($m=0;$m -lt 10;$m++){ $s=('{0,-11}' -f $mets[$m]); for($f=0;$f -lt 10;$f++){ $s+=('{0,10:N3}' -f [double]$sm.Cells.Item(4+$m,2+$f).Value2) }; Write-Output $s }
    $wb.Close($true)
}catch{ Write-Output "ERR $_"; Write-Output $_.ScriptStackTrace }
finally{ $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
