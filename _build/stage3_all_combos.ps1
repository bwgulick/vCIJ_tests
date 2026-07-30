$ErrorActionPreference = 'Stop'
. 'c:\Users\bgulick\Downloads\vCIJ_tests\_build\fs_lib.ps1'
. 'c:\Users\bgulick\Downloads\vCIJ_tests\_build\cook_lib.ps1'
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

# FS solve: fits params (AI4:AI9) + lengths (G5:G25); objective AI18
function Solve-FS($xl,$ws){
    $ws.Activate()
    $target=$ws.Range('$AI$18'); $bychange=$ws.Range('$AI$4:$AI$9,$G$5:$G$25')
    $xl.Run('SolverReset')|Out-Null
    $xl.Run('SolverOk',$target,2,0,$bychange,1)|Out-Null
    try{$xl.Run('SolverOptions',0.000001,100,0.0001,$false,$false,2,1,1,0.075,$false,$false)|Out-Null}catch{}
    $res=$xl.Run('SolverSolve',$true)
    try{$xl.Run('SolverFinish',1)|Out-Null}catch{}
    return $res
}
# Cook solve: fits params only (AP4:AP9); objective AP18
function Solve-Cook($xl,$ws){
    $ws.Activate()
    $target=$ws.Range('$AP$18'); $bychange=$ws.Range('$AP$4:$AP$9')
    $xl.Run('SolverReset')|Out-Null
    $xl.Run('SolverOk',$target,2,0,$bychange,1)|Out-Null
    try{$xl.Run('SolverOptions',0.000001,100,0.0001,$false,$false,2,1,1,0.075,$false,$false)|Out-Null}catch{}
    $res=$xl.Run('SolverSolve',$true)
    try{$xl.Run('SolverFinish',1)|Out-Null}catch{}
    return $res
}

try {
    $wb = $xl.Workbooks.Open($out)
    Load-Solver $xl | Out-Null
    $align = $wb.Worksheets.Item('Align')
    $inp   = $wb.Worksheets.Item('Inputs')
    $inp.Range('M23').Value2 = [double]120   # CutoffLowBar = 120 (source window)

    # robust delete of any pre-existing combo sheets
    $targets = @('FS_NoVs111','FS_NoVp111','FS_NoVs100','FS_NoVp100',
                 'Cook_NoVs111','Cook_NoVp111','Cook_NoVs100','Cook_NoVp100')
    $existing = @(); foreach($s in $wb.Worksheets){ $existing += $s.Name }
    foreach($t in $targets){ if($existing -contains $t){ $wb.Worksheets.Item($t).Delete() } }

    $seed = @(240.0,122.0,46.0,5.4,3.1,0.29)

    # combo definitions: name, measured-Cij templates for FS (cols R/S/T from L,M,P,Q,J)
    #  measC11, measC12, measC44
    $FScombos = @(
      @{name='FS_NoVs111'; c11='=J{r}*L{r}^2';         c12='=1.5*P{r}-0.5*R{r}-2*T{r}'; c44='=J{r}*M{r}^2'},
      @{name='FS_NoVp111'; c11='=J{r}*L{r}^2';         c12='=R{r}+T{r}-3*Q{r}';         c44='=J{r}*M{r}^2'},
      @{name='FS_NoVs100'; c11='=J{r}*L{r}^2';         c12='=R{r}+T{r}-3*Q{r}';         c44='=Q{r}+0.5*P{r}-0.5*R{r}'},
      @{name='FS_NoVp100'; c11='=3*Q{r}+S{r}-T{r}';    c12='=P{r}-Q{r}-T{r}';           c44='=J{r}*M{r}^2'}
    )
    # Cook combos: vcombo2 (from ref vel G/H/I/J), measured-Cij (cols Y/Z/AA from Q,S,T,W,X)
    $Cookcombos = @(
      @{name='Cook_NoVs111'; v='=I{r}^2-(4/3)*H{r}^2';                  c11='=Q{r}*S{r}^2';      c12='=1.5*W{r}-0.5*Y{r}-2*AA{r}'; c44='=Q{r}*T{r}^2'},
      @{name='Cook_NoVp111'; v='=G{r}^2+(2/3)*H{r}^2-2*J{r}^2';         c11='=Q{r}*S{r}^2';      c12='=Y{r}+AA{r}-3*X{r}';         c44='=Q{r}*T{r}^2'},
      @{name='Cook_NoVs100'; v='=(1/3)*I{r}^2-(4/3)*J{r}^2+(2/3)*G{r}^2';c11='=Q{r}*S{r}^2';     c12='=Y{r}+AA{r}-3*X{r}';         c44='=X{r}+0.5*W{r}-0.5*Y{r}'},
      @{name='Cook_NoVp100'; v='=I{r}^2-(4/3)*H{r}^2';                  c11='=3*X{r}+Z{r}-AA{r}';c12='=W{r}-X{r}-AA{r}';           c44='=Q{r}*T{r}^2'}
    )

    $results = @()

    # ---- build + solve FS sheets ----
    $prev = $align
    foreach ($cb in $FScombos) {
        $ws = Build-FSCombo -wb $wb -afterSheet $prev -name $cb.name -measC11 $cb.c11 -measC12 $cb.c12 -measC44 $cb.c44 -seedParams $seed
        $prev = $ws
        $wb.Application.Calculate()
        $res = Solve-FS $xl $ws
        $row = [ordered]@{ name=$cb.name; res=$res }
        $lbls=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
        for($i=0;$i -lt 6;$i++){ $row[$lbls[$i]]=[double]$ws.Cells.Item(4+$i,35).Value2 }
        $row['K0']=[double]$ws.Cells.Item(11,35).Value2
        $row['GH0']=[double]$ws.Cells.Item(4,39).Value2
        $row['misfit']=[double]$ws.Cells.Item(18,35).Value2
        $results += [pscustomobject]$row
        $wb.Save()
    }

    # ---- build + solve Cook sheets ----
    foreach ($cb in $Cookcombos) {
        $ws = Build-CookCombo -wb $wb -afterSheet $prev -name $cb.name -vcombo2 $cb.v -measC11 $cb.c11 -measC12 $cb.c12 -measC44 $cb.c44 -seedParams $seed
        $prev = $ws
        $wb.Application.Calculate()
        $res = Solve-Cook $xl $ws
        $row = [ordered]@{ name=$cb.name; res=$res }
        $lbls=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
        for($i=0;$i -lt 6;$i++){ $row[$lbls[$i]]=[double]$ws.Cells.Item(4+$i,42).Value2 }
        $row['K0']=[double]$ws.Cells.Item(11,42).Value2
        $row['GH0']=[double]$ws.Cells.Item(4,40).Value2
        $row['misfit']=[double]$ws.Cells.Item(18,42).Value2
        $results += [pscustomobject]$row
        $wb.Save()
    }

    Write-Output "========== RESULTS (CutoffLow=120, High=420) =========="
    foreach($r in $results){
        Write-Output ("{0,-13} res={1}  C11={2,7:N2} C12={3,7:N2} C44={4,6:N2}  C11'={5,5:N3} C12'={6,5:N3} C44'={7,5:N3}  K0={8,6:N2} GH0={9,6:N2}  mis={10:N3}" -f `
          $r.name,$r.res,$r.C11_0,$r.C12_0,$r.C44_0,$r."C11'",$r."C12'",$r."C44'",$r.K0,$r.GH0,$r.misfit)
    }

    $wb.Save(); $wb.Close($true)
} catch {
    Write-Output "ERROR: $_"
    Write-Output $_.ScriptStackTrace
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
