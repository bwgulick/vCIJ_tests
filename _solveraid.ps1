param(
  [string]$Wb  = "C:\Users\bgulick\Downloads\vCIJ_tests\Cij analysis\Vanadium_Cij_Brian_Claude.xlsx",
  [string]$Log = "C:\Users\bgulick\Downloads\vCIJ_tests\_solveraid_log.txt"
)
$ErrorActionPreference = "Stop"

$fits = @(
  @{Sheet="No Vs 111"; Fit="CK"; P=@("AW26","AW27","AW28","AW29","AW30","AW31"); RC=@("AW","AT","AV","AU"); Row=42},
  @{Sheet="No Vs 111"; Fit="FS"; P=@("BA26","BA27","BA28","BA29","BA30","BA31"); RC=@("BJ","BU","BV","BW"); Row=68},
  @{Sheet="No Vp 111"; Fit="CK"; P=@("AY27","AY28","AY29","AY30","AY31","AY32"); RC=@("AY","AV","AX","AW"); Row=42},
  @{Sheet="No Vp 111"; Fit="FS"; P=@("BC27","BC28","BC29","BC30","BC31","BC32"); RC=@("BL","BW","BY","BX"); Row=68},
  @{Sheet="No Vs 100"; Fit="CK"; P=@("AX26","AX27","AX28","AX29","AX30","AX31"); RC=@("AX","AU","AW","AV"); Row=42},
  @{Sheet="No Vs 100"; Fit="FS"; P=@("BB26","BB27","BB28","BB29","BB30","BB31"); RC=@("BK","BW","BY","BX"); Row=68},
  @{Sheet="No Vp 100"; Fit="CK"; P=@("AZ27","AZ28","AZ29","AZ30","AZ31","AZ32"); RC=@("AZ","AW","AY","AX"); Row=42},
  @{Sheet="No Vp 100"; Fit="FS"; P=@("BD27","BD28","BD29","BD30","BD31","BD32"); RC=@("BM","BY","CA","BZ"); Row=68},
  @{Sheet="Global No Vs 111"; Fit="CK"; P=@("AW26","AW27","AW28","AW29","AW30","AW31"); RC=@("AW","AT","AV","AU"); Row=42},
  @{Sheet="Global No Vs 111"; Fit="FS"; P=@("BA26","BA27","BA28","BA29","BA30","BA31"); RC=@("BJ","BU","BV","BW"); Row=68},
  @{Sheet="Global No Vp 111"; Fit="CK"; P=@("AY27","AY28","AY29","AY30","AY31","AY32"); RC=@("AY","AV","AX","AW"); Row=42},
  @{Sheet="Global No Vp 111"; Fit="FS"; P=@("BC27","BC28","BC29","BC30","BC31","BC32"); RC=@("BL","BW","BY","BX"); Row=68},
  @{Sheet="Global No Vs 100"; Fit="CK"; P=@("AX26","AX27","AX28","AX29","AX30","AX31"); RC=@("AX","AU","AW","AV"); Row=42},
  @{Sheet="Global No Vs 100"; Fit="FS"; P=@("BB26","BB27","BB28","BB29","BB30","BB31"); RC=@("BK","BW","BY","BX"); Row=68},
  @{Sheet="Global No Vp 100"; Fit="CK"; P=@("AZ27","AZ28","AZ29","AZ30","AZ31","AZ32"); RC=@("AZ","AW","AY","AX"); Row=42},
  @{Sheet="Global No Vp 100"; Fit="FS"; P=@("BD27","BD28","BD29","BD30","BD31","BD32"); RC=@("BM","BY","CA","BZ"); Row=68}
)
$RR0 = 6; $RR1 = 16
$ANCHORCOL = 200   # 'GR' (matches already-written blocks; re-runs are idempotent)

# invert n x n (flat, row-major); returns flat n x n
function InvertFlat([double[]]$A,[int]$n){
  $w=2*$n
  $M=[double[]]::new($n*$w)
  for($i=0;$i -lt $n;$i++){ for($j=0;$j -lt $n;$j++){ $M[$i*$w+$j]=$A[$i*$n+$j] }; $M[$i*$w+$n+$i]=1.0 }
  for($col=0;$col -lt $n;$col++){
    $piv=$col; $best=[math]::Abs($M[$col*$w+$col])
    for($r=$col+1;$r -lt $n;$r++){ $ab=[math]::Abs($M[$r*$w+$col]); if($ab -gt $best){$best=$ab;$piv=$r} }
    if($best -lt 1e-300){ throw "singular matrix at col $col" }
    if($piv -ne $col){ for($j=0;$j -lt $w;$j++){ $t=$M[$piv*$w+$j];$M[$piv*$w+$j]=$M[$col*$w+$j];$M[$col*$w+$j]=$t } }
    $d=$M[$col*$w+$col]
    for($j=0;$j -lt $w;$j++){ $M[$col*$w+$j]=$M[$col*$w+$j]/$d }
    for($r=0;$r -lt $n;$r++){ if($r -ne $col){ $f=$M[$r*$w+$col]; if($f -ne 0){ for($j=0;$j -lt $w;$j++){ $M[$r*$w+$j]=$M[$r*$w+$j]-$f*$M[$col*$w+$j] } } } }
  }
  $Inv=[double[]]::new($n*$n)
  for($i=0;$i -lt $n;$i++){ for($j=0;$j -lt $n;$j++){ $Inv[$i*$n+$j]=$M[$i*$w+$n+$j] } }
  return ,$Inv
}
# quadratic form g^T S3 g  (S3 flat 3x3)
function QF($g,[double[]]$S3){ $v=0.0; for($i=0;$i -lt 3;$i++){ for($j=0;$j -lt 3;$j++){ $v+=$g[$i]*$S3[$i*3+$j]*$g[$j] } }; return [math]::Sqrt([math]::Max($v,0)) }
# 1-based column index -> letters
function ColLetter([int]$c){ $s=""; while($c -gt 0){ $m=[int](($c-1)%26); $s=[char](65+$m)+$s; $c=[int][math]::Floor(($c-1)/26) }; return $s }

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.ScreenUpdating=$false
$logLines = New-Object System.Collections.ArrayList
try {
  $book = $xl.Workbooks.Open($Wb)
  $xl.Calculation = -4135
  $ws=@{}; foreach($s in $book.Worksheets){ $ws[$s.Name]=$s }

  foreach($cfg in $fits){
    $sh=$ws[$cfg.Sheet]; $p=6
    $resAddr=@(); foreach($c in $cfg.RC){ for($r=$RR0;$r -le $RR1;$r++){ $resAddr += "$c$r" } }
    $N=$resAddr.Count
    $theta=[double[]]::new($p)
    for($k=0;$k -lt $p;$k++){ $theta[$k]=[double]$sh.Range($cfg.P[$k]).Value2 }
    $xl.Calculate()
    $r0=[double[]]::new($N)
    for($i=0;$i -lt $N;$i++){ $v=$sh.Range($resAddr[$i]).Value2; if($v -eq $null){$v=0}; $r0[$i]=[double]$v }
    $SSR=0.0; for($i=0;$i -lt $N;$i++){ $SSR+=$r0[$i]*$r0[$i] }
    $Jac=[double[]]::new($N*$p)
    for($j=0;$j -lt $p;$j++){
      $t=$theta[$j]; $d=1e-4*[math]::Max([math]::Abs($t),1.0)
      $sh.Range($cfg.P[$j]).Value2=$t+$d; $xl.Calculate()
      $rp=[double[]]::new($N)
      for($i=0;$i -lt $N;$i++){ $v=$sh.Range($resAddr[$i]).Value2; if($v -eq $null){$v=0}; $rp[$i]=[double]$v }
      $sh.Range($cfg.P[$j]).Value2=$t-$d; $xl.Calculate()
      $rm=[double[]]::new($N)
      for($i=0;$i -lt $N;$i++){ $v=$sh.Range($resAddr[$i]).Value2; if($v -eq $null){$v=0}; $rm[$i]=[double]$v }
      $sh.Range($cfg.P[$j]).Value2=$t
      for($i=0;$i -lt $N;$i++){ $Jac[$i*$p+$j]=($rp[$i]-$rm[$i])/(2.0*$d) }
    }
    $xl.Calculate()
    $JtJ=[double[]]::new($p*$p)
    for($a=0;$a -lt $p;$a++){ for($b=0;$b -lt $p;$b++){ $s=0.0; for($i=0;$i -lt $N;$i++){ $s+=$Jac[$i*$p+$a]*$Jac[$i*$p+$b] }; $JtJ[$a*$p+$b]=$s } }
    $inv=InvertFlat $JtJ $p
    $dof=$N-$p; $s2=$SSR/$dof; $rms=[math]::Sqrt($SSR/$N)
    $Cov=[double[]]::new($p*$p); $SE=[double[]]::new($p)
    for($a=0;$a -lt $p;$a++){ for($b=0;$b -lt $p;$b++){ $Cov[$a*$p+$b]=$s2*$inv[$a*$p+$b] }; $SE[$a]=[math]::Sqrt([math]::Max($Cov[$a*$p+$a],0)) }
    $C11=$theta[0]; $C12=$theta[1]; $C44=$theta[2]
    $S3=[double[]]::new(9)
    for($a=0;$a -lt 3;$a++){ for($b=0;$b -lt 3;$b++){ $S3[$a*3+$b]=$Cov[$a*$p+$b] } }
    $Dif=$C11-$C12; $den=3.0*$Dif+4.0*$C44; $den2=$den*$den
    $gK=@([double](1.0/3.0), [double](2.0/3.0), [double]0.0)
    $gGV=@([double]0.2, [double](-0.2), [double]0.6)
    $gr0=20.0*$C44*$C44/$den2; $gr2=15.0*$Dif*$Dif/$den2
    $gGR=@([double]$gr0, [double](-$gr0), [double]$gr2)
    $gGH=@([double](($gGV[0]+$gGR[0])/2.0), [double](($gGV[1]+$gGR[1])/2.0), [double](($gGV[2]+$gGR[2])/2.0))
    $sK=QF $gK $S3; $sGV=QF $gGV $S3; $sGR=QF $gGR $S3; $sGH=QF $gGH $S3

    # write per-cell: text via .Formula, numbers via .Value2 (separate COM properties -> no binder clash)
    $r0row=[int]$cfg.Row; $c0=[int]$ANCHORCOL
    $L0=ColLetter $c0; $L1=ColLetter ($c0+1); $L2=ColLetter ($c0+2)
    $labs=@("C11_0","C12_0","C44_0","C11'","C12'","C44'")
    $sh.Range("$L0$r0row").Formula="$($cfg.Fit) FIT COVARIANCE (SolverAid) - valid at current optimum"
    $hr=$r0row+1
    $sh.Range("$L0$hr").Formula="param"; $sh.Range("$L1$hr").Formula="value"; $sh.Range("$L2$hr").Formula="SE(1sigma)"
    for($k=0;$k -lt 6;$k++){ $rr=$r0row+2+$k; $sh.Range("$L0$rr").Formula=$labs[$k]; $sh.Range("$L1$rr").Value2=[double]$theta[$k]; $sh.Range("$L2$rr").Value2=[double]$SE[$k] }
    $rf=$r0row+8
    $sh.Range("$L0$rf").Formula="s^2";       $sh.Range("$L1$rf").Value2=[double]$s2
    $r=$rf+1; $sh.Range("$L0$r").Formula="RMS res";    $sh.Range("$L1$r").Value2=[double]$rms
    $r=$rf+2; $sh.Range("$L0$r").Formula="dof";        $sh.Range("$L1$r").Value2=[double]$dof
    $r=$rf+3; $sh.Range("$L0$r").Formula="N_res";      $sh.Range("$L1$r").Value2=[double]$N
    $r=$rf+4; $sh.Range("$L0$r").Formula="sigma(K0)";  $sh.Range("$L1$r").Value2=[double]$sK
    $r=$rf+5; $sh.Range("$L0$r").Formula="sigma(GV0)"; $sh.Range("$L1$r").Value2=[double]$sGV
    $r=$rf+6; $sh.Range("$L0$r").Formula="sigma(GR0)"; $sh.Range("$L1$r").Value2=[double]$sGR
    $r=$rf+7; $sh.Range("$L0$r").Formula="sigma(GH0)"; $sh.Range("$L1$r").Value2=[double]$sGH
    $mrow=$rf+9
    $sh.Range("$L0$mrow").Formula="Cov(theta) 6x6:"
    for($a=0;$a -lt 6;$a++){ $rr=$mrow+1+$a; for($b=0;$b -lt 6;$b++){ $cl=ColLetter ($c0+$b); $sh.Range("$cl$rr").Value2=[double]$Cov[$a*$p+$b] } }
    $nr=$mrow+8; $sh.Range("$L0$nr").Formula="NOTE: SEs valid at current converged params; recompute after any re-fit."

    $ln=("{0,-11} {1}: SSR={2:E4} dof={3} RMS={4:F5} | SE C11={5:F4} C12={6:F4} C44={7:F4} | sK={8:F4} sGV={9:F4} sGR={10:F4} sGH={11:F4}" -f $cfg.Sheet,$cfg.Fit,$SSR,$dof,$rms,$SE[0],$SE[1],$SE[2],$sK,$sGV,$sGR,$sGH)
    [void]$logLines.Add($ln); Write-Output $ln
  }
  $xl.Calculation=-4105; $book.Application.CalculateFull(); $book.Save(); $book.Close($true)
  [System.IO.File]::WriteAllLines($Log,$logLines)
}
finally {
  $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
