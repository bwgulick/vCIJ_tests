# Reusable Cook combo-sheet builder.
# Cook length is analytic (trapezoidal integral of per-combo integrand); only 6 Cij params are fitted.
# Column map (data rows 4..25):
#  A oil B P C tP100 D tS100 E tP111 F tS111
#  G Vp100r H Vs100r I Vp111r J Vs111r   (reference-length velocities, use Lfin)
#  K vcombo2  L integrand=4*Lfin100^2/K
#  M trap-incr  N Integral(cumul)  O Cook-length  P L111  Q rho  R f
#  S Vp100 T Vs100 U Vp111 V Vs111  W rVp111^2 X rVs111^2
#  Y C11meas Z C12meas AA C44meas
#  AB P_FS AC C11mod AD C12mod AE C44mod
#  AF dP AG dC11 AH dC12 AI dC44 AJ gated-sq
#  AK Kbulk AL GV AM GR AN GH
# Param block col AP(42): 4..9 params; 11 K0;12 K0';14..16 a-coefs;18 objective

function Build-CookCombo {
    param(
        $wb, $afterSheet, [string]$name,
        [string]$vcombo2,                                   # template {r}: uses G,H,I,J
        [string]$measC11, [string]$measC12, [string]$measC44,  # templates {r}: uses Q,S,T,W,X and each other
        [int]$n = 22, [double[]]$seedParams
    )
    $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $afterSheet)
    $ws.Name = $name

    $hdr = @('Oil bar','P (GPa)','tP100','tS100','tP111','tS111',
             'Vp100r','Vs100r','Vp111r','Vs111r','vcombo2','integrand',
             'trap','Integral','Cook L','L111','rho','f',
             'Vp100','Vs100','Vp111','Vs111','rVp111^2','rVs111^2',
             'C11 exp','C12 exp','C44 exp',
             'P_FS','C11 mod','C12 mod','C44 mod',
             'dP','dC11','dC12','dC44','gated sq',
             'K bulk','GV','GR','GH')
    for ($c=0;$c -lt $hdr.Count;$c++){ $ws.Cells.Item(3,1+$c).Value2 = $hdr[$c] }

    for ($i=0;$i -lt $n;$i++){
        $r = 4 + $i
        $ws.Cells.Item($r,1).Formula  = "=Align!A$r"
        $ws.Cells.Item($r,2).Formula  = "=Align!B$r"
        $ws.Cells.Item($r,3).Formula  = "=Align!C$r"
        $ws.Cells.Item($r,4).Formula  = "=Align!D$r"
        $ws.Cells.Item($r,5).Formula  = "=Align!F$r"
        $ws.Cells.Item($r,6).Formula  = "=Align!G$r"
        $ws.Cells.Item($r,7).Formula  = "=2*Lfin_100/C$r"    # G Vp100r
        $ws.Cells.Item($r,8).Formula  = "=2*Lfin_100/D$r"    # H Vs100r
        $ws.Cells.Item($r,9).Formula  = "=2*Lfin_111/E$r"    # I Vp111r
        $ws.Cells.Item($r,10).Formula = "=2*Lfin_111/F$r"    # J Vs111r
        $ws.Cells.Item($r,11).Formula = ($vcombo2 -replace '\{r\}',"$r")   # K vcombo2
        $ws.Cells.Item($r,12).Formula = "=4*Lfin_100^2/K$r"  # L integrand
        if ($r -eq 4) {
            $ws.Cells.Item($r,13).Value2  = [double]0        # M trap (none at first row)
            $ws.Cells.Item($r,14).Value2  = [double]0        # N Integral
        } else {
            $ws.Cells.Item($r,13).Formula = "=0.5*(B$r-B$($r-1))*(L$r+L$($r-1))"  # M trap
            $ws.Cells.Item($r,14).Formula = "=N$($r-1)+M$r"                        # N Integral cumul
        }
        $ws.Cells.Item($r,15).Formula = "=Lfin_100/(1+N$r*scal_100)"  # O Cook length
        $ws.Cells.Item($r,16).Formula = "=O$r*Lfin_111/Lfin_100"      # P L111
        $ws.Cells.Item($r,17).Formula = "=rho_100*(Lfin_100/O$r)^3"   # Q rho
        $ws.Cells.Item($r,18).Formula = "=0.5*(1-(Lfin_100/O$r)^2)" # R f  (V0/V=(L0/L)^3, so (V0/V)^(2/3)=(L0/L)^2)
        $ws.Cells.Item($r,19).Formula = "=2*O$r/C$r"         # S Vp100
        $ws.Cells.Item($r,20).Formula = "=2*O$r/D$r"         # T Vs100
        $ws.Cells.Item($r,21).Formula = "=2*P$r/E$r"         # U Vp111
        $ws.Cells.Item($r,22).Formula = "=2*P$r/F$r"         # V Vs111
        $ws.Cells.Item($r,23).Formula = "=Q$r*U$r^2"         # W rVp111^2
        $ws.Cells.Item($r,24).Formula = "=Q$r*V$r^2"         # X rVs111^2
        $ws.Cells.Item($r,25).Formula = ($measC11 -replace '\{r\}',"$r")  # Y C11
        $ws.Cells.Item($r,26).Formula = ($measC12 -replace '\{r\}',"$r")  # Z C12
        $ws.Cells.Item($r,27).Formula = ($measC44 -replace '\{r\}',"$r")  # AA C44
        # FS-form model fit (params col AP)
        $ws.Cells.Item($r,28).Formula = "=-3*`$AP`$11*((1-2*R$r)^2.5)*(1+1.5*(4-`$AP`$12)*R$r)*R$r"  # AB P_FS
        $ws.Cells.Item($r,29).Formula = "=((1-2*R$r)^3.5)*(`$AP`$4-`$AP`$14*R$r)+3*AB$r"  # AC C11mod
        $ws.Cells.Item($r,30).Formula = "=((1-2*R$r)^3.5)*(`$AP`$5-`$AP`$15*R$r)+1*AB$r"  # AD C12mod
        $ws.Cells.Item($r,31).Formula = "=((1-2*R$r)^3.5)*(`$AP`$6-`$AP`$16*R$r)+1*AB$r"  # AE C44mod
        $ws.Cells.Item($r,32).Formula = "=AB$r-B$r"   # AF dP
        $ws.Cells.Item($r,33).Formula = "=AC$r-Y$r"   # AG dC11
        $ws.Cells.Item($r,34).Formula = "=AD$r-Z$r"   # AH dC12
        $ws.Cells.Item($r,35).Formula = "=AE$r-AA$r"  # AI dC44
        $ws.Cells.Item($r,36).Formula = "=IF(AND(A$r>=CutoffLowBar,A$r<=CutoffOilBar),AF$r^2+AG$r^2+AH$r^2+AI$r^2,0)"  # AJ gated
        $ws.Cells.Item($r,37).Formula = "=(AC$r+2*AD$r)/3"                       # AK K bulk
        $ws.Cells.Item($r,38).Formula = "=(AC$r-AD$r+3*AE$r)/5"                  # AL GV
        $ws.Cells.Item($r,39).Formula = "=5*(AC$r-AD$r)*AE$r/(3*(AC$r-AD$r)+4*AE$r)"  # AM GR
        $ws.Cells.Item($r,40).Formula = "=(AL$r+AM$r)/2"                         # AN GH
    }

    $last = 3 + $n
    # param block col AP=42, labels AO=41
    $ws.Cells.Item(3,42).Value2 = 'FIT PARAMS'
    $lbl = @('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
    for ($i=0;$i -lt 6;$i++){
        $ws.Cells.Item(4+$i,41).Value2 = [string]$lbl[$i]
        $ws.Cells.Item(4+$i,42).Value2 = [double]$seedParams[$i]
    }
    $ws.Cells.Item(11,41).Value2='K0';   $ws.Cells.Item(11,42).Formula="=(`$AP`$4+2*`$AP`$5)/3"
    $ws.Cells.Item(12,41).Value2="K0'";  $ws.Cells.Item(12,42).Formula="=(`$AP`$7+2*`$AP`$8)/3"
    $ws.Cells.Item(14,41).Value2='C11a'; $ws.Cells.Item(14,42).Formula="=3*`$AP`$11*(`$AP`$7-3)-7*`$AP`$4"
    $ws.Cells.Item(15,41).Value2='C12a'; $ws.Cells.Item(15,42).Formula="=3*`$AP`$11*(`$AP`$8-1)-7*`$AP`$5"
    $ws.Cells.Item(16,41).Value2='C44a'; $ws.Cells.Item(16,42).Formula="=3*`$AP`$11*(`$AP`$9-1)-7*`$AP`$6"
    $ws.Cells.Item(18,41).Value2='OBJECTIVE (misfit)'; $ws.Cells.Item(18,42).Formula="=SUM(AJ4:AJ$last)"

    # outputs col AS=45 label, AT=46 value
    $ws.Cells.Item(3,45).Value2='OUTPUTS'
    $outs = @(
      @('C11_0','=$AP$4'), @('C12_0','=$AP$5'), @('C44_0','=$AP$6'),
      @("C11'",'=$AP$7'), @("C12'",'=$AP$8'), @("C44'",'=$AP$9'),
      @('K0','=$AP$11'), @("K0'",'=$AP$12'),
      @('GV0','=AL4'), @('GR0','=AM4'), @('GH0 (Hill)','=AN4'),
      @('misfit','=$AP$18')
    )
    for ($i=0;$i -lt $outs.Count;$i++){
        $ws.Cells.Item(4+$i,45).Value2 = [string]$outs[$i][0]
        $ws.Cells.Item(4+$i,46).Formula = $outs[$i][1]
    }
    return $ws
}
