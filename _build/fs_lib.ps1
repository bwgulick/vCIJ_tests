# Reusable finite-strain combo-sheet builder.
# Column map (data rows 4..25):
#  A oil  B P(GPa)  C tP100  D tS100  E tP111  F tS111
#  G L(fit)  H L111  I V0/V  J rho  K f
#  L Vp100 M Vs100 N Vp111 O Vs111  P rVp111^2  Q rVs111^2
#  R C11meas S C12meas T C44meas
#  U P_FS  V C11mod W C12mod X C44mod
#  Y dP  Z dC11  AA dC12  AB dC44  AC gated-sq
#  AD Kbulk AE GV AF GR AG GH
# Param block col AI: 4..9 = C11_0,C12_0,C44_0,C11',C12',C44'; 11 K0;12 K0';14..16 a-coefs;18 objective

function Build-FSCombo {
    param(
        $wb, $afterSheet, [string]$name,
        [string]$measC11, [string]$measC12, [string]$measC44,   # templates with {r}
        [int]$n = 22, [double[]]$seedParams
    )
    $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $afterSheet)
    $ws.Name = $name

    $hdr = @('Oil bar','P (GPa)','tP100','tS100','tP111','tS111',
             'L fit','L111','V0/V','rho','f',
             'Vp100','Vs100','Vp111','Vs111','rVp111^2','rVs111^2',
             'C11 meas','C12 meas','C44 meas',
             'P_FS','C11 mod','C12 mod','C44 mod',
             'dP','dC11','dC12','dC44','gated sq',
             'K bulk','GV','GR','GH')
    for ($c=0;$c -lt $hdr.Count;$c++){ $ws.Cells.Item(3,1+$c).Value2 = $hdr[$c] }

    # seed lengths: 0.704 at oil0 decreasing to ~0.689 at oil630
    for ($i=0;$i -lt $n;$i++){
        $r = 4 + $i
        $oil = 30*$i
        $seedL = 0.704 - ($oil/630.0)*0.015
        $ws.Cells.Item($r,1).Formula  = "=Align!A$r"
        $ws.Cells.Item($r,2).Formula  = "=Align!B$r"
        $ws.Cells.Item($r,3).Formula  = "=Align!C$r"
        $ws.Cells.Item($r,4).Formula  = "=Align!D$r"
        $ws.Cells.Item($r,5).Formula  = "=Align!F$r"
        $ws.Cells.Item($r,6).Formula  = "=Align!G$r"
        if ($r -eq 4) { $ws.Cells.Item($r,7).Formula = "=Lfin_100" }  # G4 fixed = zero-P reference length (anchors scale)
        else          { $ws.Cells.Item($r,7).Value2  = [double]$seedL } # G5:G25 fitted lengths (changing)
        $ws.Cells.Item($r,8).Formula  = "=G$r*Lfin_111/Lfin_100"     # H L111
        $ws.Cells.Item($r,9).Formula  = "=(`$G`$4/G$r)^3"            # I V0/V
        $ws.Cells.Item($r,10).Formula = "=rho_100*I$r"               # J rho
        $ws.Cells.Item($r,11).Formula = "=0.5*(1-I$r^(2/3))"         # K f
        $ws.Cells.Item($r,12).Formula = "=2*G$r/C$r"                 # L Vp100
        $ws.Cells.Item($r,13).Formula = "=2*G$r/D$r"                 # M Vs100
        $ws.Cells.Item($r,14).Formula = "=2*H$r/E$r"                 # N Vp111
        $ws.Cells.Item($r,15).Formula = "=2*H$r/F$r"                 # O Vs111
        $ws.Cells.Item($r,16).Formula = "=J$r*N$r^2"                 # P rVp111^2
        $ws.Cells.Item($r,17).Formula = "=J$r*O$r^2"                 # Q rVs111^2
        $ws.Cells.Item($r,18).Formula = ($measC11 -replace '\{r\}',"$r")  # R C11meas
        $ws.Cells.Item($r,19).Formula = ($measC12 -replace '\{r\}',"$r")  # S C12meas
        $ws.Cells.Item($r,20).Formula = ($measC44 -replace '\{r\}',"$r")  # T C44meas
        # model
        $ws.Cells.Item($r,21).Formula = "=-3*`$AI`$11*((1-2*K$r)^2.5)*(1+1.5*(4-`$AI`$12)*K$r)*K$r"  # U P_FS
        $ws.Cells.Item($r,22).Formula = "=((1-2*K$r)^3.5)*(`$AI`$4-`$AI`$14*K$r)+3*U$r"  # V C11mod
        $ws.Cells.Item($r,23).Formula = "=((1-2*K$r)^3.5)*(`$AI`$5-`$AI`$15*K$r)+1*U$r"  # W C12mod
        $ws.Cells.Item($r,24).Formula = "=((1-2*K$r)^3.5)*(`$AI`$6-`$AI`$16*K$r)+1*U$r"  # X C44mod
        # residuals
        $ws.Cells.Item($r,25).Formula = "=U$r-B$r"    # Y dP
        $ws.Cells.Item($r,26).Formula = "=V$r-R$r"    # Z dC11
        $ws.Cells.Item($r,27).Formula = "=W$r-S$r"    # AA dC12
        $ws.Cells.Item($r,28).Formula = "=X$r-T$r"    # AB dC44
        $ws.Cells.Item($r,29).Formula = "=IF(AND(A$r>=CutoffLowBar,A$r<=CutoffOilBar),Y$r^2+Z$r^2+AA$r^2+AB$r^2,0)"  # AC gated
        # outputs
        $ws.Cells.Item($r,30).Formula = "=(V$r+2*W$r)/3"                     # AD K bulk
        $ws.Cells.Item($r,31).Formula = "=(V$r-W$r+3*X$r)/5"                 # AE GV
        $ws.Cells.Item($r,32).Formula = "=5*(V$r-W$r)*X$r/(3*(V$r-W$r)+4*X$r)"  # AF GR
        $ws.Cells.Item($r,33).Formula = "=(AE$r+AF$r)/2"                     # AG GH
    }

    $last = 3 + $n
    # param block (col AI = 35)
    $ws.Cells.Item(3,35).Value2 = 'FIT PARAMS'
    $lbl = @('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
    for ($i=0;$i -lt 6;$i++){
        $ws.Cells.Item(4+$i,34).Value2 = [string]$lbl[$i]           # col AH label
        $ws.Cells.Item(4+$i,35).Value2 = [double]$seedParams[$i]    # col AI value (changing)
    }
    $ws.Cells.Item(11,34).Value2='K0';      $ws.Cells.Item(11,35).Formula="=(`$AI`$4+2*`$AI`$5)/3"
    $ws.Cells.Item(12,34).Value2="K0'";     $ws.Cells.Item(12,35).Formula="=(`$AI`$7+2*`$AI`$8)/3"
    $ws.Cells.Item(14,34).Value2='C11a';    $ws.Cells.Item(14,35).Formula="=3*`$AI`$11*(`$AI`$7-3)-7*`$AI`$4"
    $ws.Cells.Item(15,34).Value2='C12a';    $ws.Cells.Item(15,35).Formula="=3*`$AI`$11*(`$AI`$8-1)-7*`$AI`$5"
    $ws.Cells.Item(16,34).Value2='C44a';    $ws.Cells.Item(16,35).Formula="=3*`$AI`$11*(`$AI`$9-1)-7*`$AI`$6"
    $ws.Cells.Item(18,34).Value2='OBJECTIVE (misfit)'; $ws.Cells.Item(18,35).Formula="=SUM(AC4:AC$last)"

    # output summary block (col AL=38 label, AM=39 value) - Cij0 at P=0 come straight from params
    $ws.Cells.Item(3,38).Value2='OUTPUTS'
    $outs = @(
      @('C11_0','=$AI$4'), @('C12_0','=$AI$5'), @('C44_0','=$AI$6'),
      @("C11'",'=$AI$7'), @("C12'",'=$AI$8'), @("C44'",'=$AI$9'),
      @('K0','=$AI$11'), @("K0'",'=$AI$12'),
      @('GV0','=AE4'), @('GR0','=AF4'), @('GH0 (Hill)','=AG4'),
      @('misfit','=$AI$18')
    )
    for ($i=0;$i -lt $outs.Count;$i++){
        $ws.Cells.Item(4+$i,38).Value2 = $outs[$i][0]
        $ws.Cells.Item(4+$i,39).Formula = $outs[$i][1]
    }

    return $ws
}
