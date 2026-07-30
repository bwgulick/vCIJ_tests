$ErrorActionPreference = 'Stop'
$out = 'c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Add()
    # remove extra default sheets, keep one
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    $inp = $wb.Worksheets.Item(1)
    $inp.Name = 'Inputs'
    $align = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $inp)  # add after Inputs
    $align.Name = 'Align'

    # ---------- INPUTS SHEET ----------
    $inp.Range('A1').Value2 = 'INPUTS - paste raw measured data here (100 block cols A-D, 111 block cols F-I). Constants at cols L-M.'
    $hdr = @('Oil bar','SBR 100 (S buffer)','2TP 100 (raw)','2TS 100 (raw)')
    for ($c=0;$c -lt 4;$c++){ $inp.Cells.Item(3,1+$c).Value2 = $hdr[$c] }
    $hdr2 = @('Oil bar','SBR 111 (S buffer)','2TP 111 (raw)','2TS 111 (raw)')
    for ($c=0;$c -lt 4;$c++){ $inp.Cells.Item(3,6+$c).Value2 = $hdr2[$c] }

    # 100 raw data: oil, SBR100, 2TP100, 2TS100  (seeded from user's `data` sheet)
    $d100 = @(
      @(0,1.1602,0.229199934662983,0.510926917293233),
      @(30,1.1596,0.224745720055653,0.507233646604075),
      @(60,1.1526,0.220945720055653,0.505833646604075),
      @(90,1.1486,0.218745720055653,0.505833646604075),
      @(120,1.1442,0.216545720055653,0.505233646604075),
      @(150,1.1396,0.214745720055653,0.505233646604075),
      @(180,1.1364,0.213345720055653,0.505233646604075),
      @(210,1.1334,0.211745720055653,0.504233646604075),
      @(240,1.1296,0.210745720055653,0.504233646604075),
      @(270,1.1264,0.209145720055653,0.504633646604075),
      @(300,1.1238,0.208145720055653,0.503833646604075),
      @(330,1.1206,0.207345720055653,0.504033646604075),
      @(360,1.1178,0.206345720055653,0.504433646604075),
      @(390,1.1158,0.205545720055653,0.503833646604075),
      @(420,1.1132,0.204745720055653,0.503233646604075),
      @(450,1.1108,0.203945720055653,0.504433646604075),
      @(480,1.1094,0.203345720055653,0.505233646604075),
      @(510,1.1084,0.202345720055653,0.505633646604075),
      @(540,1.1066,0.201745720055653,0.505233646604075),
      @(570,1.1042,0.200945720055653,0.505833646604075),
      @(600,1.103,0.200345720055653,0.507033646604075),
      @(630,1.1022,0.199745720055653,0.506833646604075)
    )
    $d111 = @(
      @(0,1.1382,0.2029,0.4194),
      @(30,1.1162,0.2011,0.4126),
      @(60,1.112,0.1995,0.4052),
      @(90,1.1056,0.1969,0.4028),
      @(120,1.1034,0.1953,0.3972),
      @(150,1.0978,0.1941,0.396),
      @(180,1.0942,0.1929,0.3926),
      @(210,1.0914,0.1917,0.3894),
      @(240,1.0862,0.1905,0.3894),
      @(270,1.0832,0.1889,0.3866),
      @(300,1.0804,0.1881,0.3846),
      @(330,1.0758,0.1875,0.3844),
      @(360,1.0726,0.1863,0.383),
      @(390,1.0708,0.1859,0.3812),
      @(420,1.068,0.1851,0.38),
      @(450,1.0644,0.1845,0.3798),
      @(480,1.0626,0.1839,0.3788),
      @(510,1.0614,0.1829,0.3774),
      @(540,1.0594,0.1825,0.3763),
      @(570,1.0568,0.1813,0.3766),
      @(600,1.0552,0.1805,0.3762),
      @(630,1.0546,0.1801,0.3746)
    )
    $n = $d100.Count
    for ($r=0;$r -lt $n;$r++){
        $xr = 4 + $r
        for($c=0;$c -lt 4;$c++){
            $inp.Cells.Item($xr,1+$c).Value2 = [double]$d100[$r][$c]
            $inp.Cells.Item($xr,6+$c).Value2 = [double]$d111[$r][$c]
        }
    }

    # constants block (col L labels, col M values)
    $consts = @(
      @('Linit 100 (initial length, mm)',0.712,'Linit_100'),
      @('Lfin 100 (final/used length, mm)',0.704,'Lfin_100'),
      @('alpha 100 (thermal expansion)',0.000046,'alpha_100'),
      @('rho 100 (density g/cc)',6.09,'rho_100'),
      @('gamma 100 (Gruneisen)',1.23,'gamma_100'),
      @('T 100 (K)',300,'T_100'),
      @('SBR0 100 (0-P buffer S-time = first measured pt; source O24)',1.1602,'SBR0_100'),
      @('scalar 100 =(1+a*g*T)/(12*rho*Lfin^2)','=(1+M6*M8*M9)/(12*M7*M5^2)','scal_100'),
      @('',''),
      @('Linit 111 (initial length, mm)',0.623,'Linit_111'),
      @('Lfin 111 (final/used length, mm)',0.607,'Lfin_111'),
      @('alpha 111 (thermal expansion)',0.000046,'alpha_111'),
      @('rho 111 (density g/cc)',6.09,'rho_111'),
      @('gamma 111 (Gruneisen)',1.23,'gamma_111'),
      @('T 111 (K)',300,'T_111'),
      @('SBR0 111 (pressure_fit, 0-P buffer S-time)',1.1225,'SBR0_111'),
      @('scalar 111 =(1+a*g*T)/(12*rho*Lfin^2)','=(1+M15*M17*M18)/(12*M16*M14^2)','scal_111'),
      @('',''),
      @('Cutoff HIGH oil bar (fit uses oil <= this)',420,'CutoffOilBar'),
      @('Cutoff LOW oil bar (fit uses oil >= this)',120,'CutoffLowBar'),
      @('Normalize flag (0/1, global fits)',0,'NormalizeFlag')
    )
    $row = 4
    foreach ($ct in $consts) {
        $label = [string]$ct[0]
        $inp.Cells.Item($row,12).Value2 = $label   # col L
        $val = $ct[1]
        if ($null -ne $val -and "$val" -ne '') {
            if ($val -is [string] -and $val.StartsWith('=')) {
                $inp.Cells.Item($row,13).Formula = $val
            } else {
                $inp.Cells.Item($row,13).Value2 = [double]$val
            }
        }
        if ($ct.Count -ge 3 -and [string]$ct[2] -ne '') {
            $addr = "=Inputs!`$M`$$row"
            $wb.Names.Add([string]$ct[2], $addr) | Out-Null
        }
        $row++
    }

    $inp.Columns.Item('L').ColumnWidth = 40
    $inp.Columns.Item('A').ColumnWidth = 8

    # ---------- ALIGN SHEET ----------
    $align.Range('A1').Value2 = 'ALIGN - master pressure = 100 calibration; 111 travel times resampled onto master P via 6th-order polynomial (LINEST).'
    $ah = @('Oil bar','P master (GPa)','tP100','tS100','P111 (GPa)','tP111 aligned','tS111 aligned')
    for ($c=0;$c -lt $ah.Count;$c++){ $align.Cells.Item(3,1+$c).Value2 = $ah[$c] }

    # polynomial coefficient blocks (6th order): row2 = 2TP111 coeffs, row3 = 2TS111 coeffs, cols K..Q (c6..c0)
    $align.Cells.Item(1,11).Value2 = 'poly coeffs c6..c0 ->'
    $align.Cells.Item(2,10).Value2 = '2TP111:'
    $align.Cells.Item(3,10).Value2 = '2TS111:'
    for ($k=1;$k -le 7;$k++){
        $col = 10 + $k  # K=11 .. Q=17
        $fTP = "=INDEX(LINEST(Inputs!`$H`$4:`$H`$$(3+$n),`$E`$4:`$E`$$(3+$n)^{1,2,3,4,5,6}),1,$k)"
        $fTS = "=INDEX(LINEST(Inputs!`$I`$4:`$I`$$(3+$n),`$E`$4:`$E`$$(3+$n)^{1,2,3,4,5,6}),1,$k)"
        $align.Cells.Item(2,$col).Formula = $fTP
        $align.Cells.Item(3,$col).Formula = $fTS
    }

    for ($r=0;$r -lt $n;$r++){
        $xr = 4 + $r
        $align.Cells.Item($xr,1).Formula = "=Inputs!A$xr"
        $align.Cells.Item($xr,2).Formula = "=249.7*(1-Inputs!B$xr/SBR0_100)"
        $align.Cells.Item($xr,3).Formula = "=Inputs!C$xr"
        $align.Cells.Item($xr,4).Formula = "=Inputs!D$xr"
        $align.Cells.Item($xr,5).Formula = "=249.7*(1-Inputs!G$xr/SBR0_111)"
        $align.Cells.Item($xr,6).Formula = "=SUMPRODUCT(`$K`$2:`$P`$2,B$xr^{6,5,4,3,2,1})+`$Q`$2"
        $align.Cells.Item($xr,7).Formula = "=SUMPRODUCT(`$K`$3:`$P`$3,B$xr^{6,5,4,3,2,1})+`$Q`$3"
    }

    $wb.Names.Add('nData', "=$n") | Out-Null

    $wb.SaveAs($out, 51)  # xlOpenXMLWorkbook
    Write-Output "SAVED $out"

    # read back all rows for verification (master axis vs source, P111, aligned 111 times)
    $srcMaster = @(0,0.129132908,1.635683503,2.496569557,3.443544217,4.433563179,5.122272022,5.767936563,6.585778314,7.274487157,7.834063093,8.522771936,9.125392174,9.555835201,10.11541114,10.63194277,10.93325289,11.1484744,11.53587313,12.05240476,12.31067057,12.48284778)
    Write-Output "row  oil   Pmaster   Psrc     dMaster  P111    tP111al   tS111al"
    for ($r=0;$r -lt $n;$r++){
        $xr=4+$r
        $pm=[double]$align.Cells.Item($xr,2).Value2
        $ps=$srcMaster[$r]
        Write-Output ("{0,2}  {1,4}  {2,8:N5} {3,8:N5} {4,8:N5}  {5,6:N3}  {6:N6}  {7:N6}" -f $xr,
            [int]$align.Cells.Item($xr,1).Value2, $pm, $ps, ($pm-$ps),
            [double]$align.Cells.Item($xr,5).Value2,
            [double]$align.Cells.Item($xr,6).Value2, [double]$align.Cells.Item($xr,7).Value2)
    }

    $wb.Save()
    $wb.Close($true)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
