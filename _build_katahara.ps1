$ErrorActionPreference = 'Stop'
$path = 'C:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_finite_strain_all_equations_GlobalMin - Copy.xlsx'

# --- backup ---
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$bak = $path -replace '\.xlsx$', "_prekatahara_$stamp.xlsx"
Copy-Item -LiteralPath $path -Destination $bak -Force
Write-Host "Backup: $bak"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($path)
    $src = "'FS No Vp 100'"

    # drop existing target sheet if present
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq 'Katahara extrap') { $s.Delete() } }
    $ws = $wb.Worksheets.Add()
    $ws.Name = 'Katahara extrap'

    # ---------- parameter block (X labels / Y values) ----------
    $ws.Range('X1').Value2 = 'KATAHARA 1979  (V, 298 K)'
    $ws.Range('X2').Value2 = 'C11_0 (GPa)'; $ws.Range('Y2').Value2 = [double]230.92
    $ws.Range('X3').Value2 = 'C12_0 (GPa)'; $ws.Range('Y3').Value2 = [double]120.01
    $ws.Range('X4').Value2 = 'C44_0 (GPa)'; $ws.Range('Y4').Value2 = [double]43.36
    $ws.Range('X5').Value2 = 'dC11/dP';     $ws.Range('Y5').Value2 = [double]5.65
    $ws.Range('X6').Value2 = 'dC12/dP';     $ws.Range('Y6').Value2 = [double]3.57
    $ws.Range('X7').Value2 = 'dC44/dP';     $ws.Range('Y7').Value2 = [double]0.185
    # derived params (formulas)
    $ws.Range('X8').Value2  = "C'_0";   $ws.Range('Y8').Formula  = '=(Y2-Y3)/2'
    $ws.Range('X9').Value2  = "dC'/dP"; $ws.Range('Y9').Formula  = '=(Y5-Y6)/2'
    $ws.Range('X10').Value2 = 'K0';     $ws.Range('Y10').Formula = '=(Y2+2*Y3)/3'
    $ws.Range('X11').Value2 = "K0'";    $ws.Range('Y11').Formula = '=(Y5+2*Y6)/3'
    $ws.Range('X12').Value2 = 'C11a';   $ws.Range('Y12').Formula = '=3*Y10*(Y5-3)-7*Y2'
    $ws.Range('X13').Value2 = 'C12a';   $ws.Range('Y13').Formula = '=3*Y10*(Y6-1)-7*Y3'
    $ws.Range('X14').Value2 = 'C44a';   $ws.Range('Y14').Formula = '=3*Y10*(Y7-1)-7*Y4'

    # ---------- headers ----------
    $hdr = @('P meas (GPa)','V0/V','f','P BM3 kat',
             'C11 lin','C12 lin','C44 lin',"C' lin",'K lin',
             'C11 FS','C12 FS','C44 FS',"C' FS",'K FS',
             'C11 mine','C12 mine','C44 mine',
             'dC11 FS-mine','dC12 FS-mine','dC44 FS-mine')
    for ($i=0; $i -lt $hdr.Count; $i++) {
        $ws.Cells.Item(1, $i+1).Value2 = $hdr[$i]
    }

    # ---------- data formulas (row 2 template; Excel fills rel refs to row 23) ----------
    $map = @(
        @('A2:A23', '=''FS No Vp 100''!A2'),
        @('B2:B23', '=''FS No Vp 100''!AP2'),
        @('C2:C23', '=0.5*(1-B2^0.666666667)'),
        @('D2:D23', '=-3*$Y$10*((1-2*C2)^2.5)*(1+3*(4-$Y$11)*C2/2)*C2'),
        @('E2:E23', '=$Y$2+$Y$5*A2'),
        @('F2:F23', '=$Y$3+$Y$6*A2'),
        @('G2:G23', '=$Y$4+$Y$7*A2'),
        @('H2:H23', '=(E2-F2)/2'),
        @('I2:I23', '=(E2+2*F2)/3'),
        @('J2:J23', '=(1-2*C2)^3.5*($Y$2-$Y$12*C2)+D2*3'),
        @('K2:K23', '=(1-2*C2)^3.5*($Y$3-$Y$13*C2)+D2*1'),
        @('L2:L23', '=(1-2*C2)^3.5*($Y$4-$Y$14*C2)+D2*1'),
        @('M2:M23', '=(J2-K2)/2'),
        @('N2:N23', '=(J2+2*K2)/3'),
        @('O2:O23', '=''FS No Vp 100''!AT2'),
        @('P2:P23', '=''FS No Vp 100''!AV2'),
        @('Q2:Q23', '=''FS No Vp 100''!AU2'),
        @('R2:R23', '=J2-O2'),
        @('S2:S23', '=K2-P2'),
        @('T2:T23', '=L2-Q2')
    )
    foreach ($m in $map) { $ws.Range($m[0]).Formula = $m[1] }

    # cosmetics
    $ws.Range('A1:T1').Font.Bold = $true
    $ws.Range('X1').Font.Bold = $true
    [void]$ws.Columns.Item("A:Y").AutoFit()

    $xl.CalculateFull()
    $wb.Save()

    # sanity read-back
    Write-Host ("K0={0:N2}  K0'={1:N3}  C11a={2:N1}  C12a={3:N1}  C44a={4:N1}" -f `
        $ws.Range('Y10').Value2, $ws.Range('Y11').Value2, $ws.Range('Y12').Value2, `
        $ws.Range('Y13').Value2, $ws.Range('Y14').Value2)
    Write-Host "`n P(GPa)  C11lin C11FS C11mine | C12lin C12FS C12mine | C44lin C44FS C44mine"
    foreach ($r in 2,12,23) {
        Write-Host ("{0,6:N2}  {1,6:N1}{2,6:N1}{3,7:N1} | {4,6:N1}{5,6:N1}{6,7:N1} | {7,6:N2}{8,6:N2}{9,7:N2}" -f `
            $ws.Cells.Item($r,1).Value2, $ws.Cells.Item($r,5).Value2, $ws.Cells.Item($r,10).Value2, $ws.Cells.Item($r,15).Value2, `
            $ws.Cells.Item($r,6).Value2, $ws.Cells.Item($r,11).Value2, $ws.Cells.Item($r,16).Value2, `
            $ws.Cells.Item($r,7).Value2, $ws.Cells.Item($r,12).Value2, $ws.Cells.Item($r,17).Value2)
    }
    $wb.Close($true)
    Write-Host "`nDONE - sheet 'Katahara extrap' written."
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
