$ErrorActionPreference='Stop'
$out='c:\Users\bgulick\Downloads\vCIJ_tests\Vanadium_Cij_simplified.xlsx'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.ScreenUpdating=$false
try{
    $wb=$xl.Workbooks.Open($out)
    $existing=@(); foreach($s in $wb.Worksheets){ $existing+=$s.Name }
    if($existing -contains 'Summary'){ $wb.Worksheets.Item('Summary').Delete() }
    $sm=$wb.Worksheets.Add($wb.Worksheets.Item(1))   # first sheet
    $sm.Name='Summary'

    # metric rows and their source output-block row numbers
    $metrics=@(
      @('C11_0 (GPa)',4), @('C12_0 (GPa)',5), @('C44_0 (GPa)',6),
      @("C11'",7), @("C12'",8), @("C44'",9),
      @('K0 (GPa)',10), @("K0'",11), @('G Hill 0 (GPa)',14), @('total misfit',15)
    )
    # fit sheets: name, value-column-letter
    $fits=@(
      @('FS_NoVs111','AM'), @('Cook_NoVs111','AT'),
      @('FS_NoVp111','AM'), @('Cook_NoVp111','AT'),
      @('FS_NoVs100','AM'), @('Cook_NoVs100','AT'),
      @('FS_NoVp100','AM'), @('Cook_NoVp100','AT'),
      @('Global_FS','AT'),  @('Global_Cook','DA')
    )
    $colHdr=@('No Vs111 FS','No Vs111 CK','No Vp111 FS','No Vp111 CK','No Vs100 FS','No Vs100 CK',
              'No Vp100 FS','No Vp100 CK','GLOBAL FS','GLOBAL CK')

    $sm.Range('A1').Value2='SIMPLIFIED VANADIUM Cij — SUMMARY (all fits live-linked; edit Inputs, re-run each sheet''s Solver)'
    # header row 3
    for($c=0;$c -lt $colHdr.Count;$c++){ $sm.Cells.Item(3,2+$c).Value2=$colHdr[$c] }
    $sm.Cells.Item(3,13).Value2='Katahara 1979'
    $sm.Cells.Item(3,14).Value2='Lenkkeri 1978'
    $sm.Cells.Item(3,15).Value2='Greiner 1979'

    # metric rows starting row4
    for($m=0;$m -lt $metrics.Count;$m++){
        $rr=4+$m
        $sm.Cells.Item($rr,1).Value2=[string]$metrics[$m][0]
        $srcRow=$metrics[$m][1]
        for($f=0;$f -lt $fits.Count;$f++){
            $sheet=$fits[$f][0]; $vcol=$fits[$f][1]
            $sm.Cells.Item($rr,2+$f).Formula="='$sheet'!$vcol$srcRow"
        }
    }
    # literature (only Cij0 and Cij')
    $lit=@{
      'Katahara'=@{col=13; C11=230.9; C12=120.02; C44=43.36; C11p=5.65; C12p=3.566; C44p=0.185}
      'Lenkkeri'=@{col=14; C11=235.7; C12=126.5; C44=43.6}
      'Greiner' =@{col=15; C11=235.89; C12=120.76; C44=46.55}
    }
    foreach($k in $lit.Keys){
        $L=$lit[$k]; $c=$L.col
        if($L.ContainsKey('C11')){ $sm.Cells.Item(4,$c).Value2=[double]$L.C11 }
        if($L.ContainsKey('C12')){ $sm.Cells.Item(5,$c).Value2=[double]$L.C12 }
        if($L.ContainsKey('C44')){ $sm.Cells.Item(6,$c).Value2=[double]$L.C44 }
        if($L.ContainsKey('C11p')){ $sm.Cells.Item(7,$c).Value2=[double]$L.C11p }
        if($L.ContainsKey('C12p')){ $sm.Cells.Item(8,$c).Value2=[double]$L.C12p }
        if($L.ContainsKey('C44p')){ $sm.Cells.Item(9,$c).Value2=[double]$L.C44p }
    }
    # derived K0 = (C11+2*C12)/3 for literature columns M,N,O
    foreach($cl in @('M','N','O')){ $sm.Range($cl+'10').Formula="=($cl`4+2*$cl`5)/3" }

    $sm.Columns.Item('A').ColumnWidth=16
    $sm.Range('A1').Font.Bold=$true
    $sm.Range('A3:O3').Font.Bold=$true
    $sm.Range('A4:A13').Font.Bold=$true

    $wb.Application.Calculate()
    # read back for verification
    Write-Output "==== SUMMARY (live-linked) ===="
    $line='metric        '
    foreach($h in $colHdr){ $line += ('{0,11}' -f $h.Substring(0,[Math]::Min(11,$h.Length))) }
    Write-Output $line
    for($m=0;$m -lt $metrics.Count;$m++){
        $rr=4+$m; $s=('{0,-13}' -f $metrics[$m][0])
        for($f=0;$f -lt $fits.Count;$f++){ $s+=('{0,11:N3}' -f [double]$sm.Cells.Item($rr,2+$f).Value2) }
        Write-Output $s
    }
    $wb.Save(); $wb.Close($true)
}catch{ Write-Output "ERROR: $_"; Write-Output $_.ScriptStackTrace }
finally{ $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
