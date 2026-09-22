param(
  [string]$Wb  = "$PSScriptRoot\..\workbooks\V_Cij.xlsx",
  [string]$Src = "No Vp 100",
  [string]$Out = "Plot Data",
  [int]$R0 = 2,
  [int]$R1 = 23
)
$ErrorActionPreference = "Stop"

# Column plan: A..Q  header + per-row formula template ({r}=source row, {src}=source sheet)
# value and its sigma are always adjacent, one row per pressure point.
$cols = @(
  @{c=1;  h="Pressure (GPa)"; f="='{src}'!A{r}"},
  @{c=2;  h="Length (mm)";    f="='{src}'!M{r}"},
  @{c=3;  h="sig Length";     f="='{src}'!Q{r}"},
  @{c=4;  h="Vp100 (km/s)";   f="=2*0.712/'{src}'!F{r}"},
  @{c=5;  h="sig Vp100";      f="=D{r}*SQRT(('{src}'!Q{r}/'{src}'!M{r})^2+(0.0002/'{src}'!F{r})^2)"},
  @{c=6;  h="Vs100 (km/s)";   f="='{src}'!AE{r}"},
  @{c=7;  h="sig Vs100";      f="='{src}'!Z{r}"},
  @{c=8;  h="Vp111 (km/s)";   f="='{src}'!AC{r}"},
  @{c=9;  h="sig Vp111";      f="='{src}'!AA{r}"},
  @{c=10; h="Vs111 (km/s)";   f="='{src}'!AD{r}"},
  @{c=11; h="sig Vs111";      f="='{src}'!AB{r}"},
  @{c=12; h="C11 (GPa)";      f="='{src}'!AL{r}"},
  @{c=13; h="sig C11";        f="='{src}'!FI{r}"},
  @{c=14; h="C44 (GPa)";      f="='{src}'!AJ{r}"},
  @{c=15; h="sig C44";        f="='{src}'!FK{r}"},
  @{c=16; h="C12 (GPa)";      f="='{src}'!AK{r}"},
  @{c=17; h="sig C12";        f="='{src}'!FJ{r}"}
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
try {
  $book = $xl.Workbooks.Open($Wb)
  $xl.Calculation = -4135   # manual

  # delete existing Out sheet if present
  foreach ($s in @($book.Worksheets)) {
    if ($s.Name -eq $Out) { $s.Delete() | Out-Null }
  }
  # add new sheet at end
  $last = $book.Worksheets.Item($book.Worksheets.Count)
  $sh = $book.Worksheets.Add([System.Reflection.Missing]::Value, $last)
  $sh.Name = $Out

  # title row (row 1 merged-ish label in A1..? keep simple: A1 title, headers row 2, data 3..)
  $HDR = 2
  $sh.Cells.Item(1,1).Value2 = "PLOT DATA  (per-point values + 1-sigma error bars)   source combo: $Src"

  foreach ($col in $cols) {
    $sh.Cells.Item($HDR, $col.c).Value2 = $col.h
  }
  $outRow = $HDR
  for ($r = $R0; $r -le $R1; $r++) {
    $outRow++
    foreach ($col in $cols) {
      $ftext = $col.f.Replace("{src}", $Src).Replace("{r}", "$r")
      # formulas in E reference D{r} on the OUTPUT sheet, but our output rows are shifted.
      # Fix: E's Vp100 sigma multiplies the output-sheet Vp100 cell = D$outRow.
      $ftext = $ftext.Replace("=D{0}" -f $r, "=D$outRow")
      # (only col 5 contains a D{r} self-reference; guard by exact token)
      $sh.Cells.Item($outRow, $col.c).Formula = $ftext
    }
  }

  # note under the table
  $noteRow = $outRow + 2
  $sh.Cells.Item($noteRow,1).Value2 = "Notes:"
  $sh.Cells.Item($noteRow+1,1).Value2 = "Pressure=col A (gold gauge). Length & sig from Cook integration (M, Q)."
  $sh.Cells.Item($noteRow+2,1).Value2 = "Velocities V=2*0.712/t; sigV=V*sqrt((sigL/L)^2+(sigt/t)^2), sigt=0.0002 us. Vs100/Vp111/Vs111 sig taken from source cols Z/AA/AB; Vp100 sig computed here (its source col Y is a pre-existing #REF!)."
  $sh.Cells.Item($noteRow+3,1).Value2 = "C11/C44/C12 experimental points from '$Src' cols AL/AJ/AK; per-point 1-sigma from analytic correlation-aware block (cols FI/FK/FJ)."
  $sh.Cells.Item($noteRow+4,1).Value2 = "For an XY chart: use the value column as Y and the adjacent 'sig' column as the custom +/- error bar amount."

  # cosmetics
  $sh.Rows.Item(1).Font.Bold = $true
  $sh.Rows.Item($HDR).Font.Bold = $true
  $sh.Columns("A:Q").ColumnWidth = 12

  $xl.Calculation = -4105   # automatic
  $book.Application.CalculateFull()
  $book.Save()
  Write-Output ("Created '$Out' sheet: $($R1-$R0+1) data rows (source rows $R0..$R1), header row $HDR.")
  $book.Close($true)
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
