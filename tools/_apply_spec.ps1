param(
  [string]$Wb   = "$PSScriptRoot\..\workbooks\V_Cij.xlsx",
  [string]$Spec = "$PSScriptRoot\_uncert_step1_spec.tsv"
)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
try {
  $book = $xl.Workbooks.Open($Wb)
  $xl.Calculation = -4135   # xlCalculationManual (after a workbook exists)
  # cache sheets
  $ws = @{}
  foreach ($s in $book.Worksheets) { $ws[$s.Name] = $s }

  $n = 0
  foreach ($line in [System.IO.File]::ReadAllLines($Spec)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $p = $line.Split("`t")
    $sheet = $p[0]; $cell = $p[1]; $kind = $p[2]
    $content = if ($p.Count -ge 4) { $p[3] } else { "" }
    $rng = $ws[$sheet].Range($cell)
    if ($kind -eq "F") {
      $rng.Formula = $content
    } else {
      if ([string]::IsNullOrEmpty($content)) { $rng.ClearContents() | Out-Null }
      else { $rng.Value2 = $content }
    }
    $n++
  }
  $xl.Calculation = -4105   # xlCalculationAutomatic
  $book.Application.CalculateFull()
  $book.Save()
  Write-Output ("Applied $n cell ops and saved.")
  $book.Close($true)
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
