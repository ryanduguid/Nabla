# Write every calculated cell value in a workbook to a tab-separated file, so a checker can
# compare what Excel produces against what the file has cached.
#
# Cells are keyed by tab position rather than by $ws.Index, which intermittently returns an
# empty value over COM and silently attributes a sheet's cells to the wrong sheet.
param([string]$Path, [string]$Out)

$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path $Path).Path

function Invoke-Excel([scriptblock]$op, [int]$tries = 12) {
    for ($i = 1; $i -le $tries; $i++) {
        try { return & $op }
        catch { if ($i -eq $tries) { throw }; Start-Sleep -Milliseconds (250 * $i) }
    }
}

for ($w = 0; $w -lt 10 -and @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0; $w++) {
    Start-Sleep -Milliseconds 500
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Output 'Excel is already running. Close it first: this drives Excel over COM and quits it when done.'
    exit 2
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
$wb = $null; $exit = 0
$sw = [System.IO.StreamWriter]::new($Out, $false, (New-Object System.Text.UTF8Encoding $false))
try {
    $wb = Invoke-Excel { $xl.Workbooks.Open($Path, 0, $false) }
    Invoke-Excel { $xl.CalculateFullRebuild() }
    # CalculationState comes back over COM as the enum's NAME, "xlDone", not as 0, so a
    # test against 0 is true forever and the wait below never ends early. Excel very
    # occasionally reports itself still busy after a rebuild that has in fact finished, so
    # ask a second time before giving up rather than failing a run for a transient state.
    $settled = $false
    for ($attempt = 1; $attempt -le 2 -and -not $settled; $attempt++) {
        for ($k = 0; $k -lt 300 -and "$($xl.CalculationState)" -notin @('xlDone', '0'); $k++) {
            Start-Sleep -Milliseconds 50
        }
        $settled = "$($xl.CalculationState)" -in @('xlDone', '0')
        if (-not $settled) { Invoke-Excel { $xl.CalculateFullRebuild() } }
    }
    if (-not $settled) { throw 'Excel did not finish calculating within 30 seconds' }
    $pos = 0
    foreach ($ws in $wb.Worksheets) {
        $pos++
        $ur = $ws.UsedRange
        $r0 = $ur.Row; $c0 = $ur.Column
        # Not Invoke-Excel: PowerShell flattens a two-dimensional array returned through a
        # function, which turns the whole sheet into one row and hides every difference.
        $vals = $ur.Value2
        if ($null -eq $vals) { continue }
        if ($vals -isnot [object[,]]) {
            $sw.WriteLine("$pos`t$r0`t$c0`t" + (([string]$vals) -replace "`r`n", '\n' -replace "`n", '\n' -replace "`r", '\n'))
            continue
        }
        for ($i = 1; $i -le $vals.GetLength(0); $i++) {
            for ($j = 1; $j -le $vals.GetLength(1); $j++) {
                $v = $vals[$i, $j]
                if ($null -eq $v) { continue }
                # A cell may hold newlines. Fold them so one cell stays one line.
                $s = ([string]$v) -replace "`r`n", '\n' -replace "`n", '\n' -replace "`r", '\n'
                $sw.WriteLine("$pos`t" + ($r0 + $i - 1) + "`t" + ($c0 + $j - 1) + "`t" + $s)
            }
        }
    }
    Invoke-Excel { $wb.Close($false) }; $wb = $null
}
catch { Write-Output ('FAIL: ' + $_.Exception.Message); $exit = 1 }
finally {
    $sw.Close()
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $exit
