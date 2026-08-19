# Recalculate ozzit.xlsx in Excel and save, so the values in the file are the values its
# own formulas produce.
#
# The build edits the workbook as XML and has no formula engine, so anything downstream of
# a value it changes keeps the answer it had before. Shifting the sample dates forward two
# years left 3,193 cached cells across 43 sheets holding numbers their formulas no longer
# produce. Excel replaced them on open, so no reader ever saw one, but a file that
# disagrees with itself cannot be checked by anything except Excel.
#
# This is the only step in the pipeline that needs Excel, and it changes no formula: all
# 4,506 formula cells come out byte-identical. Run tools/verify_cache.py afterwards, which
# is the gate that proves it.
param([string]$Path = 'ozzit.xlsx')

$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path $Path).Path

function Invoke-Excel([scriptblock]$op, [int]$tries = 12) {
    for ($i = 1; $i -le $tries; $i++) {
        try { return & $op }
        catch { if ($i -eq $tries) { throw }; Start-Sleep -Milliseconds (250 * $i) }
    }
}

# Excel's COM server is shared and this script quits it when it finishes, so refuse to run
# rather than close someone's open workbooks underneath them.
for ($w = 0; $w -lt 10 -and @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0; $w++) {
    Start-Sleep -Milliseconds 500
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Output 'Excel is already running. Close it first: this drives Excel over COM and quits it when done.'
    exit 2
}

$before = (Get-Item $Path).Length
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 1
$wb = $null; $exit = 0
try {
    $wb = Invoke-Excel { $xl.Workbooks.Open($Path, 0, $false) }
    Invoke-Excel { $xl.CalculateFullRebuild() }
    # xlDone = 0. Saving while the engine is still working would write half an answer.
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
    Invoke-Excel { $wb.Save() }
    Invoke-Excel { $wb.Close($false) }; $wb = $null
    $after = (Get-Item $Path).Length
    Write-Output ("refreshed {0}: {1} -> {2} bytes" -f (Split-Path $Path -Leaf), $before, $after)
}
catch { Write-Output ('FAIL: ' + $_.Exception.Message); $exit = 1 }
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $exit
