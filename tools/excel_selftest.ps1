# End-to-end check of nabla.xlsx in a real Excel calculation engine.
#
#   powershell -ExecutionPolicy Bypass -File tools\excel_selftest.ps1 [path\to\nabla.xlsx]
#
# Requires Excel with LAMBDA support (Microsoft 365, or Excel 2024 and later). GitHub's
# runners have no Excel, so CI runs tools/verify_workbook.py instead and this stays a
# local gate: verify_workbook.py checks the file's structure, this checks its arithmetic.
#
# The workbook is opened, fully recalculated, scanned for error cells, probed with a
# temporary sheet of assertions, then closed WITHOUT saving. It is never modified.

param([string]$Path = "$PSScriptRoot\..\nabla.xlsx")

$ErrorActionPreference = 'Stop'
$L = [char]0x03BB                 # this file stays pure ASCII so encoding cannot corrupt it
$Path = (Resolve-Path $Path).Path

$checks = @()
function Check($id, $formula) { $script:checks += @{ id = $id; f = $formula } }

# Each formula must evaluate to exactly "OK". Anything else is printed as the failure.
function Near($id, $expr, $want, $tol = '0.000001') {
    Check $id "=LET(v, $expr, IF(ISERROR(v), `"ERROR`", IF(ABS(v-($want))<$tol, `"OK`", `"got `"&TEXT(v,`"0.00####`"))))"
}
function Same($id, $expr, $want) {
    Check $id "=LET(v, $expr, IF(ISERROR(v), `"ERROR`", IF(v=`"$want`", `"OK`", `"got [`"&v&`"]`")))"
}

$dv = "nb.DiminishingValue$L"
$pc = "nb.PrimeCost$L"
$ga = "nb.GSTAdd$L"
$ge = "nb.GSTExtract$L"
$fy = "nb.FinancialYear$L"

# --- Depreciation: a schedule must always sum to cost, whatever the effective life.
# ATO effective lives are frequently fractional (3 1/3, 6 2/3, 13 1/3), and a life of
# 2 years or less drives the diminishing-value rate to its 100% cap.
foreach ($life in '1', '1.5', '2', '2.5', '3', '10/3', '4', '5', '20/3', '8', '10', '40/3', '15', '20', '25', '40') {
    $tag = $life.Replace('/', 'over')
    Near "DV sums to cost, life $life"  "SUM($dv(1000,$life))"  '1000'
    Near "PC sums to cost, life $life"  "SUM($pc(1000,$life))"  '1000'
    Near "DV period count, life $life"  "COLUMNS($dv(1000,$life))"  "MAX(1,ROUNDUP($life,0))"
    Near "PC period count, life $life"  "COLUMNS($pc(1000,$life))"  "MAX(1,ROUNDUP($life,0))"
    Near "DV never negative, life $life" "SUMPRODUCT(--($dv(1000,$life)<0))" '0'
    Near "PC never negative, life $life" "SUMPRODUCT(--($pc(1000,$life)<0))" '0'
}
# The worked examples printed in each function's own inline help.
Same 'DV documented example' "TEXTJOIN(`",`",FALSE,TEXT($dv(1000,5),`"0.00`"))" '400.00,240.00,144.00,86.40,129.60'
Same 'PC documented example' "TEXTJOIN(`",`",FALSE,TEXT($pc(1000,5),`"0.00`"))" '200.00,200.00,200.00,200.00,200.00'
# Diminishing value must fall period on period. The final period is excluded because it
# carries the residual write-off, which is deliberately larger than the period before it.
Near 'DV declines over time' "SUMPRODUCT(--(INDEX($dv(1000,10),1,SEQUENCE(,8))<INDEX($dv(1000,10),1,SEQUENCE(,8)+1)))" '0'
Near 'DV residual is the tail' "INDEX($dv(1000,5),1,5)-129.6" '0'
Same 'DV help with no args'  "INDEX($dv(),1,1)" 'FUNCTION:'
Same 'PC help with no args'  "INDEX($pc(),1,1)" 'FUNCTION:'

# --- GST
Near 'GST add, default rate'      "$ga(100)"          '110'
Near 'GST add, blank rate cell'   "$ga(100,Z1)"       '110'
Near 'GST add, explicit zero'     "$ga(100,0)"        '100'
Near 'GST add, explicit 15%'      "$ga(100,0.15)"     '115'
Near 'GST add, array'             "SUM($ga({100;250}))" '385'
Near 'GST extract, default rate'  "$ge(110)"          '10'
Near 'GST extract, blank rate'    "$ge(110,Z1)"       '10'
Near 'GST extract, explicit zero' "$ge(110,0)"        '0'
Near 'GST extract, array'         "SUM($ge({110;220}))" '30'
Near 'GST round trip'             "$ge($ga(100))"     '10'
Same 'GST add keeps blanks blank'     "`"[`"&$ga(Z1)&`"]`"" '[]'
Same 'GST extract keeps blanks blank' "`"[`"&$ge(Z1)&`"]`"" '[]'
Same 'GST add help with no args'      "INDEX($ga(),1,1)" 'FUNCTION:'
Same 'GST extract help with no args'  "INDEX($ge(),1,1)" 'FUNCTION:'

# --- Financial year. The headline use is labelling a COLUMN of dates, so the array
# path matters more than the scalar one.
Same 'FY 30 Jun 2026'          "$fy(DATE(2026,6,30))"  'FY2026'
Same 'FY 1 Jul 2026'           "$fy(DATE(2026,7,1))"   'FY2027'
Same 'FY 31 Dec 2026'          "$fy(DATE(2026,12,31))" 'FY2027'
Same 'FY 29 Feb 2028'          "$fy(DATE(2028,2,29))"  'FY2028'
Same 'FY blank cell'           "`"[`"&$fy(Z1)&`"]`""    '[]'
Same 'FY array spans boundary' "TEXTJOIN(`"|`",FALSE,$fy(DATE(2026,6,30)+{0;1;2}))" 'FY2026|FY2027|FY2027'
Same 'FY array, whole year'    "TEXTJOIN(`"|`",FALSE,$fy(DATE(2026,1,1)+{0;200;400}))" 'FY2026|FY2027|FY2027'
Same 'FY January start'        "$fy(DATE(2026,7,1),1)"  'FY2026'
Same 'FY January start, Dec'   "$fy(DATE(2026,12,31),1)" 'FY2026'
Same 'FY April start, 31 Mar'  "$fy(DATE(2026,3,31),4)" 'FY2026'
Same 'FY April start, 1 Apr'   "$fy(DATE(2026,4,1),4)"  'FY2027'
Same 'FY December start'       "$fy(DATE(2026,12,1),12)" 'FY2027'
Same 'FY array, January start' "TEXTJOIN(`"|`",FALSE,$fy(DATE(2026,6,30)+{0;1},1))" 'FY2026|FY2026'
Same 'FY help with no args'    "INDEX($fy(),1,1)" 'FUNCTION:'

# --- the Australian tax worksheet, as the reader actually sees it
$au = "'Australian tax'!"
Near 'Sheet: DV total equals cost' "${au}B10-${au}A6" '0'
Near 'Sheet: PC total equals cost' "${au}B11-${au}A6" '0'
Near 'Sheet: GST added'            "${au}B16" '1100'
Near 'Sheet: GST extracted'        "${au}B19" '100'
Same 'Sheet: FY 30 Jun 2026'       "${au}B23" 'FY2026'
Same 'Sheet: FY 1 Jul 2026'        "${au}B24" 'FY2027'
Same 'Sheet: FY 15 Aug 2026'       "${au}B25" 'FY2027'
Same 'Sheet: FY 31 Dec 2026'       "${au}B26" 'FY2027'

# --- Debt sculpting. The debt module has never had a numeric check of any kind, which is
# how DebtSculptVariableLRV shipped from v1.2.0 to v2.2.0 adding each period's interest
# back into a balance the same period's cash had already paid. These are balance
# identities rather than expected figures: a schedule that satisfies all of them cannot
# be double-counting, whatever the inputs.
$lrv = "nb.DebtSculptVariableLRV$L"
$dsf = "nb.DebtSculptFixed$L"
$dsv = "nb.DebtSculptVariable$L"
$ilrv = "nb.InterestLRV$L"

# 1,000 drawn in period 1, 300 of cash a period, 1.2 times covered, 6% a year, 5 years.
# Rows are opening balance, interest, MINUS the principal repayment, closing balance.
$sched = "$lrv(, {1000,0,0,0,0}, {300,300,300,300,300}, {1.2,1.2,1.2,1.2,1.2}, {0.06,0.06,0.06,0.06,0.06}, 12)"

Near 'Debt: repayments retire the principal exactly' "SUM(INDEX($sched, 3, 0))" '-1000' '0.0000001'
Near 'Debt: schedule ends at zero'                   "INDEX($sched, 4, 5)"      '0'     '0.0000001'
Near 'Debt: closing = opening less repayment' `
     "SUMPRODUCT(ABS(INDEX($sched,4,0) - INDEX($sched,1,0) - INDEX($sched,3,0)))" '0' '0.0000001'
# No new debt after period 1, so every opening must be the previous closing. Line the two
# rows up by dropping the first opening and the last closing rather than by position.
Near 'Debt: each opening = the last closing' `
     "SUMPRODUCT(ABS(DROP(INDEX($sched,1,0),,1) - DROP(INDEX($sched,4,0),,-1)))" '0' '0.0000001'
Near 'Debt: cash used never exceeds CFADS/DSCR' `
     "MAX(INDEX($sched,2,0) - INDEX($sched,3,0)) - 250" '0' '0.0000001'
Near 'Debt: balance never goes negative'             "MIN(0, MIN(INDEX($sched, 4, 0)))" '0' '0.0000001'

# Cash well over the debt must clear it in one period, not leave twice the interest behind.
Near 'Debt: surplus cash clears the balance' `
     "INDEX($lrv(, {1000,0}, {1800,1800}, {1.2,1.2}, {0.06,0.06}, 12), 4, 1)" '0' '0.0000001'
# No cash must capitalise one period of interest, not two.
Near 'Debt: no cash capitalises interest once' `
     "LET(s, $lrv(, {1000,0}, {0,0}, {1.2,1.2}, {0.06,0.06}, 12), INDEX(s,4,1) - 1000 - INDEX(s,2,1))" '0' '0.0000001'

# The other two sculpting functions pay the whole debt service, so their balances differ,
# but the same roll-forward has to hold: closing = opening + interest - debt service.
foreach ($fn in $dsf, $dsv) {
    $arg = if ($fn -eq $dsf) { '1.2, 0.06' } else { '{1.2,1.2,1.2}, {0.06,0.06,0.06}' }
    $s2 = "$fn(, {1000,0,0}, {300,300,300}, $arg, 12)"
    Near "Debt: roll-forward holds for $fn" `
         "SUMPRODUCT(ABS(INDEX($s2,4,0) - INDEX($s2,1,0) - INDEX($s2,2,0) - INDEX($s2,3,0)))" '0' '0.0000001'
}

# The row a reader is told to label. Only the LRV function reports a principal repayment.
Same 'Debt: LRV row 3 is principal repayment' "INDEX($lrv(), 4, 2)" 'Principal repayments'
Same 'Debt: fixed row 3 is debt service'      "INDEX($dsf(), 4, 2)" 'Debt service (interest and principal)'
Same 'Debt: variable row 3 is debt service'   "INDEX($dsv(), 4, 2)" 'Debt service (interest and principal)'

# The debt module's one worked example.
Near 'Debt: InterestLRV worked example' "$ilrv(6666.37, 3.50, 90000, 0.03/12)" '222.90' '0.005'

# --- PeriodStart. A period anchored on a month end is defined by EDATE: the anchor's day
# of the month where the target month has one, the month's own end where it does not, so
# 31 January monthly runs 31 Jan, 28 Feb, 31 Mar. Up to v2.3.0 the function walked the
# calendar and took a single day back off whatever overflowed, which put 5 March in a
# period starting 2 March. Each grid check compares 180 dates at 13-day steps against
# that EDATE schedule, which is stated here rather than borrowed from the function.
$ps = "nb.PeriodStart$L"
$psWant = 'MAP(ds, LAMBDA(d, LET(s, EDATE(a, SEQUENCE(121,1,-60) * m), MAX(IF(s <= d, s, 0)))))'

foreach ($day in '31', '30', '29', '28', '15', '1') {
    foreach ($m in '1', '3', '12') {
        Near "PeriodStart: anchor 2026-01-$day, $m-month periods" `
             ("LET(a, DATE(2026,1,$day), m, $m, ds, DATE(2024,1,1) + SEQUENCE(1,180,0,13), " +
              "SUMPRODUCT(--(MAP(ds, LAMBDA(d, $ps(a, m, d))) <> $psWant)))") '0'
    }
}
Near 'PeriodStart: anchor 29 Feb 2028, monthly' `
     ("LET(a, DATE(2028,2,29), m, 1, ds, DATE(2024,1,1) + SEQUENCE(1,180,0,13), " +
      "SUMPRODUCT(--(MAP(ds, LAMBDA(d, $ps(a, m, d))) <> $psWant)))") '0'

# The reported case, its neighbours on either side, and the documented example.
Same 'PeriodStart: 31 Jan monthly, 5 Mar' `
     "TEXT($ps(`"31/1/2026`", 1, `"5/3/2026`"), `"yyyy-mm-dd`")" '2026-02-28'
Same 'PeriodStart: 31 Jan monthly, 27 Feb' `
     "TEXT($ps(`"31/1/2026`", 1, `"27/2/2026`"), `"yyyy-mm-dd`")" '2026-01-31'
Same 'PeriodStart: 31 Jan monthly, 28 Feb' `
     "TEXT($ps(`"31/1/2026`", 1, `"28/2/2026`"), `"yyyy-mm-dd`")" '2026-02-28'
Same 'PeriodStart: 31 Jan monthly, 31 Mar' `
     "TEXT($ps(`"31/1/2026`", 1, `"31/3/2026`"), `"yyyy-mm-dd`")" '2026-03-31'
Same 'PeriodStart: 31 Dec monthly, 29 Feb 2028' `
     "TEXT($ps(`"31/12/2025`", 1, `"29/2/2028`"), `"yyyy-mm-dd`")" '2028-02-29'
Same 'PeriodStart: 30 Nov monthly, 15 Feb' `
     "TEXT($ps(`"30/11/2025`", 1, `"15/2/2026`"), `"yyyy-mm-dd`")" '2026-01-30'
Same 'PeriodStart: documented example' `
     "TEXT($ps(`"2026-01-01`", 3, `"2026-04-15`"), `"yyyy-mm-dd`")" '2026-04-01'
Same 'PeriodStart: date before the anchor' `
     "TEXT($ps(DATE(2026,1,1), 3, DATE(2025,11,15)), `"yyyy-mm-dd`")" '2025-10-01'
Same 'PeriodStart: date on the anchor' `
     "TEXT($ps(DATE(2026,1,15), 3, DATE(2026,1,15)), `"yyyy-mm-dd`")" '2026-01-15'
Same 'PeriodStart: help with no args' "INDEX($ps(),1,1)" 'FUNCTION:'

# --- TimelineOffset. The interval is read off the timeline's first two dates and, up to
# v2.3.0, converted to whole months and divided by. A daily, weekly or fortnightly
# timeline rounds to no months at all, so every one of them returned #DIV/0!.
$to = "nb.TimelineOffset$L"
$daily = 'DATE(2026,1,1) + SEQUENCE(1,60,0,1)'
$weekly = 'DATE(2026,1,1) + SEQUENCE(1,60,0,7)'
$fortnightly = 'DATE(2026,1,1) + SEQUENCE(1,60,0,14)'

Near 'TimelineOffset: daily, 10 days in'        "$to(DATE(2026,1,11), $daily)"        '10'
Near 'TimelineOffset: daily, 3 days before'     "$to(DATE(2025,12,29), $daily)"       '-3'
Near 'TimelineOffset: weekly, 20 days in'       "$to(DATE(2026,1,21), $weekly)"       '2'
Near 'TimelineOffset: weekly, 3 days before'    "$to(DATE(2025,12,29), $weekly)"      '-1'
Near 'TimelineOffset: weekly, one whole period before' "$to(DATE(2025,12,25), $weekly)" '-1'
Near 'TimelineOffset: fortnightly, 30 days in'  "$to(DATE(2026,1,31), $fortnightly)"  '2'

# A sub-monthly period is a fixed number of days, so the offset is the day difference
# floored by that count. 200 dates at 3-day steps, starting three months before the
# timeline does, so the negative side is covered too.
foreach ($tl in @('daily', $daily, '1'), @('weekly', $weekly, '7'), @('fortnightly', $fortnightly, '14')) {
    Near "TimelineOffset: $($tl[0]) counts whole periods" `
         ("LET(t, $($tl[1]), ds, DATE(2025,10,1) + SEQUENCE(1,200,0,3), " +
          "SUMPRODUCT(--(MAP(ds, LAMBDA(d, $to(d, t))) <> " +
          "MAP(ds, LAMBDA(d, INT((d - INDEX(t,1)) / $($tl[2])))))))") '0'
}

# and the month path must answer exactly what it always did, month ends included.
foreach ($tl in @('monthly', '1', '15'), @('quarterly', '3', '15'), @('yearly', '12', '15'),
                @('monthly off a 31st', '1', '31')) {
    Near "TimelineOffset: $($tl[0]) unchanged" `
         ("LET(b, DATE(2026,1,$($tl[2])), mpp, $($tl[1]), t, EDATE(b, SEQUENCE(1,40,0,mpp)), " +
          "ds, DATE(2024,1,1) + SEQUENCE(1,200,0,11), " +
          "SUMPRODUCT(--(MAP(ds, LAMBDA(d, $to(d, t))) <> " +
          "MAP(ds, LAMBDA(d, LET(ks, SEQUENCE(121,1,-60), s, EDATE(b, ks * mpp), " +
          "MAX(IF(s <= d, ks, -9999))))))))") '0'
}
Same 'TimelineOffset: help with no args' "INDEX($to(),1,1)" 'FUNCTION:'

# The two worked examples printed in the function's own help, which a reader is meant to
# copy. Neither could be run as printed: the call was missing its two closing brackets.
Near 'TimelineOffset: documented example, inside the timeline' `
     "$to(`"15/2/2026`", EDATE(`"1/1/2026`", SEQUENCE( , 12, 0)))" '1'
Near 'TimelineOffset: documented example, before the timeline' `
     "$to(`"15/2/2025`", EDATE(`"1/1/2026`", SEQUENCE( , 12, 0)))" '-11'

$xl = $null; $wb = $null; $tmp = $null; $exit = 0

# Excel rejects incoming COM calls while it is mid-calculation (RPC_E_CALL_REJECTED),
# so anything that can arrive at a busy moment gets retried with a backoff.
function Invoke-Excel([scriptblock]$op, [int]$tries = 12) {
    for ($i = 1; $i -le $tries; $i++) {
        try { return & $op }
        catch {
            if ($i -eq $tries) { throw }
            Start-Sleep -Milliseconds (250 * $i)
        }
    }
}

# Excel's COM server is shared: New-Object attaches to an already running EXCEL.EXE, and
# this script calls Quit() when it finishes. Refuse to run rather than close someone's
# open workbooks underneath them.
# A previous run's instance can take a moment to exit, so give it a few seconds first.
for ($w = 0; $w -lt 10 -and @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0; $w++) {
    Start-Sleep -Milliseconds 500
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Output 'Excel is already running. Close it first: this test drives Excel over COM and quits it when done.'
    exit 2
}

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1; $xl.AskToUpdateLinks = $false; $xl.EnableEvents = $false

    # Excel's COM server intermittently refuses the first call after a prior instance quits.
    $wb = Invoke-Excel { $xl.Workbooks.Open($Path, 0, $false) }
    if ($null -eq $wb) { throw "could not open $Path" }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    Invoke-Excel { $xl.CalculateFullRebuild() }
    $sw.Stop()

    Write-Output ("Excel {0} build {1}" -f $xl.Version, $xl.Build)
    Write-Output ("{0}: {1} sheets, {2} defined names, full rebuild {3:N2}s" -f
        (Split-Path $Path -Leaf), $wb.Worksheets.Count, $wb.Names.Count, $sw.Elapsed.TotalSeconds)
    Write-Output ''

    # --- every cell in the workbook, after a real recalculation
    $errors = @(); $formulaCount = 0
    foreach ($ws in $wb.Worksheets) {
        $used = $ws.UsedRange
        try { $formulaCount += $used.SpecialCells(-4123).Count } catch {}
        foreach ($kind in -4123, 2) {
            try { $bad = $used.SpecialCells($kind, 16) } catch { continue }
            foreach ($cell in $bad) {
                $errors += ("{0}!{1} = {2}   <= {3}" -f $ws.Name, $cell.Address($false, $false),
                            $cell.Text, $cell.Formula2)
            }
        }
    }
    Write-Output ("Recalculated cells: {0} formulas, {1} in error" -f $formulaCount, $errors.Count)
    if ($errors.Count) { $errors | ForEach-Object { Write-Output ("  " + $_) }; $exit = 1 }

    # --- assertions on a scratch sheet.
    # Written and read as whole ranges: a call per cell gets rejected while Excel is
    # busy recalculating, and 130 round-trips is slow besides.
    $tmp = Invoke-Excel { $wb.Worksheets.Add() }
    Invoke-Excel { $tmp.Name = 'zz_selftest' }
    Invoke-Excel { $tmp.Range('Z1:Z3').ClearContents() }   # the deliberate blank cells

    $grid = New-Object 'object[,]' $checks.Count, 1
    for ($i = 0; $i -lt $checks.Count; $i++) { $grid[$i, 0] = $checks[$i].f }
    $addr = 'A1:A' + $checks.Count
    # Hold calculation off until every probe is in place, so Excel is not recalculating
    # the sheet underneath the write.
    Invoke-Excel { $xl.Calculation = -4135 }        # xlCalculationManual
    Invoke-Excel { $tmp.Range($addr).Formula2 = $grid }
    Invoke-Excel { $xl.Calculation = -4105 }        # xlCalculationAutomatic
    Invoke-Excel { $xl.CalculateFullRebuild() }
    # leading comma stops PowerShell flattening the 2-D range value on the way out
    $vals = Invoke-Excel { , $tmp.Range($addr).Value2 }

    $failed = @()
    for ($i = 0; $i -lt $checks.Count; $i++) {
        # Value2 hands back a 1-based 2-D array; GetValue avoids PowerShell's index parsing
        $got = if ($vals.Rank -eq 2) { [string]$vals.GetValue($i + 1, 1) } else { [string]$vals[$i] }
        if ($got -ne 'OK') { $failed += ("{0}: {1}" -f $checks[$i].id, $got) }
    }
    Write-Output ("Assertions: {0} run, {1} failed" -f $checks.Count, $failed.Count)
    if ($failed.Count) { $failed | ForEach-Object { Write-Output ("  FAIL " + $_) }; $exit = 1 }

    Invoke-Excel { $tmp.Delete() }; $tmp = $null
    Invoke-Excel { $wb.Close($false) }; $wb = $null

    Write-Output ''
    Write-Output $(if ($exit -eq 0) { 'PASS' } else { 'FAIL' })
}
catch {
    Write-Output ('FATAL: ' + $_.Exception.Message)
    Write-Output $_.ScriptStackTrace
    $exit = 1
}
finally {
    if ($tmp) { try { $tmp.Delete() } catch {} }
    if ($wb)  { try { $wb.Close($false) } catch {} }
    if ($xl)  { try { $xl.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $exit
