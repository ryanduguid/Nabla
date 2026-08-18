# Changelog

## Unreleased

- **`nb.OverLapDaysλ` compared its dates as text.** It converts all four arguments, so that
  a date written as text becomes a serial number, and then compared the raw arguments
  anyway. The four conversions were never read. Two text dates therefore compared as
  strings, which ranks `"7/1/2025"` after `"17/1/2025"`, so the function picked the wrong
  start and end and the subtraction that follows coerced them back to dates. Its own third
  example claimed 12 shared days for two January 2025 periods that share 2. The comparison
  now reads the converted values and the example says 2. Numbers
  and real dates are unaffected, because converting one returns it unchanged, which is why
  the rental schedule on the function's own demonstration sheet was right all along.
- **`nb.Periodsλ` could not return a negative.** Its description promises "Returns negative
  values if Date1 is after Date2" and two of its four examples show one, but the difference
  was floored at 1 before `SIGN` saw it, so the sign was always `+1`.
  `=nb.Periodsλ("15/5/2025", "31/3/2025")` returned 1 where it should return -1. The sign
  now comes from the difference itself. Equal dates give a sign of 0 rather than 1, which
  changes nothing: the interval count is 0 either way.
- **`nb.VDBλ` ignored its `No_Switch` argument.** It declares the argument and defaults it
  to `FALSE`, then called Excel's `VDB` without it, so passing `TRUE` did nothing. Its help
  demonstrates the function with `No_Switch` set to `TRUE` and printed the `FALSE` answer,
  which is why nothing looked wrong. `=nb.VDBλ(1000, 100, 5, 1.5, TRUE)` now gives 300.00,
  210.00, 147.00, 102.90 and 72.03, against the 121.50, 121.50 tail that switching to
  straight line produces. Calls that omit the argument are unaffected, including
  `nb.Depreciateλ`'s `VDB` method, which never passed one.

  All three came from the same external audit of v1.2.6, and each was confirmed in Excel
  before it was fixed. Two of them were hidden by their own worked example, which printed
  the answer the bug produces. Every other cached value in the workbook is byte-identical
  to v2.0.0; the one cell that moved is the corrected example result on
  `nb.OverLapDaysλ`'s sheet.

  Three defects in the same functions are **not** fixed here, because each changes results
  for anyone already relying on them and none is a one-line correction. `nb.Periodsλ`
  counts complete intervals where its examples count boundaries crossed, so its forward
  example returns 1 against the 2 it claims, and its `W` example 52 against 53. It also
  returns `#VALUE!` for the range arguments its help says it takes, which is what the
  demonstration sheet has been caching. And four more functions ignore their date
  conversions the way `nb.OverLapDaysλ` did: `nb.PeriodLabelλ`, `nb.ScheduleRatesλ`,
  `nb.ScheduleValuesλ` and `nb.Timelineλ`, six dead conversions between them, all in Dates.

- **The depreciation helpers no longer claim to be ATO methods.** `nb.DiminishingValueλ` described itself as the ATO 200% diminishing value method and `nb.PrimeCostλ` as the ATO prime cost method. Neither is one. They take a cost and an effective life and nothing else: no acquisition date, no income year, no days held, no disposal. `nb.DiminishingValueλ` also writes the entire undeducted residual off in its final period, so for a cost of 1,000 over five years it returns 400, 240, 144, 86.40 and 129.60, where a diminishing-balance calculation would deduct 51.84 that year and carry the rest forward. The schedules are unchanged and still useful for modelling; what changes is that they are now described as modelling schedules rather than tax calculations, in the function help, the method codes on `nb.Depreciateλ`, the Data Validation sheet, the Australian tax worksheet and the README. The worksheet now says so on its face, and the README says plainly not to use them to prepare a return.

  Raised by an external audit of v1.2.6. Its arithmetic checks out: the fifth-year deduction under a plain diminishing-balance calculation is 51.84, not the 129.60 residual the function returns.

- **Three worked examples called a different function than the one they document.** `nb.RollingMinλ`'s example called `nb.RollingMaxλ`, `nb.RollingSumλ`'s called `nb.RollingMinλ`, and `nb.ScheduleValuesByItemsλ`'s called `nb.ScheduleRatesByItemsλ`. The examples are the lines a reader copies, so copying one as printed ran a different function. Each claimed result was correct for the function it wrongly called, which is exactly why nothing looked wrong.
- **`nb.RollingAvgλ` claimed a result belonging to `nb.RollingSumλ`,** the same copy going the other way: its call was right and its answer, `1,3,6,9,12`, was the running total rather than the running average. Excel gives `1,1.5,2,3,4`, which is what it now says.
- **`nb.Amortiseλ`'s help pointed at a function that does not exist,** `LableAmortiseλ`, a transposition of `LabelAmortiseλ`.
- **The build's help corrections can no longer reach a live formula.** Correcting help text also refreshes any copy of that help already spilled and cached on a demonstration sheet, and that refresh was a blanket text replace across the whole worksheet. Fixing `nb.ScheduleValuesByItemsλ`'s example rewrote a real formula on a neighbouring sheet, changing which function it called and spilling `#SPILL!` across it. The Excel gate caught it before it shipped. The refresh is now confined to cached values, never formulas, and a parameter label must match a whole cell rather than appear anywhere inside one.
- **`tools/verify_signatures.py` now reads the worked examples too.** It already checked the signature and the parameter table; it now also requires that a function's examples call that function, and that every function named anywhere in a help is one the library declares. Both are pure text, so both run in CI. Run against the source before these fixes it reports all four defects above by name. 76 example blocks are checked.

- **`previous_name` is now pinned to a released baseline rather than to the build's own intermediate names.** The column shipped correct in v2.0.0, but it was derived from the `nabla.<module>.` names the build uses internally before it flattens them, which happen to match what v1.2.6 shipped only because nothing has been added since. The next function added would have published a predecessor that no release ever carried, and the build's count check could not have caught it, because both sides of that count come from the same build. The 130 names v1.2.6 shipped are now recorded in `tools/released-names-v1.2.6.txt`, and a function whose predecessor is not in that file records nothing rather than a plausible-looking guess. A baseline name that no function claims now stops the build and says which name, because a function disappearing without a forwarding address is the one thing this column exists to prevent.

### Added

- `tools/verify_previous_names.py`, a fifth gate and the fourth in CI. The build's asserts guarantee the column is right when it is generated, but the committed file is what people read, so this checks the published index against the published baseline. Exercised against four deliberate breakages: a fabricated predecessor, a dropped function, one old name claimed twice, and two unrelated functions with their predecessors swapped. It names the offending function in each case rather than reporting a count.

  That last case is why the check does more than count. Claiming every baseline name exactly once proves the mapping is a bijection, which is not the same as proving it is the right one: swap two functions' predecessors and every count still balances while a reader is sent to the wrong function. The rename only ever appended to a bare name, adding a `B`, `E` or `U` tag or a module word, so each new bare name must begin with the old one. All 130 satisfy that, and a swap does not.

## v2.0.0, 18 August 2026, one namespace

**This release renames every function and breaks every formula written against v1.2.x.**
`nabla.f.Amortiseλ` is now `nb.Amortiseλ`. Models built on the old names keep working
only if they keep the old workbook, which stays available at the `v1.2.6` tag.

Six module prefixes became one. The gain is five fewer characters on every call and, more
to the point, autocomplete that works: typing `=nb.` narrows to this library instead of
requiring you to remember which of six modules a function lived in first.

Of the 130 functions, 106 change prefix and nothing else. Seventeen bare names existed in
more than one module: the fuller implementation keeps the plain name and the other 19 take
a one-letter tag, `B` for debt, `E` for essentials, `U` for utilities. The five About
tables take words, so `nabla.f.Aboutλ` is now `nb.AboutFinancialλ` and the bare name
`Aboutλ` is gone.

### Finding a renamed function

`functions.csv` now carries a `previous_name` column, one entry for every function, so
every old name has a documented replacement. The full table is in the release notes.

Checked against the index v1.2.6 shipped: 130 names then, 130 now, one to one, with no
name dropped and none invented.

### Also

- The `src/` round-trip has a known limitation, now documented in the README. Importing a
  module through the Advanced Formula Environment recreates its functions under the
  container's name, `Dates.CountDOWλ` rather than `nb.CountDOWλ`, because the AFE takes
  the prefix from the container. The workbook is the authority for the `nb.` names.
- `nabla.xlsx` is attached to this release as a downloadable asset, which earlier releases
  did not offer.
- The entry below for the namespace work said the tag scheme covered 23 functions. It
  covers 19; the other 4 were the About renames, counted twice.

## Unreleased

- **`tools/verify_signatures.py` now reads the parameter tables too.** It checked the FUNCTION line and stopped there, which is why the table defects fixed in v1.2.6 had to be found by hand. The two are independent pieces of hand-written text, so the checks are independent: 117 signatures and 122 parameter tables, the extra five being the debt module's functions, which have never carried a FUNCTION line but do carry tables. A row whose label ends in `!` is an aside rather than a parameter, and the internal `DoNotUse` counter may be documented or omitted. Run against v1.2.5 it reports all four table divergences; against v1.2.3, twenty-four problems across both checks.

## v1.2.6, 18 Aug 2026, parameter tables that describe their own function

Every function's help repeats its parameters as a table below the signature. v1.2.4 corrected the signatures; this corrects the tables, which are a separate piece of hand-written text and had drifted on their own. Every parameter table in the library now lists exactly what its function declares.

- **`nabla.r.EquityRatioλ` documented a different function's arguments.** Its parameter table listed `OperatingIncome` and `InterestExpenses`, which belong to `InterestCoverageRatioλ`, while the function takes `ShareholdersEquity`, `TotalAssets` and `IntangibleAssets`. This is the second half of the copy that gave it the wrong name until v1.2.3: the name and the table came across together and only the name was corrected then. Its three real parameters had never been described at all. They are now, in the wording its neighbours already use for the same quantities.
- **`nabla.r.EquityMultiplierλ` dropped a word from its second parameter.** Its table called it `ShareholdersEquity`; the LAMBDA and the signature above it both say `TotalShareholdersEquity`. `ShareholdersEquity` is what the neighbouring `DebtToEquityRatioλ` genuinely takes, which is where the shortened name came from.
- **`IsBetweenλ` documented a parameter it does not have, and described the wrong limit.** Its parameter table called the second argument `Lo` where the LAMBDA and the signature above it both say `Low`, and the row for `Hi` read "The lower limit that the value must be less than", copied from the row above it, so the function's own help described its upper bound as a lower one. The `Inclusive` row referred to `Lo` as well. All three are corrected in the Essentials and Utilities copies, which are clones of each other; the Dates module's own `IsBetweenλ` has always had both right and its wording, "the higher limit that the value must be less than", is what the other two now use. Corrected in the module source, the defined name, and the demonstration sheet that had cached the old text.

## v1.2.5, 18 Aug 2026, a colon back where it belongs

- **`EXAMPLES       :` in four help tables.** The label column is written as the label, then padding, then the arrow that separates it from the second column, so every arrow lines up. Four labels put the colon after the padding instead of before it, one character wider than every other row, and since the help is built with `TRIM()` the reader saw `EXAMPLES :` with the colon adrift. It affected `IsBetweenλ` and `IsInListλ`, each in both the Essentials and Utilities modules; those four are clones of one another and every other `EXAMPLES` label in the library was already correct. Corrected in the module source, the defined name, and the one demonstration sheet that had cached the old label.

## v1.2.4, 18 Aug 2026, help you can read

### `IsInListλ` built its help sideways

`TEXTSPLIT` takes the text, then a column delimiter, then an optional row delimiter. This one supplied a single delimiter: the arrow that should have separated the two columns was left concatenated onto the end of the text, and the pilcrow that should have ended each row became the column delimiter. Calling `nb.IsInListλ()` for help therefore returned one row of eleven columns and spilled sideways across the sheet instead of down it, with each label and its explanation run together in a single cell. The other 125 functions that build help this way supply both delimiters; these two were the only ones that did not. Both copies now do, and their help returns the 11-row, 2-column table it was always written to be, read back out of Excel to confirm it.

### Help that describes its own parameters

Every function's help opens with a signature, and repeats the same parameters as a table three rows below. Where the two disagreed, the table was right every time: it matches what the LAMBDA declares. Thirteen signatures did not, and four of them documented a neighbouring function's arguments outright, which is how the wrong-name defect fixed in v1.2.3 got in as well. Both were the same copy: name and parameter list came across together, and only the name was corrected.

| function | its signature said | its signature says now |
|---|---|---|
| `nb.CorkScrewReversalλ` | `Opening, Flow1, ...` | `Opening, ReversalFlags, Flow1, ...` |
| `nb.Movementλ` | `BeginningValue` | `BeginningValues` |
| `nb.LabelAmortiseλ` | `[LoanNames]` alone | all four parameters |
| `nb.Depreciateλ` | `[Factor]` | `[Factors]` |
| `nb.DBλ` | `[Month]` | `[Months]` |
| `nb.TimelineOffsetλ` | `ArrayStart` | `Date` |
| `nb.FilterContainsλ` | `Text` | `FilterByText` |
| `nb.QuickRatioλ` | `LiquidAssets` | `QuickAssets` |
| `nb.WorkingCapitalTurnoverRatioλ` | `CostOfGoodsSold, AverageInventory` | `NetAnnualSales, WorkingCapital` |
| `nb.DSCRλ` | `Totaldebtservice` | `TotalDebtService` |
| `nb.CashFlowMarginλ` | `NetIncome` | `CashFlowFromOperatingActivities` |
| `nb.PriceToBookRatioλ` | `BookValuePerShareBvps` | `BookValuePerShare` |
| `nb.PriceToCashRatioλ` | `SalesPerShare` | `OperatingCashFlowPerShare` |

Three parameter tables disagreed with their own LAMBDA too, and were corrected the same way: `LoanAPR` and `LoanTerm` in `LabelAmortiseλ`, and `TotaldebtService` in `DSCRλ`.

`nb.IsInListλ` and `nb.IsInListUλ` were the one case where the declaration was the odd one out. It shouts `LIST`, while the signature, the parameter table and one of the function's own two references all write `List`. Excel resolves identifiers case-insensitively, so the parameter is renamed to match the rest of the library rather than the help being made to shout back. Both functions were exercised in Excel afterwards, including the branch that calls `ISOMITTED()` on the renamed parameter.

Each correction is applied in three places: the module source `src/` is exported from, the defined name Excel installs, and the help already spilled and cached on the demonstration sheets. Five sheets carried a stale copy. Every corrected signature was then read back out of a running Excel rather than trusted from the file.

### Added

- `tools/verify_signatures.py` reads every function's help signature and compares it against the LAMBDA's own declaration, character for character, since case is exactly the kind of difference that goes unnoticed. It accounts for every declaration in every module and prints the tally, and fails if it parsed too few, because a checker that reads nothing passes everything. Square brackets are ignored: upstream declares every parameter optional so a function called with no arguments can return its own help, so the declaration says nothing about which arguments a caller may omit. Run against the previous release it reports all 15 divergences. Now runs in CI.

- **`FLow1` in the corkscrew signatures.** `nb.Corkscrewλ` and `nb.CorkScrewReversalλ` both spelled their second argument `FLow1` on the FUNCTION line of their help, with a capital L. The parameter table three rows below spelled it `Flow1`, and so did the LAMBDA, so anyone copying the signature was copying a name the function does not have. Corrected in the module source, in the defined name, in `functions.csv`, and in the help output already cached on the demonstration sheet, which would otherwise have kept showing the typo until something forced a recalculation. Read back out of Excel afterwards, both functions now report `( Opening, Flow1, ...)`.

## 2026-08-18, later still

### One namespace: every function is now `nb.`

Six module prefixes became one. `nabla.f.Amortiseλ` is now `nb.Amortiseλ`, which is five
fewer characters on every call and, more to the point, makes formula autocomplete useful:
typing `=nb.` narrows to this library instead of dumping 130 entries behind a prefix you
had to spell out first.

Collisions were the only real obstacle. Eighteen base names existed in more than one
module, forty-one functions in total. Nothing was dropped: the fuller implementation keeps the
plain name and the other takes a one-letter tag.

- `B` for debt, `E` for essentials, `U` for utilities: `nb.AmortiseBλ`, `nb.IsBetweenEλ`,
  `nb.SumRowsUλ` and so on, 19 functions in all.
- The five About tables take words rather than letters, because `nb.AboutRλ` tells a
  reader nothing: `nb.AboutDatesλ`, `nb.AboutEssentialsλ`, `nb.AboutFinancialλ`,
  `nb.AboutRatiosλ`, `nb.AboutUtilitiesλ`.
- The Utilities group is a copy of Essentials; 16 of its 17 functions are byte-identical
  after normalising the namespace. They were kept, tagged, rather than merged away.

The rename is done in the build, not by editing the artefact, so it cannot drift. That
also fixed something a direct edit would have missed: the Excel Labs project store holds
the LAMBDA source as base64 UTF-16 JSON, including a `projectNames` index of all 130
functions. Patching the workbook leaves 365 stale tokens in there. Building leaves none.

Also in this release:

- The Advanced Formula Environment modules are named for what they hold (Dates,
  Essentials, Financial, Ratios, Utilities, Debt) rather than by prefix, since six
  containers all called `nb` would collide. `src/` follows the same names.
- `functions.csv` keeps a `module` column, now filled from the group a function came from
  rather than parsed out of its name.
- `tools/verify_workbook.py` checks `nb.` tokens resolve, and the retired `nabla.<module>.`
  namespace joins the banned-token list so it cannot come back.
- Two Windows-only build bugs: the script died printing a λ to a cp1252 console, and
  `os.makedirs` raised on a bare output filename.

Verified by rebuilding from upstream and comparing every cell against the previous build:
85,647 cells, no numeric change, and the only text differences are the renamed functions,
the rewritten cover paragraph, the corrected About tables, and the product name capitalised
from `nabla` to `Nabla` in 29 places across worksheet strings, drawing callouts and the
document title.

## 2026-08-18, help that names itself

Four functions announced a neighbour's name on the first line of their own inline help, because each was written by copying that neighbour and the name was never updated. The description and parameter list underneath were always correct, so only the leading name was wrong, but calling `nabla.f.DDBλ()` for help and being told you are looking at `DBλ` is worse than no help at all.

| function | its help said | its help says now |
|---|---|---|
| `nabla.e.AvgColsλ` and `nabla.u.AvgColsλ` | `SumColsλ` | its own name |
| `nabla.f.CorkScrewReversalλ` | `Corkscrewλ` | `CorkScrewReversalλ` |
| `nabla.f.DDBλ` | `DBλ` | `DDBλ` |
| `nabla.r.EquityRatioλ` | `InterestCoverageRatioλ` | `EquityRatioλ` |

Applied to the module source and the defined name alike, so `src/` still reproduces what ships. No function in the library now names a different one in its own help. `functions.csv` picks the corrections up, since it reads its signature column from that line.

## 2026-08-18, importable sources

`src/` exists so the library can be read, diffed and loaded back into Excel. Testing that last part for the first time found that one module could not be loaded at all.

- **The five `nabla.debt.*` functions could not be imported.** The exporter wrote them by stripping the internal prefixes off the stored definitions, which destroyed two things. `_xlop.Name` marks an **optional** parameter; stripping the prefix leaves a required one, and since every one of those functions calls `ISOMITTED()` on its parameters, Excel rejected each definition outright. `[0]!` is the internal token for "a name in this workbook" and is not something you can type back in. The exporter now maps them to `[Name]` and to a bare reference. Measured in Excel: **nil of five accepted before, five of five after**, and swapping the published definitions into the workbook in place of the shipped ones leaves all 17,003 numeric cells identical.
- **`nabla.f.SumDepreciateλ` shipped a later revision than its own published source**, inherited from upstream: the installed function carries a blank help row and a different, behaviour-identical test for its omitted argument. The source is brought up to the version that ships.

### Added

- `tools/verify_sources.py` compares every function in `src/` against the defined name that ships, and now runs in CI. It maps the four conventions that separate the stored form from the typed form rather than ignoring them: the `_xlfn.`/`_xlpm.`/`_xlws.` markers, `_xlop.Name` against `[Name]`, `[0]!Name` against a bare reference, and `SINGLE(x)` against `@x`. Mapping the parameter marker rather than stripping it is the point: stripping is what hid the defect above. Run against the previous release it reports all six divergences.

## 2026-08-18, TOC filter

- **The table of contents opened filtered.** Upstream saved it with the Type slicer restricted to `Worksheet`, so 16 of the 66 entries, every one describing a table, were hidden on open with nothing to indicate they existed. The filter criteria and the row visibility stored alongside them also disagreed, because the row retyped from Worksheet to Function in the first round kept its old visibility. Both are cleared: the workbook now opens showing all 66 entries with every slicer button selected.
- The slicer itself was exercised in Excel and was never at fault. Each of its three buttons filters exclusively and correctly (1 Function, 16 Table, 49 Worksheet) and clearing restores all 66. The `<autoFilter>` element is retained, since the slicer binds to it; it simply carries no criteria now.

## 2026-08-18, later

### Run in Excel for the first time

Every prior release was reasoned about statically. Opening `nabla.xlsx` in Excel 365 and executing the functions found four defects that no amount of file inspection would have caught. The workbook itself recalculated cleanly: 1,129 formulas, no error cells.

- `DiminishingValueλ` returned `#NUM!` for an effective life of two years or less. Capping the rate at 100% drives `(1-Rate)` to zero, and Excel evaluates `0^0` as `#NUM!` rather than 1. Two years is an ordinary ATO effective life, so the function was unusable for a common case.
- `DiminishingValueλ` and `PrimeCostλ` silently under-depreciated any fractional effective life. `SEQUENCE` and `EXPAND` truncate a length of 6⅔ to six periods, and the residual write-off tested for the final period by comparing against the life rather than the period count, so it never fired. A 1,000 asset over 6⅔ years wrote off 882.34 and 899.96 respectively instead of 1,000, with no error shown. ATO effective lives are routinely fractional.
- `FinancialYearλ` gave the wrong answer for a range of dates, which is its headline use. `AND` is an aggregate: it collapsed the whole column to a single true or false, so every date inherited the first one's financial year. Labelling a column spanning 30 June produced one financial year for all of it.
- The published module source in `src/` had drifted from the compiled function. `FinancialYearλ` shipped one definition inside the workbook and a different, older one in `src/nabla.d.txt`. Anyone importing the source got a function that behaved differently from the one being documented.

### Fixes

- `DiminishingValueλ` guards the first period against `0^0`, derives its period count with `ROUNDUP` so a part year gets its own period, and writes the residual off against that count. Schedules now sum to cost for every life tested from one to forty years, whole or fractional.
- `PrimeCostλ` derives its period count the same way and puts the part-year remainder in the final period.
- `FinancialYearλ` multiplies instead of using `AND`, so each date is evaluated on its own.
- `GSTAddλ` and `GSTExtractλ` return blank for a blank amount, so a part-filled column no longer fills with zeros.
- The compiled name and the readable module source are now generated from one expression per function, and `src/` and `functions.csv` are exported from the built workbook. They regenerate byte-identical to the previous hand-written files apart from the fixes above, so the drift was confined to `FinancialYearλ`.
- Restored the missing help-column delimiter in `RollingMinλ`, an upstream defect that collapsed its signature row.

### Added

- `tools/excel_selftest.ps1`: opens the workbook in Excel, forces a full rebuild, fails on any error cell, and runs 138 assertions covering all sixteen effective lives, the GST and financial-year edge cases, and the demonstration worksheet. Every one of the four defects above fails this test on the previous build.
- The **Australian tax** worksheet now shows the totals written off by both depreciation methods, which is the property a bad effective life breaks, and labels its dates with a single spilled call so the array path is demonstrated rather than sidestepped by four separate scalar calls.

## 2026-08-18

### Verification round
- Added an **Australian tax** worksheet demonstrating all five Australian functions, with a table-of-contents entry and its own tab colour. They previously had no on-sheet presence at all.
- `DiminishingValueλ` now caps its rate at 100% (a life under two years previously wrote off more than cost) and writes the undeducted residual off in the final period, so a schedule sums to cost. The documented example changes to 400.00, 240.00, 144.00, 86.40, 129.60.
- `GSTAddλ` and `GSTExtractλ` no longer treat a blank Rate cell as nil GST; a blank now falls back to the 10% default.
- `FinancialYearλ` returns blank rather than FY1900 for empty cells, uses array-safe coercion, and handles a January financial-year start correctly.
- The `SEE ALSO` lines existed only in the compiled names, so an Excel Labs save would have silently deleted them. They are now in the module sources too.
- Fixed worked examples the two-year date shift had invalidated: `CountDOWλ` stated 2 where it now returns 3, and `PeriodLabelλ` stated 2023 results against a 2025 input.
- Fixed frozen sample data that produced uninformative demonstrations: the `Periodsλ` yearly row spanned one day short of a year and returned 0, and the onboarding dates were spaced 30 days apart against a twelve-year timeline so every customer landed in the first period.
- Replaced the 13 volatile `RANDARRAY` grids that the first performance pass missed, so no formula in the workbook is volatile except the sheet-name titles.
- Corrected upstream help defects: two missing column delimiters that collapsed a help row, and the misspelt `Liabilites` parameter.
- The GST helpers are listed under their own AUSTRALIAN TAX heading in the module index rather than inside the depreciation suite.
- Only the cover opens selected, and the table-of-contents columns were widened for the longer `nabla.*` names.

First nabla release, derived from the upstream 6 July 2024 workbook.

### Rebranding
- Renamed every function namespace, worksheet, AFE module and help reference from `BXD`/`BXE`/`BXF`/`BXR`/`BXU`/`BXLDebt` to the `nabla.*` scheme.
- Removed branded cover art, the cover video thumbnails and their YouTube link, Dropbox model links, and the source add-in credit line. A maths-citation video link in an IntOnIntλ source comment is retained as third-party credit.
- Help links to the upstream gists and site replaced with this repository's URL and relabelled "Repository"; author revision histories preserved; workbook creator metadata credits the original author.
- Workbook metadata retitled `nabla`.

### Australian English and conventions
- Spelling swept to Australian English in prose, help text and function names (`Amortiseλ`, `LabelAmortiseλ`, `SumAmortiseλ`, `AmortiseλDV`, amortisation terminology).
- Date number formats flipped to day-first (`d/m/yyyy`, `dd/mm/yyyy`, `dd/mm/yyyy h:mm`).
- Help examples and sample text dates rewritten day-first.
- Sample data Americanisms ported: currency label to AUD, `Apt.`/`Apartment` to `Unit`, `Wal*Art` to `Wool*Art`, household budget items to Pay/Home insurance/Strata levies/Petrol.
- Removed the foreign depreciation regime the upstream library carried: its function, method code, dispatch branch inside `Depreciateλ` and the special-case life, salvage and disposal handling that went with it, plus the foreign tax authority reference on the Data Validation sheet and the foreign accounting-standard paragraph on the `Depreciateλ` worksheet. The library is now Australian-only.
- Added five Australian functions, each with inline help, an AFE source module and a Name Manager description: `nabla.f.DiminishingValueλ` (ATO 200% diminishing value), `nabla.f.PrimeCostλ` (ATO prime cost), `nabla.f.GSTAddλ`, `nabla.f.GSTExtractλ` and `nabla.d.FinancialYearλ`.
- `Depreciateλ` method codes are now `SLN`, `SYD`, `DB`, `DDB`, `VDB`, `DV` and `PC`; the Data Validation sheet and the `DepreciateλDV` diagnostic list the same set.
- Worksheets print on A4.

### Dates
- Function version stamps set to 18 Aug 2026 (110 version lines across help blocks and About tables, plus the AFE modules).
- All sample and demonstration dates shifted forward two years, calendar-aware: text dates, ISO dates, every date-formatted serial cell (inputs, table data and cached outputs) and serial array constants; 29 February clamps to 28 February when the target year is not a leap year. Demo tables (rentals, loans, items) and their timeline anchors moved together, so every worked example stays internally consistent after recalculation.
- `fullCalcOnLoad` enabled so cached demo outputs refresh on first open.

### Fixes
- Defined `nabla.e.Aboutλ`; the upstream workbook called `BXE.Aboutλ` on its own worksheet without defining it.
- Replaced the undefined `Sheetλ` title formula on 46 worksheets with a self-contained `TEXTAFTER(CELL("filename",A1),"]")` title; the upstream file cached `#NAME?` in every one.
- Replaced locale-fragile text-date arguments in `RANDBETWEEN` with `DATE()` calls.
- Removed a dead table-of-contents hyperlink to a worksheet that never existed, an empty Power Query mashup, orphaned rich-value image residue, the regenerable `calcChain` cache, and a merged cell left behind by the removed cover section. The table-of-contents row for that worksheet now correctly reads Function rather than Worksheet.
- Fixed typos: `Amoritization`, `Occurence`, `preceeding`, `dynamice`, "click and worksheet name", and an upstream misspelling of the author's name.
- Repaired an inherited `#REF!` argument in the TimelinePositionλ demo timeline and the `nabla.u.Aboutλ` text that suggested the wrong module name.
- Moved the first loan on the `Amortiseλ` worksheets to 1 March 2026. Upstream started it a year before the model timeline with a ten-month term, so it was fully repaid before the first period and its six rows rendered as zeros; it now shows a partial schedule. The worksheet caption is restated to match.
- Ported the remaining foreign sample data: household budget items became Australian equivalents (Pay, Home insurance, Strata levies, Petrol), and the depreciation note on the `Depreciateλ` worksheet was rewritten without its foreign accounting-standard framing.
- Added Name Manager descriptions to every new defined name.
- Defined `nabla.d.Aboutλ`, which the upstream workbook shipped as source but never installed, so it returned #NAME?.
- Extended `tblMethods` to cover the added prime cost row, and fixed the table's own copy of the sample-date formula, which still held the pre-conversion text dates.
- Carried the two-year date shift into help examples written with two-digit years and into cached values beyond 2064, which an earlier bound had skipped.
- Fixed prose that the date shift had left stale ("that loan starts in 2020"), a doubled word introduced by the rename, and a US working-week aside in the occurrence-date help.
- Fixed further upstream typos: "equally equally", "specifice text", "Some of Years".

### Currency with Excel 365
- Six helpers whose job Excel 365 has since taken over natively (the three `RangeToDAλ` copies, `FilterContainsλ`, `SumPeriodsλ` and `SumContainsλ`) now carry a `SEE ALSO` line in their inline help pointing at `TRIMRANGE`, the `REGEX` functions and `GROUPBY`/`PIVOTBY`. Checked against Microsoft's documentation in August 2026.
- The cover sheet states the Excel requirement: Microsoft 365, or Excel 2024 and later.
- Replaced the upstream `Coming soon` webpage placeholders in 74 help blocks with the repository URL, and removed the duplicate `Website` line that repeated the `Repository` line in every About table.
- Drawing text is tagged `en-AU` so Excel stops spell-checking Australian prose against a US dictionary, and the Advanced Formula Environment project store now declares the `en-au` locale with day-first date order.

### Performance
- Replaced the 38 `RANDBETWEEN` formulas that generated the sample data with fixed values. They made 4,349 of 4,627 formula cells volatile, so every edit recalculated 93% of the workbook; one of them on the `TimelineOffsetλ` sheet alone drove 2,583 cells. Nothing but the sheet-name titles is volatile now.
- Cleared 4,243 stale always-calculate flags left on cells that no longer depend on anything volatile, and removed the table column formulas that would have re-injected the random data.
- The verification script now fails the build if a volatile function or an unjustified always-calculate flag reappears.
- Fixed sample data also makes every worked example reproducible, so the numbers match the captions on each open.

### Presentation
- Tab colours by module: dates navy, array essentials blue, financial teal, ratios green, and the cover, contents and reference sheets grey.
- Gridlines hidden on all 49 sheets.
- Every sheet opens at the top left, and the workbook opens on the cover rather than on whichever tab was last active.

### Packaging
- Removed the embedded printer configuration, which carried the original author's printer name and a foreign default paper size.
- Repacked at maximum deflate; the workbook is smaller than the upstream file despite the added functions.
- Added `functions.csv`, a generated index of all 130 functions, and a GitHub Actions check that rebuilds the verification pass on every push.

### Typography
- Calibri and Calibri Light replaced with Aptos and Aptos Display across styles, theme and rich-text runs.
