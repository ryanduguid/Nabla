# Ozzit

![Ozzit](assets/ozzit.svg)

A LAMBDA function library for building dynamic-array financial models in Excel.

`ozzit.xlsx` ships 130 named functions under a single `oz.` prefix. Every function carries inline help, and the core functions each have a demonstration worksheet with live, editable examples. The **Australian tax** worksheet demonstrates the depreciation, GST and financial-year helpers. Every function is built from native Excel functions only, so models assembled with Ozzit save as ordinary `.xlsx` workbooks with no add-ins and no macros.

**Version: 20 August 2026**

## Requirements

Excel with LAMBDA and dynamic arrays: Microsoft 365, or Excel 2024 and later. LibreOffice and older Excel versions will not evaluate the functions.

Everything in the library is built from functions that shipped with Excel 2024, so the whole library works on that baseline. Nothing here depends on a preview feature.

## Modules

Every function shares one prefix, `oz.`, so a call is three characters of namespace
rather than eight. The groupings below describe what the library covers; they are not
part of the name.

| Group | Functions | Covers |
|---|---|---|
| Dates | 13 | Dates and timelines: periods, schedules, overlaps, occurrence tests, financial-year labels |
| Essentials | 17 | Array essentials: row and column totals, averages, counts, range conversion |
| Financial | 39 | Financial building blocks: amortisation, depreciation, corkscrews, IRR, rolling sums, GST |
| Ratios | 39 | Financial ratios: liquidity, leverage, margins, returns, market multiples |
| Utilities | 17 | Standalone copies of the Essentials functions, each carrying a `U` suffix |
| Debt | 5 | Debt sculpting: amortisation schedule, fixed and variable DSCR sculpting, sculpting interest |

Where two groups shipped a function of the same name, the fuller implementation keeps
the plain name and the other takes a one-letter tag: `B` for debt, `E` for essentials,
`U` for utilities. So `oz.Amortiseλ` is the financial amortisation schedule and
`oz.AmortiseBλ` the debt one; `oz.SumRowsλ` and `oz.SumRowsUλ` are the essentials and
utilities copies. The five About tables take words instead: `oz.AboutFinancialλ`,
`oz.AboutRatiosλ`, and so on.

## Getting started

1. Open `ozzit.xlsx`.
2. Cell A1 of every visible worksheet links back to the table of contents; every name in the TOC links to its worksheet.
3. For inline help, type a function name with no arguments in an empty cell, for example `=oz.Amortiseλ()`. The help block spills syntax, parameters and worked examples.
4. Grey-shaded cells on each worksheet are inputs. Change them and watch the function respond.
5. To use the functions in your own workbook, copy a green-shaded cell across (Excel brings the named LAMBDA with it), or import the plain-text source from `src/` with the Advanced Formula Environment in the Excel Labs add-in.

   Importing `src/` that way recreates the functions under the module container's own
   name, so `Dates.txt` produces `Dates.CountDOWλ` rather than `oz.CountDOWλ`: the
   Advanced Formula Environment takes the prefix from the container, and one flat
   namespace cannot be six containers. The workbook is the authority for the `oz.`
   names. `src/` is for reading, diffing, and pasting a single definition into Name
   Manager, where the name is yours to choose.

The workbook opens showing the numbers its own formulas produce, so nothing has to recalculate before it reads correctly and Excel does not ask you to save a file you never edited. That is a property the build cannot give it: the build edits the workbook as XML with no formula engine, so `tools/refresh_cache.py` recalculates it in Excel afterwards and `tools/verify_cache.py` proves every cached value matches.

Functions with a data-validation companion (named with a `DV` suffix, such as `oz.AmortiseλDV`) diagnose argument problems when the parent function returns something unexpected.

## Australian conventions

The library is Australian-only: the foreign tax regimes and references the upstream library carried have been removed rather than relabelled, and the conventions below are Australian throughout.

- Australian English throughout (amortise, modelling, and so on), including function names.
- Dates are day-first, in text examples and in cell formats. Worksheets print on A4.
- Sample data uses AUD.
- Depreciation, GST and financial-year helpers:

| Function | Purpose |
|---|---|
| `oz.DiminishingValueλ(Cost, Life)` | Diminishing balance at 200% of the straight-line rate, writing the residual off in the final period |
| `oz.PrimeCostλ(Cost, Life)` | Straight line over whole years |
| `oz.GSTAddλ(Amounts, [Rate])` | Adds GST to GST-exclusive amounts, 10% by default |
| `oz.GSTExtractλ(Amounts, [Rate])` | Returns the GST inside GST-inclusive amounts |
| `oz.FinancialYearλ(Dates, [StartMonth])` | Labels dates with their financial year, starting 1 July |

The 10% default and one-eleventh extraction reflect the basic rule for a taxable supply in [*A New Tax System (Goods and Services Tax) Act 1999*](https://www.legislation.gov.au/C2004A00446/latest/text) ss 9-70 and 9-75 (source checked 20 August 2026). Recheck the current Act and any applicable special rule at the time of use. These helpers apply arithmetic only: they do not decide whether a supply is taxable, GST-free, input taxed, outside the GST system or subject to a special rule.

The same scope note is embedded as NOTES! rows in the `oz.GSTAddλ` and `oz.GSTExtractλ` inline help. The About table and Name Manager comments keep the one-line descriptions. `tools/postbuild/gst_help_text.py` applies it to the committed `ozzit.xlsx` and `src/`; `tools/transform_from_upstream.py` still needs the uncommitted upstream workbook and still stops at v3.0.0.

`oz.Depreciateλ` accepts the method codes `SLN`, `SYD`, `DB`, `DDB`, `VDB`, `DV` (diminishing value) and `PC` (prime cost).

**These are modelling schedules, not tax calculations.** They take a cost and an effective life and nothing else: no acquisition date, no income year, no days held, and no disposal. `oz.DiminishingValueλ` also writes the whole undeducted residual off in its final period, which a diminishing-balance calculation does not do on its own, so the schedule totals to cost by construction. For a cost of 1,000 over five years it returns 400, 240, 144, 86.40 and 129.60, where the last figure is the residual rather than a computed deduction.

Use them for modelling. Do not use them to prepare a return, and do not treat any output as a deduction the ATO would accept. Apportion for part-year ownership yourself.

## Dynamic-Array Formula Walkthrough

Ozzit functions use native Excel dynamic arrays to spill full calculation schedules from a single formula cell:

### 1. Loan Amortisation (`=oz.Amortiseλ(Principal, Rate, Periods, [Type])`)
Spills a complete 5-column schedule (`Period`, `Payment`, `Principal`, `Interest`, `Closing Balance`):

| Period | Payment | Principal | Interest | Balance |
| :---: | :---: | :---: | :---: | :---: |
| 1 | \$1,887.12 | \$1,470.46 | \$416.67 | \$98,529.54 |
| 2 | \$1,887.12 | \$1,476.58 | \$410.54 | \$97,052.96 |
| 3 | \$1,887.12 | \$1,482.73 | \$404.39 | \$95,570.22 |
| ... | ... | ... | ... | \$0.00 |

### 2. Diminishing Value Depreciation (`=oz.DiminishingValueλ(Cost, Life)`)
Calculates diminishing balance at 200% straight-line rate with exact terminal residual write-off:

| Year 1 | Year 2 | Year 3 | Year 4 | Year 5 (Residual Write-off) | Total Written Off |
| :---: | :---: | :---: | :---: | :---: | :---: |
| \$400.00 | \$240.00 | \$144.00 | \$86.40 | \$129.60 | \$1,000.00 |

### 3. GST Arithmetic (`=oz.GSTExtractλ(Amounts)` & `=oz.GSTAddλ(Amounts)`)
Extracts or appends GST across full dynamic arrays according to *ANTS(GST)A 1999* ss 9-70/75:
- `=oz.GSTExtractλ(1100)` -> returns `$100.00`
- `=oz.GSTAddλ(1000)` -> returns `$1,100.00`

## Modern Excel

Excel 365 has gained functions since this library's upstream release in July 2024, and a few of them do natively what some helpers here were written to work around. Where that is the case the function's own inline help carries a `SEE ALSO` line, so you find out while you are using it rather than after:

| Helper | Native equivalent in Excel 365 |
|---|---|
| `oz.RangeToDAλ`, `oz.RangeToDAEλ`, `oz.RangeToDAUλ` | `TRIMRANGE`, or trim references (`.:.`) |
| `FilterContainsλ` | `REGEXTEST`, `REGEXEXTRACT` |
| `SumPeriodsλ`, `SumContainsλ` | `GROUPBY`, `PIVOTBY` |

The helpers are kept because they still work on the Excel 2024 baseline and inside the library's own composition, and because the native functions are Microsoft 365 only. Prefer the native function when your audience is on 365. Checked against Microsoft's documentation in August 2026.

## Performance and presentation

The workbook is built to stay responsive on modest hardware. No formula in it is volatile except the sheet-name titles, so editing a cell recalculates only what depends on it rather than the whole file. That covers the random-number formulas behind the sample data in both their forms. The sample data is fixed rather than randomly generated, which also means the worked examples match their captions every time you open them.

Each module has its own tab colour, gridlines are hidden, and every sheet opens at the top left on the cover.

## Repository layout

| Path | Contents |
|---|---|
| `ozzit.xlsx` | The library and its documentation workbook |
| `src/*.txt` | Plain-text LAMBDA source per group (Dates, Essentials, Financial, Ratios, Utilities, Debt), diffable and importable |
| `ATTRIBUTION.md` | Provenance and upstream copyright |
| `functions.csv` | Machine-readable index of every function |
| `tools/` | The build and maintenance pipeline: `transform_from_upstream.py` rebuilds the v3.0.0 baseline from an uncommitted upstream workbook and stops there; later tracked postbuild passes start from the committed `ozzit.xlsx` and `src/` (see ATTRIBUTION.md), then the AFE store is synchronised, the workbook is canonicalised/polished, and the gates check it |
| `CHANGELOG.md` | What changed in this release |
| `assets/` | Logo |
| `LICENCE` | MIT, and what it does and does not cover |

`src/` and `functions.csv` are generated from `ozzit.xlsx` by the build, not edited by hand, so the published source of a function is always the definition that ships.

## Checks

Nine gates, because they answer different questions. Seven run in CI; the two that need Excel run locally.

```bash
python tools/verify_workbook.py ozzit.xlsx
```

Structure: XML well-formedness, entity-expansion declarations, executable and external-link parts, undefined names, `#REF!`, volatile functions, stray always-calculate flags, slicer-cache bindings, the file-properties sheet list against the real one, and tokens that must never reappear. That last list includes anything describing the machine that built the release: Excel stamps the workbook with the directory it was last saved from, so every release up to v2.2.0 published a path off the build machine. Runs in CI on every push to `main` and on every pull request.

```bash
python tools/verify_sources.py ozzit.xlsx src
```

Provenance: every function in `src/` must match the defined name that ships, and must be written in the form Excel accepts as typed input rather than the form the file format stores. The two differ in four ways, which it maps rather than ignores. It also requires that a function which converts a date argument goes on to read the conversion: five of them converted one and then used the raw argument, which is silent until text ordering and date ordering disagree. Also runs in CI.

```bash
python tools/verify_signatures.py src
```

Documentation: every function states its parameters twice, once as the signature on its help's FUNCTION line and again as a table below it, and ends with worked examples a reader is meant to copy. All three are hand-written text inside a string literal, so nothing else in the build ever reads them and they drift separately: one function's table described a different function's arguments for six releases, and three worked examples called a neighbouring function rather than the one they were printed under. It reads 117 signatures, 122 parameter tables and 119 example blocks, checks that every function named in a help is one the library declares, accounts for every declaration in every module, and refuses to pass if it parsed too few. Also runs in CI.

```bash
python tools/verify_previous_names.py functions.csv
```

Migration: v2.0.0 renamed every function, so `functions.csv` carries a `previous_name` column and `tools/released-names-v1.2.6.txt` records the 130 names the last release before that rename shipped. This checks the published index against the published baseline: every old name is claimed by exactly one function, no function claims a name that never shipped, and anything added since records nothing rather than a plausible-looking guess. Also runs in CI.

```bash
python tools/verify_index.py ozzit.xlsx src functions.csv
```

Index: re-derives every module, signature and description in `functions.csv` from `src/` and the defined names in the workbook. This catches a stale or hand-edited machine-readable index, including the truncated long signature and escaped line-break descriptions that shipped before the export was fixed. Also runs in CI.

```bash
python tools/verify_afe.py ozzit.xlsx src
```

Advanced Formula Environment: the AFE task pane is the documented editing surface, and it has its own copy of the library inside `customXml/item1.xml`. This gate requires the five non-recursive module texts to equal `src/*.txt` and every shipped non-Debt name to be listed. Debt must stay out: the module is recursive and `src/Debt.txt` documents the Name Manager import path. Also runs in CI.

```bash
python -m unittest discover -s tools/tests -v
```

Tooling: exercises the workbook sanitiser against clean and simulated Excel-saved files, deterministic compression, failed atomic replacement, malformed input and privacy leaks; proves the cache refresh delegates to the same implementation; and regression-tests the source and signature parsers against new modules and LAMBDAs without a top-level `LET`. Runs in CI so maintenance edits cannot quietly weaken the gates or corrupt a workbook while replacing it.

```bash
powershell -ExecutionPolicy Bypass -File tools/excel_selftest.ps1
```

Arithmetic: opens the workbook in a real Excel, forces a full rebuild, fails on any error cell, then runs 259 assertions over the Australian functions, the worksheet that demonstrates them, the timeline and allocation helpers on period lengths from a day to a year, and the balance identities of the debt sculpting schedules. Needs Excel with LAMBDA support, so it cannot run on GitHub's runners and stays a local gate. It opens Excel over COM and quits it when finished, so it refuses to start if Excel is already running rather than closing your workbooks; it never saves the file it tests.

```bash
python tools/verify_cache.py ozzit.xlsx
```

Cached values: an `.xlsx` stores a formula and the answer Excel last got from it, and nothing keeps them in step. The build edits values as XML with no formula engine, so every cell downstream of an edit keeps the answer it had before: shifting the sample dates forward two years left 3,193 such cells across 43 sheets, and five cells shipped a saved `#VALUE!` from v1.2.0 to v2.2.0. Excel replaces them all on open, which is exactly why it needs a gate: the file can be wrong in a way only a second tool can see, and everything that reads an `.xlsx` without a formula engine reads the cached answer. This opens the workbook, recalculates, and compares all 20,228 cached values against what the formulas produce. Needs Excel, so it is a local gate. Run `python tools/refresh_cache.py` to fix what it reports.

Sanitising is a fixer, not a gate. Any save through Excel, not only the pipeline's, adds parts that do not belong in a distributed file: `printerSettings` binaries on machines with a printer installed, an `x15ac:absPath` recording the save directory, always-calculate flags on non-volatile cells, and empty worksheet rels. `python tools/sanitise_workbook.py ozzit.xlsx` strips all of it and rewrites the archive canonically, so two saves of the same content produce the same bytes. `refresh_cache.py` delegates its own save to the same sanitiser; run it directly after every other time Excel touched the file.

Presentation is reproducible too: `python tools/polish_workbook.py ozzit.xlsx` applies the Luma-neutral help highlights, accessible strapline colour, locale-safe date styles and the reviewed function-sheet descriptions. The pass is idempotent and the CI presentation contract checks each of those rules.

## Worksheet catalogue

| Worksheet | What it demonstrates |
|---|---|
| `oz.CountDOWλ` | Count instances of a specific day of the week between two dates |
| `oz.IsBetweenλ` | Determine if a value is between a lower and upper limit |
| `oz.IsOccurrenceDateλ` | Determine if a date passed is when a potentially repeating event happens |
| `oz.OverLapDaysλ` | Return how many days overlap two period ranges. |
| `oz.Periodsλ` | Determine the number of periods from Starts to Ends inclusive |
| `oz.PeriodLabelλ` | Creates a label for a date based on period interval |
| `oz.ScheduleRatesλ` | Schedule rates that persist until replaced in a timeline. |
| `oz.ScheduleRatesByItemsλ` | Schedule rates that persist until replaced in a timeline for each item in a list. |
| `oz.ScheduleValuesλ` | Schedules values in a timeline. |
| `oz.ScheduleValuesByItemsλ` | Schedules values in a timeline for each item in a list. |
| `oz.Timelineλ` | Creates a horizontal list of start or end dates for a timeline |
| `oz.Amortiseλ` | Creates a corkscrew amortisation schedule. |
| `oz.LabelAmortiseλ` | Create row labels for Amortiseλ result |
| `oz.SumAmortiseλ` | Create totals for payments, interest, and principal portion in Amortiseλ results |
| `oz.Corkscrewλ` | Creates a simple corkscrew where the closing balance is the sum of independent flows plus opening balance |
| `oz.Cumulativeλ` | Creates a row or column of cumulative totals from a total row or column |
| `oz.Depreciateλ` | Create a block of CAPEX, Opening Balance, Depreciation Values, and Book Value for each asset |
| `oz.LabelDepreciateλ` | Create row labels for Depreciateλ result |
| `oz.SumDepreciateλ` | Create row totals for CAPEX, Depreciation, Book Value, Salvage Value, and Disposal costs in Depreciateλ results |
| `oz.SumContains` | Creates a row of totals for each row in an array where its labels contain a unique letter, word, or phrase. |
| `oz.IntOnIntλ` | Calculate Interest on Interest. Use to determine amount needed to cover debt plus interest on debt |
| `oz.IRRλ` | Calculates IRR, correcting for when the first investment is not in the first period |
| `oz.Reversalλ` | Create a row that reverses input values in the next period. |
| `oz.Movementλ` | Create a row of differences from column to column |
| `oz.RollingSumλ` | Creates totals for preceding values of a set size moving from beginning to end over a row of values. |
| `oz.SumPeriodsλ` | Groups and totals all columns in a Values array by period resulting in one column for each period. |
| `oz.TimelineOffsetλ` | Determines how many columns a date is offset from a timeline's first date |
| `oz.TimelinePositionλ` | Places an array or value appropriately within a model's timeline. |
| `oz.AboutEssentialsλ` | About the Ozzit Array Essentials library |
| `oz.CountCλ` | Count the number of times one or more characters appear in a string |
| `oz.SumRowsλ` | Creates totals for each row in array. |
| `oz.SumColsλ` | Creates totals for each column in array. |
| `oz.AvgRowsλ` | Gets the average of each row in an array |
| `oz.AvgColsλ` | Gets the average of each column in an array |
| `oz.MinRowsλ` | Gets the minimum of each row in an array |
| `oz.MinColsλ` | Gets the minimum of each column in an array |
| `oz.MaxRowsλ` | Gets the maximum of each row in an array |
| `oz.MaxColsλ` | Gets the maximum of each column in an array |
| `oz.CountRowsλ` | Count the number of numbers in each row of an array |
| `oz.CountColsλ` | Count the number of numbers in each column of an array |
| `oz.CountARowsλ` | Count all non-empty cells in each row of a range. NOTE! Dynamic Arrays always fill each cell. |
| `oz.CountAColsλ` | Count all non-empty cells in each column of a range. NOTE! Dynamic Arrays always fill each cell. |
| `oz.IsBetweenEλ` | Determine if a value is between a lower and upper limit |
| `oz.RangeToDAEλ` | Convert a static range into a dynamic array |
| `oz.FinancialRatios` | Three dozen financial Ratios |

The other 85 functions (all of Ratios, Utilities and Debt, the depreciation-method, GST and rolling-statistic helpers in Financial, `oz.FinancialYearλ`, and the module About tables) have no dedicated worksheet; call any of them with no arguments for inline help, `oz.FinancialRatios` demonstrates the ratio suite on one worksheet, and [functions.csv](functions.csv) lists every function with its signature.

## Attribution and licence

Ozzit is a renamed and reworked derivative of Craig Hatmaker's Financial Starter Pack (BXL / 5g). See [ATTRIBUTION.md](ATTRIBUTION.md) for provenance, upstream copyright and the full list of changes. This repository adds Australian GST, financial-year and modelling-depreciation helpers; it is not an individual-tax or Division 7A engine.

[LICENCE](LICENCE) is MIT and covers only what was written for this repository: `tools/`, `.github/`, the Markdown files and `assets/`. It does not cover `ozzit.xlsx`, `src/` or `functions.csv`, which derive from the upstream workbook. No licence for that work could be located, so its author retains all rights in the original material.
