# Nabla

![Nabla](assets/nabla.svg)

A LAMBDA function library for building dynamic-array financial models in Excel.

`nabla.xlsx` ships 130 named functions under a single `nb.` prefix. Every function carries inline help, and the core functions each have a demonstration worksheet with live, editable examples. The **Australian tax** worksheet demonstrates the ATO depreciation, GST and financial-year helpers. Every function is built from native Excel functions only, so models assembled with Nabla save as ordinary `.xlsx` workbooks with no add-ins and no macros.

**Version: 18 August 2026**

## Requirements

Excel with LAMBDA and dynamic arrays: Microsoft 365, or Excel 2024 and later. LibreOffice and older Excel versions will not evaluate the functions.

Everything in the library is built from functions that shipped with Excel 2024, so the whole library works on that baseline. Nothing here depends on a preview feature.

## Modules

Every function shares one prefix, `nb.`, so a call is three characters of namespace
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
`U` for utilities. So `nb.Amortiseλ` is the financial amortisation schedule and
`nb.AmortiseBλ` the debt one; `nb.SumRowsλ` and `nb.SumRowsUλ` are the essentials and
utilities copies. The five About tables take words instead: `nb.AboutFinancialλ`,
`nb.AboutRatiosλ`, and so on.

## Getting started

1. Open `nabla.xlsx`.
2. Cell A1 of every visible worksheet links back to the table of contents; every name in the TOC links to its worksheet.
3. For inline help, type a function name with no arguments in an empty cell, for example `=nb.Amortiseλ()`. The help block spills syntax, parameters and worked examples.
4. Grey-shaded cells on each worksheet are inputs. Change them and watch the function respond.
5. To use the functions in your own workbook, copy a green-shaded cell across (Excel brings the named LAMBDA with it), or import the plain-text source from `src/` with the Advanced Formula Environment in the Excel Labs add-in.

   Importing `src/` that way recreates the functions under the module container's own
   name, so `Dates.txt` produces `Dates.CountDOWλ` rather than `nb.CountDOWλ`: the
   Advanced Formula Environment takes the prefix from the container, and one flat
   namespace cannot be six containers. The workbook is the authority for the `nb.`
   names. `src/` is for reading, diffing, and pasting a single definition into Name
   Manager, where the name is yours to choose.

The workbook recalculates fully on first open, so demonstration outputs (including random sample data) refresh and Excel will offer to save the result.

Functions with a data-validation companion (named with a `DV` suffix, such as `nb.AmortiseλDV`) diagnose argument problems when the parent function returns something unexpected.

## Australian conventions

The library is Australian-only. Tax content follows ATO practice, and the foreign tax regimes and references the upstream library carried have been removed rather than relabelled.

- Australian English throughout (amortise, modelling, and so on), including function names.
- Dates are day-first, in text examples and in cell formats. Worksheets print on A4.
- Sample data uses AUD.
- Australian tax helpers:

| Function | Purpose |
|---|---|
| `nb.DiminishingValueλ(Cost, Life)` | ATO diminishing value method (200% declining balance), writing the residual off in the final period |
| `nb.PrimeCostλ(Cost, Life)` | ATO prime cost method (straight line) |
| `nb.GSTAddλ(Amounts, [Rate])` | Adds GST to GST-exclusive amounts, 10% by default |
| `nb.GSTExtractλ(Amounts, [Rate])` | Returns the GST inside GST-inclusive amounts |
| `nb.FinancialYearλ(Dates, [StartMonth])` | Labels dates with their financial year, starting 1 July |

`nb.Depreciateλ` accepts the method codes `SLN`, `SYD`, `DB`, `DDB`, `VDB`, `DV` (diminishing value) and `PC` (prime cost).

The depreciation helpers return full-year amounts. Apportion the first year yourself if the asset was held for part of it.

## Modern Excel

Excel 365 has gained functions since this library's upstream release in July 2024, and a few of them do natively what some helpers here were written to work around. Where that is the case the function's own inline help carries a `SEE ALSO` line, so you find out while you are using it rather than after:

| Helper | Native equivalent in Excel 365 |
|---|---|
| `nb.RangeToDAλ`, `nb.RangeToDAEλ`, `nb.RangeToDAUλ` | `TRIMRANGE`, or trim references (`.:.`) |
| `FilterContainsλ` | `REGEXTEST`, `REGEXEXTRACT` |
| `SumPeriodsλ`, `SumContainsλ` | `GROUPBY`, `PIVOTBY` |

The helpers are kept because they still work on the Excel 2024 baseline and inside the library's own composition, and because the native functions are Microsoft 365 only. Prefer the native function when your audience is on 365. Checked against Microsoft's documentation in August 2026.

## Performance and presentation

The workbook is built to stay responsive on modest hardware. No formula in it is volatile except the sheet-name titles, so editing a cell recalculates only what depends on it rather than the whole file. That covers the random-number formulas behind the sample data in both their forms. The sample data is fixed rather than randomly generated, which also means the worked examples match their captions every time you open them.

Each module has its own tab colour, gridlines are hidden, and every sheet opens at the top left on the cover.

## Repository layout

| Path | Contents |
|---|---|
| `nabla.xlsx` | The library and its documentation workbook |
| `src/*.txt` | Plain-text LAMBDA source per group (Dates, Essentials, Financial, Ratios, Utilities, Debt), diffable and importable |
| `ATTRIBUTION.md` | Provenance and upstream copyright |
| `functions.csv` | Machine-readable index of every function |
| `tools/` | The build pipeline: rebuilds the workbook from upstream and checks it |
| `CHANGELOG.md` | What changed in this release |
| `assets/` | Logo |

`src/` and `functions.csv` are generated from `nabla.xlsx` by the build, not edited by hand, so the published source of a function is always the definition that ships.

## Checks

Four gates, because they answer different questions.

```bash
python tools/verify_workbook.py nabla.xlsx
```

Structure: XML well-formedness, undefined names, `#REF!`, volatile functions, stray always-calculate flags, slicer-cache bindings, and tokens that must never reappear. Runs in CI on every push.

```bash
python tools/verify_sources.py nabla.xlsx src
```

Provenance: every function in `src/` must match the defined name that ships, and must be written in the form Excel accepts as typed input rather than the form the file format stores. The two differ in four ways, which it maps rather than ignores. Also runs in CI.

```bash
python tools/verify_signatures.py src
```

Documentation: every function states its parameters twice, once as the signature on its help's FUNCTION line and again as a table below it, and both must match what the LAMBDA declares, character for character. Both are hand-written text inside a string literal, so nothing else in the build ever reads them and they drift separately: one function's table described a different function's arguments for six releases. It reads 117 signatures and 122 parameter tables, accounts for every declaration in every module, and refuses to pass if it parsed too few. Also runs in CI.

```bash
powershell -ExecutionPolicy Bypass -File tools/excel_selftest.ps1
```

Arithmetic: opens the workbook in a real Excel, forces a full rebuild, fails on any error cell, then runs 138 assertions over the Australian functions and the worksheet that demonstrates them. Needs Excel with LAMBDA support, so it cannot run on GitHub's runners and stays a local gate. It opens Excel over COM and quits it when finished, so it refuses to start if Excel is already running rather than closing your workbooks; it never saves the file it tests.

## Worksheet catalogue

| Worksheet | What it demonstrates |
|---|---|
| `nb.CountDOWλ` | Count instances of a specific day of the week between two dates |
| `nb.IsBetweenλ` | Determine if a value is between a lower and upper limit |
| `nb.IsOccurrenceDateλ` | Determine if a date passed is when a potentially repeating event happens |
| `nb.OverLapDaysλ` | Return how many days overlap two period ranges. |
| `nb.Periodsλ` | Determine the number of periods from Starts to Ends inclusive |
| `nb.PeriodLabelλ` | Creates a label for a date based on period interval |
| `nb.ScheduleRatesλ` | Schedule rates that persist until replaced in a timeline. |
| `nb.ScheduleRatesByItemsλ` | Schedule rates that persist until replaced in a timeline for each item in a list. |
| `nb.ScheduleValuesλ` | Schedules values in a timeline. |
| `nb.ScheduleValuesByItemsλ` | Schedules values in a timeline for each item in a list. |
| `nb.Timelineλ` | Creates a horizontal list of start or end dates for a timeline |
| `nb.Amortiseλ` | Creates a corkscrew amortisation schedule. |
| `nb.LabelAmortiseλ` | Create row labels for Amortiseλ result |
| `nb.SumAmortiseλ` | Create totals for payments, interest, and principal portion in Amortiseλ results |
| `nb.Corkscrewλ` | Creates a simple corkscrew where the closing balance is the sum of independent flows plus opening balance |
| `nb.Cumulativeλ` | Creates a row or column of cumulative totals from a total row or column |
| `nb.Depreciateλ` | Create a block of CAPEX, Opening Balance, Depreciation Values, and Book Value for each asset |
| `nb.LabelDepreciateλ` | Create row labels for Depreciateλ result |
| `nb.SumDepreciateλ` | Create row totals for CAPEX, Depreciation, Book Value, Salvage Value, and Disposal costs in Depreciateλ results |
| `nb.SumContains` | Creates a row of totals for each row in an array where its labels contain a unique letter, word, or phrase. |
| `nb.IntOnIntλ` | Calculate Interest on Interest. Use to determine amount needed to cover debt plus interest on debt |
| `nb.IRRλ` | Calculates IRR, correcting for when the first investment is not in the first period |
| `nb.Reversalλ` | Create a row that reverses input values in the next period. |
| `nb.Movementλ` | Create a row of differences from column to column |
| `nb.RollingSumλ` | Creates totals for preceding values of a set size moving from beginning to end over a row of values. |
| `nb.SumPeriodsλ` | Groups and totals all columns in a Values array by period resulting in one column for each period. |
| `nb.TimelineOffsetλ` | Determines how many columns a date is offset from a timeline's first date |
| `nb.TimelinePositionλ` | Places an array or value appropriately within a model's timeline. |
| `nb.AboutEssentialsλ` | About the Nabla Array Essentials library |
| `nb.CountCλ` | Count the number of times one or more characters appear in a string |
| `nb.SumRowsλ` | Creates totals for each row in array. |
| `nb.SumColsλ` | Creates totals for each column in array. |
| `nb.AvgRowsλ` | Gets the average of each row in an array |
| `nb.AvgColsλ` | Gets the average of each column in an array |
| `nb.MinRowsλ` | Gets the minimum of each row in an array |
| `nb.MinColsλ` | Gets the minimum of each column in an array |
| `nb.MaxRowsλ` | Gets the maximum of each row in an array |
| `nb.MaxColsλ` | Gets the maximum of each column in an array |
| `nb.CountRowsλ` | Count the number of numbers in each row of an array |
| `nb.CountColsλ` | Count the number of numbers in each column of an array |
| `nb.CountARowsλ` | Count all non-empty cells in each row of a range. NOTE! Dynamic Arrays always fill each cell. |
| `nb.CountAColsλ` | Count all non-empty cells in each column of a range. NOTE! Dynamic Arrays always fill each cell. |
| `nb.IsBetweenEλ` | Determine if a value is between a lower and upper limit |
| `nb.RangeToDAEλ` | Convert a static range into a dynamic array |
| `nb.FinancialRatios` | Three dozen financial Ratios |

The other 85 functions (all of Ratios, Utilities and Debt, the depreciation-method, GST and rolling-statistic helpers in Financial, `nb.FinancialYearλ`, and the module About tables) have no dedicated worksheet; call any of them with no arguments for inline help, `nb.FinancialRatios` demonstrates the ratio suite on one worksheet, and [functions.csv](functions.csv) lists every function with its signature.

## Attribution

Nabla is a renamed and reworked derivative of an existing LAMBDA library. See [ATTRIBUTION.md](ATTRIBUTION.md) for provenance, upstream copyright and the full list of changes.
