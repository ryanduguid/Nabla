# nabla

![nabla](assets/nabla.svg)

A LAMBDA function library for building dynamic-array financial models in Excel.

`nabla.xlsx` ships 126 named functions across six modules. Every function carries inline help, and the core functions each have a demonstration worksheet with live, editable examples. Every function is built from native Excel functions only, so models assembled with nabla save as ordinary `.xlsx` workbooks with no add-ins and no macros.

**Version: 18 August 2026**

## Requirements

Excel with LAMBDA and dynamic arrays (Microsoft 365, or Excel 2024 and later). LibreOffice and older Excel versions will not evaluate the functions.

## Modules

| Module | Functions | Covers |
|---|---|---|
| `nabla.d` | 11 | Dates and timelines: periods, schedules, overlaps, occurrence tests |
| `nabla.e` | 17 | Array essentials: row and column totals, averages, counts, range conversion |
| `nabla.f` | 37 | Financial building blocks: amortisation, depreciation, corkscrews, IRR, rolling sums |
| `nabla.r` | 39 | Financial ratios: liquidity, leverage, margins, returns, market multiples |
| `nabla.u` | 17 | Utilities mirroring `nabla.e` for standalone use |
| `nabla.debt` | 5 | Debt sculpting: amortisation schedule, fixed and variable DSCR sculpting, sculpting interest |

## Getting started

1. Open `nabla.xlsx`.
2. Cell A1 of every visible worksheet links back to the table of contents; every name in the TOC links to its worksheet.
3. For inline help, type a function name with no arguments in an empty cell, for example `=nabla.f.Amortiseλ()`. The help block spills syntax, parameters and worked examples.
4. Grey-shaded cells on each worksheet are inputs. Change them and watch the function respond.
5. To use the functions in your own workbook, copy a green-shaded cell across (Excel brings the named LAMBDA with it), or import the plain-text source from `src/` with the Advanced Formula Environment in the Excel Labs add-in.

The workbook recalculates fully on first open, so demonstration outputs (including random sample data) refresh and Excel will offer to save the result.

Functions with a data-validation companion (named with a `DV` suffix, such as `nabla.f.AmortiseλDV`) diagnose argument problems when the parent function returns something unexpected.

## Australian conventions

- Australian English throughout (amortise, modelling, and so on), including function names.
- Dates are day-first, in text examples and in cell formats.
- Sample data uses AUD.
- `nabla.f.DiminishingValueλ(Cost, Life)` implements the ATO 200% diminishing value method. The US MACRS method remains available as a documented legacy option inside `nabla.f.Depreciateλ`.

## Repository layout

| Path | Contents |
|---|---|
| `nabla.xlsx` | The library and its documentation workbook |
| `src/nabla.*.txt` | Plain-text LAMBDA source per module, diffable and importable |
| `ATTRIBUTION.md` | Provenance and upstream copyright |
| `CHANGELOG.md` | What changed in this release |
| `assets/` | Logo |

## Worksheet catalogue

| Worksheet | What it demonstrates |
|---|---|
| `nabla.d.CountDOWλ` | Count instances of a specific day of the week between two dates |
| `nabla.d.IsBetweenλ` | Determine if a value is between a lower and upper limit |
| `nabla.d.IsOccurrenceDateλ` | Determine if a date passed is when a potentially repeating event happens |
| `nabla.d.OverLapDaysλ` | Return how many days overlap two period ranges. |
| `nabla.d.Periodsλ` | Determine the number of periods from Starts to Ends inclusive |
| `nabla.d.PeriodLabelλ` | Creates a label for a date based on period interval |
| `nabla.d.ScheduleRatesλ` | Schedule rates that persist until replaced in a timeline. |
| `nabla.d.ScheduleRatesByItemsλ` | Schedule rates that persist until replaced in a timeline for each item in a list. |
| `nabla.d.ScheduleValuesλ` | Schedules values in a timeline. |
| `nabla.d.ScheduleValuesByItemsλ` | Schedules values in a timeline for each item in a list. |
| `nabla.d.Timelineλ` | Creates a horizontal list of start or end dates for a timeline |
| `nabla.f.Amortiseλ` | Creates a corkscrew amortisation schedule. |
| `nabla.f.LabelAmortiseλ` | Create row labels for Amortiseλ result |
| `nabla.f.SumAmortiseλ` | Create totals for payments, interest, and principal portion in Amortiseλ results |
| `nabla.f.Corkscrewλ` | Creates a simple corkscrew where the closing balance is the sum of independent flows plus opening balance |
| `nabla.f.Cumulativeλ` | Creates a row or column of cumulative totals from a total row or column |
| `nabla.f.Depreciateλ` | Create a block of CAPEX, Opening Balance, Depreciation Values, and Book Value for each asset |
| `nabla.f.LabelDepreciateλ` | Create row labels for Depreciateλ result |
| `nabla.f.SumDepreciateλ` | Create row totals for CAPEX, Depreciation, Book Value, Salvage Value, and Disposal costs in Depreciateλ results |
| `nabla.f.SumContains` | Creates a row of totals for each row in an array where its labels contain a unique letter, word, or phrase. |
| `nabla.f.IntOnIntλ` | Calculate Interest on Interest. Use to determine amount needed to cover debt plus interest on debt |
| `nabla.f.IRRλ` | Calculates IRR, correcting for when the first investment is not in the first period |
| `nabla.f.Reversalλ` | Create a row that reverses input values in the next period. |
| `nabla.f.Movementλ` | Create a row of differences from column to column |
| `nabla.f.RollingSumλ` | Creates totals for preceding values of a set size moving from beginning to end over a row of values. |
| `nabla.f.SumPeriodsλ` | Groups and totals all columns in a Values array by period resulting in one column for each period. |
| `nabla.f.TimelineOffsetλ` | Determines how many columns a date is offset from a timeline's first date |
| `nabla.f.TimelinePositionλ` | Places an array or value appropriately within a model's timeline. |
| `nabla.e.Aboutλ` | About the nabla Array Essentials library |
| `nabla.e.CountCλ` | Count the number of times one or more characters appear in a string |
| `nabla.e.SumRowsλ` | Creates totals for each row in array. |
| `nabla.e.SumColsλ` | Creates totals for each column in array. |
| `nabla.e.AvgRowsλ` | Gets the average of each row in an array |
| `nabla.e.AvgColsλ` | Gets the average of each column in an array |
| `nabla.e.MinRowsλ` | Gets the minimum of each row in an array |
| `nabla.e.MinColsλ` | Gets the minimum of each column in an array |
| `nabla.e.MaxRowsλ` | Gets the maximum of each row in an array |
| `nabla.e.MaxColsλ` | Gets the maximum of each column in an array |
| `nabla.e.CountRowsλ` | Count the number of numbers in each row of an array |
| `nabla.e.CountColsλ` | Count the number of numbers in each column of an array |
| `nabla.e.CountARowsλ` | Count all non-empty cells in each row of a range. NOTE! Dynamic Arrays always fill each cell. |
| `nabla.e.CountAColsλ` | Count all non-empty cells in each column of a range. NOTE! Dynamic Arrays always fill each cell. |
| `nabla.e.IsBetweenλ` | Determine if a value is between a lower and upper limit |
| `nabla.e.RangeToDAλ` | Convert a static range into a dynamic array |
| `nabla.r.FinancialRatios` | Three dozen financial Ratios |

The remaining functions (all of `nabla.r`, `nabla.u` and `nabla.debt`, the depreciation-method and rolling-statistic helpers in `nabla.f`, and the module `Aboutλ` tables) have no dedicated worksheet; call any of them with no arguments for inline help, and `nabla.r.FinancialRatios` demonstrates the ratio suite on one worksheet.

## Attribution

nabla is a renamed and reworked derivative of an existing LAMBDA library. See [ATTRIBUTION.md](ATTRIBUTION.md) for provenance, upstream copyright and the full list of changes.
