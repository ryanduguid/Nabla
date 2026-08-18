# Changelog

## 2026-08-18

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
