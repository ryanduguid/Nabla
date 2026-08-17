# Changelog

## 2026-08-18

First nabla release, derived from the upstream 6 July 2024 workbook.

### Rebranding
- Renamed every function namespace, worksheet, AFE module and help reference from `BXD`/`BXE`/`BXF`/`BXR`/`BXU`/`BXLDebt` to the `nabla.*` scheme.
- Removed branded cover art, the cover video thumbnails and their YouTube link, Dropbox model links, and the source add-in credit line. A maths-citation video link in an IRRλ source comment is retained as third-party credit.
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
- Function version stamps set to 18 Aug 2026 (106 version lines across help blocks and About tables, plus the AFE modules).
- All sample and demonstration dates shifted forward two years, calendar-aware: text dates, ISO dates, every date-formatted serial cell (inputs, table data and cached outputs) and serial array constants; 29 February clamps to 28 February when the target year is not a leap year. Demo tables (rentals, loans, items) and their timeline anchors moved together, so every worked example stays internally consistent after recalculation.
- `fullCalcOnLoad` enabled so cached demo outputs refresh on first open.

### Fixes
- Defined `nabla.e.Aboutλ`; the upstream workbook called `BXE.Aboutλ` on its own worksheet without defining it.
- Replaced the undefined `Sheetλ` title formula on 46 worksheets with a self-contained `TEXTAFTER(CELL("filename",A1),"]")` title; the upstream file cached `#NAME?` in every one.
- Replaced locale-fragile text-date arguments in `RANDBETWEEN` with `DATE()` calls.
- Removed a dead table-of-contents hyperlink to a worksheet that never existed, a stale `Slicer_Type` name, an empty Power Query mashup and orphaned rich-value image residue.
- Fixed typos: `Amoritization`, `Occurence`, `preceeding`, `dynamice`, "click and worksheet name", and an upstream misspelling of the author's name.
- Repaired an inherited `#REF!` argument in the TimelinePositionλ demo timeline and the `nabla.u.Aboutλ` text that suggested the wrong module name.
- Ported the remaining foreign sample data: household budget items became Australian equivalents (Pay, Home insurance, Strata levies, Petrol), and the depreciation note on the `Depreciateλ` worksheet was rewritten without its foreign accounting-standard framing.
- Added Name Manager descriptions to every new defined name.
- Fixed further upstream typos: "equally equally", "specifice text", "Some of Years".

### Packaging
- Removed the embedded printer configuration, which carried the original author's printer name and a foreign default paper size.
- Repacked at maximum deflate; the workbook is smaller than the upstream file despite the added functions.
- Added `functions.csv`, a generated index of all 129 functions, and a GitHub Actions check that rebuilds the verification pass on every push.

### Typography
- Calibri and Calibri Light replaced with Aptos and Aptos Display across styles, theme and rich-text runs.
