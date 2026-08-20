# Changelog

## Unreleased, help copy-paste

- **`oz.CorkscrewλDV` described `oz.Depreciateλ`.** Its Name Manager comment, source
  header and `functions.csv` row were copied from the neighbouring DV function. The
  LAMBDA takes Opening and Flow1–4; the comment now says Corkscrewλ.
- **`oz.AboutFinancialλ` described `oz.RollingAvgλ` as a maximum.** The About table
  reused `oz.RollingMaxλ`'s sentence. It now matches the dedicated RollingAvgλ help:
  averages of a moving window, not a maximum.
- **`oz.CountColsλ`, `oz.CountColsUλ`, `oz.CountAColsλ` and `oz.CountAColsUλ` were
  indexed as row functions.** Source help already said column; the workbook comments
  and `functions.csv` still said "each row", which is what `CountRowsλ` correctly
  says. The index now says column.
- **`oz.Corkscrewλ` and `oz.CorkScrewReversalλ` still advertised `<Coming Soon>` as
  their WEBPAGE.** Every other function points at the repository; they now do too.
- **A Corkscrewλ revision comment dated the help rewrite 1 May 2924.** That is a
  mistyped 2024. Source comments are stripped before the src-vs-workbook check, so
  only the comment changes.

## v3.1.0, 20 August 2026, FY27 examples, Luma styling, a cleaner file

- **Every worked example now starts on 1 July 2026**, the start of FY27. The demo inputs
  on 141 cells moved forward, each sheet shifted by whole months so the relationships
  between its dates are the ones they always were. Where an example's documented result
  depends on the dates rather than the arithmetic, the shift is an exact multiple of the
  recurrence: `CountDOWλ` moves from 22/3/2012–10/4/2012 to 23/7/2026–11/8/2026 and still
  answers 3. The help text carries the new dates too, rewritten to the same character
  count so the columns in all 130 help blocks stay aligned.

  The one date deliberately left behind is 30/06/2026 on the Australian tax sheet. It is
  there to sit beside 01/07/2026 and show FY2026 turning into FY2027, so moving it would
  delete the thing it demonstrates.

- **The workbook is styled to the Luma Advisors palette**, taken from the published logo
  and stylesheet rather than guessed: purple `#5C2D91` as the single accent, near-black
  `#04001F` and a warm grey for everything else. The legacy accents inherited from
  upstream are gone: two greens, two blues and a maroon in the help blocks fold into the
  brand purple, and the mint, yellow and pink cell fills fold into the neutral greys.
  Thirteen font colours become eight. `assets/ozzit.svg` moves off its teal to match.

- **The workbook no longer carries a copy of my filesystem.** Saving through Excel adds
  parts that have no business in a distributed workbook: 50 `printerSettings` binaries
  pinned to whatever printer was installed, an `x15ac:absPath` recording the directory the
  file was last saved from, and always-calculate flags on 83 cells that compute nothing
  volatile. All three are stripped, along with five worksheet relationship parts left
  empty once the printer settings went. The zip is rebuilt deterministically, so two
  builds of the same content produce the same bytes, and the part count is back to the
  211 this workbook shipped with.

  Net size is 443,284 bytes against 438,263 before, 1.1% larger: the styling additions
  cost slightly more than the stripped parts saved. Against the copy in circulation that
  was last saved through Excel, which carried all of the above, it is 13.6% smaller.

- **`RangeToDAλ`, `RangeToDAEλ` and `RangeToDAUλ` keep OFFSET, on purpose.** OFFSET is
  volatile and replacing it looked like the obvious speed win, but INDEX cannot do what
  these functions do. Their whole point is to grow a reference past the cell you hand
  them, and `INDEX(A1,12,1)` on a one-cell reference is `#REF!`, verified in Excel rather
  than assumed. There is no non-volatile way to expand a reference, so the volatility
  stays and the functions keep working.

  The measurement that prompted this is worth recording: a full rebuild of all 1,129
  formulas takes 0.21s, and a volatile-only recalculation takes under 2ms. There was no
  performance problem to fix.

- **The clean-up is now a tool, not a session.** `tools/sanitise_workbook.py` removes what
  any Excel save adds — printer settings, the recorded save path, stray always-calculate
  flags, empty worksheet rels — and rewrites the archive deterministically. The pipeline's
  own `refresh_cache.py` already stripped its share after every refresh; the sanitiser
  covers every save that is not the pipeline's, so a workbook touched by hand no longer
  fails `verify_workbook.py` on `x15ac:absPath` the next time it is committed.

- **The clean-up has one implementation and a regression suite.** `refresh_cache.py` now
  calls the same sanitiser as a manual save instead of carrying a smaller, second copy of
  the XML surgery. CI regression tests cover no-op byte identity, deterministic output,
  printer settings, cache-refresh delegation, atomic replacement failure, local-path and
  XML-parser guards, LAMBDA signatures without top-level `LET`, and unexpected source
  modules. The replacement is genuinely atomic on Windows as well as POSIX: it writes a
  sibling temporary archive, calls `os.replace`, and removes the temporary file if
  replacement fails.

- **The shipped archive is 3,894 bytes smaller without changing any part payload.** The old
  writer handed `ZipInfo` objects to `writestr`, which silently ignored the `ZipFile`'s
  level-9 setting. The canonical writer sets the level on every entry, fixes host-dependent
  zip metadata, and verifies byte stability on a second run. All 211 OOXML part payloads
  are byte-identical to the previous workbook.

- **CI now checks the published index and the workbook's executable-part inventory.** A
  fabricated module or description in `functions.csv` used to pass because only the
  migration column was verified; all 130 rows are now re-derived from `src/` and the
  workbook. A workbook with an injected `vbaProject.bin` also used to pass; macros,
  ActiveX, OLE embeddings and external-link parts are now rejected explicitly.

- **The workflow's third-party actions are pinned to full commit SHAs.** Dependabot checks
  those pins weekly, so supply-chain integrity does not trade away updates. Repository-wide
  LF normalisation and explicit ignores for Ruff and Superpowers scratch data also stop a
  maintainer's Git configuration or local tools from creating whole-file or accidental
  public diffs.

- **The final legacy palette and copy defects are gone.** The mint formula highlights on
  109 cells become pale lavender `#DED9E8`; the 44 visible function straplines move from
  sub-AA `#808080` to the existing `#6E6862` neutral; 23 real date cells stop using the
  reader's locale-dependent short-date format; and eight copied or missing A2 descriptions
  now state what their sheets demonstrate. `tools/polish_workbook.py` makes the pass
  reproducible and idempotent, with an executable presentation contract in CI.

- **The AFE authoring view now carries the same library the workbook calculates.** Four
  modules in the Advanced Formula Environment store still said 19 Aug 2026 after the
  shipped names and `src/` moved to 20 Aug; six stale strings are synchronised. A new
  gate requires all five AFE-compatible modules to equal `src/` and keeps recursive Debt
  out, matching the documented import path.

- **The deterministic postbuild passes are now tracked, idempotent tools.** The FY27
  help-text swaps and the Luma palette that produced v3.1.0 live in `tools/postbuild/`,
  each with asserted hit counts and a byte no-op on the current workbook, covered by the
  CI test suite. The COM date shift is deliberately not ported: it is Excel-state-dependent
  and cannot be reproduced byte-for-byte, which ATTRIBUTION.md now says plainly. One dead
  bold font entry (Aptos Narrow, pre-palette colour) and the TOC's rich-text `Totals` run,
  both of which the original restyle missed, are folded into the palette by the new pass.

## v3.0.0, 19 August 2026, the library is now Ozzit

- **Every function's prefix changes from `nb.` to `oz.`, and the library is renamed from
  Nabla to Ozzit.** Nabla sits one edit from an acronym nobody wants beside their name on
  a public repository, which is reason enough on its own. It was also a poor description:
  the nabla operator means a gradient, and there is no calculus anywhere in this library.
  Ozzit says what the library is instead, which is Australian.

  This breaks every formula written against `nb.`. The break is deliberate and it is taken
  now rather than later, while the repository has no dependants to strand.

  Nothing else changes. The bare names are untouched, so `nb.Amortiseλ` becomes
  `oz.Amortiseλ` and `nb.SumRowsUλ` becomes `oz.SumRowsUλ`: the migration is a
  three-character find and replace, and no argument, result or rounding moves with it. The
  259 assertions in the Excel self-test answer exactly what they answered before.

  The rename is length-preserving throughout, because the help block every function
  carries is aligned in columns: `Nabla` and `Ozzit` are both five characters, `nb.` and
  `oz.` both three, and the bare `nb` the help text names in prose is two, like `oz`.
  `tools/rebrand_to_ozzit.py` asserts that property on every string it rewrites rather
  than trusting it, so not one of the 130 help blocks needed re-aligning.

  What the rename does not touch is the record of what earlier releases shipped.
  `tools/released-names-v1.2.6.txt` and the `previous_name` column of `functions.csv`
  still read `nabla.d.Aboutλ` and the rest, because somebody migrating off an older
  workbook has to be able to look up a name that existed rather than one that never did.
  The older entries in this file are left alone for the same reason.

- `nabla.xlsx` is now `ozzit.xlsx`, and `assets/nabla.svg` is now `assets/ozzit.svg`
  carrying a new mark. The old one drew the nabla operator, which the library never
  implemented.

- `tools/transform_from_upstream.py` now names the upstream functions `ozzit.*` on the way
  through and lands them in `oz.`, so a rebuild from the upstream workbook produces this
  release rather than the previous one.

- The version stamp each function carries reads 19 August 2026, as does the README.

## v2.6.0, 19 August 2026, any period length

- **`nb.Amortiseλ` returned `#DIV/0!` on every timeline shorter than a month.** It reads
  the period length off the timeline's first two dates and rounds it to whole months, which
  is nought for anything under about a fortnight, and the next two lines divide by it.
  Daily, weekly and fortnightly timelines all failed.

  The schedule is solved monthly whatever the timeline, because the function's own
  description says the payments are monthly, and only then folded into the timeline's
  periods. Flooring the count at one month is not enough on its own: the folded block is
  laid down one period at a time, so a weekly timeline would date month two a week after
  month one and report a year of interest inside a quarter. Each month's figures now go in
  the one period that contains that month's start, and the periods between hold nothing,
  which is what `nb.Depreciateλ` has always done on the same timelines. The two balance
  rows are dated the same way rather than divided, because they are balances.

  A twelve-month loan drawn on 1 January 2026 now reports the same money on a weekly, a
  fortnightly and a daily timeline as it does on a monthly one, to the cent, and puts the
  first month in period one and the second in the week that holds 1 February rather than
  the week after the first.

  Two details are worth stating because the obvious versions of both are wrong. The by-date
  path runs whenever the period is under 28 days, not whenever the period rounds to nought
  months: no month is shorter than 28 days, and a period of 16 to 27 days is no whole number
  of months either but used to round to one and be laid out as though it were a month long.
  And a period ends where the next one opens, rather than a fixed number of days after it
  opens, so a timeline whose periods are not all the same length still counts every month
  once. Only the last period has no successor to ask, and it runs on as far as the period
  before it did. Whole-month timelines are untouched and answer exactly what
  they answered before, six-monthly ones included: two- and six-month intervals were never
  affected, measured in Excel, because the schedule arithmetic is generic in the count.
  `PpY` divided twelve by that count, nothing read it, and it is gone.

- **`nb.Depreciateλ` dropped the last period of any timeline shorter than a month.** Every
  period but the last takes its end date from the next period's start. The last has no
  successor and was given `EDATE(` its own start`, MpP) - 1`, which at nought months is the
  day *before* it opens, so nothing could fall inside it. Forty-eight weekly periods from
  1 January 2026 reported **1,833.37** of a 2,000.00 year: December went missing outright,
  while forty-nine weekly periods and twelve monthly ones both reported 2,000.00. The last
  period now ends one period on from its own start, counted in months where the timeline is
  monthly or longer and in days where it is shorter.

  Nothing else about it was broken. The published source carried two `SWITCH` lookups that
  listed only monthly, quarterly and yearly and returned `#N/A` for anything else, and this
  changelog has twice named them as a reason the function fails on other intervals. That
  was wrong, and Excel says so: each is read by exactly one binding, nothing reads those
  two bindings, and Excel never evaluates them. Two-, four- and six-month timelines have
  always returned correct schedules. All four bindings are deleted rather than generalised,
  and the parameter table no longer names three intervals as though they were the only ones.

- **`nb.Depreciateλ` stopped responding when a life in years was not a life.** Its
  arguments read `InitialValues`, `InServiceDates`, `LifeInYears`, `Timeline`. Transposing
  the middle two puts a date serial where the life belongs, and 1 January 2026 is 46,023,
  so the function is asked for a schedule 552,276 months long and Excel stops answering
  rather than erroring. It now checks that every life is a number over 0 and no more than
  100 and returns a message naming the argument order, and clamps the life it uses as well
  as reporting it, so the arrays are never built. A life of 100 years still works; 101, 0,
  a date and a word are all refused. A life that arrives as text rather than as a number
  still works, because Excel coerced it before and refusing it now would be a regression:
  the check reads the value rather than the type.

- **`nb.Allocateλ` rebuilt its whole answer on every pass.** It accumulated with `HSTACK`
  inside a `REDUCE`, so each new group copied everything already built and the work was
  quadratic in the number of amounts. It is quick at any sane input, and it is what turned
  the transposed call above from an error into a workbook that stops responding: 46,023
  amounts spread over 552,276 months is of the order of 10^10 element copies. The answer is
  a closed form, so the row is sized once and each column computed from its own index.

  Every published result is unchanged to the digit, the deliberate last-column adjustment
  included: it is still the amount less the `SUM` of the same array of equal parts, in the
  same order, so even the floating-point residue is identical. The demonstration sheets
  prove it, since `nb.Depreciateλ` allocates through this function on every one of them and
  not one depreciation figure moved.

- **`nb.InterestLRVλ` understated the interest in the period a debt is retired.** It solves
  for the interest on the average balance over a period, taking the principal repayment as
  the cash available for debt service less that interest. Nobody repays more than they owe,
  and its caller says so: `nb.DebtSculptVariableLRVλ` caps the payment at the principal.
  This function did not, so wherever the cap binds it solved a repayment larger than the
  debt, put the average balance below half the opening balance, and reported too little
  interest. On 1,000 at 5% a period with 1,200 of cash it returned **20.51** where the
  balance runs from 1,000 to nil and the interest is **25.00**. The repayment is now capped
  inside the iteration, both where it is assumed and where it is solved.

  It converges faster, not slower: once the cap binds, the assumed and solved repayments
  are both the principal and the first pass is the last. The published example is nowhere
  near the cap and still prints 222.90. No balance anywhere moves, because a closing
  balance is the principal less the payment and never read the interest, which is exactly
  why the six balance identities in the self-test could not see this.

- **Two help-text misspellings, the last of their kind.** `nb.TimelinePositionλ`'s
  parameter table spelled the word the function is named after "timline", which v2.5.0
  fixed in `nb.TimelineOffsetλ` and named as still outstanding here. `nb.LabelDepreciateλ`'s
  said "teh".

- **A licence, for the parts of this repository that can carry one.** [LICENCE](LICENCE) is
  MIT and covers `tools/`, `.github/`, the Markdown files and `assets/`. It does not extend
  to `nabla.xlsx`, `src/` or `functions.csv`: those derive from a workbook whose author
  retains all rights, as `ATTRIBUTION.md` has said since v1.0.0, and nothing here can grant
  what is not ours to grant.

  `ATTRIBUTION.md` now also records the build input. The upstream workbook is not in this
  repository and cannot be, so it names the file's size, its sha256 and its part count,
  which is enough to tell a rebuild from the same input apart from a rebuild from a
  different one.

- `tools/excel_selftest.ps1` goes from 197 assertions to 259, and waits for Excel to finish
  calculating before reading them. It did not: it asked for the answers immediately after
  ordering a full rebuild, and a range Excel is still calculating hands back a null rather
  than a partial answer. With the lighter assertion set that never lost the race. The new
  ones are heavy enough to lose it, and the failure looks like a crash in the harness
  rather than a slow workbook, so it now waits and says so if the answers never arrive.

  Twenty-eight cells changed across the whole workbook, all of them text, all on the four
  sheets that display the help of the functions above. No numeric cell anywhere differs
  from v2.5.0.

- A note for whoever adds a line of help next: a function's help spills down its own
  demonstration sheet, and the sheets are laid out with between nought and seven free rows
  underneath. `nb.Amortiseλ` has one. Two added rows blocked its spill and the whole help
  block came back an error, which `tools/verify_cache.py` caught and the self-test's
  error-cell scan would have caught too.

## v2.5.0, 19 August 2026, an example that runs

- **`nb.TimelineOffsetλ`'s worked example could not be run as printed.** The call was
  missing the two closing brackets that finish `EDATE` and the function call itself, so
  copying the one line a reader is meant to copy got a syntax error rather than an answer.
  The Result column beside it was empty, where every other example in the library prints
  what it returns, which is how the missing brackets went unseen: there was no answer to
  disagree with. Upstream wrote it against 2/15/2022 and a timeline starting 1/1/2023, and
  the date sweep moved both forward two years with everything else, which still left the
  example two years behind the 1 January 2026 timeline the demonstration sheet builds.

  It is now two rows against that same timeline, one date inside it and one before it,
  because a date falling before a model's timeline is what the function's own discussion
  comment is written for. `tools/excel_selftest.ps1` runs both and holds them to the
  printed results, 1 and -11, so the example cannot drift from what it claims again. The
  parameter table above it also spelled "timline", which is now the word the function is
  named after.

  Five cells changed, all of them on the sheet that displays this function's own help. No
  other sheet, no formula and no defined name's behaviour differs from v2.4.0.
  `nb.TimelinePositionλ`'s parameter table carries the same "timline" spelling and is
  untouched.

## v2.4.0, 19 August 2026, period starts that exist

- **`nb.PeriodStartλ` returned a date that is not a period start whenever the anchor is a
  month end.** It walked the calendar to the period's month, rebuilt the date with the
  anchor's day of the month, and corrected whatever overflowed by exactly one day. Anchored
  on 31 January, monthly, that asks for 31 February: Excel reads it as 3 March, one day back
  off it is 2 March, and 2 March is in neither the right month nor on any period boundary.
  28 February, which is where that period does start, was unreachable.

  A schedule anchored on a month end is what `EDATE` describes: the anchor's day of the
  month where the target month has one, the month's own end where it does not, so 31 January
  monthly runs 31 Jan, 28 Feb, 31 Mar. The procedure now counts the whole periods from the
  anchor to the date of interest and steps the anchor on by that many, then steps back one
  period where a truncated quotient overshoots.

  Measured against that schedule over 9,900 cases, being eleven anchors by five period
  lengths by 180 dates at 13-day steps from January 2024, the old procedure was wrong 169
  times. Every one had an anchor on the 29th, 30th or 31st, 29 February included; anchors on
  the 1st, 15th and 28th were already right and answer exactly as before.
  `("31/1/2026", 1, "5/3/2026")` gave 2 March 2026 and now gives 28 February 2026.

  The new procedure compares two dates rather than taking months and days out of them, and
  Excel orders every number before any text, so a date written as text would never compare
  as a date. Both date arguments are converted first, the same way the Dates module's
  functions do it, and `tools/verify_sources.py` already fails a conversion that is bound
  and never read.

- **`nb.TimelineOffsetλ` divided by zero on every daily, weekly and fortnightly timeline.**
  It reads the interval off the timeline's first two dates and rounds it to whole months,
  which is zero for anything shorter than about a fortnight, and the next line divides by
  it. Every such call returned `#DIV/0!`.

  A sub-monthly period is a fixed number of days, which is what makes it easy: the offset is
  now the day difference floored by that count, so a date 20 days into a weekly timeline is
  in period 2 and one three days before it starts is in period -1. The month path is
  untouched and answers exactly what it answered before on monthly, quarterly and yearly
  timelines, month-end anchored ones included, over 200 dates at 11-day steps.

  This does not make `nb.Amortiseλ` work on a sub-monthly timeline. It calls
  `nb.TimelineOffsetλ` on whatever timeline it is handed, so it did inherit this failure,
  but it rounds the same interval to months itself and then divides twelve by it, so a
  weekly timeline still fails there on its own arithmetic. `nb.Depreciateλ` holds a third
  copy of the rounding and turns nought months into `#N/A`. Neither is changed here.

`tools/excel_selftest.ps1` now runs 195 assertions rather than 152. Of the 43 new ones, 19
compare `nb.PeriodStartλ` against an `EDATE` schedule written out in the test rather than
borrowed from the function, and 4 hold `nb.TimelineOffsetλ`'s month path to the same
comparison, so a future change to either cannot quietly redefine what a period start is.

Only `xl/workbook.xml` and the formula-environment store changed. No worksheet part differs
from v2.3.0 and all 20,221 cached values are identical: the sheet that demonstrates
`nb.TimelineOffsetλ` uses a monthly timeline, and no sheet calls `nb.PeriodStartλ`.

## v2.3.0, 19 August 2026, a file that agrees with itself

Two things this release does that no earlier one could. The workbook now holds the numbers
its own formulas produce, rather than whatever the build last left in it: 3,193 cached
cells across 43 of the 50 sheets were stale, and Excel hid every one of them by
recalculating on open. And it no longer carries a path off the machine that built it,
which every release since v1.2.0 has published.

The rest came from a full audit: thirteen readers across the build, the gates, the
workbook, the LAMBDA sources, the prose and the CI, each finding put to a second pass that
tried to refute it. Two functions were comparing a raw date against a converted one, a
worked example called a different function, `functions.csv` disagreed with the workbook it
is generated from, and nine comment banners named or described the wrong function.

Two gates are new or widened as a result, and one stopped asking for something that is no
longer true.

- **Every release since v1.2.0 published a path off the build machine.** Excel stamps
  `xl/workbook.xml` with the directory a file was last saved from, so `nabla.xlsx` carried
  `C:\Users\<account>\Downloads\` and nothing in the pipeline knew the field
  existed. On this machine the
  account renders as a placeholder, so no name went out, but the same field built anywhere
  else carries the account name into a public artefact. The build strips it, the refresh
  strips it again because Excel puts it back on every save, and it joins the banned tokens
  so CI fails rather than trusting either step to have run.

- **Two more functions compared a raw date against a converted one,** the defect their
  non-ByItems siblings had until v2.2.0. The check added then could not see these two: each
  does read its conversion, but only to pass it to the recursive call, so the first row of
  the result came from the raw argument and every row below it from the converted one, off
  the same call. Given period bounds written as text, `nb.ScheduleValuesByItemsλ` returned
  0 where the answer is 100, and `nb.ScheduleRatesByItemsλ` returned the 20 January rate
  where the 1 January rate applies. Both now match their real-date answers exactly.

- **`nb.DebtSculptVariableLRVλ`'s worked example called its sibling** and was missing its
  closing bracket, so the one line a reader copies ran the other schedule, which after
  v2.2.0 no longer means the same thing. `tools/verify_signatures.py` was added in v2.1.0
  to catch exactly this and structurally could not: it skipped the text on the `EXAMPLES`
  row itself and read only the rows below, which left **43 of the 119 example blocks
  unread**, including every function that puts its example on the label row. It now reads
  all 119, and against v2.2.0 it names this defect by function.

- **`functions.csv` disagreed with the workbook it is generated from.** `nb.Depreciateλ`
  published a signature cut off mid-parameter list with a bracket left open, because its
  help wraps onto a second row and the exporter read only the first. 31 descriptions
  carried `_x000a_`, the raw OOXML escape for a line break out of a Name Manager comment,
  and two began with the help table's column delimiter. All fixed at the source, and
  `tools/verify_previous_names.py` now checks the two columns nothing checked: brackets
  balance, no raw escapes, no leading delimiter, nothing empty.

- **`docProps/app.xml` described a 49-sheet workbook that has 50.** The build registered
  the Australian tax worksheet in the content types, the relationships, the workbook, the
  contents table and its table range, and not in the file-properties list, which is what a
  reader's properties dialogue and any tool that trusts it will show. The build now lists
  it, and `tools/verify_workbook.py` compares that list against the real sheets in order.

- **`ATTRIBUTION.md` still called the depreciation helpers ATO methods.** v2.1.0 removed
  that claim from the function help, the method codes, the Data Validation sheet, the
  Australian tax worksheet and the README, and missed the one file that describes the
  derivative to a stranger.

- **Nine comment banners named or described the wrong function.** Every function is
  introduced by a comment naming it and saying what it does. They are stripped before
  anything reaches the workbook, so no part of the build and no gate has ever read them,
  and they are the one piece of the published source that drifted with nothing watching.
  Three misnamed their function: `→SumContainsλ` with the help table's arrow on it,
  `InterestCoverateRatioλ` and `PriceToCashsRatioλ`. Six described a different one:
  `CountColsλ` and `CountAColsλ` in both modules said "row" where they count columns, and
  `IsInListλ` in both said it tests whether a value falls between two limits. Each is
  anchored to the name above it, because the same wording is correct above `CountRowsλ`
  and above `IsBetweenλ`.

- **`nb.MaxColsλ` told the reader it returns the minimum,** copied from `nb.MinColsλ` and
  never changed, in both the Essentials and Utilities copies. `nb.CountColsλ` published its
  signature as `CountColsλ( Array,)`, with an empty second argument in it.

  The last three entries came from an outside review, checked one by one against the
  source. Six of its nine banner claims landed on wording that is correct where it sits, so
  only the three it named plus the six that are genuinely wrong were changed.

- **The workbook now ships the numbers its own formulas produce.** An `.xlsx` stores two
  things for every calculated cell, the formula and the answer Excel last got from it, and
  nothing keeps them in step. The build edits the workbook as XML and has no formula
  engine, so every cell downstream of a value it changes keeps the answer it had before.
  Shifting the sample dates forward two years left **3,193 cached cells across 43 of the 50
  sheets** holding numbers their own formulas no longer produce. `nb.IsOccurrenceDateλ`
  alone accounted for 978 of them.

  Nobody ever saw one. Excel recalculates on open and replaces the lot, which is precisely
  what made this worth finding: the file could be wrong in a way only a second tool could
  see, and everything that reads an `.xlsx` without a formula engine, from a diff to a
  converter to a web preview, reads the cached answer. It is the same defect as the five
  saved `#VALUE!` cells fixed in v2.2.0, three orders of magnitude wider.

  `tools/refresh_cache.py` recalculates the workbook in Excel and saves it, then clears the
  always-calculate flags Excel puts back. `fullCalcOnLoad` comes off with them: that flag
  exists to force a recalculation past stale values, and there are none left. The workbook
  now opens reading correctly without recalculating, so Excel no longer asks to save a file
  the reader never edited.

  Nothing about the functions changes. Measured against the build's own output: **all 4,506
  worksheet formula cells byte-identical, all 130 defined names byte-identical, the
  Advanced Formula Environment store byte-identical.** The file grows by 247 bytes, from
  435,267 to 435,514. Excel's own save is 75 KB larger, because a recalculated value is
  longer than the stale integer it replaces and Excel adds a calculation chain, but the
  refresh rewrites the archive afterwards and compresses it harder than Excel does, which
  gives all of that back.

### Added

- `tools/verify_cache.py`, a sixth gate and the second that needs Excel. It opens the
  workbook, recalculates, and compares every one of the 20,221 cached values against what
  the formula produces. Run against the previous build it reports all 3,193 by sheet, row
  and column. It refuses to pass on fewer than 15,000 comparisons, so a run that silently
  read almost nothing fails rather than reporting success.
- `tools/refresh_cache.ps1` and `tools/dump_values.ps1`, the two Excel steps the pair above
  drive. Both fold in a lesson that cost real time: `CalculationState` comes back over COM
  as the name `xlDone` rather than as `0`, so the obvious test against `0` is true forever
  and a wait built on it never ends early.

### Changed

- CI stopped installing `openpyxl`, which nothing has ever imported: the build is
  deliberately pure zip and XML surgery, and the dependency was installed on every run to
  be ignored. The workflow now also cancels a run its own successor supersedes, and asks
  for read-only access to the repository rather than the default write scope.
- `tools/verify_cache.py` fails when a sheet holding cached values contributes no
  comparisons at all. A floor on the total is not enough on its own: one sheet dropping out
  of the dump would still leave twenty thousand comparisons and look like a pass.
- `tools/verify_workbook.py` no longer demands `fullCalcOnLoad`. There are two honest ways
  for a reader to see the right numbers, and that flag is only one of them. It now fails
  when a workbook has neither `fullCalcOnLoad` nor the calculation chain Excel leaves
  behind when it saves, which is the case where the file would open showing whatever the
  build last left in it.


- **`nb.DebtSculptVariableLRVλ` added each period's interest back into a balance the same
  period's cash had already paid.** It worked out the payment as the debt service less the
  interest, which is only right if the interest is paid, and then set the closing balance to
  the principal plus the interest less the payment, which is only right if it is not. Both
  lines were there, so every closing balance carried one period's interest too much and the
  error compounded into the next opening balance.

  On 1,000 of debt at 6% with CFADS of 300 and a DSCR of 1.2 over five periods, it repaid
  1,078.60 of a 1,000 loan and still reported 92.80 outstanding. The cash paid was 1,250
  against 1,171.40 actually owed; the 78.60 overpaid plus the 92.80 still owing came to
  171.40, which is the whole interest charge counted a second time.

  The payment is the principal repayment, so it is now capped at the principal rather than
  at the principal less the interest, and it comes off the balance rather than being added
  to it. A negative repayment needs no special case: when the cash cannot cover the
  interest, subtracting a negative capitalises the shortfall, which is what should happen.
  The same schedule now repays exactly 1,000 and closes at 0. With no cash at all the debt
  grows by one period of interest rather than two, and 1,500 of cash against 1,000 of debt
  clears it in one period rather than leaving 30.93 behind.

- **Two sculpting functions told the reader to label the wrong row.** `nb.DebtSculptFixedλ`
  and `nb.DebtSculptVariableλ` both suggest "Principal repayments" for their third row,
  which holds the whole debt service: on the figures above that row reads 250 where the
  principal repaid is 190. Their arithmetic was never wrong, only the label, which now reads
  "Debt service (interest and principal)". `nb.DebtSculptVariableLRVλ` keeps the original
  label, because with the fix above its third row really is the principal repayment.

- **The debt module had never had a numeric check of any kind,** which is how this shipped
  from v1.2.0 to v2.2.0. `tools/excel_selftest.ps1` now runs 152 assertions rather than 138.
  The fourteen new ones are balance identities rather than expected figures, so a schedule
  that satisfies them cannot be double-counting whatever the inputs: repayments retire the
  principal exactly, closing equals opening less repayment, each opening is the previous
  closing, the cash used never exceeds CFADS over DSCR, the balance never goes negative, and
  the same roll-forward holds for the two functions that pay the whole debt service. Run
  against v2.2.0, seven of the fourteen fail, and they report the figures above by name.

  Not fixed, and worth knowing before relying on the final period of a sculpted schedule:
  `nb.InterestLRVλ` computes interest on the average balance over the period, but it is not
  told about the repayment cap, so in the one period where the cap binds it still assumes
  the larger uncapped repayment and understates the interest. In the five-period schedule
  above that is 0.97 in the final period against 4.22; with 1,500 of cash against 1,000 of
  debt it is 15.46 against 30.00. The balance is right either way, because the repayment is
  capped at the principal.

  The five debt functions also remain absent from the workbook's embedded Advanced Formula
  Environment store, so the workbook still cannot be rebuilt from that store alone. They are
  the only self-recursive functions in the library, each calling itself by name, and an
  imported module takes its prefix from its container, so the recursion would call a name
  that does not exist there. Upstream leaves the same five out for the same reason. `src/`
  carries all 130 and `tools/verify_sources.py` checks all 130 on every push.

## v2.2.0, 18 August 2026, counting the way the help counts

v2.1.0 listed what it knew was still wrong and left it, because each of these changes
moves results for anyone already relying on them. This release makes them right.

`nb.Periodsλ` now counts the way its own four examples count, which returns one more
period wherever an end date falls part way through one. `nb.ScheduleValuesλ` and
`nb.ScheduleRatesλ` now read the date conversions they were computing and throwing away,
so both answer correctly when given dates written as text. Nothing is renamed, and
formulas written against v2.1.0 keep working. Across all 20,813 recalculated cells in the
workbook, four moved, all four on the Periods demonstration sheet.

- **`nb.Periodsλ` returned one period fewer than its own examples claim.** It counted whole
  intervals, which is what `DATEDIF` returns, while every one of its four worked examples
  counts the period starts crossed between the two dates. Its description says as much: it
  lists "End Date is inclusive" as one of its differences from `DATEDIF`, and then the
  procedure called `DATEDIF`. From 31 March to 15 May is one whole month and two month
  starts, and the help says 2. It now counts ordinals, the same way for all five intervals,
  so a part period at the end counts once and a whole one does not count twice.

  `D` is unchanged, because a day ordinal is the serial number itself. `W` follows
  `nb.PeriodLabelλ`'s own week numbering, which restarts each 1 January and so labels every
  year with 53 weeks, the last of them one or two days long. That is what makes the help's
  fourth example -53 rather than the 52 whole weeks its two dates are apart. All four
  examples now hold: 2, -2, -12 and -53.

  This changes results. Anything measured in months, quarters, weeks or years returns one
  more than it did wherever the end date falls part way through a period, which is most of
  the time. The two demonstration cells that move are the quarters row, 4 to 5, and the
  weeks row, 52 to 53. If you were relying on whole-interval counts, `DATEDIF` is still
  there and still does that.

  The third example was also the one line in this help that could not be copied: it passed
  four arguments to a function that takes three. The `-12` it claims is right without the
  fourth.

- **Four more functions threw away their date conversions,** the defect `nb.OverLapDaysλ`
  carried until v2.1.0. Two of them gave wrong answers for it:

  - `nb.ScheduleValuesλ` compared text period starts against converted effective dates, so
    nothing ever fell inside a period. It returned 0 where it should return 100.
  - `nb.ScheduleRatesλ` looked up text dates in a text array, so `XLOOKUP` matched them in
    dictionary order. Asked for the rate in force on `"5/1/2026"` against starts of
    `"1/1/2026"` and `"10/1/2026"`, it returned the January 10 rate, because `"5/1/2026"`
    sorts after `"10/1/2026"` as text.
  - `nb.PeriodLabelλ` and `nb.Timelineλ` were saved by Excel: both feed the date straight
    into `TEXT`, `YEAR`, `EDATE` or `SEQUENCE`, all of which coerce a text date themselves.
    Checked against every case tried, they returned the right answer before and after. The
    fix is a guard, not a repair.

- **The workbook stopped shipping cached errors it does not reproduce.** Five cells on the
  Periods demonstration sheet were saved holding `#VALUE!`. They come from the upstream
  workbook, which carries 65 such cells, and they have been in every tagged release from
  v1.2.0 on. Excel replaces them with the right answers as soon as the file opens, so no
  reader ever saw them and no recalculation reproduces them, but a file that disagrees with
  itself is a file nobody can check. They now hold the answers the formula gives. One
  cached error is left, `#NAME?` on the Essentials About sheet, which resolves on open the
  same way; its cell holds a spilled table rather than a number, so it needs the table
  built rather than a value written.

### Added

- `tools/verify_sources.py` now requires that a function which converts a date argument
  goes on to read the conversion. It is the check that would have caught all five of these,
  and it could not be added before they were fixed, because it fails on them. Run against
  v2.1.0 it names all six bindings and the function each belongs to. A general unused-value
  check would need a real parser and would report a great deal more; this one asks a
  narrower question and gets a clean answer.

## v2.1.0, 18 August 2026, functions that do what they say

v2.0.0 renamed every function. This one fixes what three of them compute, and stops two
more from claiming to be something they are not.

Nothing is renamed, so formulas written against v2.0.0 keep working. Three functions do
return different numbers, and only for the inputs they were handling wrongly: dates passed
as text to `nb.OverLapDaysλ`, a start date after the end date in `nb.Periodsλ`, and
`No_Switch` set to `TRUE` in `nb.VDBλ`.

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
  example returns 1 against the 2 it claims, and its `W` example 52 against 53. Its
  demonstration sheet also ships five cached `#VALUE!` cells, which Excel replaces with
  the right answers the moment the file opens. (This entry first said the function returns
  `#VALUE!` for range arguments. It does not; that claim came from a probe that named a
  worksheet which does not exist. The cached errors are stale, not live. Corrected the
  same day, here and in the published release notes.) And four more functions ignore their
  date
  conversions the way `nb.OverLapDaysλ` did: `nb.PeriodLabelλ`, `nb.ScheduleRatesλ`,
  `nb.ScheduleValuesλ` and `nb.Timelineλ`, six dead conversions between them, all in Dates.

- **The depreciation helpers no longer claim to be ATO methods.** `nb.DiminishingValueλ` described itself as the ATO 200% diminishing value method and `nb.PrimeCostλ` as the ATO prime cost method. Neither is one. They take a cost and an effective life and nothing else: no acquisition date, no income year, no days held, no disposal. `nb.DiminishingValueλ` also writes the entire undeducted residual off in its final period, so for a cost of 1,000 over five years it returns 400, 240, 144, 86.40 and 129.60, where a diminishing-balance calculation would deduct 51.84 that year and carry the rest forward. The schedules are unchanged and still useful for modelling; what changes is that they are now described as modelling schedules rather than tax calculations, in the function help, the method codes on `nb.Depreciateλ`, the Data Validation sheet, the Australian tax worksheet and the README. The worksheet now says so on its face, and the README says plainly not to use them to prepare a return.

  Raised by an external audit of v1.2.6. Its arithmetic checks out: the fifth-year deduction under a plain diminishing-balance calculation is 51.84, not the 129.60 residual the function returns.

- **Three worked examples called a different function than the one they document.** `nb.RollingMinλ`'s example called `nb.RollingMaxλ`, `nb.RollingSumλ`'s called `nb.RollingMinλ`, and `nb.ScheduleValuesByItemsλ`'s called `nb.ScheduleRatesByItemsλ`. The examples are the lines a reader copies, so copying one as printed ran a different function. Each claimed result was correct for the function it wrongly called, which is exactly why nothing looked wrong.
- **`nb.RollingAvgλ` claimed a result belonging to `nb.RollingSumλ`,** the same copy going the other way: its call was right and its answer, `1,3,6,9,12`, was the running total rather than the running average. Excel gives `1,1.5,2,3,4`, which is what it now says.
- **`nb.Amortiseλ`'s help pointed at a function that does not exist,** `LableAmortiseλ`, a transposition of `LabelAmortiseλ`.
- **The build's help corrections can no longer reach a live formula.** Correcting help text also refreshes any copy of that help already spilled and cached on a demonstration sheet, and that refresh was a blanket text replace across the whole worksheet. Fixing `nb.ScheduleValuesByItemsλ`'s example rewrote a real formula on a neighbouring sheet, changing which function it called and spilling `#SPILL!` across it. The Excel gate caught it before it shipped. The refresh is now confined to cached values, never formulas, and a parameter label must match a whole cell rather than appear anywhere inside one.
- **`tools/verify_signatures.py` now reads the worked examples too.** It already checked the signature and the parameter table; it now also requires that a function's examples call that function, and that every function named anywhere in a help is one the library declares. Both are pure text, so both run in CI. Run against the source before these fixes it reports all four defects above by name. 76 example blocks are checked.

- **v2.0.0's own release text carried three defects the rename left behind.** The flat-namespace substitution ran over the prose as well as the code: the README's coverage note became "all of `nb`, `nb` and `nb`" where it meant Ratios, Utilities and Debt, and this changelog named `nb.IsInListλ` twice where the second is `nb.IsInListUλ`. It also said 89 functions change prefix and nothing else; 89 is only the count that never collided, and the true figure including the 17 collision winners is 106, checked against `functions.csv`. The cell-comparison claim omitted the capitalisation of the product name from `nabla` to `Nabla` in 29 places, done in the same build. Corrected here and in the published v2.0.0 release notes.
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
