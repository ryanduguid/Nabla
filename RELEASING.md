# Releasing Ozzit

## Scope and history

This policy applies to future releases after it is merged into the default
branch. It does not alter or certify historical releases, their tags or their
assets.

Correct a published or verified candidate under a new version and tag. Never move or retarget an existing tag or overwrite or replace a verified asset.
A draft upload is a candidate only; it is not approval to publish a release.

## Candidate and source ownership

Start from a clean, isolated checkout of the exact candidate commit. Record the
full commit SHA and the pre-verification SHA-256 and byte length of
`ozzit.xlsx`.

`ozzit.xlsx` is the shipped authority. `src/*.txt`, the Advanced Formula
Environment (AFE) store and `functions.csv` are bound publication views that
must agree with the exact workbook in the same tagged commit. Do not approve
changes to `src/` or `functions.csv` in isolation. A formula change must update
the compiled defined names and every publication view together through a
documented source-owning change process.

The v3.0.0 tracked builder starts from its disclosed upstream input, which
is not committed, and stops at the v3.0.0 artefacts. Post-v3.0.0 passes start
from the committed `ozzit.xlsx` and `src/` recorded in `ATTRIBUTION.md`. They
do not regenerate the current workbook from upstream. The current result also
includes a one-off Excel-state-dependent date shift. The consistency checks do not prove
byte-for-byte regeneration of the current workbook. The v3.0.0
builder, the postbuild limits and any future one-off transformation remain
disclosed in `ATTRIBUTION.md` and `CHANGELOG.md`. Do not broaden a
reproducibility claim without a complete regenerator and a byte comparison.

MIT covers `tools/`, `.github/`, Markdown files and `assets/` written for this repository. MIT does not cover `ozzit.xlsx`, `src/` or `functions.csv`.
No open-source licence was located for the derived upstream material, so its author retains the relevant rights. Every source archive must include `ATTRIBUTION.md` and `LICENCE` and state this licence split. Do not label the whole workbook and source bundle as MIT, and do not generate an SPDX or CycloneDX document that assigns MIT to the derived material.

## Approval and tag

A human maintainer explicitly approves the exact candidate commit and release version after all nine gates pass.
Create an annotated, cryptographically signed tag for that exact commit. Run
`git verify-tag` and record the tag object SHA, peeled commit SHA and
verification result.

Do not let `gh release create` silently create a lightweight tag from a branch.
Tag signing, remote tag publication and release publication are separate authorised actions; this document does not approve any of them.

## Future release bundle

Use stable names derived from the approved tag:

1. `ozzit.xlsx` — the consumer workbook from the tagged tree.
2. `Ozzit-<version>-source.zip` — a `git archive` of the exact signed tag, including the tagged workbook, `src/`, `functions.csv`, tools, policy, `ATTRIBUTION.md` and `LICENCE`.
3. `Ozzit-<version>-verification.txt` — captured gate evidence, Excel version and build, formula, error and assertion counts, cached-value comparison count, candidate commit and workbook before-and-after hashes.
4. `release-manifest.json` — schema version, release version, tag object and peeled commit SHAs, workbook and source-archive hashes and sizes, build-input provenance, licence scopes, every payload asset name and media type, gate results and verifier identity and date.
5. A genuine SPDX or CycloneDX inventory only when applicable and generated from real distributable components or dependencies. A cosmetic or all-MIT inventory is prohibited. The release manifest is the appropriate inventory for the present workbook-only project.
6. `SHA256SUMS` — SHA-256 for every payload asset except `SHA256SUMS` itself, including the manifest, verification evidence and any genuine SBOM.

The standalone workbook, archived workbook and tagged workbook must have the same SHA-256.
The independent verifier must record and compare GitHub's API `digest` for the uploaded checksum file. If a detached signature is added later, define whether it sits outside the checksum set and verify it separately; do not create a checksum and signature cycle.

## Verification gates

Run the following seven static and repository gates from the clean candidate:

```powershell
python tools/verify_workbook.py ozzit.xlsx
python tools/verify_sources.py ozzit.xlsx src
python tools/verify_signatures.py src
python tools/verify_previous_names.py functions.csv
python tools/verify_index.py ozzit.xlsx src functions.csv
python tools/verify_afe.py ozzit.xlsx src
python -m unittest discover -s tools/tests -v
```

All seven commands must exit zero. Record their substantive function,
signature, table, example, index, module and test counts rather than only their
exit status.

Before either native gate, prove that no user Excel process is running. Do not
close or attach to a user's Excel session. Then run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\excel_selftest.ps1 -Path .\ozzit.xlsx
python tools/verify_cache.py ozzit.xlsx
```

The current acceptance baseline is 1,129 formulas recalculated with zero error cells and 259 assertions run with zero failures.
The cached-value gate must pass and report its actual comparison count. Record
the Excel version and build. Hash the workbook immediately before and after
both native gates and require the bytes to be byte-identical.

If a reviewed workbook change deliberately changes a formula, assertion or
cached-value count, update the documented expected count in the same candidate
and review that change. Do not silently weaken a floor.

## Packaging, independent verification and immutability

1. Build every asset from the signed tag into a fresh staging directory.
2. Compute the manifest and `SHA256SUMS` only after every payload asset has its final bytes.
3. Upload the assets to a draft release. A draft upload is not approval, publication or evidence of immutability.
4. An independent verifier downloads every asset from GitHub rather than reading the builder's local staging directory.
5. The independent verifier checks every checksum, API digest, filename and size; verifies the signed annotated tag; proves the standalone, archived and tagged workbooks have the same SHA-256; inspects the source archive; and repeats the static and native gates on the downloaded candidate where the required Excel environment is available.
6. Record the independent result in the release evidence. Only after it passes may an authorised maintainer publish the release.
7. Re-query the published release. Claim immutability only when GitHub reports `immutable=true`. If the feature is unavailable, record that limit and retain the no-overwrite and new-version rule.
8. After verification, do not delete or replace assets, move the tag or edit evidence to make a failed candidate appear successful. Correct an error under a new version and tag.
