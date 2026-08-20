import difflib
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SECURITY = ROOT / "SECURITY.md"
RELEASING = ROOT / "RELEASING.md"
VERIFY_WORKFLOW = ROOT / ".github" / "workflows" / "verify.yml"
DEPENDABOT = ROOT / ".github" / "dependabot.yml"


# These reviewed documents are deliberately exact. A policy change must update
# both the owning document and its canonical contract in the same review.
SECURITY_CANONICAL = """# Security policy

## Supported versions

Only the latest published GitHub release is supported for security fixes.
Older releases and unreleased branches are not supported release lines.
Publishing a newer release supersedes the previously supported release.

## Reporting a vulnerability

Report a suspected vulnerability through [GitHub private vulnerability
reporting](https://github.com/ryanduguid/Ozzit/security/advisories/new). The
form's availability depends on the live GitHub private vulnerability reporting
setting for this repository.

Do not disclose a suspected vulnerability in a public issue, discussion, pull request or commit before coordinated disclosure.

Include the affected release, impact, reproduction steps, suggested mitigation
and a minimal synthetic reproduction. Do not upload client workbooks, real
client or production data, credentials, access tokens, private keys, session
material, private URLs, .env files or other sensitive files. The public
`ozzit.xlsx` release artefact is not a client workbook, but a user workbook or
an extract containing client information remains sensitive.

## What this library does and does not do

Ozzit is an Excel LAMBDA workbook plus Python tools that verify and rebuild it.
The tools read local files and make no network call. They hold no credentials
and have no runtime dependencies outside the Python standard library.

Do not run the tools with elevated privileges. Treat `ozzit.xlsx` and `src/`
as published artefacts from this repository, not as untrusted user input.
"""


RELEASING_CANONICAL = r"""# Releasing Ozzit

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

The v3.0.0 tracked builder starts from its disclosed upstream input. The
current v3.1.0 result also includes a one-off Excel-state-dependent date shift
and direct workbook and source changes. The consistency checks do not prove
byte-for-byte regeneration of the current workbook. The v3.0.0 builder, the
v3.1.0 postbuild limits and any future one-off transformation remain disclosed
in `ATTRIBUTION.md` and `CHANGELOG.md`. Do not broaden a reproducibility claim
without a complete regenerator and a byte comparison.

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
"""


VERIFY_WORKFLOW_CANONICAL = """name: Verify workbook

on:
  push:
    branches: [main]
  pull_request:
  workflow_dispatch:

# A push that supersedes an earlier one cancels it rather than paying for both.
concurrency:
  group: verify-${{ github.ref }}
  cancel-in-progress: true

# The gates only read files. Nothing here needs write access to the repository.
permissions:
  contents: read

jobs:
  verify:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1
        with:
          persist-credentials: false
      - uses: actions/setup-python@5fda3b95a4ea91299a34e894583c3862153e4b97 # v7.0.0
        with:
          python-version: '3.12'
      - name: Check workbook integrity
        run: python tools/verify_workbook.py ozzit.xlsx
      - name: Check src/ reproduces the shipped functions
        run: python tools/verify_sources.py ozzit.xlsx src
      - name: Check each function's help describes its own parameters
        run: python tools/verify_signatures.py src
      - name: Check the published index records what each function replaced
        run: python tools/verify_previous_names.py functions.csv
      - name: Check the published index matches the current artefacts
        run: python tools/verify_index.py ozzit.xlsx src functions.csv
      - name: Check the Advanced Formula Environment store matches src/
        run: python tools/verify_afe.py ozzit.xlsx src
      - name: Run tool regression tests
        run: python -m unittest discover -s tools/tests -v
"""


DEPENDABOT_CANONICAL = """version: 2
updates:
  - package-ecosystem: github-actions
    directory: /
    schedule:
      interval: weekly
    cooldown:
      default-days: 7
    open-pull-requests-limit: 2
    groups:
      # init and analyze are separate dependencies to Dependabot but must
      # move together: a job running one release against the other's config
      # fails outright. Grouped so the pair can never arrive split.
      codeql-action:
        patterns:
          - "github/codeql-action*"
"""


def read_utf8(path: Path) -> str:
    return path.read_bytes().decode("utf-8")


def normalise_line_endings(text: str) -> str:
    """Treat CRLF and LF as equivalent, without changing any other byte text."""
    return text.replace("\r\n", "\n")


def assert_canonical_document(actual: str, expected: str, label: str) -> None:
    actual = normalise_line_endings(actual)
    expected = normalise_line_endings(expected)
    if actual == expected:
        return
    difference = "".join(
        difflib.unified_diff(
            expected.splitlines(keepends=True),
            actual.splitlines(keepends=True),
            fromfile=f"canonical/{label}",
            tofile=label,
        )
    )
    raise AssertionError(f"{label} differs from its reviewed canonical text:\n{difference}")


def replace_once(text: str, old: str, new: str) -> str:
    count = text.count(old)
    if count != 1:
        raise AssertionError(
            f"expected one canonical fixture occurrence of {old!r}, found {count}"
        )
    return text.replace(old, new, 1)


def security_mutants(text: str) -> dict[str, str]:
    tilde_fence = replace_once(
        text,
        "## Reporting a vulnerability\n\n",
        "## Reporting a vulnerability\n\n~~~text`example\n",
    )
    tilde_fence = replace_once(
        tilde_fence,
        "\n## What this library does and does not do",
        "\n~~~\n\n## What this library does and does not do",
    )
    return {
        "unclosed HTML comment": replace_once(
            text,
            "## Reporting a vulnerability\n\n",
            "## Reporting a vulnerability\n\n<!--\n",
        ),
        "tilde fence with backtick info": tilde_fence,
        "depends on no setting": replace_once(
            text,
            "form's availability depends on the live GitHub private vulnerability reporting",
            "form's availability depends on no live GitHub private vulnerability reporting",
        ),
        "seven-day guarantee": replace_once(
            text,
            "\n## What this library does and does not do",
            "\nA valid report will be acknowledged in seven days.\n\n"
            "## What this library does and does not do",
        ),
        "please upload client workbooks": replace_once(
            text,
            "\n## What this library does and does not do",
            "\nPlease upload client workbooks when they help.\n\n"
            "## What this library does and does not do",
        ),
        "extra default-branch support": replace_once(
            text,
            "\n## Reporting a vulnerability",
            "\nSecurity fixes also apply to the default branch.\n\n"
            "## Reporting a vulnerability",
        ),
    }


def releasing_mutants(text: str) -> dict[str, str]:
    return {
        "exact v3.1 regeneration": replace_once(
            text,
            "\n## Approval and tag",
            "\nThe tracked builder exactly regenerates v3.1.0.\n\n"
            "## Approval and tag",
        ),
        "checksum includes itself": replace_once(
            text,
            "\n## Verification gates",
            "\n`SHA256SUMS` contains a checksum for `SHA256SUMS`.\n\n"
            "## Verification gates",
        ),
        "draft immutability after download": replace_once(
            text,
            "4. An independent verifier downloads every asset from GitHub rather than reading the builder's local staging directory.",
            "4. An independent verifier downloads every asset from GitHub rather than reading the builder's local staging directory. The draft upload is immutable.",
        ),
        "published asset replacement": replace_once(
            text,
            "\n## Candidate and source ownership",
            "\nPublished assets may be replaced in place.\n\n"
            "## Candidate and source ownership",
        ),
        "workbook MIT": replace_once(
            text,
            "\n## Approval and tag",
            "\n`ozzit.xlsx` is MIT-licensed.\n\n## Approval and tag",
        ),
        "optional native gates": replace_once(
            text,
            "\n## Packaging, independent verification and immutability",
            "\nThe native Excel gates are optional.\n\n"
            "## Packaging, independent verification and immutability",
        ),
        "historical certification": replace_once(
            text,
            "\n## Candidate and source ownership",
            "\nThis policy certifies all historical releases.\n\n"
            "## Candidate and source ownership",
        ),
    }


class RepositoryPolicyCanonicalTests(unittest.TestCase):
    def test_security_and_release_documents_match_reviewed_canonical_text(self):
        for path, expected in (
            (SECURITY, SECURITY_CANONICAL),
            (RELEASING, RELEASING_CANONICAL),
        ):
            with self.subTest(path=path.name):
                assert_canonical_document(read_utf8(path), expected, path.name)

    def test_security_contract_rejects_every_residual_variant(self):
        for name, mutant in security_mutants(SECURITY_CANONICAL).items():
            with self.subTest(name=name), self.assertRaises(AssertionError):
                assert_canonical_document(mutant, SECURITY_CANONICAL, SECURITY.name)

    def test_release_contract_rejects_every_residual_variant(self):
        for name, mutant in releasing_mutants(RELEASING_CANONICAL).items():
            with self.subTest(name=name), self.assertRaises(AssertionError):
                assert_canonical_document(mutant, RELEASING_CANONICAL, RELEASING.name)

    def test_truthful_reproducibility_limit_is_a_valid_canonical_value(self):
        truth = "The tracked builder cannot reproduce current v3.1.0 byte-for-byte.\n"
        assert_canonical_document(
            truth.replace("\n", "\r\n"),
            truth,
            "truthful-reproducibility-probe.md",
        )
        self.assertIn(
            "The consistency checks do not prove\n"
            "byte-for-byte regeneration of the current workbook.",
            RELEASING_CANONICAL,
        )

    def test_line_ending_normalisation_does_not_hide_other_changes(self):
        with self.assertRaises(AssertionError):
            assert_canonical_document("policy\r", "policy\n", "line-ending-probe.md")
        with self.assertRaises(AssertionError):
            assert_canonical_document("policy\n\n", "policy\n", "spacing-probe.md")

    def test_workflow_and_dependabot_match_reviewed_controls(self):
        for path, expected in (
            (VERIFY_WORKFLOW, VERIFY_WORKFLOW_CANONICAL),
            (DEPENDABOT, DEPENDABOT_CANONICAL),
        ):
            with self.subTest(path=path.name):
                assert_canonical_document(read_utf8(path), expected, path.name)


if __name__ == "__main__":
    unittest.main()
