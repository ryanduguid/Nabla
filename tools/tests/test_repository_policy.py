import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SECURITY = ROOT / "SECURITY.md"
RELEASING = ROOT / "RELEASING.md"
VERIFY_WORKFLOW = ROOT / ".github" / "workflows" / "verify.yml"
DEPENDABOT = ROOT / ".github" / "dependabot.yml"

CHECKOUT = (
    "actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1"
)


def read_optional(path: Path) -> str:
    return path.read_text(encoding="utf-8") if path.exists() else ""


def normalise_markdown_prose(text: str) -> str:
    """Treat Markdown line wrapping as whitespace for prose contracts."""
    return re.sub(r"\s+", " ", text).strip()


def markdown_section(text: str, heading: str) -> str:
    """Return one Markdown section through the next peer or parent heading."""
    lines = text.splitlines()
    matches = []
    pattern = re.compile(rf"^(#{{1,6}})\s+{re.escape(heading)}\s*$")
    for index, line in enumerate(lines):
        match = pattern.match(line)
        if match:
            matches.append((index, len(match.group(1))))
    if len(matches) != 1:
        raise AssertionError(
            f"expected exactly one {heading!r} heading, found {len(matches)}"
        )

    start, level = matches[0]
    end = len(lines)
    for index in range(start + 1, len(lines)):
        match = re.match(r"^(#{1,6})\s+", lines[index])
        if match and len(match.group(1)) <= level:
            end = index
            break
    return "\n".join(lines[start + 1 : end]).strip()


def workflow_steps(text: str) -> list[dict]:
    """Parse action steps and bind each step's with keys to that step only."""
    lines = text.splitlines()
    steps = []
    for index, line in enumerate(lines):
        match = re.match(r"^(\s*)steps:\s*(?:#.*)?$", line)
        if not match:
            continue
        steps_indent = len(match.group(1))
        index += 1
        while index < len(lines):
            current = lines[index]
            if current.strip() and len(current) - len(current.lstrip()) <= steps_indent:
                break
            item = re.match(r"^(\s*)-\s+(.*)$", current)
            if not item or len(item.group(1)) <= steps_indent:
                index += 1
                continue
            item_indent = len(item.group(1))
            block = [current]
            index += 1
            while index < len(lines):
                candidate = lines[index]
                indent = len(candidate) - len(candidate.lstrip())
                if candidate.strip() and indent <= steps_indent:
                    break
                if re.match(rf"^\s{{{item_indent}}}-\s+", candidate):
                    break
                block.append(candidate)
                index += 1

            uses = None
            with_values = {}
            with_indent = None
            for offset, block_line in enumerate(block):
                uses_match = re.match(r"^\s*(?:-\s+)?uses:\s*([^\s#]+)", block_line)
                if uses_match:
                    uses = uses_match.group(1).strip("\"'")
                with_match = re.match(r"^(\s*)with:\s*(?:#.*)?$", block_line)
                if with_match:
                    with_indent = len(with_match.group(1))
                    for child in block[offset + 1 :]:
                        if not child.strip():
                            continue
                        child_indent = len(child) - len(child.lstrip())
                        if child_indent <= with_indent:
                            break
                        scalar = re.match(r"^\s*([^:#]+):\s*([^#]+?)\s*$", child)
                        if scalar:
                            with_values[scalar.group(1).strip()] = scalar.group(2).strip(
                                "\"' "
                            )
                    break
            steps.append({"uses": uses, "with": with_values, "text": "\n".join(block)})
    return steps


def dependabot_updates(text: str) -> list[dict]:
    """Parse each Dependabot update as one bound ecosystem/directory/schedule item."""
    lines = text.splitlines()
    updates_index = next(
        (index for index, line in enumerate(lines) if re.match(r"^updates:\s*$", line)),
        None,
    )
    if updates_index is None:
        return []

    entries = []
    index = updates_index + 1
    while index < len(lines):
        item = re.match(r"^(\s*)-\s+package-ecosystem:\s*([^#]+?)\s*$", lines[index])
        if not item:
            index += 1
            continue
        item_indent = len(item.group(1))
        block = [lines[index]]
        index += 1
        while index < len(lines):
            candidate = lines[index]
            if re.match(rf"^\s{{{item_indent}}}-\s+package-ecosystem:", candidate):
                break
            if candidate.strip() and len(candidate) - len(candidate.lstrip()) <= item_indent:
                break
            block.append(candidate)
            index += 1

        root_indent = item_indent + 2
        directory = None
        interval = None
        schedule_indent = None
        for offset, block_line in enumerate(block):
            directory_match = re.match(
                rf"^\s{{{root_indent}}}directory:\s*([^#]+?)\s*$", block_line
            )
            if directory_match:
                directory = directory_match.group(1).strip("\"' ")
            schedule_match = re.match(
                rf"^(\s{{{root_indent}}})schedule:\s*(?:#.*)?$", block_line
            )
            if schedule_match:
                schedule_indent = len(schedule_match.group(1))
                for child in block[offset + 1 :]:
                    if not child.strip():
                        continue
                    child_indent = len(child) - len(child.lstrip())
                    if child_indent <= schedule_indent:
                        break
                    interval_match = re.match(r"^\s*interval:\s*([^#]+?)\s*$", child)
                    if interval_match:
                        interval = interval_match.group(1).strip("\"' ")
                        break
        entries.append(
            {
                "ecosystem": item.group(2).strip("\"' "),
                "directory": directory,
                "interval": interval,
                "text": "\n".join(block),
            }
        )
    return entries


def release_contract(text: str) -> dict[str, int]:
    section = markdown_section(text, "Packaging, independent verification and immutability")
    patterns = {
        "draft": r"upload[^.\n]*\bdraft release\b",
        "independent": r"independent verifier[^.\n]*downloads?[^.\n]*from GitHub",
        "publication": r"authorised maintainer[^.\n]*publish(?:es)? the release",
        "immutability": r"GitHub reports\s+`?immutable=true`?",
    }
    positions = {}
    for name, pattern in patterns.items():
        match = re.search(pattern, section, re.IGNORECASE)
        if not match:
            raise AssertionError(f"release process is missing the {name} stage")
        positions[name] = match.start()
    return positions


def assert_checkout_credentials_disabled(text: str) -> None:
    checkout_steps = [step for step in workflow_steps(text) if step["uses"] == CHECKOUT]
    if len(checkout_steps) != 1:
        raise AssertionError(f"expected one exact checkout step, found {len(checkout_steps)}")
    if checkout_steps[0]["with"].get("persist-credentials", "").lower() != "false":
        raise AssertionError("checkout's own with mapping must disable persisted credentials")


def assert_github_actions_dependabot(text: str) -> None:
    matches = [
        entry
        for entry in dependabot_updates(text)
        if entry["ecosystem"] == "github-actions"
        and entry["directory"] == "/"
        and entry["interval"] in {"daily", "weekly"}
    ]
    if len(matches) != 1:
        raise AssertionError(
            "one GitHub Actions update entry must bind root directory to an at-least-weekly schedule"
        )


def assert_private_reporting_contract(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_section(text, "Reporting a vulnerability")
    )
    required = (
        r"https://github\.com/ryanduguid/Ozzit/security/advisories/new",
        r"availability[^.\n]*depends on[^.\n]*live GitHub[^.\n]*setting",
        r"do not[^.\n]*(?:public issue|issue)[^.\n]*discussion[^.\n]*pull request[^.\n]*commit",
    )
    for pattern in required:
        if not re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(f"reporting section is missing contract: {pattern}")


def assert_release_licence_limits(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_section(text, "Candidate and source ownership")
    )
    for sentence in re.split(r"(?<=[.!?])\s+", section):
        all_mit_claim = re.search(
            r"(?:entire|whole|all)[^.]*?(?:workbook|source)[^.]*?(?:is|are|as) MIT",
            sentence,
            re.IGNORECASE,
        )
        prohibition = re.match(
            r"(?:do not|never|must not|cannot)\b", sentence, re.IGNORECASE
        )
        if all_mit_claim and not prohibition:
            raise AssertionError(
                "the derived workbook/source bundle must not be labelled MIT"
            )
    required = (
        r"MIT[^.\n]*covers[^.\n]*`tools/`[^.\n]*`.github/`[^.\n]*Markdown[^.\n]*`assets/`",
        r"MIT[^.\n]*does not cover[^.\n]*`ozzit\.xlsx`[^.\n]*`src/`[^.\n]*`functions\.csv`",
        r"no open-source licence[^.\n]*located[^.\n]*author retains",
        r"source archive[^.\n]*`ATTRIBUTION\.md`[^.\n]*`LICENCE`",
    )
    for pattern in required:
        if not re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(f"licence boundary is missing: {pattern}")


def assert_checksum_contract(text: str) -> None:
    section = markdown_section(text, "Future release bundle")
    if re.search(
        r"`?SHA256SUMS`?[^.\n]*(?:hash(?:es)?|include(?:s)?)[^.\n]*(?:itself|`SHA256SUMS`)",
        section,
        re.IGNORECASE,
    ) and not re.search(
        r"every payload asset except `SHA256SUMS` itself", section, re.IGNORECASE
    ):
        raise AssertionError("SHA256SUMS cannot stably hash itself")
    required = (
        r"every payload asset except `SHA256SUMS` itself",
        r"independent verifier[^.\n]*record[^.\n]*compare[^.\n]*GitHub(?:'s)? API `?digest`?[^.\n]*checksum file",
    )
    for pattern in required:
        if not re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(f"checksum contract is missing: {pattern}")


def assert_no_overwrite_contract(text: str) -> None:
    section = markdown_section(text, "Scope and history")
    if not re.search(r"new version and tag", section, re.IGNORECASE):
        raise AssertionError("corrections require a new version and tag")
    if not re.search(
        r"never[^.\n]*(?:move|retarget)[^.\n]*tag[^.\n]*(?:overwrite|replace)[^.\n]*asset",
        section,
        re.IGNORECASE,
    ):
        raise AssertionError("scope must prohibit moving tags and overwriting assets")


class RepositoryPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.security = read_optional(SECURITY)
        cls.releasing = read_optional(RELEASING)
        cls.workflow = VERIFY_WORKFLOW.read_text(encoding="utf-8")
        cls.dependabot = DEPENDABOT.read_text(encoding="utf-8")

    def test_security_supported_versions_section_is_explicit(self):
        section = markdown_section(self.security, "Supported versions")
        self.assertRegex(section, r"(?i)only the latest published GitHub release")
        self.assertRegex(section, r"(?i)older releases[^.]*not supported")
        self.assertRegex(section, r"(?i)unreleased branches[^.]*not supported")
        self.assertRegex(section, r"(?i)newer release[^.]*supersedes")
        self.assertNotRegex(section, r"(?i)latest version on the default branch")

    def test_security_private_reporting_section_is_conditional_and_private(self):
        assert_private_reporting_contract(self.security)

    def test_security_reporting_section_requires_synthetic_non_sensitive_evidence(self):
        section = normalise_markdown_prose(
            markdown_section(self.security, "Reporting a vulnerability")
        )
        self.assertRegex(section, r"(?i)minimal synthetic reproduction")
        for phrase in (
            "client workbooks",
            "real client or production data",
            "credentials",
            "access tokens",
            "private keys",
            "session material",
            "private URLs",
            ".env files",
        ):
            self.assertIn(phrase.lower(), section.lower())
        for phrase in (
            "affected release",
            "impact",
            "reproduction steps",
            "suggested mitigation",
        ):
            self.assertIn(phrase, section.lower())
        self.assertRegex(
            section,
            r"(?i)public `ozzit\.xlsx`[^.]*not[^.]*client workbook",
        )

    def test_release_ownership_section_binds_workbook_source_afe_and_index(self):
        section = normalise_markdown_prose(
            markdown_section(self.releasing, "Candidate and source ownership")
        )
        self.assertRegex(section, r"(?i)`ozzit\.xlsx`[^.]*shipped authority")
        self.assertRegex(
            section,
            r"(?i)`src/\*\.txt`[^.]*(?:Advanced Formula Environment \(AFE\)|AFE) store[^.]*`functions\.csv`[^.]*bound publication views",
        )
        self.assertRegex(
            section,
            r"(?i)do not approve[^.]*`src/`[^.]*`functions\.csv`[^.]*in isolation",
        )
        self.assertRegex(
            section,
            r"(?i)consistency checks[^.]*do not prove[^.]*byte-for-byte regeneration",
        )

    def test_release_policy_preserves_reproducibility_and_licence_limits(self):
        section = markdown_section(self.releasing, "Candidate and source ownership")
        self.assertRegex(section, r"(?i)v3\.0\.0[^.]*tracked builder")
        self.assertRegex(section, r"(?i)v3\.1\.0[^.]*one-off[^.]*Excel")
        self.assertRegex(section, r"(?i)`ATTRIBUTION\.md`[^.]*`CHANGELOG\.md`")
        assert_release_licence_limits(self.releasing)

    def test_release_bundle_section_names_every_required_payload(self):
        section = markdown_section(self.releasing, "Future release bundle")
        required = (
            "`ozzit.xlsx`",
            "`Ozzit-<version>-source.zip`",
            "`Ozzit-<version>-verification.txt`",
            "`release-manifest.json`",
            "`SHA256SUMS`",
        )
        for name in required:
            self.assertIn(name, section)
        self.assertRegex(section, r"(?i)genuine SPDX or CycloneDX[^.]*only when applicable")
        self.assertRegex(
            section,
            r"(?i)standalone workbook[^.]*archived workbook[^.]*tagged workbook[^.]*same SHA-256",
        )
        assert_checksum_contract(self.releasing)

    def test_release_verification_section_contains_all_nine_exact_commands(self):
        section = markdown_section(self.releasing, "Verification gates")
        commands = (
            "python tools/verify_workbook.py ozzit.xlsx",
            "python tools/verify_sources.py ozzit.xlsx src",
            "python tools/verify_signatures.py src",
            "python tools/verify_previous_names.py functions.csv",
            "python tools/verify_index.py ozzit.xlsx src functions.csv",
            "python tools/verify_afe.py ozzit.xlsx src",
            "python -m unittest discover -s tools/tests -v",
            "powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\\tools\\excel_selftest.ps1 -Path .\\ozzit.xlsx",
            "python tools/verify_cache.py ozzit.xlsx",
        )
        for command in commands:
            self.assertEqual(section.count(command), 1, command)
        self.assertRegex(section, r"1,129 formulas[^.]*zero error cells")
        self.assertRegex(section, r"259 assertions[^.]*zero failures")
        self.assertRegex(section, r"(?i)cached-value[^.]*actual comparison count")
        self.assertRegex(section, r"(?i)Excel version and build")
        self.assertRegex(section, r"(?i)hash[^.]*before and after[^.]*byte-identical")

    def test_release_approval_requires_signed_annotated_exact_commit_tag(self):
        section = markdown_section(self.releasing, "Approval and tag")
        self.assertRegex(
            section,
            r"(?i)human maintainer[^.]*approves[^.]*exact candidate commit[^.]*release version[^.]*nine gates",
        )
        self.assertRegex(
            section,
            r"(?i)annotated, cryptographically signed tag[^.]*exact commit",
        )
        self.assertIn("git verify-tag", section)
        self.assertRegex(section, r"(?i)tag object SHA[^.]*peeled commit SHA[^.]*verification result")
        self.assertRegex(section, r"(?i)`gh release create`[^.]*lightweight tag")
        self.assertRegex(
            section,
            r"(?i)tag signing[^.]*remote tag publication[^.]*release publication[^.]*separate authorised actions",
        )

    def test_release_orders_draft_download_verification_before_immutability(self):
        positions = release_contract(self.releasing)
        self.assertLess(positions["draft"], positions["independent"])
        self.assertLess(positions["independent"], positions["publication"])
        self.assertLess(positions["publication"], positions["immutability"])

    def test_release_history_is_prospective_and_forbids_overwrite(self):
        section = markdown_section(self.releasing, "Scope and history")
        self.assertRegex(section, r"(?i)future releases[^.]*after[^.]*merged")
        self.assertRegex(section, r"(?i)does not alter or certify historical releases")
        assert_no_overwrite_contract(self.releasing)

    def test_checkout_step_disables_persisted_credentials(self):
        assert_checkout_credentials_disabled(self.workflow)

    def test_github_actions_dependabot_entry_remains_at_least_weekly(self):
        assert_github_actions_dependabot(self.dependabot)


class RepositoryPolicyMutationTests(unittest.TestCase):
    def test_reporting_contract_accepts_semantic_markdown_line_wraps(self):
        policy = """# Security
## Reporting a vulnerability
Use https://github.com/ryanduguid/Ozzit/security/advisories/new. The form's
availability depends on the live GitHub
setting.
Do not disclose this in a public issue, discussion, pull request or commit.
"""
        assert_private_reporting_contract(policy)

    def test_licence_guard_accepts_an_explicit_all_mit_prohibition(self):
        policy = """# Release
## Candidate and source ownership
MIT covers `tools/`, `.github/`, Markdown and `assets/`.
MIT does not cover `ozzit.xlsx`, `src/` or `functions.csv`.
No open-source licence was located and the author retains the rights.
The source archive includes `ATTRIBUTION.md` and `LICENCE`.
Do not label the whole workbook and source bundle as MIT.
"""
        assert_release_licence_limits(policy)

    def test_checkout_rejects_credentials_setting_on_a_different_step(self):
        mutations = {
            "different step": f"""jobs:
  verify:
    steps:
      - uses: {CHECKOUT}
      - uses: actions/setup-python@abc
        with:
          persist-credentials: false
""",
            "comment": f"""jobs:
  verify:
    steps:
      # persist-credentials: false
      - uses: {CHECKOUT}
""",
            "job environment": f"""jobs:
  verify:
    env:
      persist-credentials: false
    steps:
      - uses: {CHECKOUT}
""",
        }
        for name, workflow in mutations.items():
            with self.subTest(name=name), self.assertRaises(AssertionError):
                assert_checkout_credentials_disabled(workflow)

    def test_dependabot_rejects_monthly_actions_with_unrelated_weekly_entry(self):
        config = """updates:
  - package-ecosystem: github-actions
    directory: /
    schedule:
      interval: monthly
  - package-ecosystem: pip
    directory: /
    schedule:
      interval: weekly
"""
        with self.assertRaises(AssertionError):
            assert_github_actions_dependabot(config)

    def test_dependabot_rejects_directory_and_cadence_split_across_entries(self):
        config = """updates:
  - package-ecosystem: github-actions
    directory: /
    schedule:
      interval: monthly
  - package-ecosystem: github-actions
    directory: /ci
    schedule:
      interval: weekly
"""
        with self.assertRaises(AssertionError):
            assert_github_actions_dependabot(config)

    def test_reporting_rejects_private_url_outside_its_own_section(self):
        policy = """# Security
## Reporting a vulnerability
The form availability depends on the live GitHub setting.
Do not use a public issue, discussion, pull request or commit.

## Links
https://github.com/ryanduguid/Ozzit/security/advisories/new
"""
        with self.assertRaises(AssertionError):
            assert_private_reporting_contract(policy)

    def test_licence_rejects_mit_claim_for_the_whole_derived_bundle(self):
        policy = """# Release
## Candidate and source ownership
MIT covers `tools/`, `.github/`, Markdown and `assets/`.
MIT does not cover `ozzit.xlsx`, `src/` or `functions.csv`.
No open-source licence was located and the author retains the rights.
The source archive includes `ATTRIBUTION.md` and `LICENCE`.
The entire workbook and source bundle is MIT.
"""
        with self.assertRaises(AssertionError):
            assert_release_licence_limits(policy)

    def test_release_order_rejects_immutability_before_download_verification(self):
        policy = """# Release
## Packaging, independent verification and immutability
Upload the assets to a draft release.
GitHub reports `immutable=true`.
The independent verifier downloads the assets from GitHub.
An authorised maintainer publishes the release.
"""
        positions = release_contract(policy)
        self.assertFalse(
            positions["draft"]
            < positions["independent"]
            < positions["publication"]
            < positions["immutability"]
        )

    def test_checksum_rejects_a_self_hash_cycle(self):
        policy = """# Release
## Future release bundle
`SHA256SUMS` hashes every payload asset including `SHA256SUMS` itself.
The independent verifier records and compares GitHub's API `digest` for the checksum file.
"""
        with self.assertRaises(AssertionError):
            assert_checksum_contract(policy)

    def test_history_rejects_new_version_wording_without_no_overwrite_rule(self):
        policy = """# Release
## Scope and history
Correct an error under a new version and tag.
"""
        with self.assertRaises(AssertionError):
            assert_no_overwrite_contract(policy)


if __name__ == "__main__":
    unittest.main()
