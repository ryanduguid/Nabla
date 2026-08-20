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


def strip_html_comments(text: str) -> str:
    return re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)


def strip_fenced_blocks(text: str) -> str:
    """Remove fenced examples so they cannot satisfy prose requirements."""
    kept = []
    fence_character = None
    fence_length = 0
    for line in text.splitlines():
        fence = re.match(r"^\s*(`{3,}|~{3,})(?:[^`]*)$", line)
        if fence_character is None:
            if fence:
                marker = fence.group(1)
                fence_character = marker[0]
                fence_length = len(marker)
                continue
            kept.append(line)
            continue
        if re.match(
            rf"^\s*{re.escape(fence_character)}{{{fence_length},}}\s*$", line
        ):
            fence_character = None
            fence_length = 0
    return "\n".join(kept)


def replace_fixture_once(text: str, old: str, new: str) -> str:
    if text.count(old) != 1:
        raise AssertionError(
            f"expected one fixture occurrence of {old!r}, found {text.count(old)}"
        )
    return text.replace(old, new, 1)


def wrap_fixture_section(text: str, heading: str, opener: str, closer: str) -> str:
    pattern = re.compile(
        rf"(?ms)^(## {re.escape(heading)}\s*\n)(.*?)(?=^## |\Z)"
    )
    result, count = pattern.subn(
        lambda match: f"{match.group(1)}{opener}\n{match.group(2).rstrip()}\n{closer}\n\n",
        text,
        count=1,
    )
    if count != 1:
        raise AssertionError(f"could not wrap fixture section {heading!r}")
    return result


def markdown_section(text: str, heading: str) -> str:
    """Return one Markdown section through the next peer or parent heading."""
    lines = strip_html_comments(text).splitlines()
    headings = []
    fence_character = None
    fence_length = 0
    for index, line in enumerate(lines):
        fence = re.match(r"^\s*(`{3,}|~{3,})(?:[^`]*)$", line)
        if fence_character is None and fence:
            marker = fence.group(1)
            fence_character = marker[0]
            fence_length = len(marker)
            continue
        if fence_character is not None:
            if re.match(
                rf"^\s*{re.escape(fence_character)}{{{fence_length},}}\s*$", line
            ):
                fence_character = None
                fence_length = 0
            continue
        match = re.match(r"^(#{1,6})\s+(.+?)\s*$", line)
        if match:
            headings.append((index, len(match.group(1)), match.group(2)))

    matches = []
    for index, level, title in headings:
        if title == heading:
            matches.append((index, level))
    if len(matches) != 1:
        raise AssertionError(
            f"expected exactly one {heading!r} heading, found {len(matches)}"
        )

    start, level = matches[0]
    end = len(lines)
    for index, candidate_level, _ in headings:
        if index > start and candidate_level <= level:
            end = index
            break
    return "\n".join(lines[start + 1 : end]).strip()


def markdown_prose_section(text: str, heading: str) -> str:
    return strip_fenced_blocks(markdown_section(text, heading))


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
    section = normalise_markdown_prose(
        markdown_prose_section(
            text, "Packaging, independent verification and immutability"
        )
    )
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

    before_download_verification = section[: positions["independent"]]
    premature_claims = (
        r"\bdraft(?: upload| release)?\s+(?:is|becomes|remains)\s+(?:an?\s+)?immutable\b",
        r"\bdraft(?: upload| release)?[^.]*\b(?:proves|constitutes|confirms)\s+(?:final\s+)?approval\b",
        r"\bdraft(?: upload| release)?[^.]*\b(?:is|becomes)\s+(?:proof of\s+)?approval\b",
    )
    for pattern in premature_claims:
        if re.search(pattern, before_download_verification, re.IGNORECASE):
            raise AssertionError(
                "draft release claims approval or immutability before downloaded-asset verification"
            )
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


def assert_supported_versions_contract(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Supported versions")
    )
    required = (
        r"only the latest published GitHub release",
        r"older releases[^.]*not supported",
        r"unreleased branches[^.]*not supported",
        r"newer release[^.]*supersedes",
    )
    for pattern in required:
        if not re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(f"supported-version contract is missing: {pattern}")

    for sentence in re.split(r"(?<=[.!?])\s+", section):
        if not re.search(
            r"default branch|older releases?|unreleased branches?",
            sentence,
            re.IGNORECASE,
        ) or not re.search(r"support", sentence, re.IGNORECASE):
            continue
        denies_support = re.search(
            r"(?:not|no longer) supported|unsupported|does not support",
            sentence,
            re.IGNORECASE,
        )
        if not denies_support:
            raise AssertionError(
                f"unsupported release line receives support: {sentence}"
            )


def assert_private_reporting_contract(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Reporting a vulnerability")
    )
    required = (
        r"https://github\.com/ryanduguid/Ozzit/security/advisories/new",
        r"availability\s+depends\s+on[^.]*live GitHub[^.]*setting",
        r"do not[^.\n]*(?:public issue|issue)[^.\n]*discussion[^.\n]*pull request[^.\n]*commit",
    )
    for pattern in required:
        if not re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(f"reporting section is missing contract: {pattern}")

    contradictions = (
        r"availability\s+(?:never|does not|doesn't)\s+depend",
        r"availability[^.]*independent of[^.]*setting",
        r"availability[^.]*regardless of[^.]*setting",
    )
    for pattern in contradictions:
        if re.search(pattern, section, re.IGNORECASE):
            raise AssertionError("private-reporting availability was made unconditional")

    response_sla = re.search(
        r"\b(?:acknowledg\w*|respond\w*|fix\w*|remediat\w*|resolv\w*)[^.]*"
        r"\bwithin\s+(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+"
        r"(?:business\s+)?(?:hours?|days?|weeks?)\b",
        section,
        re.IGNORECASE,
    )
    if response_sla:
        raise AssertionError("security policy must not invent a response or remediation SLA")


def assert_sensitive_reporting_contract(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Reporting a vulnerability")
    )
    if not re.search(r"minimal synthetic reproduction", section, re.IGNORECASE):
        raise AssertionError("reporting requires a minimal synthetic reproduction")

    sensitive_phrases = (
        "client workbooks",
        "real client or production data",
        "credentials",
        "access tokens",
        "private keys",
        "session material",
        "private URLs",
        ".env files",
    )
    prohibition = next(
        (
            sentence
            for sentence in re.split(r"(?<=[.!?])\s+", section)
            if re.search(r"(?:do not|never|must not)\s+upload", sentence, re.IGNORECASE)
        ),
        None,
    )
    if prohibition is None:
        raise AssertionError("sensitive evidence is not bound to an upload prohibition")
    for phrase in sensitive_phrases:
        if phrase.lower() not in prohibition.lower():
            raise AssertionError(
                f"sensitive evidence is outside the upload prohibition: {phrase}"
            )

    permission_patterns = (
        r"(?:^|[.!?]\s+)upload\s+client workbooks?",
        r"\b(?:may|can|should|are allowed to|are permitted to)\s+upload\s+"
        r"(?:client workbooks?|real client|production data|credentials|access tokens?)",
    )
    for pattern in permission_patterns:
        if re.search(pattern, section, re.IGNORECASE):
            raise AssertionError("security policy permits sensitive evidence uploads")


def assert_release_licence_limits(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Candidate and source ownership")
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
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Future release bundle")
    )
    positive_self_hashes = (
        r"`?SHA256SUMS`?\s+(?:must\s+|will\s+|shall\s+|does\s+|can\s+|may\s+)?hash(?:es)?[^.]*\bitself\b",
        r"\binclude(?:s|d|ing)?\s+`?SHA256SUMS`?\s+itself\b",
        r"`?SHA256SUMS`?[^.]*\bself-hash(?:es|ing)?\b",
    )
    for pattern in positive_self_hashes:
        if re.search(pattern, section, re.IGNORECASE):
            raise AssertionError("SHA256SUMS cannot stably hash itself")
    required = (
        r"every payload asset except `SHA256SUMS` itself",
        r"independent verifier[^.\n]*record[^.\n]*compare[^.\n]*GitHub(?:'s)? API `?digest`?[^.\n]*checksum file",
    )
    for pattern in required:
        if not re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(f"checksum contract is missing: {pattern}")


def assert_no_overwrite_contract(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Scope and history")
    )
    if not re.search(r"new version and tag", section, re.IGNORECASE):
        raise AssertionError("corrections require a new version and tag")
    if not re.search(
        r"never[^.\n]*(?:move|retarget)[^.\n]*tag[^.\n]*(?:overwrite|replace)[^.\n]*asset",
        section,
        re.IGNORECASE,
    ):
        raise AssertionError("scope must prohibit moving tags and overwriting assets")
    permissions = (
        r"\bverified assets?\s+(?:may|can|should|must|will|are allowed to|are permitted to)\s+(?:be\s+)?(?:overwritten|replaced)\b",
        r"\b(?:may|can|should|must|are allowed to|are permitted to)\s+(?:overwrite|replace)\s+(?:a\s+)?verified assets?\b",
        r"\bpermission to\s+(?:overwrite|replace)\s+(?:a\s+)?verified assets?\b",
    )
    for pattern in permissions:
        if re.search(pattern, section, re.IGNORECASE):
            raise AssertionError("verified assets are permitted to be overwritten")


def assert_reproducibility_limits(text: str) -> None:
    section = normalise_markdown_prose(
        markdown_prose_section(text, "Candidate and source ownership")
    )
    false_current_claims = (
        r"\btracked builder[^.]*\breproduc\w*[^.]*\bcurrent v?3\.1\.0\b[^.]*\bbyte-for-byte\b",
        r"\bcurrent v?3\.1\.0\b[^.]*\bbyte-for-byte reproduc\w*\b",
        r"\bcurrent workbook[^.]*\bbyte-for-byte reproduc\w*\b",
    )
    for pattern in false_current_claims:
        if re.search(pattern, section, re.IGNORECASE):
            raise AssertionError(
                "release policy falsely claims current-workbook byte reproducibility"
            )


class RepositoryPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.security = read_optional(SECURITY)
        cls.releasing = read_optional(RELEASING)
        cls.workflow = VERIFY_WORKFLOW.read_text(encoding="utf-8")
        cls.dependabot = DEPENDABOT.read_text(encoding="utf-8")

    def test_security_supported_versions_section_is_explicit(self):
        section = markdown_prose_section(self.security, "Supported versions")
        assert_supported_versions_contract(self.security)
        self.assertNotRegex(section, r"(?i)latest version on the default branch")

    def test_security_private_reporting_section_is_conditional_and_private(self):
        assert_private_reporting_contract(self.security)

    def test_security_reporting_section_requires_synthetic_non_sensitive_evidence(self):
        section = normalise_markdown_prose(
            markdown_prose_section(self.security, "Reporting a vulnerability")
        )
        assert_sensitive_reporting_contract(self.security)
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
            markdown_prose_section(self.releasing, "Candidate and source ownership")
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
        section = normalise_markdown_prose(
            markdown_prose_section(self.releasing, "Candidate and source ownership")
        )
        self.assertRegex(section, r"(?i)v3\.0\.0[^.]*tracked builder")
        self.assertRegex(section, r"(?i)v3\.1\.0[^.]*one-off[^.]*Excel")
        self.assertRegex(section, r"(?i)`ATTRIBUTION\.md`[^.]*`CHANGELOG\.md`")
        assert_reproducibility_limits(self.releasing)
        assert_release_licence_limits(self.releasing)

    def test_release_bundle_section_names_every_required_payload(self):
        section = markdown_prose_section(self.releasing, "Future release bundle")
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
        section = markdown_prose_section(self.releasing, "Approval and tag")
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
        section = markdown_prose_section(self.releasing, "Scope and history")
        self.assertRegex(section, r"(?i)future releases[^.]*after[^.]*merged")
        self.assertRegex(section, r"(?i)does not alter or certify historical releases")
        assert_no_overwrite_contract(self.releasing)

    def test_checkout_step_disables_persisted_credentials(self):
        assert_checkout_credentials_disabled(self.workflow)

    def test_github_actions_dependabot_entry_remains_at_least_weekly(self):
        assert_github_actions_dependabot(self.dependabot)


class RepositoryPolicyMutationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.security = SECURITY.read_text(encoding="utf-8")
        cls.releasing = RELEASING.read_text(encoding="utf-8")
        cls.workflow = VERIFY_WORKFLOW.read_text(encoding="utf-8")
        cls.dependabot = DEPENDABOT.read_text(encoding="utf-8")

    def assert_policy_method_rejects(
        self,
        method: str,
        *,
        security: str | None = None,
        releasing: str | None = None,
    ) -> None:
        case = RepositoryPolicyTests(method)
        case.security = self.security if security is None else security
        case.releasing = self.releasing if releasing is None else releasing
        case.workflow = self.workflow
        case.dependabot = self.dependabot
        with self.assertRaises(AssertionError):
            getattr(case, method)()

    def test_security_contracts_ignore_comments_and_fenced_examples(self):
        for label, opener, closer in (
            ("HTML comment", "<!--", "-->"),
            ("fenced example", "```text", "```"),
        ):
            hidden = wrap_fixture_section(
                self.security, "Supported versions", opener, closer
            )
            hidden = wrap_fixture_section(
                hidden, "Reporting a vulnerability", opener, closer
            )
            with self.subTest(label=label):
                for method in (
                    "test_security_supported_versions_section_is_explicit",
                    "test_security_private_reporting_section_is_conditional_and_private",
                    "test_security_reporting_section_requires_synthetic_non_sensitive_evidence",
                ):
                    self.assert_policy_method_rejects(method, security=hidden)

    def test_reporting_rejects_negated_pvr_conditionality(self):
        mutated = replace_fixture_once(
            self.security,
            "availability depends on the live GitHub",
            "availability never depends on the live GitHub",
        )
        self.assert_policy_method_rejects(
            "test_security_private_reporting_section_is_conditional_and_private",
            security=mutated,
        )

    def test_reporting_rejects_permission_to_upload_client_workbooks(self):
        mutated = replace_fixture_once(
            self.security,
            "Do not upload client workbooks",
            "Upload client workbooks",
        )
        self.assert_policy_method_rejects(
            "test_security_reporting_section_requires_synthetic_non_sensitive_evidence",
            security=mutated,
        )

    def test_reporting_rejects_an_unauthorised_acknowledgement_sla(self):
        mutated = replace_fixture_once(
            self.security,
            "\n## What this library does and does not do",
            "\nA valid report will be acknowledged within seven days.\n\n"
            "## What this library does and does not do",
        )
        self.assert_policy_method_rejects(
            "test_security_private_reporting_section_is_conditional_and_private",
            security=mutated,
        )

    def test_supported_versions_rejects_an_extra_default_branch_line(self):
        mutated = replace_fixture_once(
            self.security,
            "\n## Reporting a vulnerability",
            "\nThe current default branch is also a supported release line.\n\n"
            "## Reporting a vulnerability",
        )
        self.assert_policy_method_rejects(
            "test_security_supported_versions_section_is_explicit",
            security=mutated,
        )

    def test_ownership_rejects_false_current_v31_byte_reproducibility(self):
        mutated = replace_fixture_once(
            self.releasing,
            "\n## Approval and tag",
            "\nThe tracked builder reproduces the current v3.1.0 workbook byte-for-byte.\n\n"
            "## Approval and tag",
        )
        self.assert_policy_method_rejects(
            "test_release_policy_preserves_reproducibility_and_licence_limits",
            releasing=mutated,
        )

    def test_packaging_rejects_draft_immutability_and_approval_claims(self):
        mutated = replace_fixture_once(
            self.releasing,
            "3. Upload the assets to a draft release. A draft upload is not approval, publication or evidence of immutability.",
            "3. Upload the assets to a draft release. A draft upload is not approval, publication or evidence of immutability. The draft upload is immutable and proves approval.",
        )
        self.assert_policy_method_rejects(
            "test_release_orders_draft_download_verification_before_immutability",
            releasing=mutated,
        )

    def test_history_rejects_permission_to_overwrite_verified_assets(self):
        mutated = replace_fixture_once(
            self.releasing,
            "Correct a published or verified candidate under a new version and tag. Never move or retarget an existing tag or overwrite or replace a verified asset.",
            "Correct a published or verified candidate under a new version and tag. Never move or retarget an existing tag or overwrite or replace a verified asset. Verified assets may be overwritten.",
        )
        self.assert_policy_method_rejects(
            "test_release_history_is_prospective_and_forbids_overwrite",
            releasing=mutated,
        )

    def test_bundle_rejects_a_contradictory_checksum_self_hash(self):
        mutated = replace_fixture_once(
            self.releasing,
            "6. `SHA256SUMS` — SHA-256 for every payload asset except `SHA256SUMS` itself, including the manifest, verification evidence and any genuine SBOM.",
            "6. `SHA256SUMS` — SHA-256 for every payload asset except `SHA256SUMS` itself, including the manifest, verification evidence and any genuine SBOM. `SHA256SUMS` hashes itself.",
        )
        self.assert_policy_method_rejects(
            "test_release_bundle_section_names_every_required_payload",
            releasing=mutated,
        )

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
