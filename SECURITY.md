# Security policy

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
