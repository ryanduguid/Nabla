# Security policy

## Supported versions

Security fixes are applied to the latest version on the default branch.

## Reporting a vulnerability

Please use this repository's private vulnerability reporting feature. Do not
open a public issue for a suspected security vulnerability. Include a clear
description, reproduction steps, impact, and any suggested mitigation.

A valid report will be acknowledged within seven days, and the fix and
disclosure timeline will be agreed with the reporter.

## What this library does and does not do

Ozzit is an Excel LAMBDA workbook plus Python tools that verify and rebuild it.
The tools read local files and make no network call. They hold no credentials
and have no runtime dependencies outside the Python standard library.

Do not run the tools with elevated privileges. Treat `ozzit.xlsx` and `src/`
as published artefacts from this repository, not as untrusted user input.
