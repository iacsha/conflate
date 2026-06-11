# Changelog

## Unreleased

### Security
- Output spreadsheets are hardened against formula (CSV) injection: any cell value starting with `=`, `+`, `-`, `@`, tab, or carriage return is prefixed with an apostrophe so it is treated as text. Applies to the Decisions, Flagged, Clusters, Raw-scan, and Write-back files.
- SQL UPDATE generation now escapes string values (NUL stripped, single quotes doubled) and validates table/column identifiers against injection. The generated file header documents the assumed ANSI / T-SQL dialect and the MySQL/MariaDB backslash caveat.
- New per-column **Sensitive** tag: values from tagged columns are masked as `[REDACTED]` in `Conflate.log`. Output files still contain the real values.
- The open-folder action no longer builds a shell command string.

## v1.0 (2026-06-11)

First public release of Conflate, a fuzzy-match deduplication and master mapping tool for Excel and CSV data.

### Features
- **Dedupe Mode:** find near-duplicate records within a single spreadsheet.
- **Master Mode:** map records to a trusted master list, with many-to-many column mappings.
- **Two engines:** RapidFuzz (best under ~10,000 rows) and TF-IDF (best for larger files).
- **Keyboard-driven review** with a match-score histogram, live stats bar, and time estimate.
- **Canonical registry:** reuse earlier decisions and chain-update superseded ones.
- **Structured-code columns:** digits must match exactly for fields tagged as codes.
- **Resumable sessions** with automatic progress saving every 10 decisions.
- **Outputs:** Decisions workbook (Decisions + Flagged + Clusters sheets), optional write-back to the source file, and optional SQL UPDATE statements.
- Ships as a standalone Windows executable, no Python required.

### Notes
- The build is unsigned, so Windows SmartScreen may warn on first launch. Click More info, then Run anyway.
