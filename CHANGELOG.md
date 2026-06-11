# Changelog

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
