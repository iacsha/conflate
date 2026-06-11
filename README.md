# Conflate

Fuzzy-match deduplication and master mapping for Excel and CSV data.
No Python required to run.

## What it does
- **Dedupe Mode** — finds near-duplicate records within a single spreadsheet
- **Master Mode** — maps your data to a trusted master list (many-to-many column mappings)
- Handles typos, abbreviations, and word-order differences
- **Two engines** — RapidFuzz (best under ~10,000 rows) or TF-IDF (best above that)
- **Human-in-the-loop review** — keyboard-driven (`←/→` retain, `Space` skip, `F` flag, `Ctrl+Z` undo)
- **Canonical registry** — reuse earlier decisions and retro-update superseded ones
- **Structured-code matching** — tag code columns so part/SKU numbers must match exactly
- **Resumable** — progress is saved automatically; close and pick up where you left off
- Exports: a decisions workbook (Decisions + Flagged + Clusters sheets), an optional
  write-back to your source file, and optional SQL `UPDATE` statements — all with a full audit trail

## Download
Grab the latest build from the [Releases page](https://github.com/iacsha/conflate/releases) —
unzip and double-click `Conflate.exe`. No install needed (Windows).

<!-- Screenshots: add a few PNGs of the setup and review screens here once available. -->

## Building from source
Requirements: Python 3.10+, Windows. Run from a plain Command Prompt (not Anaconda Prompt).

    git clone https://github.com/iacsha/conflate
    cd conflate
    build.bat

The script creates a clean virtual environment, installs pinned dependencies
(scikit-learn 1.5.2), and packages a standalone `Conflate.exe` via PyInstaller.

## Tech stack
Python · customtkinter · pandas · RapidFuzz · scikit-learn · PyInstaller

## License
MIT — see [LICENSE](LICENSE).
