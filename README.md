# Conflate

Fuzzy-match deduplication and master mapping for Excel and CSV data.
No Python required to run.

## What it does
- Finds near-duplicate records within a spreadsheet (Dedupe Mode)
- Maps your data to a trusted master list (Master Mode)
- Handles typos, abbreviations, word-order differences
- Human-in-the-loop review with keyboard shortcuts
- Exports a clean decisions file with full audit trail

## Download
[Conflate_v1.zip](releases) — unzip and double-click. No install needed.

## Screenshots
(add a few screenshots here)

## Building from source
Requirements: Python 3.10+, Windows

    git clone https://github.com/yourname/conflate
    cd conflate
    build.bat   # run from plain CMD, not Anaconda Prompt

## Tech stack
Python · customtkinter · pandas · RapidFuzz · scikit-learn · PyInstaller

## License
MIT
