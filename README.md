# Conflate

Fuzzy-match deduplication and master mapping for Excel and CSV data.
No Python required to run.

## What it does
- **Dedupe Mode** - finds near-duplicate records within a single spreadsheet
- **Master Mode** - maps your data to a trusted master list (many-to-many column mappings)
- Handles typos, abbreviations, and word-order differences
- **Two engines** - RapidFuzz (best under ~10,000 rows) or TF-IDF (best above that)
- **Human-in-the-loop review** - keyboard-driven (`←/→` retain, `Space` skip, `F` flag, `Ctrl+Z` undo)
- **Canonical registry** - reuse earlier decisions and retro-update superseded ones
- **Structured-code matching** - tag code columns so part/SKU numbers must match exactly
- **Resumable** - progress is saved automatically; close and pick up where you left off
- Exports: a decisions workbook (Decisions + Flagged + Clusters sheets), an optional
  write-back to your source file, and optional SQL `UPDATE` statements - all with a full audit trail

## Download
Grab the latest build from the [Releases page](https://github.com/iacsha/conflate/releases) -
unzip and double-click `Conflate.exe`. No install needed (Windows).

## Quick start
1. **Select Primary Data** - your Excel/CSV file to clean. (Optionally add a **Master List** to map against.)
2. **Check the columns** to match (e.g. `Supplier Name`); set Match Strictness (85% is a good default).
3. **Start Scan**, review the score histogram, then **Proceed to Review**.
4. Decide each pair with the keyboard - `←` retain left, `→` retain right, `Space` skip, `F` flag, `Ctrl+Z` undo.
5. On finish, pick your outputs: a **Decisions** workbook, a **write-back** to your source file, and/or **SQL UPDATE** statements.

## Example: merging duplicate delivery customers

A pizza shop's order system has the same customers entered several ways over time, with small differences in name and address:

| Cust_ID | Name | Address |
|---------|------|---------|
| C-1001 | Jonathan Meyer | 1428 Elm Street, Apt 3B |
| C-1002 | Jon Meyer | 1428 Elm St #3B |
| C-1003 | Maria Gonzalez | 76 Lakeview Dr |
| C-1004 | Maria Gonzales | 76 Lake View Drive |
| C-1005 | Antonio Russo | 9 Park Pl |

Load the file, check both **Name** and **Address** as search columns, set `Cust_ID` as the unique ID, and set Match Strictness to 85%:

<img width="946" height="1093" alt="Conflate setup screen with the customer file loaded and the Name and Address columns selected" src="https://github.com/user-attachments/assets/d899200f-361d-424d-abb2-abe02dde7f4c" />

After the scan, Conflate shows how the match scores are distributed so you can gauge data quality before reviewing:

<img width="640" height="430" alt="Match Score Distribution histogram for the sample customer scan" src="https://github.com/user-attachments/assets/fc052373-9e9e-4a58-8c42-6a21bce70e24" />

It surfaces the likely same-person duplicates as candidate pairs:

| Score | Item A | Item B |
|-------|--------|--------|
| 95% | Jonathan Meyer \| 1428 Elm Street, Apt 3B | Jon Meyer \| 1428 Elm St #3B |
| 92% | Maria Gonzalez \| 76 Lakeview Dr | Maria Gonzales \| 76 Lake View Drive |

Step through each pair in the review screen and keep the cleaner record as canonical with `←`:

<img width="952" height="1092" alt="Conflate review screen comparing two customer records side by side with decision buttons" src="https://github.com/user-attachments/assets/a0d7983f-a996-4192-b08b-d895557b7e33" />

The exported **Decisions** sheet is a join-ready audit trail, use `Primary_ID` / `Duplicate_ID` to VLOOKUP the merges back into your order data:

| Primary_ID | Duplicate_ID | Final Selection | Action | Score | Note |
|------------|--------------|-----------------|--------|-------|------|
| C-1002 | C-1001 | Jonathan Meyer \| 1428 Elm Street, Apt 3B | Retained Left | 95 | same household |
| C-1004 | C-1003 | Maria Gonzalez \| 76 Lakeview Dr | Retained Left | 92 | |

## Documentation
📖 **[Full User Guide](docs/USER_GUIDE.md)** - setup, every option, review workflow, outputs, troubleshooting, and glossary.

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
MIT - see [LICENSE](LICENSE).
