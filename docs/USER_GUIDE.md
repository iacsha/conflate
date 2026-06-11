# Conflate - User Guide

**v1.0 · Data Deduplication & Master Mapping**
*Clean your data. Trust your results.*

---

## 1. What is Conflate?

Conflate is a desktop application for cleaning and standardising Excel and CSV data. It finds records that refer to the same real-world thing but are spelled differently, and lets you decide how to resolve them - without writing a single formula or macro.

There are two ways to use it:

- **Deduplication Mode** - finds duplicate or near-duplicate records *within a single spreadsheet* and lets you merge or discard them.
- **Master Mapping Mode** - compares your data against a trusted master list (e.g. a canonical supplier list) and maps each of your records to its correct master entry.

Conflate uses **fuzzy matching**, which means it catches variations like `"St. Mary's Hospital"` vs `"Saint Marys Hospital"` vs `"St Marys Hosp"` - things that an exact lookup would miss entirely.

---

## 2. Getting Started

### System Requirements
- Windows 10 or 11 (64-bit)
- Minimum 4 GB RAM recommended for large files (10,000+ rows)

### Installation
Conflate ships as a single `.exe` file. No installer is required.

1. Download `Conflate_v1.0.zip` from the [Releases page](https://github.com/iacsha/conflate/releases).
2. Right-click the zip and select **Extract All**. Choose a permanent location such as `C:\Tools\Conflate`.
3. Double-click `Conflate.exe` to launch the application.

> **Note:** If you received a folder-mode build, do not move `Conflate.exe` out of its folder - all the supporting files must stay together.

### First Launch
When you first run Conflate, Windows may show a **SmartScreen** warning saying the publisher is unknown. This is normal for unsigned applications. Click **More info → Run anyway** to proceed.

A log file called `Conflate.log` is created in the same folder as the exe. It records everything the app does and is useful for troubleshooting. Open it any time with the **📋 View Log** button in the top-right of the setup screen.

---

## 3. The Setup Screen

When Conflate opens you see the **Setup** screen, where you configure your scan before it runs. Work through the steps from top to bottom.

### Step 1 - Load Your Files

- **Primary Data (required)** - Click **1. Select Primary Data** and choose your Excel (`.xlsx`, `.xls`) or CSV file. This is your working data - the file you want to clean.
- **Master List (optional)** - Click **1b. Select Master List (Optional)** to map your data against a trusted reference file. Leave this blank to run in **Deduplication Mode** instead.

> **Tip:** Loading a Master List switches the app into **Master Mapping Mode** and reveals extra options for configuring how columns are matched between the two files.

### Step 2 - Select Columns to Search
Check the column(s) that contain the text you want to match. For example, if your file has a `Supplier Name` column, check it. You can check multiple columns; Conflate combines them into a single search string.

> **Note:** Selecting too many columns can reduce match quality. Focus on the columns that most uniquely identify each record - usually a name or description field.

**The "Code" tag (structured codes).** Next to each primary column is a small **Code** button. Tag a column as a code when it holds a part number, SKU, or account code whose digits must match exactly. When tagged, Conflate suppresses any match where the digit sequences differ - so `10-42` and `10_42` still match (same digits, different separators), but `10-42` and `10-43` do **not**.

### Step 2b - Select Unique ID Columns
These dropdowns specify which column holds the unique identifier for each record (such as a `Supplier ID` or `Item Code`). When set, both IDs appear in every output row so you can use VLOOKUP / XLOOKUP to join your results back to the original data.

Both default to *- none -*. They are optional but strongly recommended if your data has a primary key column.

### Column Mappings (Master Mode only)
When a Master List is loaded, a **Column Mappings** section appears below the ID selectors. Here you define which primary-file column corresponds to which master-file column.

One mapping pair is shown by default. Click **+ Add Pair** to compare multiple fields (e.g. both `Supplier Name` and `City`); each pair contributes to the overall match score. Click the red **✕** to remove a pair.

> **Tip:** Many-to-many mappings are supported. You can map several primary columns to different master columns in any combination.

### Step 3 - Configure Match Settings

| Setting | Description |
|---------|-------------|
| **Match Strictness** | How similar two strings must be before they are flagged as a potential match. **85%** is a good default. Lower it if you are missing obvious matches; raise it to reduce false positives. |
| **Max Matches per Item** | Maximum number of match candidates returned per record. Default **5**. Increase for exploratory work; decrease to keep the review queue manageable. |
| **Test Mode** | Processes only the first 100 rows. Use it to verify your settings before a full scan. |
| **Processing Engine** | **RapidFuzz** - best for files under 10,000 rows. **TF-IDF (Heavy Duty)** - best for larger files; significantly faster at scale. |

---

## 4. Running a Scan

1. Complete all setup steps above.
2. Click **3. Start Scan**. A progress bar and status message appear while the scan runs.
3. When the scan finishes, a **Match Score Distribution** chart appears showing how your results spread across score bands.
4. Review the histogram to gauge data quality - a large spike at 90-100% suggests mostly clear duplicates; an even spread suggests more ambiguous matches.
5. Click **Proceed to Review** to begin making decisions, or **Cancel** to return to setup.

> **Note:** A raw scan results file is saved automatically to your primary file's folder as soon as the scan completes - every candidate pair found, regardless of your review decisions.

To stop a scan in progress, click **Stop Scan**. Any matches found up to that point are kept.

---

## 5. Reviewing Matches

The Review screen presents each candidate pair one at a time. For each pair you see the match score, the two items being compared, and the full row context from both source files.

### The Stats Bar
A live bar at the top shows progress at a glance:
- **Approved** - decisions where you kept or overrode a record
- **Flagged** - pairs set aside for an expert to decide later
- **Skipped** - pairs you passed over without deciding
- **Remaining** - pairs left in the queue
- **Est. Time Left** - a rolling estimate based on your current pace

### The Match Score
The score is the similarity of the two strings as a percentage. The colour signals confidence:

| Score | Meaning |
|-------|---------|
| **95% and above** | Very likely the same record. Green. |
| **85% - 94%** | Probably the same; check context. Orange. |
| **Below 85%** | Lower confidence - review carefully. Red. |

### Making Decisions
Use the buttons or keyboard shortcuts to act on each pair:

| Action | Shortcut | What it does |
|--------|----------|--------------|
| **Retain Left** | `←` or `1` | Keeps Item A (your data) as the final value. |
| **Retain Right** *(Master Mode: "Retain Right (Master)")* | `→` or `2` | Replaces Item A with Item B. In Master Mode, maps your record to the master entry. |
| **Skip** | `Space` or `3` | Moves to the next pair without recording a decision. The pair is not written to the output. |
| **Flag for Review** | `F` or `4` | Marks the pair for expert review. Saved in a separate **Flagged for Review** sheet. |
| **Use Canonical** | `C` or `5` | Applies a previously-approved canonical value (see below). Only available when a suggestion is showing. |
| **Undo Last** | `Ctrl + Z` | Steps back one decision. Press repeatedly to undo a chain. |

### Canonical Suggestions
As you approve decisions, Conflate remembers the value you chose for each item. If a later pair contains a value you have already resolved, a green **canonical banner** appears showing the previously-approved value. Press **Use Canonical [C/5]** to apply it for consistency across the whole dataset.

If a new decision supersedes a value that earlier decisions already used as their final selection, Conflate asks whether to **chain-update** those earlier decisions to the new canonical value - keeping your output internally consistent.

### Decision Notes
Below the buttons is a free-text **Note** field. Type anything relevant before pressing a decision key - e.g. *"Vendor confirmed same supplier"* or *"Different department, do not merge"*. The note is saved alongside the decision.

> **Note:** Hotkeys are paused while your cursor is in the Note field. Click outside it or press **Tab** to restore keyboard shortcuts.

### Full Row Context
The large area at the bottom shows every column from both Item A and Item B, so you can inspect the full record before deciding. In Master Mode with column mappings configured, a **Column-Pair Comparison** section shows each mapped field side by side with its own score.

### Saving Progress and Resuming
Click **Save Progress & Exit** any time to save your place. Next time you load the same primary file and click **Start Scan**, Conflate detects the saved session and asks whether to **resume** or **start fresh**.

> **Note:** Progress is also saved automatically every 10 decisions - you will never lose more than 10 decisions if the app closes unexpectedly.

---

## 6. Understanding the Output

When you finish reviewing (or click **Save Progress & Exit** after a full review), an **Export Options** dialog lets you choose what to generate. All files are written to the same folder as your primary data file.

### Decisions File *(always available)*
Named `Final_Internal_Dedupe_[filename]_[timestamp].xlsx` or `Final_Master_Mapping_[filename]_[timestamp].xlsx` depending on mode. Contains up to three sheets:

- **Decisions** - every pair you resolved (Retain Left/Right or Use Canonical): Primary ID, matched record ID, final selected value, action taken, match score, and any note.
- **Flagged for Review** - every pair you flagged, in the same format, ready for a subject-matter expert.
- **Clusters** - a summary grouping every variant under its canonical value, with a variant count, sorted by cluster size.

> **Tip:** The `Primary_ID` and `Master_ID` / `Duplicate_ID` columns are designed for VLOOKUP. Use them to join the decisions back to your original spreadsheet.

### Write-back to Original *(optional)*
Adds a `Canonical_Name` column to a copy of your source file (`Writeback_[filename]_[timestamp].xlsx`), mapping each original value to the canonical value you chose - handy when you want the cleaned data in the original layout.

### SQL UPDATE Statements *(optional)*
Generates `SQL_Updates_[filename]_[timestamp].sql` - a set of `UPDATE` statements that apply your canonical values to a database table. You configure the table name, the `WHERE` column, and one or more `SET` column → value-source mappings. Only rows whose value actually changed produce a statement.

### Raw Scan File *(automatic)*
`Dedupe_Scan_[filename]_[timestamp].xlsx` or `Master_Scan_[filename]_[timestamp].xlsx`, written automatically before review begins. Contains every candidate pair the engine found, with all original column values from both records prefixed `A_` and `B_`. A complete audit trail independent of your review decisions.

### Log File
`Conflate.log` in the application folder records every scan, every decision (what was decided, the score, any note), and any errors - useful for auditing changes over time or troubleshooting. Open it with **📋 View Log** in the app.

---

## 7. Tips for Best Results

**Choosing the right strictness.** Start at **85%**. Too many false positives → raise to 90%+. Missing expected matches → lower to 75-80%. After a scan, the **Match Score Distribution** histogram is your best guide: a cluster in the 90-100% band means relatively clean data; an even spread means more variation and a lower threshold with more manual review.

**Column selection.** Select only the columns that most uniquely identify a record. A single well-chosen name column usually beats five columns where most are empty or numeric. Numeric columns (IDs, codes, quantities) generally don't help fuzzy matching - tag those as **Code** instead so their digits must match exactly.

**Test Mode.** Always run Test Mode first on a new dataset. It processes 100 rows so you can confirm your settings before a full scan that could take minutes.

**Large files.** For more than 10,000 unique values, switch the engine to **Heavy Duty (TF-IDF)**. On 30,000 rows, TF-IDF typically finishes in under 30 seconds; RapidFuzz on the same data can take 10 minutes or more.

**Reviewing efficiently.** Use the keyboard throughout - `←` / `→` are faster than the mouse. After typing a note, press **Tab** to restore hotkeys. High-confidence matches (95%+) are sorted to the top of the queue, so work the top first and **Skip** lower-confidence pairs for later.

---

## 8. Troubleshooting

| Problem | Solution |
|---------|----------|
| App won't open - SmartScreen warning | Click **More info → Run anyway**. Normal for unsigned apps. |
| No matches found | Lower the Match Strictness slider. Confirm you selected the right columns - name/description fields work best. |
| Too many false positives | Raise Match Strictness to 90%+. Remove columns that aren't meaningful identifiers; tag code columns as **Code**. |
| Scan is very slow | Switch to **Heavy Duty (TF-IDF)**. Enable **Test Mode** first to confirm settings. |
| Item B is blank in review | Make sure you selected at least one column in the correct panel. Dedupe Mode → Primary columns only. Master Mode → select Master columns (right panel) or configure Column Mappings. |
| Output file not found | Output goes to your **primary data file's folder**, not the Conflate app folder. Check the completion dialog for the exact path. |
| Crash or unexpected error | Check `Conflate.log` (or **View Log** in the app) - it captures the full error with context. |

---

## 9. Glossary

| Term | Definition |
|------|-----------|
| **Fuzzy Matching** | Finding strings that are similar but not identical - accounts for typos, abbreviations, and formatting differences. |
| **Match Score** | A percentage of how similar two strings are; 100% means identical. |
| **Deduplication** | Identifying and resolving duplicate records within a single dataset. |
| **Master Mapping** | Linking records in your data to their correct equivalents in a trusted reference list. |
| **RapidFuzz** | The default engine. Best under 10,000 rows. Uses token-sort-ratio scoring, which handles word-order differences well. |
| **TF-IDF** | Alternative engine using character n-gram analysis. Significantly faster on large datasets (10,000+ rows). |
| **Canonical Value** | The single approved value chosen to represent a group of variants. |
| **Cluster** | A canonical value plus all the variants that were merged into it. |
| **Structured Code** | A column tagged **Code** whose digit sequence must match exactly for a pair to be considered a match. |
| **Primary ID** | A column uniquely identifying each primary record, retained in the output for joining back to source data. |
| **Column Mapping** | A pairing of a primary-file column to a master-file column, defining what gets compared. |
| **Raw Scan File** | The auto-generated Excel file of every candidate pair the engine found, written before review begins. |
| **Flag for Review** | A decision that marks a pair as needing expert judgment; saved in a dedicated sheet. |

---

*Conflate v1.0 - Data Deduplication & Master Mapping. For support, open `Conflate.log` via the **View Log** button in the application.*
