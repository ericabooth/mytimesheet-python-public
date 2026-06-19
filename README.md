# Texas 2036 Time Sheet System
**Author:** Antigravity (AI Pair Programmer)  
**For:** Eric Booth, Senior Research Analyst, Texas 2036  
**Branding Theme:** Texas 2036 Navy (`#1B2D55`) & Orange (`#D44500`)  
**Design Philosophy:** Tufte-Inspired Analytical Visual Design ("Whisper, don't shout")  

This folder contains a self-contained, automated, and manually-extensible timesheet system. It analyzes file modifications across your workspace directories to estimate active work hours and workdays.

---

## 📂 Files in this Folder

1. **`timesheet_generator.py`**: The reusable Python script that scans your workspace directories to collect and aggregate file modifications.
2. **`timesheet.xlsx`**: The generated Excel workbook populated with your actual activity data from **June 1 to June 18, 2026**, driven by live formulas.
3. **`README.md`**: This guide.

---

## ⏱️ Methodology & Logic

### 1. Workday Shift (2:00 AM Boundary)
To handle workdays extending past midnight, any activity that occurs **between 12:00 AM and 2:00 AM** is shifted back by one day and counted towards the *previous* calendar day.
- *Example:* A file modified on **June 9 at 1:30 AM** counts as activity on the **June 8 workday**.
- *Example:* A file modified on **June 9 at 2:30 AM** counts as activity on the **June 9 workday**.

### 2. Folder-Level Activity
The spreadsheet tracks activity at the **folder level** (the immediate directory containing modified files, e.g., `_code`, `03_output/tables`, `PriceTransparency`), rather than individual files. This groups your activities by project task.
- Paths are automatically shortened and beautified (e.g. relative to your Google Drive root or home folder `~/`) to keep the sheet readable.

### 3. Date Span Calculation
For each workday, the system:
1. Identifies the **earliest start time** across all active folders.
2. Identifies the **latest end time** across all active folders.
3. Computes the total **Estimated Work Day Length** as the span between the earliest start time and latest end time (`(Latest - Earliest) * 24` hours).

---

## 📊 Spreadsheet Structure (`timesheet.xlsx`)

The workbook is split into two sheets:

### 1. `Summary` (Daily & Weekly Dashboard)
A clean, professionally styled sheet matching Texas 2036 brand colors.
* **Daily Log (Columns A–G)**:
  * **Date**: Calendar dates from June 1 to June 18. Stored as native datetime objects at midnight to guarantee robust comparisons across spreadsheet engines.
  * **Day**: Automatically populated via `=TEXT(A4, "ddd")`.
  * **Earliest Start Time**: Live formulas querying the `Activity` tab:
    ```excel
    =MINIFS(Activity!$D$2:$D$5000, Activity!$A$2:$A$5000, A4)
    ```
    *Formatted with `yyyy-mm-dd hh:mm AM/PM;;` to hide zero values when no work occurred.*
  * **Latest End Time**:
    ```excel
    =MAXIFS(Activity!$E$2:$E$5000, Activity!$A$2:$A$5000, A4)
    ```
    *Formatted with `yyyy-mm-dd hh:mm AM/PM;;` to hide zero values when no work occurred.*
  * **Work Hours (Span)**: Calculates the hours difference:
    ```excel
    =IF(C4=0, 0, (D4-C4)*24)
    ```
  * **Is Workday**: Checks if active: `=IF(E4>0, 1, 0)`.
  * **Week #**: Extracts week number: `=WEEKNUM(A4, 2)` (Monday-start).
* **Weekly Summary Table (Columns I–K)**:
  * Aggregates total hours and workdays for ISO weeks **23**, **24**, and **25** using `=SUMIFS($E$4:$E$100, $G$4:$G$100, I4)` and `=SUMIFS($F$4:$F$100, $G$4:$G$100, I4)`.
  * **Grand Total Row (Row 7)**: Sums up hours and workdays for the entire June 1 – June 18 period using `=SUM(J4:J6)`.

### 2. `Activity` (Raw Data & Folder Metrics)
A database table containing the raw records aggregated from the system scan:
* **Raw Folder Activity (Columns A–F)**:
  * **Work Date** (shifted date object)
  * **Folder Name**
  * **Folder Path** (beautified)
  * **Earliest Action** (actual timestamp)
  * **Latest Action** (actual timestamp)
  * **File Modifications**: Count of file edits that occurred in this folder on this day.
* **Folder Name Summary (Columns H–I)**:
  * Pre-populated with the unique list of folders worked on.
  * **Total Edits**: Live formula summing total file modifications in each folder: `=SUMIF($B$2:$B$5000, H2, $F$2:$F$5000)`.
  * **Grand Total Row**: Standard accounting summary row aggregating all edits: `=SUM(...)`.

---

## 🚀 How to Run and Reuse the Script

You can re-run the python script at any time to refresh the spreadsheet with new data or scan a different date range.

### Prerequisites
Make sure `openpyxl` is installed:
```bash
pip install openpyxl
```

### Usage
From your terminal, navigate to this folder and run:
```bash
python3 timesheet_generator.py --start 2026-06-01 --end 2026-06-18 --output timesheet.xlsx
```

### Command Line Arguments:
* `--start YYYY-MM-DD`: The start date of the period (default: `2026-06-01`).
* `--end YYYY-MM-DD`: The end date of the period (default: `2026-06-18`).
* `--output filename.xlsx`: Name of the output Excel file (default: `timesheet.xlsx`).

---

## ✍️ How to Continue Updating Manually

If you prefer to update the spreadsheet manually without running the script:
1. Open `timesheet.xlsx`.
2. Go to the **`Activity`** worksheet.
3. Scroll to the bottom of the raw log (Columns A–F) and add a new row:
   * **Work Date**: `YYYY-MM-DD` (e.g. `2026-06-19`)
   * **Folder Name**: e.g., `mytimesheet`
   * **Folder Path**: e.g., `My Drive/StataPackageIdeas/mytimesheet`
   * **Earliest Action**: `YYYY-MM-DD hh:mm AM/PM` (e.g. `2026-06-19 09:15 AM`)
   * **Latest Action**: `YYYY-MM-DD hh:mm AM/PM` (e.g. `2026-06-19 05:45 PM`)
   * **File Modifications**: e.g. `5`
4. Go to the **`Summary`** worksheet.
5. Extend the dates in column A. You can drag the formulas in columns B–G down to match the new dates. The formulas will automatically scan the `Activity` worksheet, update the daily start/end times, calculate the work hours, and adjust the weekly summaries.
