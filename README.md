# CAB Report Generator

## Overview

This utility generates a formatted report for weekly CAB (Change Advisory Board) communications using data exported from Tableau.

The goal is to reduce manual effort, improve consistency, and standardize formatting for CAB email summaries, while still allowing for required manual review and judgment (e.g., UCPath impact).

---

## What It Does

* Parses Tableau-exported Excel data
* Filters for relevant Production changes within the CAB window
* Groups changes by change owner
* Formats output into Outlook-compatible format (RTF and HTML support)
* Produces:

  * A ready-to-copy-and-paste email body
  * A manual review file for edge cases (e.g., missing environment data)

---

## Inputs

* Tableau-exported Excel file (CAB report data)

### Expected Fields

The script automatically detects headers but expects fields such as:

* Change#
* Short Description
* Start Date
* End Date
* Environments Impacted
* Assignment Group / Change Owner Group
* Requested By
* Downtime
* State

---

## Outputs

### 1. Email Body

`weekly_change_email.<format>`

* Outlook-compatible formatting
* Grouped by assignment group
* Includes:

  * Change number and description
  * Change window (formatted for readability)
  * Downtime (human-readable format)
  * Requestor contact line

### 2. Manual Review File

`manual_review.txt`

* Contains records requiring manual validation:

  * Missing environment
  * Mixed environment (e.g., Production + Test)

---

## How to Run

```bash
python make_change_report.py \
  --input "<path to Excel file>" \
  --cab-date YYYY-MM-DD \
  --output-html weekly_change_email.<rtf,html> \
  --output-review manual_review.txt
```

### Example

```bash
python make_change_report.py \
  --input "./CAB Email Sheet.xlsx" \
  --cab-date 2026-04-07
```

---

## Workflow

1. Export CAB data from Tableau using Excel/Crosstab option and format.
2. Run the script using the exported file.
3. Open the output file.
4. Copy and paste into Outlook email (use web-based OWA for the most consistent results)
5. Perform manual review:

   * Highlight UCPath-related changes (identified during CAB)
   * Validate any entries in `manual_review.txt`
6. Send finalized CAB communication/draft

---

## Key Formatting Rules

* Bullets and use of bold and highlighted fonts is not consistent between your HTML/RTF viewer and OWA.  Review and adjust as necessary.
* Change window uses readable format:

  * Example: `Sat., 4/11 10:00 PM to Sun., 4/12 6:00 AM`
* Downtime is humanized:

  * "There will be no downtime"
  * "There will be 20 minutes of downtime"
  * "There will be 1 hour and 30 minutes of downtime"
* Uses Outlook-safe RTF/HTML (no Markdown, minimal styling)

---

## Known Limitations

* UCPath impact is not programmatically determined
* Must be identified and highlighted manually after CAB
* Requires Tableau export (no direct ServiceNow integration)
* Assumes relatively consistent export structure
* Designed for Outlook desktop compatibility (Windows + Mac)

---

## Dependencies

* Python 3.x
* pandas
* openpyxl

How to Install Python 3 dependencies:

```bash
pip install pandas openpyxl
```

---

## Ownership & Maintenance

* Script owner: (add name here)
* Updates should be version-controlled
* Business logic changes (grouping, formatting rules, etc.) should be documented

---

## Future Enhancements (Optional)

* ServiceNow integration (native report generation)
* Additional structured fields (e.g., UCPath impact flag)
* Automated email generation
* Scheduling / shared runtime execution
* Improved validation and error handling

---

## Notes

This tool is intended to support and streamline CAB reporting—not replace judgment or CAB discussion outcomes.

Manual review remains a required step in the process.

