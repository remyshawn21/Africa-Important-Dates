# Africa Important Dates

TVS NSE1 Africa Region event calendar dashboard. A single self-contained
`index.html` file, generated from an Excel workbook.

**Live dashboard:** https://YOUR-USERNAME.github.io/Africa-Important-Dates/
*(update this link once GitHub Pages is enabled)*

## Files in this folder

| File | Purpose |
|---|---|
| `index.html` | The dashboard itself. Open it in any browser. Share this file directly — no server needed. |
| `Africa Important Dates.xlsx` | Source data. Two sheets: **All Important Dates** and **Big Ticket Activities**. |
| `update_dashboard.py` | Reads the Excel file and regenerates `index.html`. Run this after every Excel edit. |
| `tvs_logo.jpg` | Logo shown in the dashboard header (already embedded in `index.html`; kept here for reference). |
| `consolidate.py` | Optional helper — see warning below. Not part of the normal day-to-day workflow. |

## Normal workflow

1. Open `Africa Important Dates.xlsx` and edit either sheet — add rows, update
   statuses, fill in descriptions, etc. Keep the column headers exactly as they are:

   `Country | Date | Event | Status | Description | Category of Event | Requires PR`

2. Save and close the file.
3. Open a terminal in this folder and run:
   ```
   python update_dashboard.py
   ```
   It prints a summary (event counts per sheet, and a warning if anything
   landed in the "Unscheduled / TBC" bucket because its date couldn't be parsed).
4. Open `index.html` to check it looks right.
5. Commit and push:
   ```
   git add -A
   git commit -m "update events"
   git push
   ```
   If GitHub Pages is enabled on this repo, the live dashboard updates
   automatically within a minute or two of the push — no extra step needed.

## Date formats `update_dashboard.py` understands

- Exact dates: `20-Feb-2026` (or a real Excel date cell)
- Day ranges: `20-21 Nov 2026` → anchored to the 20th, shown as "20–21 Nov 2026"
- Month only, no day: `September` → shown as "~Sep 2026", sorted after exact-dated events in that month
- `TBC` (or anything else unparseable) → placed in a separate "Unscheduled / TBC" section at the end of the calendar

## ⚠️ About `consolidate.py`

This script merges separate per-country Excel files (dropped into a
`country_files/` folder) into the master `Africa Important Dates.xlsx`,
**overwriting the entire "All Important Dates" sheet** with whatever it finds
in `country_files/`.

**This is not the workflow currently in use.** Since the switch to a single
shared workbook, events are added by editing `Africa Important Dates.xlsx`
directly. Only run `consolidate.py` if you deliberately want to rebuild the
main sheet from a fresh set of per-country files — otherwise it will discard
your manual edits. It never touches the "Big Ticket Activities" sheet, which
is curated by hand.
