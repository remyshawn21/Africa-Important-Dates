#!/usr/bin/env python3
"""
NSE1 Dashboard Generator
=========================
Reads an Excel workbook with two sheets:
  - "All Important Dates"     (the main Africa Region event calendar)
  - "Big Ticket Activities"   (curated high-priority events)

Both sheets must use these columns (in any order):
  Country | Date | Event | Status | Description | Category of Event | Requires PR

...and regenerates index.html (the dashboard) in place.

USAGE
-----
1. Put this script, the Excel file, and the current index.html in the same folder.
2. Run:  python update_dashboard.py
3. Share/open the updated index.html.

If more than one .xlsx file is in the folder, edit EXCEL_FILENAME below to
pin down which one to use.

DATE HANDLING
-------------
- Normal dates ("20-Feb-2026" or a real Excel date) parse exactly.
- Day ranges like "20-21 Nov 2026" are placed on the first day (20 Nov 2026)
  and displayed as "20–21 Nov 2026".
- Month-only values ("September", "October"...) are placed in that month
  (year defaults to DEFAULT_YEAR below) with no specific day, and shown
  abbreviated ("~Sep 2026") with a tilde marking them as approximate.
- "TBC" (or any other unparseable value) is placed in a separate
  "Unscheduled / TBC" bucket at the end of the calendar.
"""

import json
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency. Run: pip install openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
EXCEL_FILENAME = "Africa Important Dates.xlsx"   # <-- update if your filename differs
TEMPLATE_FILENAME = "index.html"                 # existing dashboard, used as template
OUTPUT_FILENAME = "index.html"
DEFAULT_YEAR = 2026                              # year assumed for month-only dates

SHEET_MAIN = "All Important Dates"
SHEET_BIGTICKET = "Big Ticket Activities"

REQUIRED_COLUMNS = ["Country", "Date", "Event", "Status", "Description",
                     "Category of Event", "Requires PR"]

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
MONTH_INDEX = {m.lower(): i + 1 for i, m in enumerate(MONTHS)}
MONTH_ABBR = {i + 1: m[:3] for i, m in enumerate(MONTHS)}

TBC_BUCKET = "Unscheduled / TBC"


# ─────────────────────────────────────────────────────────────────────────────
def find_excel_file():
    here = Path(__file__).parent
    explicit = here / EXCEL_FILENAME
    if explicit.exists():
        return explicit
    candidates = sorted(p for p in here.glob("*.xlsx") if not p.name.startswith("~$"))
    if not candidates:
        sys.exit(f"No .xlsx file found next to {Path(__file__).name}")
    if len(candidates) > 1:
        print(f"Multiple .xlsx files found — using '{candidates[0].name}'. "
              f"Set EXCEL_FILENAME in the script to pin a specific file.")
    return candidates[0]


def parse_event_date(raw):
    """Returns dict: date_str, sort_date ('YYYY-MM-DD' or None), no_day, is_tbc."""
    if raw is None:
        return {"date_str": "TBC", "sort_date": None, "no_day": True, "is_tbc": True}

    if isinstance(raw, datetime):
        return {"date_str": raw.strftime("%d %b %Y"),
                "sort_date": raw.strftime("%Y-%m-%d"),
                "no_day": False, "is_tbc": False}

    s = str(raw).strip()
    if not s or s.upper() == "TBC":
        return {"date_str": "TBC", "sort_date": None, "no_day": True, "is_tbc": True}

    # Standard "20-Feb-2026"
    m = re.match(r"^(\d{1,2})-([A-Za-z]{3})-(\d{4})$", s)
    if m:
        try:
            dt = datetime.strptime(s, "%d-%b-%Y")
            return {"date_str": dt.strftime("%d %b %Y"), "sort_date": dt.strftime("%Y-%m-%d"),
                    "no_day": False, "is_tbc": False}
        except ValueError:
            pass

    # Day range, e.g. "20-21 Nov 2026" — anchor to the first day
    m = re.match(r"^(\d{1,2})-(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})$", s)
    if m:
        d1, d2, mon_str, year = m.groups()
        try:
            dt = datetime.strptime(f"{d1} {mon_str[:3]} {year}", "%d %b %Y")
            return {"date_str": f"{d1}\u2013{d2} {dt.strftime('%b')} {year}",
                    "sort_date": dt.strftime("%Y-%m-%d"), "no_day": False, "is_tbc": False}
        except ValueError:
            pass

    # Month name only, e.g. "September" — no specific day given
    if s.lower() in MONTH_INDEX:
        mon_num = MONTH_INDEX[s.lower()]
        dt = datetime(DEFAULT_YEAR, mon_num, 1)
        return {"date_str": f"{MONTH_ABBR[mon_num]} {DEFAULT_YEAR}",
                "sort_date": dt.strftime("%Y-%m-%d"), "no_day": True, "is_tbc": False}

    # Anything else unrecognized — bucket as TBC but keep the original text visible
    return {"date_str": s, "sort_date": None, "no_day": True, "is_tbc": True}


def load_sheet(ws):
    header_row = [c.value.strip() if isinstance(c.value, str) else c.value for c in ws[1]]
    missing = [c for c in REQUIRED_COLUMNS if c not in header_row]
    if missing:
        sys.exit(f"Sheet '{ws.title}' is missing required column(s): {', '.join(missing)}")
    col = {name: header_row.index(name) for name in REQUIRED_COLUMNS}

    events = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        country = row[col["Country"]]
        event_name = row[col["Event"]]
        if not country or not event_name:
            continue  # blank / spacer row
        parsed = parse_event_date(row[col["Date"]])
        events.append({
            "Country": str(country).strip(),
            "Event": str(event_name).strip(),
            "Status": str(row[col["Status"]] or "").strip(),
            "Description": str(row[col["Description"]] or "").strip(),
            "CategoryOfEvent": str(row[col["Category of Event"]] or "").strip(),
            "RequiresPR": str(row[col["Requires PR"]] or "No").strip(),
            **parsed,
        })
    return events


def build_raw_structure(events, today):
    """Builds the {months: [...], data: {month: {country: [events]}}} shape."""
    dated = [e for e in events if not e["is_tbc"]]
    tbc = [e for e in events if e["is_tbc"]]

    data = {}
    month_keys = set()
    for e in dated:
        dt = datetime.strptime(e["sort_date"], "%Y-%m-%d")
        my = f"{MONTHS[dt.month - 1]} {dt.year}"
        e["_dt"] = dt
        data.setdefault(my, {}).setdefault(e["Country"], []).append(e)
        month_keys.add((dt.year, dt.month, my))

    months_list = [my for (_, _, my) in sorted(month_keys)]

    # Sort within each country: exact-day events first (chronological), then
    # no-day/approximate events after them (alphabetical by event name).
    for my, countries in data.items():
        for country, evs in countries.items():
            evs.sort(key=lambda e: (e["no_day"], e["_dt"], e["Event"]))

    def finalize(e):
        dt = e["_dt"]
        if e["no_day"]:
            date_passed = (dt.year, dt.month) < (today.year, today.month)
        else:
            date_passed = dt.date() < today.date()
        return {
            "DateStr": e["date_str"], "Event": e["Event"], "Status": e["Status"],
            "Description": e["Description"], "DatePassed": date_passed,
            "CategoryOfEvent": e["CategoryOfEvent"], "RequiresPR": e["RequiresPR"],
            "Country": e["Country"], "NoDay": e["no_day"], "SortDate": e["sort_date"],
        }

    final_data = {my: {c: [finalize(e) for e in evs] for c, evs in countries.items()}
                  for my, countries in data.items()}

    if tbc:
        months_list.append(TBC_BUCKET)
        bucket = {}
        for e in tbc:
            bucket.setdefault(e["Country"], []).append({
                "DateStr": "TBC" if e["date_str"] == "TBC" else e["date_str"],
                "Event": e["Event"], "Status": e["Status"], "Description": e["Description"],
                "DatePassed": False, "CategoryOfEvent": e["CategoryOfEvent"],
                "RequiresPR": e["RequiresPR"], "Country": e["Country"],
                "NoDay": True, "SortDate": None,
            })
        for c in bucket:
            bucket[c].sort(key=lambda e: e["Event"])
        final_data[TBC_BUCKET] = bucket

    return {"months": months_list, "data": final_data}


def replace_js_const(html, const_name, new_json):
    pattern = rf'const {re.escape(const_name)} = \{{.*?\}};\n'
    replacement = f'const {const_name} = {new_json};\n'
    new_html, count = re.subn(pattern, replacement, html, count=1, flags=re.S)
    if count == 0:
        sys.exit(f"Could not find '{const_name}' in {TEMPLATE_FILENAME} — "
                  f"has the template's script section been renamed or restructured?")
    return new_html


def main():
    excel_path = find_excel_file()
    print(f"Reading '{excel_path.name}' ...")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    for sheet in (SHEET_MAIN, SHEET_BIGTICKET):
        if sheet not in wb.sheetnames:
            sys.exit(f"Sheet '{sheet}' not found in workbook. "
                      f"Sheets present: {', '.join(wb.sheetnames)}")

    main_events = load_sheet(wb[SHEET_MAIN])
    bigticket_events = load_sheet(wb[SHEET_BIGTICKET])

    today = datetime.now()
    main_raw = build_raw_structure(main_events, today)
    bigticket_raw = build_raw_structure(bigticket_events, today)

    template_path = Path(__file__).parent / TEMPLATE_FILENAME
    if not template_path.exists():
        sys.exit(f"Template file '{TEMPLATE_FILENAME}' not found. "
                  f"Keep the current dashboard file in this folder before running the script.")
    html = template_path.read_text(encoding="utf-8")

    html = replace_js_const(html, "MAIN_DATA", json.dumps(main_raw, ensure_ascii=False))
    html = replace_js_const(html, "BIGTICKET_DATA", json.dumps(bigticket_raw, ensure_ascii=False))

    timestamp = datetime.now().strftime("%d %b %Y, %H:%M")
    html, n = re.subn(r"Last updated: <strong>.*?</strong>",
                       f"Last updated: <strong>{timestamp}</strong>", html, count=1)
    if n == 0:
        print("Warning: could not find the 'Last updated' timestamp to refresh.")

    out_path = Path(__file__).parent / OUTPUT_FILENAME
    out_path.write_text(html, encoding="utf-8")

    def count_events(raw):
        return sum(len(evs) for countries in raw["data"].values() for evs in countries.values())

    print(f"\nDone — wrote {out_path.name}")
    print(f"  All Important Dates:   {count_events(main_raw)} events "
          f"across {len(main_raw['months'])} month bucket(s)")
    print(f"  Big Ticket Activities: {count_events(bigticket_raw)} events "
          f"across {len(bigticket_raw['months'])} month bucket(s)")
    if TBC_BUCKET in main_raw["data"]:
        n_tbc = sum(len(v) for v in main_raw["data"][TBC_BUCKET].values())
        print(f"  ⚠️  {n_tbc} event(s) landed in the '{TBC_BUCKET}' bucket "
              f"(no parseable month/date given).")


if __name__ == "__main__":
    main()
