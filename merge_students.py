#!/usr/bin/env python3
"""
Student records automation (run by Jenkins).

1. Reads the OLD records from     data/students.xlsx
2. Reads the NEW records from     every .xlsx file in incoming/
3. Checks each new record BEFORE adding it:
     - RollNo already in the old records      -> duplicate, skipped
     - RollNo repeated inside the new records -> duplicate, skipped
4. Saves the result in a DIFFERENT Excel file: output/merged_students.xlsx
     Sheet "All Students"       - old records + new records that were added
     Sheet "New Records Added"  - only the new records that were added
     Sheet "Duplicates Found"   - new records that were skipped, with the reason
     Sheet "Summary"            - counts

The original files are never changed.
"""
import glob
import os
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

MAIN_FILE = "data/students.xlsx"
INCOMING_FILES = "incoming/*.xlsx"
OUTPUT_FILE = "output/merged_students.xlsx"
COLUMNS = ["RollNo", "Name", "Department", "Marks", "Grade"]


def fail(message):
    print("ERROR: " + message)
    sys.exit(1)


def read_records(path):
    """Read one Excel file and return its rows as a list of dictionaries."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    rows = list(workbook.worksheets[0].iter_rows(values_only=True))
    workbook.close()

    if not rows:
        fail(f"{path} is empty")

    # Check the column names in row 1
    header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    missing = [col for col in COLUMNS if col not in header]
    if missing:
        fail(f"{path} is missing columns {missing}. Row 1 must contain: {COLUMNS}")
    position = {col: header.index(col) for col in COLUMNS}

    records = []
    for row_number, row in enumerate(rows[1:], start=2):
        # Skip completely empty rows
        if all(value is None or str(value).strip() == "" for value in row):
            continue

        record = {col: row[position[col]] if position[col] < len(row) else None for col in COLUMNS}

        roll_no = str(record["RollNo"] or "").strip().upper()
        if not roll_no:
            fail(f"{path} row {row_number}: RollNo is empty")
        record["RollNo"] = roll_no

        try:
            marks = float(record["Marks"])
        except (TypeError, ValueError):
            fail(f"{path} row {row_number} ({roll_no}): Marks '{record['Marks']}' is not a number")
        if not 0 <= marks <= 100:
            fail(f"{path} row {row_number} ({roll_no}): Marks {marks} must be between 0 and 100")
        record["Marks"] = int(marks) if marks.is_integer() else marks

        for col in ("Name", "Department", "Grade"):
            record[col] = str(record[col]).strip() if record[col] is not None else ""

        record["row"] = row_number
        records.append(record)
    return records


def values(record, extra=()):
    return tuple(record[col] for col in COLUMNS) + tuple(extra)


def main():
    if not os.path.exists(MAIN_FILE):
        fail(f"{MAIN_FILE} not found")

    # ---- Step 1: old records ----
    old_records = read_records(MAIN_FILE)
    existing = {}
    for record in old_records:
        if record["RollNo"] in existing:
            fail(f"{MAIN_FILE} itself has a duplicate RollNo: {record['RollNo']}")
        existing[record["RollNo"]] = record
    print(f"Old records in {MAIN_FILE}: {len(old_records)}")

    # ---- Step 2 and 3: new records, checked one by one ----
    new_files = sorted(glob.glob(INCOMING_FILES))
    if not new_files:
        print("No new files found in incoming/")

    added, duplicates = [], []
    added_from = {}          # RollNo -> where it was first added from
    received = 0

    for path in new_files:
        file_name = os.path.basename(path)
        records = read_records(path)
        received += len(records)
        print(f"\nChecking {file_name}: {len(records)} records")

        for record in records:
            roll_no = record["RollNo"]

            if roll_no in existing:
                old = existing[roll_no]
                if values(old) == values(record):
                    reason = "Exact duplicate of an old record"
                else:
                    reason = (f"RollNo already exists for {old['Name']} "
                              f"({old['Department']}, Marks {old['Marks']})")
                duplicates.append(values(record, (file_name, record["row"], reason)))
                print(f"  DUPLICATE skipped: {roll_no} {record['Name']} - {reason}")

            elif roll_no in added_from:
                reason = f"RollNo repeated in new records (first in {added_from[roll_no]})"
                duplicates.append(values(record, (file_name, record["row"], reason)))
                print(f"  DUPLICATE skipped: {roll_no} {record['Name']} - {reason}")

            else:
                added.append(values(record, (file_name,)))
                added_from[roll_no] = f"{file_name} row {record['row']}"
                print(f"  Added: {roll_no} {record['Name']}")

    # ---- Step 4: build the output sheets ----
    all_students = [values(r) for r in old_records] + [row[:len(COLUMNS)] for row in added]
    summary = [
        ("Old records (students.xlsx)", len(old_records)),
        ("New records received", received),
        ("New records added", len(added)),
        ("Duplicates skipped", len(duplicates)),
        ("Total records after merge", len(all_students)),
        ("Files processed", ", ".join(os.path.basename(p) for p in new_files) or "none"),
    ]
    sheets = {
        "All Students": [tuple(COLUMNS)] + all_students,
        "New Records Added": [tuple(COLUMNS) + ("Source File",)] + added,
        "Duplicates Found": [tuple(COLUMNS) + ("Source File", "Row", "Reason")] + duplicates,
        "Summary": [("Item", "Value")] + summary,
    }

    print("\n========== SUMMARY ==========")
    for item, value in summary:
        print(f"{item}: {value}")

    if output_is_same(sheets):
        print(f"\n{OUTPUT_FILE} is already up to date - not rewritten")
        return

    save_output(sheets)
    print(f"\nSaved {OUTPUT_FILE}")


def output_is_same(sheets):
    """True if the existing output already has exactly this data (avoids a needless commit)."""
    if not os.path.exists(OUTPUT_FILE):
        return False
    try:
        workbook = load_workbook(OUTPUT_FILE, read_only=True)
        same = workbook.sheetnames == list(sheets) and all(
            [tuple(r) for r in workbook[name].iter_rows(values_only=True)] == rows
            for name, rows in sheets.items()
        )
        workbook.close()
        return same
    except Exception:
        return False


def save_output(sheets):
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")

    for name, rows in sheets.items():
        sheet = workbook.create_sheet(name)
        for row in rows:
            sheet.append(row)
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
        sheet.freeze_panes = "A2"
        for i in range(1, len(rows[0]) + 1):
            longest = max(len(str(row[i - 1])) for row in rows if len(row) >= i)
            sheet.column_dimensions[get_column_letter(i)].width = min(longest + 3, 60)

    workbook.save(OUTPUT_FILE)


if __name__ == "__main__":
    main()
