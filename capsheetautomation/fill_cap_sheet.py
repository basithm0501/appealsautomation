import csv
import sys
import os

def fill_cap_sheet(row_id, raw_csv):
    # Read the raw data
    with open(raw_csv, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    # Find the row by SubmissionId only
    row = None
    for r in rows:
        if r.get('SubmissionId') == str(row_id):
            row = r
            break
    if row is None:
        raise ValueError('Row not found by SubmissionId')

    # Auto name the output file
    output_csv = f"cap_sheet_{row_id}.csv"

    # --- WRITING SECTION ---


  

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python fill_cap_sheet.py <submission_id> <raw.csv>')
        sys.exit(1)
    fill_cap_sheet(sys.argv[1], sys.argv[2])
