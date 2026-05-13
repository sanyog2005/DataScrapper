"""Backfill the three Placement Officer columns for already-scraped rows.

For each row in scraped_results.xlsx where Website Found is set but the
placement columns are empty, visit the college's placement page (subdomain
or /placement / /tpo / /cdc path) and extract:

    Placement Officer Name
    Placement Officer Phone
    Placement Officer Email

The main spreadsheet's other columns and the Faculty sheet are preserved.

Usage:
    venv\\Scripts\\python.exe backfill_placement.py
    venv\\Scripts\\python.exe backfill_placement.py --start-row 100 --limit 50
"""
import argparse
import os
import time

import pandas as pd

import aggregators

INPUT_OUTPUT = 'scraped_results.xlsx'

PLACEMENT_FIELDS = ['Placement Officer Name',
                    'Placement Officer Phone',
                    'Placement Officer Email']


def _is_empty(v):
    return pd.isna(v) or str(v).strip() == ''


def main():
    parser = argparse.ArgumentParser(
        description='Fill the three Placement Officer columns for rows that '
                    'already have a Website Found URL.')
    parser.add_argument('--start-row', type=int, default=1,
                        help='Start from this 1-based data row (default: 1)')
    parser.add_argument('--limit', type=int, default=0,
                        help='Process at most N rows from start (default: all)')
    args = parser.parse_args()

    if not os.path.exists(INPUT_OUTPUT):
        print(f'No {INPUT_OUTPUT} found. Run main.py first.')
        return

    df = pd.read_excel(INPUT_OUTPUT, sheet_name='Colleges', dtype=str)
    try:
        faculty_df = pd.read_excel(INPUT_OUTPUT, sheet_name='Faculty', dtype=str)
    except Exception:
        faculty_df = None

    # Ensure the placement columns exist on the sheet
    for col in PLACEMENT_FIELDS:
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].astype(object)

    todo = []
    for index, row in df.iterrows():
        if index + 1 < args.start_row:
            continue
        url = row.get('Website Found')
        if _is_empty(url) or str(url).strip().lower() == 'not found':
            continue
        # Skip rows that already have any of the placement fields filled
        if not all(_is_empty(row.get(c)) for c in PLACEMENT_FIELDS):
            continue
        todo.append(index)
        if args.limit and len(todo) >= args.limit:
            break

    print(f'{len(todo)} rows to process (start-row={args.start_row}, '
          f'limit={args.limit or "all"})')
    if not todo:
        return

    filled = 0
    for i, index in enumerate(todo, start=1):
        row = df.loc[index]
        college = row.get('College Name')
        url = row.get('Website Found')
        pl = aggregators.find_placement_contact(str(url))
        if pl:
            wrote = []
            if pl.get('name'):
                df.at[index, 'Placement Officer Name'] = pl['name']
                wrote.append('Name')
            if pl.get('phone'):
                df.at[index, 'Placement Officer Phone'] = pl['phone']
                wrote.append('Phone')
            if pl.get('email'):
                df.at[index, 'Placement Officer Email'] = pl['email']
                wrote.append('Email')
            if wrote:
                filled += 1
                print(f'[{i}/{len(todo)}] row {index + 1} {str(college)[:50]:50s} '
                      f'-> {",".join(wrote)} email={pl.get("email", "-")}')
            else:
                print(f'[{i}/{len(todo)}] row {index + 1} {str(college)[:50]:50s} -> (nothing)')
        else:
            print(f'[{i}/{len(todo)}] row {index + 1} {str(college)[:50]:50s} -> (no placement page)')

        if i % 25 == 0:
            with pd.ExcelWriter(INPUT_OUTPUT, engine='openpyxl') as w:
                df.to_excel(w, sheet_name='Colleges', index=False)
                if faculty_df is not None:
                    faculty_df.to_excel(w, sheet_name='Faculty', index=False)
            print('  ...checkpoint saved')
        time.sleep(0.4)

    with pd.ExcelWriter(INPUT_OUTPUT, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='Colleges', index=False)
        if faculty_df is not None:
            faculty_df.to_excel(w, sheet_name='Faculty', index=False)
    print(f'\nDone. Filled placement contact info on {filled} rows.')


if __name__ == '__main__':
    main()
