"""Backfill the `Website Found` column for already-scraped rows that are missing
a URL. Touches only that column — does NOT re-scrape any other field.

Usage:
    venv\\Scripts\\python.exe backfill_urls.py

The script:
  1. Loads scraped_results.xlsx (Colleges + Faculty sheets).
  2. For each row in Colleges where Website Found is empty / 'Not Found',
     runs a single DDG search using main.get_website_url().
  3. Writes the URL back and checkpoints every 25 rows.
  4. Preserves the Faculty sheet untouched.
"""
import os
import re
import time

import pandas as pd
import main

INPUT_OUTPUT = 'scraped_results.xlsx'


def _is_empty_url(v):
    if pd.isna(v):
        return True
    s = str(v).strip()
    return not s or s.lower() in {'not found', 'nan', '-'}


_ILLEGAL_XLSX_CHARS_RE = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')


def _sanitize_for_excel(v):
    if isinstance(v, str):
        return _ILLEGAL_XLSX_CHARS_RE.sub('', v)
    return v


def _sanitize_dataframe_for_excel(df):
    if df is None or df.empty:
        return df
    return df.apply(lambda col: col.map(_sanitize_for_excel))


def main_loop():
    if not os.path.exists(INPUT_OUTPUT):
        print(f'No {INPUT_OUTPUT} found. Run main.py first.')
        return

    df = pd.read_excel(INPUT_OUTPUT, sheet_name='Colleges', dtype=str)
    # Faculty sheet (if present) — we read it just so we can write it back
    # unchanged when we save.
    try:
        faculty_df = pd.read_excel(INPUT_OUTPUT, sheet_name='Faculty', dtype=str)
    except Exception:
        faculty_df = None

    if 'Website Found' not in df.columns:
        df['Website Found'] = ''
    df['Website Found'] = df['Website Found'].astype(object)

    todo = df[df['Website Found'].apply(_is_empty_url)]
    print(f'{len(todo)} rows missing a Website Found URL (of {len(df)} total).')
    if len(todo) == 0:
        return

    filled = 0
    failed = 0
    for i, (index, row) in enumerate(todo.iterrows(), start=1):
        college = row.get('College Name')
        state = row.get('State')
        if pd.isna(college) or not str(college).strip():
            continue
        url = main.get_website_url(college, state)
        if url:
            df.at[index, 'Website Found'] = url
            filled += 1
            print(f'[{i}/{len(todo)}] {str(college)[:60]:60s} -> {url}')
        else:
            df.at[index, 'Website Found'] = 'Not Found'
            failed += 1
            print(f'[{i}/{len(todo)}] {str(college)[:60]:60s} -> (not found)')

        # Checkpoint every 25 rows so a crash doesn't lose progress.
        if i % 25 == 0:
            safe_df = _sanitize_dataframe_for_excel(df)
            safe_faculty_df = _sanitize_dataframe_for_excel(faculty_df)
            with pd.ExcelWriter(INPUT_OUTPUT, engine='openpyxl') as w:
                safe_df.to_excel(w, sheet_name='Colleges', index=False)
                if safe_faculty_df is not None:
                    safe_faculty_df.to_excel(w, sheet_name='Faculty', index=False)
            print(f'  ...checkpoint saved')
        time.sleep(0.4)

    safe_df = _sanitize_dataframe_for_excel(df)
    safe_faculty_df = _sanitize_dataframe_for_excel(faculty_df)
    with pd.ExcelWriter(INPUT_OUTPUT, engine='openpyxl') as w:
        safe_df.to_excel(w, sheet_name='Colleges', index=False)
        if safe_faculty_df is not None:
            safe_faculty_df.to_excel(w, sheet_name='Faculty', index=False)
    print(f'\nDone. Filled {filled} URLs, marked {failed} as Not Found.')


if __name__ == '__main__':
    main_loop()
