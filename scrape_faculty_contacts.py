"""Extract faculty/admin contact details from a list of college website URLs.

The script reuses the official-site crawler and text extractors from
aggregators.py, so it can collect:
  - email addresses, including obfuscated forms like name[at]domain[dot]com
  - phone numbers
  - a likely leadership/faculty contact person when present
  - the pages visited while crawling

Input formats supported:
  - .txt: one URL per line
  - .csv: first column or a column named URL / Website / Website Found
  - .xlsx / .xls: first sheet, first column or a URL-like column

Usage examples:
  python scrape_faculty_contacts.py urls.txt -o faculty_contacts.xlsx
  python scrape_faculty_contacts.py colleges.xlsx -o faculty_contacts.xlsx --limit 50
"""
from __future__ import annotations

import argparse
import os
from urllib.parse import urlparse
import time

import pandas as pd

import aggregators
import main


URL_COLUMNS = ('url', 'website', 'website found', 'website_url', 'website found url')
NAME_COLUMNS = ('college name', 'name', 'institution name', 'institute name')
STATE_COLUMNS = ('state', 'state name')
DEFAULT_INPUT_FILES = ('colleges_list.xlsx')


def _normalize_url(value):
    if pd.isna(value):
        return ''
    url = str(value).strip()
    if not url:
        return ''
    if not url.startswith(('http://', 'https://')):
        url = 'https://' + url.lstrip('/')
    return url


def _looks_like_url(value):
    if not value:
        return False
    try:
        parsed = urlparse(str(value).strip())
        return bool(parsed.scheme and parsed.netloc)
    except Exception:
        return False


def _pick_url_column(columns):
    lowered = {str(col).strip().lower(): col for col in columns}
    for candidate in URL_COLUMNS:
        if candidate in lowered:
            return lowered[candidate]
    for col in columns:
        sample = str(col).strip().lower()
        if 'url' in sample or 'website' in sample:
            return col
    return None


def _pick_name_column(columns):
    lowered = {str(col).strip().lower(): col for col in columns}
    for candidate in NAME_COLUMNS:
        if candidate in lowered:
            return lowered[candidate]
    return None


def _pick_state_column(columns):
    lowered = {str(col).strip().lower(): col for col in columns}
    for candidate in STATE_COLUMNS:
        if candidate in lowered:
            return lowered[candidate]
    return None


def load_urls(input_path):
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.txt':
        with open(input_path, 'r', encoding='utf-8', errors='ignore') as handle:
            return [
                _normalize_url(line)
                for line in handle
                if _normalize_url(line)
            ]

    if ext in {'.csv'}:
        df = pd.read_csv(input_path, dtype=str, encoding_errors='ignore', low_memory=False)
    else:
        df = pd.read_excel(input_path, dtype=str)

    url_col = _pick_url_column(df.columns)
    if url_col is not None:
        urls = [_normalize_url(v) for v in df[url_col].tolist()]
        valid_urls = [url for url in urls if _looks_like_url(url)]
        if valid_urls:
            print(f'Using {url_col} column: {len(valid_urls)} URLs found')
            return valid_urls

    name_col = _pick_name_column(df.columns)
    state_col = _pick_state_column(df.columns)
    if name_col and state_col and _pick_url_column(df.columns) is None:
        print(f'No URL column found; will search for URLs using {name_col} + {state_col}')
        urls = []
        for idx, (name, state) in enumerate(zip(df[name_col], df[state_col]), start=1):
            if pd.isna(name) or str(name).strip() == '':
                continue
            college_name = str(name).strip()
            college_state = str(state).strip() if not pd.isna(state) else ''
            print(f'  [{idx}] Searching for {college_name}...')
            url = main.get_website_url(college_name, college_state)
            if url:
                urls.append(_normalize_url(url))
                print(f'    -> {url}')
            time.sleep(0.4)
        return urls

    first_col = df.columns[0]
    urls = [_normalize_url(v) for v in df[first_col].tolist()]
    return [url for url in urls if _looks_like_url(url)]


def extract_from_url(url, max_pages=4):
    text, visited = aggregators.crawl_official_site(url, max_pages=max_pages)
    extracted = aggregators.extract_contacts_from_text(text)
    admin = aggregators.extract_admin_contact(text, homepage_url=url)
    result = {
        'Source URL': url,
        'Final URL': visited[0] if visited else '',
        'Pages Visited': len(visited),
        'Visited Pages': ' | '.join(visited),
        'Email ID': extracted.get('Email ID', ''),
        'Phone Number': extracted.get('Phone Number', ''),
        'POC Name': extracted.get('POC Name', ''),
        'Approx Strength': extracted.get('Approx Strength', ''),
        'Contact Name': admin.get('name', '') if admin else '',
        'Contact Role': admin.get('role', '') if admin else '',
        'Contact Email': admin.get('email', '') if admin else '',
        'Status': 'ok' if text else 'empty',
    }
    return result


def main():
    parser = argparse.ArgumentParser(description='Extract faculty/admin contact details from website URLs.')
    parser.add_argument('input_file', nargs='?', help='TXT, CSV, XLSX, or XLS file containing website URLs')
    parser.add_argument('-o', '--output', default='faculty_contacts.xlsx', help='Output Excel file')
    parser.add_argument('--limit', type=int, default=0, help='Process only the first N URLs')
    parser.add_argument('--max-pages', type=int, default=4, help='Maximum internal pages to crawl per site')
    args = parser.parse_args()

    if not args.input_file:
        for candidate in DEFAULT_INPUT_FILES:
            if os.path.exists(candidate):
                args.input_file = candidate
                print(f'No input file given; using {candidate}')
                break

    if not args.input_file:
        parser.error('input_file is required unless urls.txt, colleges_list.xlsx, or scraped_results.xlsx exists in the current folder')

    urls = load_urls(args.input_file)
    if args.limit and args.limit > 0:
        urls = urls[:args.limit]

    if not urls:
        print(f'No valid URLs found in {args.input_file}')
        return

    print(f'Loaded {len(urls)} URLs')
    rows = []
    for idx, url in enumerate(urls, start=1):
        print(f'[{idx}/{len(urls)}] {url}')
        try:
            rows.append(extract_from_url(url, max_pages=args.max_pages))
        except Exception as exc:
            rows.append({
                'Source URL': url,
                'Final URL': '',
                'Pages Visited': 0,
                'Visited Pages': '',
                'Email ID': '',
                'Phone Number': '',
                'POC Name': '',
                'Approx Strength': '',
                'Contact Name': '',
                'Contact Role': '',
                'Contact Email': '',
                'Status': f'error: {exc}',
            })

    out_df = pd.DataFrame(rows)
    out_df.to_excel(args.output, index=False)
    print(f'Wrote {len(out_df)} rows to {args.output}')


if __name__ == '__main__':
    main()