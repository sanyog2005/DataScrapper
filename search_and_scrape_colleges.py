"""End-to-end pipeline: search for college URLs, then crawl for faculty contacts.

This unified script combines the URL-search step from main.py with the
faculty-contact crawl from scrape_faculty_contacts.py:

  1. Load college names + states from input file
  2. Search for each college's official website URL
  3. Crawl the site + contact pages for faculty/admin details
  4. Extract emails (including obfuscated forms), phones, POC names
  5. Write everything to Excel

Usage:
  python search_and_scrape_colleges.py colleges_list.xlsx -o results.xlsx
  python search_and_scrape_colleges.py colleges_list.xlsx --limit 10
"""
from __future__ import annotations

import argparse
import os
import time

import pandas as pd

import aggregators
import main


NAME_COLUMNS = ('college name', 'name', 'institution name', 'institute name')
STATE_COLUMNS = ('state', 'state name')


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


def load_colleges(input_path):
    """Load college name + state pairs from input file."""
    ext = os.path.splitext(input_path)[1].lower()
    
    if ext in {'.csv'}:
        df = pd.read_csv(input_path, dtype=str, encoding_errors='ignore', low_memory=False)
    elif ext in {'.xlsx', '.xls'}:
        df = pd.read_excel(input_path, dtype=str)
    else:
        raise ValueError(f'Unsupported file format: {ext}')
    
    name_col = _pick_name_column(df.columns)
    state_col = _pick_state_column(df.columns)
    
    if not name_col:
        raise ValueError(f'Could not find a college name column. Available: {list(df.columns)}')
    if not state_col:
        raise ValueError(f'Could not find a state column. Available: {list(df.columns)}')
    
    colleges = []
    for idx, (name, state) in enumerate(zip(df[name_col], df[state_col])):
        if pd.isna(name) or str(name).strip() == '':
            continue
        colleges.append({
            'name': str(name).strip(),
            'state': str(state).strip() if not pd.isna(state) else '',
            'row_index': idx,
        })
    
    return colleges


def search_and_crawl_college(college_name, college_state, max_pages=6):
    """Search for a college's website, then crawl it for faculty contacts.
    
    Returns a dict with all extracted fields + search/crawl metadata.
    """
    print(f'  Searching for {college_name}...')
    url = main.get_website_url(college_name, college_state)
    
    result = {
        'College Name': college_name,
        'State': college_state,
        'Website Found': '',
        'Email ID': '',
        'Phone Number': '',
        'POC Name': '',
        'Approx Strength': '',
        'Contact Name': '',
        'Contact Role': '',
        'Contact Email': '',
        'Pages Visited': 0,
        'Visited Pages': '',
        'Status': 'not_found',
    }
    
    if not url:
        result['Status'] = 'url_not_found'
        return result
    
    result['Website Found'] = url
    print(f'    -> {url}')
    print(f'  Crawling {url}...')
    
    try:
        text, visited = aggregators.crawl_official_site(url, max_pages=max_pages)
        if not text:
            result['Status'] = 'crawl_empty'
            return result
        
        extracted = aggregators.extract_contacts_from_text(text)
        admin = aggregators.extract_admin_contact(text, homepage_url=url)
        
        result['Email ID'] = extracted.get('Email ID', '')
        result['Phone Number'] = extracted.get('Phone Number', '')
        result['POC Name'] = extracted.get('POC Name', '')
        result['Approx Strength'] = extracted.get('Approx Strength', '')
        
        if admin:
            result['Contact Name'] = admin.get('name', '')
            result['Contact Role'] = admin.get('role', '')
            result['Contact Email'] = admin.get('email', '')
        
        result['Pages Visited'] = len(visited)
        result['Visited Pages'] = ' | '.join(visited)
        result['Status'] = 'ok'
        
    except Exception as e:
        result['Status'] = f'crawl_error: {str(e)[:50]}'
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description='Search for college URLs, then crawl them for faculty contacts.'
    )
    parser.add_argument('input_file', help='XLSX or CSV file with college names and states')
    parser.add_argument('-o', '--output', default='search_and_scrape_results.xlsx', 
                        help='Output Excel file')
    parser.add_argument('--limit', type=int, default=0, 
                        help='Process only the first N colleges')
    parser.add_argument('--max-pages', type=int, default=4, 
                        help='Maximum internal pages to crawl per site')
    args = parser.parse_args()
    
    if not os.path.exists(args.input_file):
        print(f'Error: {args.input_file} not found')
        return
    
    colleges = load_colleges(args.input_file)
    if args.limit and args.limit > 0:
        colleges = colleges[:args.limit]
    
    if not colleges:
        print('No colleges found in input file')
        return
    
    print(f'Loaded {len(colleges)} colleges')
    print(f'Will write results to {args.output}\n')
    
    results = []
    for idx, college in enumerate(colleges, start=1):
        print(f'[{idx}/{len(colleges)}] {college["name"]} ({college["state"]})')
        try:
            result = search_and_crawl_college(
                college['name'], 
                college['state'], 
                max_pages=args.max_pages
            )
            results.append(result)
        except Exception as e:
            results.append({
                'College Name': college['name'],
                'State': college['state'],
                'Website Found': '',
                'Email ID': '',
                'Phone Number': '',
                'POC Name': '',
                'Approx Strength': '',
                'Contact Name': '',
                'Contact Role': '',
                'Contact Email': '',
                'Pages Visited': 0,
                'Visited Pages': '',
                'Status': f'error: {str(e)[:80]}',
            })
        time.sleep(0.5)
    
    out_df = pd.DataFrame(results)
    out_df.to_excel(args.output, index=False)
    
    ok_count = sum(1 for r in results if r['Status'] == 'ok')
    found_count = sum(1 for r in results if r['Website Found'])
    
    print(f'\n✓ Done. {len(results)} colleges processed')
    print(f'  - URLs found: {found_count}/{len(results)}')
    print(f'  - Crawled successfully: {ok_count}/{len(results)}')
    print(f'  - Output: {args.output}')


if __name__ == '__main__':
    main()
