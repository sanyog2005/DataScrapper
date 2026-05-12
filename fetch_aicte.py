"""Download the full AICTE approved-institutes list for the chosen academic year.

Hits the public dashboard endpoint that the Angular SPA at
facilities.aicte-india.org/dashboard/pages/angulardashboard.php uses internally:

    php/approvedinstituteserver.php?method=fetchdata&year=YYYY-YYYY&state=...&...

The endpoint returns a JSON array of arrays. Column order (observed):
    [AICTE_ID, Institute_Name, Address, District, Institution_Type, Women, Minority, PID]

We iterate every Indian state + UT and write a single Excel that main.py can
fuzzy-match against to fill District / Location for each college.

Notes:
- The API does NOT expose Email / Phone / POC. main.py's web-scrape fallback
  still has to fill those.
- The API also does not expose Stream / Intake directly here. To get those you
  would call php/approvedcourse.php?aicteid=...&course=1&year=... per institute,
  which is ~thousands of calls. We skip that by default (see ENRICH_COURSES).
"""
import json
import time
import sys
import requests
import pandas as pd

YEAR = '2024-2025'
OUTPUT = 'aicte_institutes.xlsx'

URL = 'https://facilities.aicte-india.org/dashboard/pages/php/approvedinstituteserver.php'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://facilities.aicte-india.org/dashboard/pages/angulardashboard.php',
    'X-Requested-With': 'XMLHttpRequest',
}

STATES = [
    'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chhattisgarh',
    'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jharkhand',
    'Karnataka', 'Kerala', 'Madhya Pradesh', 'Maharashtra', 'Manipur',
    'Meghalaya', 'Mizoram', 'Nagaland', 'Odisha', 'Punjab',
    'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura',
    'Uttar Pradesh', 'Uttarakhand', 'West Bengal',
    'Andaman and Nicobar Islands', 'Chandigarh',
    'Dadra and Nagar Haveli', 'Daman and Diu',
    'Delhi', 'Jammu and Kashmir', 'Ladakh', 'Lakshadweep', 'Puducherry',
]

# Column order returned by the API for each row
COLS = ['AICTE ID', 'Institute Name', 'Address', 'District',
        'Institution Type', 'Women', 'Minority', 'PID']


def fetch_state(state, year=YEAR, retries=3, backoff=2.0):
    params = {
        'method': 'fetchdata',
        'year': year,
        'program': '1',
        'level': '1',
        'institutiontype': '1',
        'Women': '1',
        'Minority': '1',
        'state': state,
        'course': '1',
    }
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(URL, params=params, headers=HEADERS, timeout=60)
            if r.status_code != 200:
                print(f'  [{r.status_code}] {state} attempt {attempt}')
                time.sleep(backoff * attempt)
                continue
            try:
                data = r.json()
            except json.JSONDecodeError:
                # The endpoint returns 200 with an empty body for states with zero rows
                if not r.text.strip():
                    return []
                print(f'  [bad json] {state}: {r.text[:200]}')
                return []
            return data if isinstance(data, list) else []
        except requests.RequestException as e:
            print(f'  [error] {state} attempt {attempt}: {e}')
            time.sleep(backoff * attempt)
    print(f'  [GIVE UP] {state}')
    return []


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else YEAR
    print(f'Fetching AICTE approved institutes for year {year}...')
    all_rows = []
    for state in STATES:
        rows = fetch_state(state, year=year)
        print(f'  {state:35s} -> {len(rows):5d} institutes')
        for row in rows:
            # Pad/truncate row to expected length
            row = list(row) + [''] * (len(COLS) - len(row))
            row = row[:len(COLS)]
            all_rows.append([state] + row)
        time.sleep(0.5)  # be polite

    df = pd.DataFrame(all_rows, columns=['State'] + COLS)
    # Mirror the address into a "Location" column so main.py's auto-detection
    # picks it up for the Location field.
    df['Location'] = df['Address']
    df.to_excel(OUTPUT, index=False)
    print(f'\nDone. {len(df)} rows -> {OUTPUT}')


if __name__ == '__main__':
    main()
