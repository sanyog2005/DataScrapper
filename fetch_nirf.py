"""Build a master index of NIRF-ranked institutes for a given year.

NIRF publishes per-category ranking pages at
    https://www.nirfindia.org/Rankings/<YEAR>/<Category>Ranking.html
where each row has the institute's Institute ID (e.g. IR-E-U-0456), Name,
City, State, Score, Rank. We scrape every category, dedupe by ID, and write
nirf_index.xlsx so the main pipeline can fuzzy-match user colleges against it.

For each institute, the submission "Data PDF" lives at
    https://www.nirfindia.org/nirfpdfcdn/<YEAR>/pdf/<Category>/<INSTITUTE_ID>.pdf
which is recorded in the index for fetching later.
"""
import re
import sys
import time

import pandas as pd
import requests
from bs4 import BeautifulSoup

YEAR = '2024'
OUTPUT = 'nirf_index.xlsx'

# Order matters: dedup by NIRF ID keeps the FIRST match, so we put the more
# specific streams up top. An institute ranked in Engineering should be
# tagged "Engineering" rather than the catch-all "Overall".
CATEGORIES = [
    'Engineering', 'Management', 'Pharmacy', 'Medical', 'Dental', 'Law',
    'Architecture', 'Research', 'Agriculture', 'College', 'University',
    'Overall',
]

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}

ID_RE = re.compile(r'^IR-[A-Z]-[A-Z]-\d{2,6}$')


def scrape_category(year, category):
    """Return list of dicts: {ID, Name, City, State, Score, Rank, Category, PDF_URL}."""
    url = f'https://www.nirfindia.org/Rankings/{year}/{category}Ranking.html'
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
    except requests.RequestException as e:
        print(f'  [err] {category}: {e}')
        return []
    if r.status_code != 200:
        print(f'  [skip] {category}: HTTP {r.status_code}')
        return []
    soup = BeautifulSoup(r.text, 'html.parser')

    rows = []
    for tr in soup.find_all('tr'):
        tds = tr.find_all('td')
        if len(tds) < 5:
            continue
        first = tds[0].get_text(strip=True)
        if not ID_RE.match(first):
            continue
        # Name is in the second td but it contains nested "More Details" markup
        # — pulling just direct text avoids that.
        name = tds[1].get_text(separator=' ', strip=True)
        # Strip the trailing "More Details" / "Close" UI text if present.
        name = re.sub(r'\s*More Details.*$', '', name).strip()

        # Schema for the remaining cells: ...<td>City</td><td>State</td><td>Score</td><td>Rank</td>
        # (the middle tds vary because of nested score-breakdown tables)
        city = tds[-4].get_text(strip=True)
        state = tds[-3].get_text(strip=True)
        score = tds[-2].get_text(strip=True)
        rank = tds[-1].get_text(strip=True)

        pdf_url = (f'https://www.nirfindia.org/nirfpdfcdn/{year}/pdf/'
                   f'{category}/{first}.pdf')
        rows.append({
            'NIRF ID':  first,
            'Name':     name,
            'City':     city,
            'State':    state,
            'Score':    score,
            'Rank':     rank,
            'Category': category,
            'Year':     year,
            'PDF URL':  pdf_url,
        })
    return rows


def main():
    year = sys.argv[1] if len(sys.argv) > 1 else YEAR
    print(f'Building NIRF index for {year}...')
    all_rows = []
    for cat in CATEGORIES:
        rows = scrape_category(year, cat)
        print(f'  {cat:14s} -> {len(rows):4d} institutes')
        all_rows.extend(rows)
        time.sleep(0.5)

    df = pd.DataFrame(all_rows)
    # Dedup: keep the *first* category we found each institute in (Overall has
    # priority since it appears first in CATEGORIES).
    df_first = df.drop_duplicates(subset=['NIRF ID'], keep='first')
    df_first.to_excel(OUTPUT, index=False)
    print(f'\nDone. {len(df)} ranked entries, {len(df_first)} unique institutes -> {OUTPUT}')


if __name__ == '__main__':
    main()
