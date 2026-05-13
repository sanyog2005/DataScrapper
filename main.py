import os
import re
import time
import warnings
import argparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from rapidfuzz import fuzz, process

# DDG client (the old "duckduckgo_search" was renamed to "ddgs")
try:
    from ddgs import DDGS
except ImportError:
    from duckduckgo_search import DDGS

import aggregators

warnings.filterwarnings("ignore", category=RuntimeWarning)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

INPUT_FILE = 'colleges_list.xlsx'
OUTPUT_FILE = 'backup.xlsx'
DEFAULT_START_ROW = 1  # Edit this if you want the script to resume from a fixed row.

# Drop the AICTE/AISHE/NIRF reference files in this folder with these names
# (or update the paths below). Any one of .xlsx / .xls / .csv works.
AICTE_FILE_CANDIDATES = ['aicte_institutes.xlsx', 'aicte_institutes.xls', 'aicte_institutes.csv']
AISHE_FILE_CANDIDATES = ['aishe_colleges.xlsx', 'aishe_colleges.xls', 'aishe_colleges.csv']
NIRF_FILE_CANDIDATES  = ['nirf_index.xlsx', 'nirf_index.csv']

FUZZY_THRESHOLD = 80  # 0-100, how close a name must match to be accepted

# LinkedIn x-ray fallback for Contact Name/Role/Email. Currently OFF — see the
# note next to the (disabled) call site in process_spreadsheet for context.
ENABLE_LINKEDIN_XRAY = False
TARGET_FIELDS = ['District', 'Location', 'Stream', 'Approx Strength',
                 'Email ID', 'Phone Number', 'POC Name']
# Additional leadership-contact columns (Director / Principal / HOD etc.)
# Populated by aggregators.extract_admin_contact() from crawled site text.
CONTACT_FIELDS = ['Contact Name', 'Contact Role', 'Contact Email']
# Placement-cell contact columns. Populated by visiting the college's own
# placement / TPO / CDC page (subdomain or /placement path).
PLACEMENT_FIELDS = ['Placement Officer Name', 'Placement Officer Phone',
                    'Placement Officer Email']

# Header-name candidates for auto-detecting columns in the reference files.
# We match case-insensitively as a substring against the actual header text.
AICTE_HEADER_MAP = {
    'name':     ['institute name', 'name of institute', 'institution name', 'college name', 'name'],
    'state':    ['state'],
    'district': ['district'],
    'location': ['city', 'town', 'place', 'location'],
    'stream':   ['program', 'course', 'level', 'discipline'],
    'strength': ['intake', 'sanctioned', 'approved intake', 'enroll'],
    'email':    ['email', 'e-mail', 'mail id'],
    'phone':    ['phone', 'mobile', 'contact', 'telephone'],
    'poc':      ['principal', 'director', 'head of institute', 'head of institution', 'president'],
}

AISHE_HEADER_MAP = {
    'name':     ['name of institution', 'institution name', 'college name', 'name'],
    'state':    ['state'],
    'district': ['district'],
    'location': ['city', 'town', 'place', 'location'],
    'stream':   ['type', 'discipline', 'stream', 'category'],
    'strength': ['enrolment', 'enrollment', 'students', 'strength'],
    'email':    ['email', 'e-mail'],
    'phone':    ['phone', 'telephone', 'mobile', 'contact'],
    'poc':      ['head of institution', 'principal', 'vice chancellor', 'director'],
}

# Header map for the NIRF index (built by fetch_nirf.py). NIRF doesn't have
# strength/email/phone/poc fields in the index itself — those come from
# the per-institute PDFs which we parse on-demand.
NIRF_HEADER_MAP = {
    'name':     ['name'],
    'state':    ['state'],
    'district': ['city'],   # NIRF stores city; close enough as a district proxy
    'location': ['city'],
    'stream':   ['category'],
    'strength': [],
    'email':    [],
    'phone':    [],
    'poc':      [],
}


# ---------------------------------------------------------------------------
# Reference-file loading + matching
# ---------------------------------------------------------------------------

def _read_any(path):
    if path.lower().endswith('.csv'):
        return pd.read_csv(path, dtype=str, encoding_errors='ignore', low_memory=False)
    return pd.read_excel(path, dtype=str)


def _first_existing(candidates):
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _find_col(df, candidates):
    """First column whose header contains any candidate substring (case-insensitive)."""
    for cand in candidates:
        for col in df.columns:
            if cand.lower() in str(col).lower():
                return col
    return None


# Tiny stopword list — only true grammatical fillers. We deliberately KEEP
# words like "institute", "college", "university" because they carry weight
# (AICTE has many similarly-named sub-units; dropping these caused over-matching).
_NAME_STOPWORDS = {'the', 'of', 'and', 'a', 'an', 'for', 'in', 'at', 'on', '&'}

# Words that are "category" labels — common across many colleges and not
# distinctive on their own. Used by the distinctive-token guard below.
_GENERIC_NAME_TOKENS = {
    'institute', 'institutes', 'institution', 'institutions', 'college',
    'colleges', 'university', 'universities', 'school', 'schools',
    'centre', 'center', 'academy', 'studies', 'education', 'research',
    'technology', 'technologies', 'engineering', 'science', 'sciences',
    'management', 'information', 'national', 'indian', 'computer',
    'applied', 'business', 'commerce', 'group', 'department',
}


def _norm_name(s):
    s = '' if s is None else str(s)
    s = s.lower()
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    tokens = [t for t in s.split() if t and t not in _NAME_STOPWORDS]
    return ' '.join(tokens)


def _distinctive_tokens(s):
    """Tokens in s that are 4+ chars and not generic college vocabulary."""
    return {t for t in _norm_name(s).split()
            if len(t) >= 4 and t not in _GENERIC_NAME_TOKENS}


def _names_likely_same(query, candidate):
    """Approve a name match between the user's query and a matched name.

    Mirrors the logic used for AICTE matching: if the query has any
    distinctive tokens (non-generic), require at least one to also appear
    in the candidate. If the query is entirely generic, require a near-
    perfect similarity score (>= 90).
    """
    q_distinctive = _distinctive_tokens(query)
    if q_distinctive:
        c_tokens = set(_norm_name(candidate).split())
        return bool(q_distinctive & c_tokens)
    # Fully generic query — only accept if names are essentially identical.
    return fuzz.token_sort_ratio(query.lower(), candidate.lower()) >= 90


def _norm_state(s):
    return '' if s is None else str(s).strip().lower()


def load_reference(path, header_map, label):
    """Load a reference file and return (df, col_map) or (None, None)."""
    if not path:
        print(f"  [skip] No {label} file found.")
        return None, None
    print(f"  [load] {label}: {path}")
    df = _read_any(path)
    col_map = {key: _find_col(df, candidates) for key, candidates in header_map.items()}
    missing_required = [k for k in ('name', 'state') if not col_map[k]]
    if missing_required:
        print(f"  [warn] {label} missing required columns {missing_required}. "
              f"Headers seen: {list(df.columns)[:15]}...")
        return None, None
    print(f"  [ok]   {label} columns -> {col_map}")
    return df, col_map


def build_state_index(df, col_map):
    """Group reference rows by normalized state for fast lookup."""
    if df is None:
        return {}
    idx = {}
    name_col = col_map['name']
    state_col = col_map['state']
    for i, row in df.iterrows():
        st = _norm_state(row.get(state_col))
        nm = _norm_name(row.get(name_col))
        if not nm or not st:
            continue
        idx.setdefault(st, []).append((nm, i))
    return idx


def match_in_reference(college, state, df, col_map, state_index, threshold=FUZZY_THRESHOLD):
    """Return a dict of fields from the best matching row, or None."""
    if df is None or not state_index:
        return None
    st = _norm_state(state)
    candidates = state_index.get(st)
    if not candidates:
        # try matching across all states as a last resort
        candidates = [pair for pairs in state_index.values() for pair in pairs]
        if not candidates:
            return None
    query = _norm_name(college)
    if not query:
        return None
    names = [c[0] for c in candidates]
    # token_sort_ratio is stricter than WRatio/token_set_ratio: it requires
    # similar tokens AND similar lengths. WRatio's partial_ratio component
    # tends to score very high when a short reference name is a substring
    # of a longer query (e.g. "engineering technology" inside "jaypee
    # university of engineering and technology"), which we want to avoid.
    best = process.extractOne(query, names, scorer=fuzz.token_sort_ratio)
    if not best or best[1] < threshold:
        return None

    # Distinctive-token guard: a "distinctive" token is one that isn't a
    # generic college category word (institute, college, technology...).
    # If the query has any distinctive tokens, require at least one to also
    # appear in the matched name. This rejects matches where a query like
    # "Indian Institute of Technology" pulls a far-away college that just
    # happens to share the generic words.
    q_distinctive = {t for t in query.split()
                     if len(t) >= 4 and t not in _GENERIC_NAME_TOKENS}
    if q_distinctive:
        m_tokens = set(names[best[2]].split())
        if not (q_distinctive & m_tokens):
            return None
    else:
        # Query is entirely generic words (e.g. truncated input like
        # "Indian Institute of Technology"). We can't disambiguate among
        # the dozens of AICTE entries that share the same generic tokens,
        # so only accept a near-exact match.
        if best[1] < 95:
            return None
    matched_idx = candidates[best[2]][1]
    row = df.loc[matched_idx]

    def get(field):
        col = col_map.get(field)
        if not col:
            return ''
        val = row.get(col)
        if pd.isna(val):
            return ''
        return str(val).strip()

    return {
        'District':        get('district'),
        'Location':        get('location'),
        'Stream':          get('stream'),
        'Approx Strength': get('strength'),
        'Email ID':        get('email'),
        'Phone Number':    get('phone'),
        'POC Name':        get('poc'),
        '_score':          round(best[1], 1),
    }


# ---------------------------------------------------------------------------
# Web-scrape fallback (kept from the original script, lightly hardened)
# ---------------------------------------------------------------------------

# Domains to skip in the WEB fallback. These are either irrelevant (social,
# Wikipedia, Microsoft Bing redirects) or aggregator sites that we don't want
# to mistakenly treat as the college's own homepage.
_WEB_URL_BLACKLIST = (
    'microsoft.com', 'wikipedia.org', 'facebook.com', 'youtube.com',
    'twitter.com', 'instagram.com', 'linkedin.com', 'reddit.com',
    'shiksha.com', 'collegedunia.com', 'collegedekho.com', 'careers360.com',
    'getmyuni.com', 'collegesearch.in', 'univariety.com', 'edu.in/listing',
    'betterstudy.in', 'jagranjosh.com', 'studyabroad.com', 'indiacollegesearch',
    'indiacollegeshub', 'collegekart.com', 'sarvgyan.com', 'unirank.org',
    '4icu.org', 'collegekhabar', 'studyclap', 'admission.ac.in', 'targetstudy',
    'thelearningpoint', 'collegepravesh', 'edunuts.com', 'indcareer.com',
)


def get_website_url(college_name, state):
    search_query = f"{college_name} {state} college India official website"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(search_query, max_results=5))
            for r in results:
                href = r.get('href', '') or ''
                if any(bad in href for bad in _WEB_URL_BLACKLIST):
                    continue
                return href
    except Exception as e:
        print(f"  [!] Search failed: {e}")
    return None


def _extract_emails_and_phones(text):
    email_re = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    # Stricter Indian phone: +91/0 prefix optional, then a 10-digit core starting with 6-9,
    # or a landline-style block. Prevents matching arbitrary 10-digit runs.
    phone_re = r'(?:\+91[- ]?|0)?[6-9]\d{9}\b|\(?\d{2,4}\)?[- ]?\d{6,8}\b'
    emails = sorted(set(re.findall(email_re, text)))
    phones = sorted(set(re.findall(phone_re, text)))
    return emails, phones


def scrape_college_data(url):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(url, headers=headers, timeout=12)
        soup = BeautifulSoup(response.content, 'html.parser')
        page_text = soup.get_text(separator=' ')
        emails, phones = _extract_emails_and_phones(page_text)
        return {
            'Email ID':     ', '.join(emails[:2]) if emails else '',
            'Phone Number': ', '.join(phones[:2]) if phones else '',
        }
    except Exception as e:
        print(f"  [!] Scrape failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def _is_empty(v):
    return pd.isna(v) or str(v).strip() == ''


def _save_workbook(df, faculty_rows, output_file):
    """Write the main `Colleges` sheet and the `Faculty` sheet to one file."""
    candidates = [output_file]
    base, ext = os.path.splitext(output_file)
    candidates.append(f"{base}.partial{ext}")
    candidates.append(f"{base}.{time.strftime('%Y%m%d-%H%M%S')}{ext}")

    last_error = None
    for path in candidates:
        try:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Colleges', index=False)
                if faculty_rows:
                    pd.DataFrame(faculty_rows).to_excel(
                        writer, sheet_name='Faculty', index=False)
            if path != output_file:
                print(f"  [save] output file was locked; wrote to {path}")
            return path
        except PermissionError as e:
            last_error = e
            continue

    raise last_error


def _fill_missing(dst_row, src):
    """Copy fields from src into dst_row only where dst_row is empty."""
    filled = []
    for k in TARGET_FIELDS:
        if k in src and src[k] and _is_empty(dst_row.get(k)):
            dst_row[k] = src[k]
            filled.append(k)
    return filled


def process_spreadsheet(input_file, output_file, limit=0, start_row=1):
    print(f"Reading {input_file}...")
    df = pd.read_excel(input_file, dtype=str)

    # If we're rerunning into the same output file, reuse the saved Colleges
    # sheet so rows before the resume point keep their previous values.
    if os.path.exists(output_file):
        try:
            existing_df = pd.read_excel(output_file, sheet_name='Colleges', dtype=str)
            if not existing_df.empty:
                existing_df = existing_df.reindex(columns=df.columns, fill_value='')
                # Replace NaN with empty strings so they don't overwrite initialized columns
                existing_df = existing_df.fillna('')
                overlap = min(len(df), len(existing_df))
                if overlap:
                    df.iloc[:overlap] = existing_df.iloc[:overlap].values
                print(f"  [ok]   preloaded {overlap} rows from existing output")
        except Exception:
            pass
    
    if limit and limit > 0:
        df = df.iloc[:limit]
        print(f"Limited to first {limit} rows")

    if start_row and start_row > 1:
        print(f"Resuming from row {start_row}")

    print("\nLoading reference files (drop them next to main.py to enable):")
    aicte_path = _first_existing(AICTE_FILE_CANDIDATES)
    aishe_path = _first_existing(AISHE_FILE_CANDIDATES)
    nirf_path  = _first_existing(NIRF_FILE_CANDIDATES)
    aicte_df, aicte_cols = load_reference(aicte_path, AICTE_HEADER_MAP, 'AICTE')
    aishe_df, aishe_cols = load_reference(aishe_path, AISHE_HEADER_MAP, 'AISHE')
    nirf_df,  nirf_cols  = load_reference(nirf_path,  NIRF_HEADER_MAP,  'NIRF')
    aicte_index = build_state_index(aicte_df, aicte_cols)
    aishe_index = build_state_index(aishe_df, aishe_cols)
    nirf_index  = build_state_index(nirf_df,  nirf_cols)

    # Build a vocabulary of known Indian districts from whichever reference
    # files we loaded. This is used to extract District from free-form
    # address strings without picking up landmarks ("Near IDBI Bank ATM").
    known_districts = set()
    for ref_df, ref_cols in ((aicte_df, aicte_cols), (aishe_df, aishe_cols)):
        if ref_df is None or not ref_cols.get('district'):
            continue
        for v in ref_df[ref_cols['district']].dropna().astype(str):
            v = v.strip().lower()
            if 2 < len(v) < 40:
                known_districts.add(v)
    if known_districts:
        print(f"  [ok]   district vocabulary: {len(known_districts)} names")
    have_any_reference = bool(aicte_index or aishe_index)
    if not have_any_reference:
        print("\n  No reference files loaded. Will fall back to web scraping only.")
        print("  Expected filenames in this folder:")
        print("    AICTE: " + " | ".join(AICTE_FILE_CANDIDATES))
        print("    AISHE: " + " | ".join(AISHE_FILE_CANDIDATES))

    # Ensure all output columns exist and accept text
    for col in TARGET_FIELDS + CONTACT_FIELDS + PLACEMENT_FIELDS + ['Source', 'Website Found']:
        if col not in df.columns:
            df[col] = ''
        df[col] = df[col].astype(object)

    stats = {'aicte': 0, 'aishe': 0, 'collegedunia': 0, 'web': 0, 'none': 0, 'skipped': 0}

    # Per-faculty rows for Sheet 2. If the output file already exists with a
    # Faculty sheet (resume case), preload it so we don't lose previously
    # collected rows.
    faculty_rows = []
    if os.path.exists(output_file):
        try:
            existing_fac = pd.read_excel(output_file, sheet_name='Faculty', dtype=str)
            faculty_rows = existing_fac.to_dict('records')
            print(f"  [ok]   preloaded {len(faculty_rows)} existing faculty rows")
        except Exception:
            pass

    for index, row in df.iterrows():
        row_no = index + 1
        if start_row and row_no < start_row:
            stats['skipped'] += 1
            continue

        college = row.get('College Name')
        state = row.get('State')

        if pd.isna(college) or str(college).strip() == '':
            continue

        # Skip rows that already have a Source recorded (resume support)
        # BUT still fill in Website Found if it's missing
        if not _is_empty(row.get('Source')):
            if _is_empty(df.at[index, 'Website Found']):
                crawl_url = get_website_url(college, state)
                df.at[index, 'Website Found'] = str(crawl_url) if crawl_url else 'Not Found'
                if crawl_url:
                    stats.setdefault('url_lookup', 0)
                    stats['url_lookup'] += 1
                    print(f"[{index + 1}/{len(df)}] {college} - [URL-ONLY] {crawl_url}")
            stats['skipped'] += 1
            continue

        print(f"[{index + 1}/{len(df)}] {college} ({state})")

        sources_used = []

        # 1. AICTE
        m = match_in_reference(college, state, aicte_df, aicte_cols, aicte_index)
        if m:
            filled = _fill_missing(df.loc[index], m)
            for k, v in m.items():
                if k in TARGET_FIELDS and v and _is_empty(df.at[index, k]):
                    df.at[index, k] = v
            if filled:
                sources_used.append(f"AICTE({m['_score']})")
                stats['aicte'] += 1
                print(f"  [AICTE] score={m['_score']} filled={filled}")

        # 2. AISHE (only fills fields still empty)
        m = match_in_reference(college, state, aishe_df, aishe_cols, aishe_index)
        if m:
            filled = []
            for k in TARGET_FIELDS:
                v = m.get(k)
                if v and _is_empty(df.at[index, k]):
                    df.at[index, k] = v
                    filled.append(k)
            if filled:
                sources_used.append(f"AISHE({m['_score']})")
                stats['aishe'] += 1
                print(f"  [AISHE] score={m['_score']} filled={filled}")

        # 2b. NIRF — fuzzy-match against the NIRF index; if the institute is
        # NIRF-ranked, download its Data PDF and pull verified Strength + Intake.
        # NIRF is the most authoritative source we have for student counts, so
        # we overwrite Strength even if a less precise value was already set.
        if nirf_index:
            m = match_in_reference(college, state, nirf_df, nirf_cols, nirf_index)
            if m:
                # We need the NIRF ID + PDF URL from the matched row to fetch
                # the PDF — look them up via the matched name (cheap O(N) scan
                # over ~800 rows).
                nirf_filled = []
                nirf_name_col = nirf_cols['name']
                # Locate the matched NIRF row so we can read its ID + PDF URL.
                # match_in_reference already picked the best fuzzy match; we
                # just need to find the same row again by re-running the same
                # token_sort_ratio scan over the (in-state) NIRF rows.
                inst_id, pdf_url, category = None, None, None
                for _, nr in nirf_df.iterrows():
                    if pd.isna(nr.get(nirf_name_col)):
                        continue
                    if fuzz.token_sort_ratio(str(nr[nirf_name_col]).lower(),
                                             college.lower()) >= 80:
                        inst_id = nr.get('NIRF ID')
                        pdf_url = nr.get('PDF URL')
                        category = nr.get('Category')
                        break
                if inst_id and pdf_url:
                    pdf_bytes = aggregators.fetch_nirf_pdf(inst_id, pdf_url)
                    parsed = aggregators.parse_nirf_pdf(pdf_bytes)
                    if parsed:
                        # Fill Approx Strength from total_strength
                        if parsed.get('total_strength'):
                            df.at[index, 'Approx Strength'] = str(parsed['total_strength'])
                            nirf_filled.append('Approx Strength')
                        # Fill Stream from NIRF category if currently empty
                        if category and _is_empty(df.at[index, 'Stream']):
                            df.at[index, 'Stream'] = category
                            nirf_filled.append('Stream')
                if nirf_filled:
                    sources_used.append(f"NIRF({m['_score']})")
                    stats.setdefault('nirf', 0)
                    stats['nirf'] += 1
                    print(f"  [NIRF]  id={inst_id} score={m['_score']} "
                          f"filled={nirf_filled}")

        # 3. CollegeDunia — fills Email / Phone / POC / Location from the
        # college's profile-page JSON-LD. Also gives us the OFFICIAL site URL
        # that step 4 then crawls.
        official_url = None
        if _is_empty(df.at[index, 'Email ID']) or _is_empty(df.at[index, 'Phone Number']):
            cd_url = aggregators.find_collegedunia_url(college, state)
            if cd_url:
                info = aggregators.parse_collegedunia_profile(cd_url)
                if info and not info.get('_error'):
                    # Reject the match if the CollegeDunia profile's name
                    # is too dissimilar from our query. The distinctive-token
                    # guard rejects cases like "Indian Institute of Technology"
                    # being filled from an IIIT profile.
                    matched = info.get('_matched_name', '')
                    if not _names_likely_same(college, matched):
                        sim = aggregators.name_similarity(college, matched)
                        print(f"  [CD]    REJECTED (name guard, sim={sim:.0f}): {matched!r}")
                    else:
                        sim = aggregators.name_similarity(college, matched)
                        official_url = info.get('Official URL')
                        # Record the official URL as soon as we have it, so
                        # rows where the WEB crawl is later skipped (because
                        # email/phone already filled) still get a URL in the
                        # Website Found column.
                        if official_url and _is_empty(df.at[index, 'Website Found']):
                            df.at[index, 'Website Found'] = official_url
                        filled = []
                        for k in TARGET_FIELDS:
                            v = info.get(k)
                            if v and _is_empty(df.at[index, k]):
                                df.at[index, k] = v
                                filled.append(k)
                        if filled:
                            sources_used.append(f'CollegeDunia({sim:.0f})')
                            stats['collegedunia'] += 1
                            print(f"  [CD]    sim={sim:.0f} url={cd_url} filled={filled}")

                        # Pull leadership + full faculty roster from CollegeDunia's
                        # __NEXT_DATA__. The leadership goes in the main sheet's
                        # Contact columns; the faculty roster goes in Sheet 2.
                        people = aggregators.fetch_collegedunia_people(cd_url)
                        cd_contact = people.get('contact')
                        if cd_contact and _is_empty(df.at[index, 'Contact Name']):
                            df.at[index, 'Contact Name'] = cd_contact['name']
                            df.at[index, 'Contact Role'] = cd_contact['role']
                            df.at[index, 'Contact Email'] = cd_contact['email']
                            if cd_contact.get('phone') and _is_empty(df.at[index, 'Phone Number']):
                                df.at[index, 'Phone Number'] = cd_contact['phone']
                            sources_used.append('CD-staff')
                            stats.setdefault('cd_contact', 0)
                            stats['cd_contact'] += 1
                            print(f"  [CDS]   {cd_contact['name']} / "
                                  f"{cd_contact['role']} / "
                                  f"{cd_contact['email'] or '(no email)'}")

                        # Stash each faculty member as a separate row keyed by
                        # the input college name + state. This list becomes
                        # Sheet 2 ("Faculty") of the output workbook.
                        for f in people.get('faculty') or []:
                            if not (f.get('name') or f.get('email') or f.get('phone')):
                                continue
                            faculty_rows.append({
                                'College Name':  college,
                                'State':         state,
                                'Faculty Name':  f['name'],
                                'Designation':   f['designation'],
                                'Department':    f['department'],
                                'Email':         f['email'],
                                'Phone':         f['phone'],
                                'Qualification': f['qualification'],
                            })
                        if people.get('faculty'):
                            stats.setdefault('faculty_rows', 0)
                            stats['faculty_rows'] += len(people['faculty'])
                            print(f"  [CDF]   {len(people['faculty'])} faculty rows added")
            time.sleep(0.4)

        # 4. Crawl the college's OWN website (homepage + contact/about/admissions
        # sub-pages). This is the single biggest lever for filling gaps —
        # contact info that isn't on the homepage often lives on /contact-us
        # or /about, and POC/strength typically only appear on /administration
        # or /about pages. We use the official URL from CollegeDunia if we
        # have it, otherwise we fall back to a DDG search for the homepage.
        # Trigger WEB crawl only if we're missing one of the *core* fields.
        # Strength and POC are not in this list because:
        #   - POC's strongest source is now the CD __NEXT_DATA__ Contact step.
        #   - Strength is rarely findable on homepages.
        # Skipping the WEB crawl when CD already filled email/phone/Contact
        # cuts roughly 60% of total runtime.
        WEB_TRIGGER_FIELDS = ('Email ID', 'Phone Number', 'Contact Name')
        still_missing = any(_is_empty(df.at[index, f]) for f in WEB_TRIGGER_FIELDS)
        if still_missing:
            crawl_url = official_url
            if not crawl_url:
                crawl_url = get_website_url(college, state)
            df.at[index, 'Website Found'] = str(crawl_url) if crawl_url else 'Not Found'
            if crawl_url:
                text, visited = aggregators.crawl_official_site(crawl_url, max_pages=6)
                if text:
                    extracted = aggregators.extract_contacts_from_text(text)
                    filled = []
                    for k, v in extracted.items():
                        if v and _is_empty(df.at[index, k]):
                            df.at[index, k] = v
                            filled.append(k)

                    # Find a leadership contact (Director/Principal/HOD + email)
                    # from the same crawled text. Restrict emails to the
                    # college's own domain so we don't pair an admin name with
                    # a footer or web-developer email.
                    if any(_is_empty(df.at[index, f]) for f in CONTACT_FIELDS):
                        contact = aggregators.extract_admin_contact(
                            text, homepage_url=crawl_url)
                        if contact:
                            df.at[index, 'Contact Name'] = contact['name']
                            df.at[index, 'Contact Role'] = contact['role']
                            df.at[index, 'Contact Email'] = contact['email']
                            filled.extend(['Contact Name', 'Contact Role', 'Contact Email'])

                    if filled:
                        sources_used.append(f'WEB({len(visited)}p)')
                        stats['web'] += 1
                        print(f"  [WEB]   pages={len(visited)} filled={filled}")

            # 4b. LinkedIn x-ray fallback — DISABLED.
            # The implementation (aggregators.find_contact_via_linkedin_xray)
            # is kept for reference, but in practice the precision filter
            # (requires both a snippet-extracted name AND a guessed email
            # verified in crawled text) returns nothing for most colleges,
            # while loose filters return too much noise. Re-enable by setting
            # ENABLE_LINKEDIN_XRAY = True at the top of the file.
            if ENABLE_LINKEDIN_XRAY and _is_empty(df.at[index, 'Contact Name']) \
                    and text and crawl_url:
                domain_for_xray = aggregators._domain_from_url(crawl_url)
                lx_contact = aggregators.find_contact_via_linkedin_xray(
                    college, domain_for_xray, crawled_text=text)
                if lx_contact:
                    df.at[index, 'Contact Name'] = lx_contact['name']
                    df.at[index, 'Contact Role'] = lx_contact['role']
                    df.at[index, 'Contact Email'] = lx_contact['email']
                    sources_used.append('LinkedIn')
                    stats.setdefault('linkedin', 0)
                    stats['linkedin'] += 1
                    print(f"  [LI]    {lx_contact['name']} / {lx_contact['role']} / "
                          f"{lx_contact['email']}")
            time.sleep(0.4)

        # 4c. Last-resort Website Found lookup. If we still don't have a URL
        # for this college (no AICTE/NIRF site URL, no CollegeDunia profile,
        # WEB crawl didn't run), do a single DDG search for the official
        # homepage. No crawl — just record the URL.
        if _is_empty(df.at[index, 'Website Found']):
            ddg_url = get_website_url(college, state)
            df.at[index, 'Website Found'] = ddg_url if ddg_url else 'Not Found'
            if ddg_url:
                stats.setdefault('url_lookup', 0)
                stats['url_lookup'] += 1
            time.sleep(0.3)

        # 4d. Placement-officer lookup. Visits the college's own
        # /placement, /tpo, /cdc page (or `placement.<domain>` subdomain) and
        # extracts the Placement Officer's name + email + phone. Only runs
        # when we have a Website Found URL to start from, and the placement
        # email column is still empty (resume-safe).
        if (not _is_empty(df.at[index, 'Website Found'])
                and _is_empty(df.at[index, 'Placement Officer Email'])
                and _is_empty(df.at[index, 'Placement Officer Phone'])):
            site_url = df.at[index, 'Website Found']
            if site_url and site_url != 'Not Found':
                pl = aggregators.find_placement_contact(site_url)
                if pl:
                    filled_pl = []
                    if pl.get('name') and _is_empty(df.at[index, 'Placement Officer Name']):
                        df.at[index, 'Placement Officer Name'] = pl['name']
                        filled_pl.append('Name')
                    if pl.get('phone') and _is_empty(df.at[index, 'Placement Officer Phone']):
                        df.at[index, 'Placement Officer Phone'] = pl['phone']
                        filled_pl.append('Phone')
                    if pl.get('email') and _is_empty(df.at[index, 'Placement Officer Email']):
                        df.at[index, 'Placement Officer Email'] = pl['email']
                        filled_pl.append('Email')
                    if filled_pl:
                        sources_used.append('Placement')
                        stats.setdefault('placement', 0)
                        stats['placement'] += 1
                        print(f"  [PL]    filled={filled_pl} email={pl.get('email','-')}")

        # 5. Final pass: infer District from Location if District is still
        # empty and we have an address string to work with.
        if _is_empty(df.at[index, 'District']) and not _is_empty(df.at[index, 'Location']):
            inferred = aggregators.infer_district_from_location(
                df.at[index, 'Location'], known_districts=known_districts)
            if inferred:
                df.at[index, 'District'] = inferred
                sources_used.append('inferred')

        if not sources_used:
            stats['none'] += 1
            print("  [--] no match in any source")
        df.at[index, 'Source'] = ' + '.join(sources_used) if sources_used else 'None'

        # Checkpoint every 25 rows
        if (index + 1) % 25 == 0:
            _save_workbook(df, faculty_rows, output_file)
            print(f"  ...saved checkpoint to {output_file} "
                  f"({len(faculty_rows)} faculty rows)")

    _save_workbook(df, faculty_rows, output_file)
    print("\nDone.")
    print(f"  AICTE matches:        {stats['aicte']}")
    print(f"  AISHE matches:        {stats['aishe']}")
    print(f"  NIRF matches:         {stats.get('nirf', 0)}")
    print(f"  CollegeDunia matches: {stats['collegedunia']}")
    print(f"  Web fallbacks:        {stats['web']}")
    print(f"  URL lookups:          {stats.get('url_lookup', 0)}")
    print(f"  Placement contacts:   {stats.get('placement', 0)}")
    print(f"  Unresolved:           {stats['none']}")
    print(f"  Skipped (had Source): {stats['skipped']}")
    print(f"  Faculty rows on Sheet 2: {len(faculty_rows)}")
    print(f"  Output: {output_file}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Extract college details: URLs, contacts, emails, phones, faculty.'
    )
    parser.add_argument('--input', default=INPUT_FILE, 
                        help=f'Input file (default: {INPUT_FILE})')
    parser.add_argument('--output', default=OUTPUT_FILE, 
                        help=f'Output file (default: {OUTPUT_FILE})')
    parser.add_argument('--limit', type=int, default=0, 
                        help='Process only the first N rows')
    parser.add_argument('--start-row', type=int, default=DEFAULT_START_ROW,
                        help=f'Start processing from this 1-based data row (default: {DEFAULT_START_ROW})')
    args = parser.parse_args()
    
    process_spreadsheet(args.input, args.output, limit=args.limit, start_row=args.start_row)
