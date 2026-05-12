"""Verify every row gets a Website Found URL after the new fallback."""
import os
import pandas as pd
import main

# Mix of cases: one row with a strong CD profile (URL from CD), one with no CD
# profile (truncated input — needs DDG fallback), one AICTE-only.
SAMPLE_INDICES = [
    1,    # MANIT Bhopal — CD profile, URL from CD JSON-LD
    0,    # "Indian Institute of Technology" — truncated, no CD match
    24,   # LNCT Jabalpur — AICTE + CD match
    6,    # APSU Rewa — needs DDG search
    14,   # Amity University, Gwalior — no CD likely
]
src = pd.read_excel('colleges_list.xlsx', dtype=str)
slice_df = src.iloc[SAMPLE_INDICES].reset_index(drop=True)
slice_df.to_excel('_slice_input.xlsx', index=False)
if os.path.exists('_slice_output.xlsx'):
    os.remove('_slice_output.xlsx')

main.process_spreadsheet('_slice_input.xlsx', '_slice_output.xlsx')

out = pd.read_excel('_slice_output.xlsx', dtype=str)
print('\n=== Website Found result per row ===')
for i, row in out.iterrows():
    name = row['College Name']
    url = row.get('Website Found', '')
    if pd.isna(url) or url == '':
        url = '<EMPTY>'
    src = row.get('Source', '')
    print(f'{i + 1}. {name[:55]:55s}')
    print(f'   URL:    {url}')
    print(f'   Source: {src}')
