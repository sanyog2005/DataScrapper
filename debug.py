print("--- STARTING DEBUG TEST ---")
import sys
from duckduckgo_search import DDGS

print(f"Python Version: {sys.version}")

try:
    with DDGS() as ddgs:
        print("Searching DuckDuckGo...")
        for r in ddgs.text("VJTI Mandatory Disclosure filetype:pdf", max_results=1):
            print(f"FOUND SOMETHING: {r['href']}")
except Exception as e:
    print(f"ERROR: {e}")

print("--- DEBUG TEST FINISHED ---")