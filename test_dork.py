import time
from ddgs import DDGS # Using the new package name from your warning

def test_on_colleges(domains):
    print(f"🚀 Searching for Mandatory Disclosures...")
    
    with DDGS() as ddgs:
        for domain in domains:
            # Enhanced Query: Looking for the exact PDF
            query = f'site:{domain} "Mandatory Disclosure" filetype:pdf'
            print(f"\n🔍 Querying: {query}")
            
            try:
                results = list(ddgs.text(query, max_results=5))
                if not results:
                    print(f"  ❌ No PDF found on {domain}")
                
                for r in results:
                    print(f"  ⭐ MATCH FOUND: {r['href']}")
            except Exception as e:
                print(f"  ⚠️ Error searching {domain}: {e}")
            
            time.sleep(2) # Be polite to the search engine

if __name__ == "__main__":
    # Test these specific domains
    colleges = ["vjti.ac.in", "bmsce.ac.in"]
    test_on_colleges(colleges)
    print("\n✅ All tests complete.")