import asyncio
import re
from playwright.async_api import async_playwright
from urllib.parse import urljoin, urlparse

# Keywords that usually lead to emails
FACULTY_KEYWORDS = ["faculty", "staff", "people", "management", "contact", "director", "department"]

async def dynamic_extract_emails(base_url):
    print(f"\n🚀 Target: {base_url}")
    EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    all_emails = set()
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(user_agent="Mozilla/5.0")
        page = await context.new_page()

        try:
            # 1. Visit Homepage
            await page.goto(base_url, wait_until="domcontentloaded", timeout=30000)
            homepage_content = await page.content()
            all_emails.update(EMAIL_REGEX.findall(homepage_content))

            # 2. Find all links on the page
            links = await page.query_selector_all("a")
            to_visit = set()
            base_domain = urlparse(base_url).netloc

            for link in links:
                href = await link.get_attribute("href")
                text = (await link.inner_text()).lower()
                
                if href:
                    full_url = urljoin(base_url, href)
                    # Only stay on the same website and look for keywords
                    if base_domain in urlparse(full_url).netloc:
                        if any(kw in text or kw in href.lower() for kw in FACULTY_KEYWORDS):
                            to_visit.add(full_url)

            # 3. Visit the top 5 most relevant pages found
            to_visit = list(to_visit)[:5] 
            print(f"📡 Found {len(to_visit)} potential contact pages. Scanning...")

            for url in to_visit:
                try:
                    print(f"  🔍 Checking: {url}")
                    await page.goto(url, wait_until="domcontentloaded", timeout=20000)
                    content = await page.content()
                    all_emails.update(EMAIL_REGEX.findall(content))
                except:
                    continue

        except Exception as e:
            print(f"⚠️ Error: {e}")
        finally:
            await browser.close()

    # Clean results
    final = {e.lower() for e in all_emails if not any(x in e.lower() for x in ['.png', '.jpg', 'example', 'wix'])}
    return sorted(list(final))

if __name__ == "__main__":
    url = "https://www.iehe.ac.in/" # Try any college URL here
    results = asyncio.run(dynamic_extract_emails(url))
    print(f"\n✅ RESULTS FOR {url}:")
    for email in results:
        print(f"  📧 {email}")