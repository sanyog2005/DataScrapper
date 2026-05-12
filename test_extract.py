import re
import io
import time
import pdfplumber
from playwright.sync_api import sync_playwright

print("!!! SCRIPT IS STARTING !!!")

PHONE_REGEX = re.compile(r'(?:\+91|0)?[6-9]\d{9}')
EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

def fetch_and_parse_pdf(url):
    print(f"\n🚀 Launching Browser for: {url}")
    
    with sync_playwright() as p:
        # 1. RUNNING WITH HEADLESS=FALSE 
        # This opens a real window so you can see the bypass happen.
        browser = p.chromium.launch(headless=False) 
        
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            accept_downloads=True
        )
        page = context.new_page()

        try:
            print("  🛰️ Navigating... Watch the browser window.")
            # Navigate to the URL
            response = page.goto(url, wait_until="load", timeout=60000)
            
            # Wait for 5 seconds to let any JS challenges finish
            time.sleep(5)

            # 2. THE SECRET MOVE: Instead of page.goto, we "fetch" through the browser's console
            # This bypasses almost all standard bot-detection.
            print("  📥 Attempting internal fetch...")
            content_base64 = page.evaluate("""
                async (url) => {
                    const response = await fetch(url);
                    const buffer = await response.arrayBuffer();
                    let binary = '';
                    const bytes = new Uint8Array(buffer);
                    for (let i = 0; i < bytes.byteLength; i++) {
                        binary += String.fromCharCode(bytes[i]);
                    }
                    return btoa(binary);
                }
            """, url)

            import base64
            buffer = base64.b64decode(content_base64)
            
            if not buffer.startswith(b'%PDF'):
                print("  ❌ Data is still not a PDF. Trying one last method: Screenshotting text?")
                return

            print("  ✅ PDF Successfully Grabbed!")
            
            with pdfplumber.open(io.BytesIO(buffer)) as pdf:
                text = " ".join([p.extract_text() or "" for p in pdf.pages[:5]])
                clean_text = " ".join(text.split())
                
                phones = list(set(PHONE_REGEX.findall(clean_text)))
                emails = list(set(EMAIL_REGEX.findall(clean_text)))
                
                print(f"  📞 Found Phones: {[p for p in phones if len(p)>=10][:5]}")
                print(f"  📧 Found Emails: {emails[:5]}")

        except Exception as e:
            print(f"  ⚠️ Error: {e}")
        finally:
            browser.close()