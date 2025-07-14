import asyncio
import os
from dotenv import load_dotenv
from playwright.async_api import async_playwright
import time

# === Load environment variables ===
load_dotenv()
EMAIL = os.getenv("SMartsheet_EMAIL").strip("'")
PASSWORD = os.getenv("SMartsheet_PASSWORD").strip("'")
DASHBOARD_URL = os.getenv("SKU_DASHBOARD").strip("'")
OUTPUT_PATH = r"C:/Users/panderson/OneDrive - American Bath Group/Documents/Paul_Anderson/Reports/dashboard.pdf"

async def save_dashboard_pdf():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=250)
        context = await browser.new_context()
        page = await context.new_page()

        # Step 1: Go to login page
        await page.goto("https://app.smartsheet.com/b/home", wait_until="networkidle")

        # Step 2: Fill email
        await page.fill('input[type="email"]', EMAIL)

        # Step 3: Click "Continue"
        await page.click('input[type="submit"][value="Continue"]')

        # Step 4: Click "Sign in with email and password"
        await page.wait_for_selector('text="Sign in with email and password"', timeout=10000)
        await page.click('text="Sign in with email and password"')

        # Step 5: Fill password
        await page.wait_for_selector('input[type="password"]', timeout=10000)
        await page.fill('input[type="password"]', PASSWORD)

        # Step 6: Click final "Sign in"
        await page.click('input[type="submit"][value="Sign in"]')

        # Step 7: Wait 30 seconds to allow Smartsheet to fully load
        print("⏳ Waiting 30 seconds for Smartsheet to finish login and load...")
        await asyncio.sleep(30)
        await page.screenshot(path="post-login-debug.png")
        print("🖼️ Screenshot saved as post-login-debug.png")

        # Step 8: Navigate to dashboard
        await page.goto(DASHBOARD_URL, wait_until='networkidle')

        # Step 9: Save dashboard as PDF
        await page.pdf(
            path=OUTPUT_PATH,
            format="A4",
            print_background=True
        )

        print(f"✅ Dashboard saved as PDF: {OUTPUT_PATH}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(save_dashboard_pdf())
