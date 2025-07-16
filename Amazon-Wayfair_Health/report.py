# report.py
from selenium import webdriver
import time

HOME_URL = "https://sellercentral.amazon.com/home"
DASHBOARD_URL = "https://sellercentral.amazon.com/performance/dashboard"

driver = webdriver.Chrome()

def full_page_screenshot(driver, file_path):
    S = lambda X: driver.execute_script(f'return document.body.parentNode.scroll{X}')
    driver.set_window_size(S('Width'), S('Height'))
    time.sleep(2)
    driver.save_screenshot(file_path)

try:
    # Start in fullscreen mode
    driver.maximize_window()

    # Login manually
    driver.get("https://sellercentral.amazon.com/")
    print("🔐 Please log in manually (60s)...")
    time.sleep(60)

    # Screenshot homepage
    driver.get(HOME_URL)
    time.sleep(5)
    full_page_screenshot(driver, "amazon_home.png")

    # Screenshot performance dashboard
    driver.get(DASHBOARD_URL)
    time.sleep(5)
    full_page_screenshot(driver, "amazon_dashboard.png")

    print("✅ Screenshots saved.")

finally:
    driver.quit()
