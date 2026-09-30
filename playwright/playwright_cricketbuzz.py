from playwright.sync_api import sync_playwright

#Open chromium browser in non-headless mode
# Go to https://www.cricbuzz.com/ 
# Extract the title of the page and print it
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://www.cricbuzz.com/")
    page.screenshot(path="screenshot.png")
    print(page.title())
    browser.close()


