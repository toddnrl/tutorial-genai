# from playwright.sync_api import sync_playwright

# URL = "공고_URL"

# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)
#     page = browser.new_page()

#     page.goto("https://www.jobkorea.co.kr/Recruit/GI_Read/49116111?sc=729&sn=103")
#     page.wait_for_load_state("networkidle")

#     items = page.locator(r'div.flex.gap-\[20px\]')

#     job_info = {}

#     for i in range(items.count()):

#         spans = items.nth(i).locator("span")

#         if spans.count() >= 2:
#             key = spans.nth(0).inner_text().strip()
#             value = spans.nth(1).inner_text().strip()

#             job_info[key] = value

#     for key, value in job_info.items():
#         print(f"{key} : {value}")

#     browser.close()

from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://www.jobkorea.co.kr/")
    
    cards = page.locator("div.rounded-\\[10px\\]")
    print(cards.count())

    title = cards.locator("span w-full").inner_text()
    print(title)






