# from playwright.sync_api import sync_playwright

# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)

#     page = browser.new_page()
#     page.goto("https://news.naver.com/section/105")

#     page.wait_for_timeout(3000)

#     news = page.locator("div.sa_text")

#     print(news.count())

#     for i in range(news.count()):

#         title = news.nth(i).locator("a strong").inner_text()

#         print(title)
#         print("------------------------")

#     browser.close()


# from playwright.sync_api import sync_playwright

# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)

#     page = browser.new_page()
#     page.goto("https://news.naver.com/section/105")

#     page.wait_for_timeout(3000)

#     news = page.locator("div.sa_text")

#     print(news.count())

#     for i in range(news.count()):

#         title = news.nth(i).locator("a strong").inner_text()
#         link = news.nth(i).locator("a.sa_text_title").get_attribute("href")

#         print(title)
#         print(link)
#         print("------------------------")

#     browser.close()



from playwright.sync_api import sync_playwright



with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()
    page.goto("https://news.naver.com/section/105")

    headlines = page.locator(".section_article.as_headline a.sa_text_title")
    # print("헤드라인 갯수 : ", headlines.count())


    for i in range(headlines.count()):
        news = headlines.nth(i)


        title = news.inner_text().strip()
        
        href = news.get_attribute('href')

        print(f"{i+1}. {title}.\n {href}")

    input("엔터")