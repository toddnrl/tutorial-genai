import requests
import csv
from bs4 import BeautifulSoup

from bs4 import BeautifulSoup
url = "https://books.toscrape.com/"

resp = requests.get(url)
resp.encoding = "utf-8"
soup = BeautifulSoup(resp.text, "html.parser")


books = soup.select("article.product_pod")

rating_map = {
    "One" :1, 
    "Two" :2,
    "Three" :3,
    "Four" :4,
    "Five" :5
}

with open("books2.csv", "w", encoding="utf-8") as file:
    csv_writer = csv.writer(file)
    csv_writer.writerow({"도서명", "평점", "가격"})

    for book in books:
        title = book.h3.a["title"]
        prices = book.select_one(".price_color").text
        prices = prices.replace("£", "")
        rating = book.p["class"][1]
        rating_num = rating_map[rating]
        

        print(f"도서명 :{title}, 가격 : {prices}, 평점 : {rating_num}")

        csv_writer.writerow({title, rating_num, prices})

    