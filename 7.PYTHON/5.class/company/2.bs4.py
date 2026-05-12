# pip install bs4

from bs4 import BeautifulSoup

html = "<html><head><title>hello</title></head><body><h1>Title</h1></body></html>"

soup = BeautifulSoup(html, "html.parser")

print(soup)