import requests
import csv
from bs4 import BeautifulSoup

url = "https://www.naver.com"
resp = requests.get(url)

