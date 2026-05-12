import requests

url = "https://www.exemple.com"

response = requests.get(url)

html = response.text

print(html)